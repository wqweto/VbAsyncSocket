VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncSocketHelperWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2019 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cAsyncSocketHelperWindow"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplNoVBIDESupport = (MST_NO_VBIDE_SUPPORT <> 0)

'=========================================================================
' API
'=========================================================================

Private Const INVALID_SOCKET                As Long = -1
'--- for WSADATA
Private Const WSADESCRIPTION_LEN            As Long = 257
Private Const WSASYS_STATUS_LEN             As Long = 129
'--- Windows Messages
Private Const WM_USER                       As Long = &H400
Private Const WM_SOCKET_NOTIFY              As Long = WM_USER + 2
'--- for Get/SetWindowLong
Private Const GWL_USERDATA                  As Long = -21
Private Const HWND_MESSAGE                  As Long = -3
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If Not ImplUseShared Then
    '--- for thunks
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
#End If

Private Type WSADATA
   wVersion             As Integer
   wHighVersion         As Integer
   szDescription        As String * WSADESCRIPTION_LEN
   szSystemStatus       As String * WSASYS_STATUS_LEN
   iMaxSockets          As Integer
   iMaxUdpDg            As Integer
   lpVendorInfo         As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MAX_SOCKETS           As Long = &HC000 - WM_SOCKET_NOTIFY
Private Const STR_HELPER_CLASS      As String = "STATIC"
Private Const STR_HELPER_CAPTION    As String = "AsyncSocket Helper Window"

Private m_uWsaData              As WSADATA
Private m_hWnd                  As Long
Private m_pSubclass             As IUnknown
Private m_pCleanup              As IUnknown
Private m_uData()               As UcsSocketDataType
Private m_lPos                  As Long
Private m_lCount                As Long

Private Type UcsSocketDataType
    SocketPtr           As Long
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Friend Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Function frAddSocket(Socket As cAsyncSocket, Index As Long) As Boolean
    Const FUNC_NAME     As String = "frAddSocket"
    Dim lSize           As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Index <> -1 Then
        Debug.Assert m_uData(Index).SocketPtr = ObjPtr(Socket)
        '--- success
        frAddSocket = True
        Exit Function
    End If
    lSize = UBound(m_uData) + 1
    If m_lCount >= lSize Then
        lSize = 2 * lSize
        If lSize > MAX_SOCKETS Then
            lSize = MAX_SOCKETS
        End If
        ReDim Preserve m_uData(0 To lSize - 1) As UcsSocketDataType
    End If
    For lIdx = 0 To UBound(m_uData)
        Index = (m_lPos + lIdx) Mod lSize
        If m_uData(Index).SocketPtr = 0 Then
            '--- note: weak reference
            m_uData(Index).SocketPtr = ObjPtr(Socket)
            m_lPos = Index + 1
            m_lCount = m_lCount + 1
            '--- success
            frAddSocket = True
            Exit Function
        End If
    Next
QH:
    Index = -1
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

Friend Function frRemoveSocket(Socket As cAsyncSocket, Index As Long) As Boolean
    Const FUNC_NAME     As String = "frRemoveSocket"
    
    On Error GoTo EH
    If Index <> -1 Then
        Debug.Assert m_uData(Index).SocketPtr = ObjPtr(Socket)
        m_uData(Index).SocketPtr = 0
        Index = -1
        m_lCount = m_lCount - 1
    End If
    '--- success
    frRemoveSocket = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

Private Function pvHandleNotify(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "pvHandleNotify"
    Dim lIdx            As Long
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_SOCKET_NOTIFY - 1
        For lIdx = 0 To UBound(m_uData)
            With m_uData(lIdx)
                If .SocketPtr <> 0 Then
                    If pvToSocket(.SocketPtr).frNotifyGetHostByName(wParam, lParam) Then
                        pvHandleNotify = 1
                        GoTo QH
                    End If
                End If
            End With
        Next
    Case WM_SOCKET_NOTIFY To WM_SOCKET_NOTIFY + UBound(m_uData)
        With m_uData(wMsg - WM_SOCKET_NOTIFY)
            If .SocketPtr <> 0 And wParam <> INVALID_SOCKET Then
                pvToSocket(.SocketPtr).frNotifyEvent wParam, lParam
                pvHandleNotify = 1
            End If
        End With
    Case Else
        pvHandleNotify = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    End Select
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

Private Function pvToSocket(ByVal lPtr As Long) As cAsyncSocket
    Call vbaObjSetAddref(pvToSocket, lPtr)
End Function

Private Property Get pvAddressOfSubclassProc() As cAsyncSocketHelperWindow
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 4)
End Property

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    SubclassProc = pvHandleNotify(hWnd, wMsg, wParam, lParam)
End Function

#If Not ImplUseShared Then
Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QECkAgcekESkAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBApAIvCBTwRKQCri8IFUBEpAKuLwgVgESkAq4vCBYQRKQCri8IFjBEpAKuLwgWUESkAq4vCBZwRKQCri8IFpBEpALn5BwAAq4PABOL6i8dfgcJQECkAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 22.3.2019 10:17:01
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwENQAV1aLdCQUg8YIgz4AdC+L+oHH4BHUAIvCBQQR1ACri8IFQBHUAKuLwgVQEdQAq4vCBXwR1ACruQkAAADzpYHC4BHUAFJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4FBLUAC1wENQAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfxiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBABmkFWL7ItVGIsKi0EshcB0J1L/0FqD+AF1N4sKUv9RMFqFwHUsiwpSavD/cST/UShaqQAAAAh1GTPAUFT/dRT/dRD/dQz/dQj/cgz/UhBY6xGLCv91FP91EP91DP91CP9RIF3CGAA=" ' 22.3.2019 10:17:35
    Const THUNK_SIZE    As Long = 420
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        #If Not ImplNoVBIDESupport Then
            If InIde Then
                aParams(7) = hIdeOwner
                aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
                aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
                aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
            End If
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitCleanupThunk(ByVal hHandle As Long, sModuleName As String, sProcName As String) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepQEPkAV1aLdCQUgz4AdCeL+oHHPBH5AIvCBcwQ+QCri8IFCBH5AKuLwgUYEfkAq7kCAAAA86WBwjwR+QBSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBH5AC1QEPkAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfxL/cgj/UgyLVCQEiwpS/1EQM8DCBAA=" ' 22.3.2019 10:19:28
    Const THUNK_SIZE    As Long = 256
    Static hThunk       As Long
    Dim aParams(0 To 1) As Long
    Dim pfnCleanup      As Long
    Dim lSize           As Long
    
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(0) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(1) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
    End If
    pfnCleanup = GetProcAddress(GetModuleHandle(sModuleName), sProcName)
    If pfnCleanup <> 0 Then
        lSize = CallWindowProc(hThunk, hHandle, pfnCleanup, VarPtr(aParams(0)), VarPtr(InitCleanupThunk))
        Debug.Assert lSize = THUNK_SIZE
    End If
End Function

Private Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Private Property Get hIdeOwner() As Long

End Property
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    If WSAStartup(&H101, m_uWsaData) = 0 Then
        m_hWnd = CreateWindowEx(0, STR_HELPER_CLASS, App.hInstance & ":" & App.ThreadID & ":" & STR_HELPER_CAPTION, _
            0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal 0)
        If m_hWnd <> 0 Then
            Call SetWindowLong(m_hWnd, GWL_USERDATA, ObjPtr(Me))
            Set m_pSubclass = InitSubclassingThunk(m_hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0))
            Set m_pCleanup = InitCleanupThunk(m_hWnd, "user32", "DestroyWindow")
        End If
    End If
    ReDim m_uData(0 To 511) As UcsSocketDataType
End Sub

Private Sub Class_Terminate()
    If m_hWnd <> 0 Then
        Call SetWindowLong(m_hWnd, GWL_USERDATA, 0)
        Set m_pSubclass = Nothing
        Set m_pCleanup = Nothing
    End If
    If m_uWsaData.wVersion <> 0 Then
        Call WSACleanup
    End If
End Sub
