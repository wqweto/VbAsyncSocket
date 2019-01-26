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
' VbAsyncSocket Project (c) 2018 by wqweto@gmail.com
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

'=========================================================================
' API
'=========================================================================

Private Const INVALID_SOCKET                As Long = -1
'--- for WSADATA
Private Const WSADESCRIPTION_LEN            As Long = 257
Private Const WSASYS_STATUS_LEN             As Long = 129
'--- Windows messages
Private Const WM_USER                       As Long = &H400
Private Const WM_SOCKET_NOTIFY              As Long = WM_USER + 2
'--- for VirtualAlloc/Free
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
Private Const MEM_RELEASE                   As Long = &H8000
'--- for Get/SetWindowLong
Private Const GWL_USERDATA                  As Long = -21
Private Const HWND_MESSAGE                  As Long = -3

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal lpAddress As Long, ByRef RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal lpAddress As Long, ByRef RetVal As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal lpAddress As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal lpAddress As Long, ByVal NewVal As Currency)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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
Private m_hSubclassThunk        As Long
Private m_uData()               As UcsSocketDataType
Private m_lPos                  As Long
Private m_lCount                As Long

Private Type UcsSocketDataType
    Socket          As cAsyncSocket
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical: " & Err.Description & " (" & MODULE_NAME & "." & sFunction & ")"
End Sub

'=========================================================================
' Properties
'=========================================================================

Friend Property Get hWnd() As Long
    hWnd = m_hWnd
    #If False Then '--- touch function (silence MZ-Tools->Review Source Code)
        sc_WndProc 0, 0, 0, 0, 0
    #End If
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
        Debug.Assert m_uData(Index).Socket Is Socket
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
        If m_uData(Index).Socket Is Nothing Then
            '--- note: weak reference
            Call CopyMemory(m_uData(Index).Socket, Socket, 4)
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
        Debug.Assert m_uData(Index).Socket Is Socket
        Call CopyMemory(m_uData(Index).Socket, 0&, 4)
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

Private Function pvHandleNotify(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "pvHandleNotify"
    Dim lIdx            As Long
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_USER + 1
        For lIdx = 0 To UBound(m_uData)
            With m_uData(lIdx)
                If Not .Socket Is Nothing Then
                    If .Socket.frNotifyGetHostByName(wParam, lParam) Then
                        GoTo QH
                    End If
                End If
            End With
        Next
    Case WM_SOCKET_NOTIFY To WM_SOCKET_NOTIFY + UBound(m_uData)
        With m_uData(wMsg - WM_SOCKET_NOTIFY)
            If .Socket Is Nothing Or wParam = INVALID_SOCKET Then
                GoTo QH
            End If
            .Socket.frNotifyEvent wParam, lParam
        End With
    End Select
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

'= shared ================================================================

#If Not ImplUseShared Then
Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
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
            m_hSubclassThunk = sc_Subclass(m_hWnd)
        End If
    End If
    ReDim m_uData(0 To 511) As UcsSocketDataType
End Sub

Private Sub Class_Terminate()
    Dim lIdx            As Long
    
    For lIdx = 0 To UBound(m_uData)
        If Not m_uData(lIdx).Socket Is Nothing Then
            '--- note: clear weak ref
            Call CopyMemory(m_uData(lIdx).Socket, 0&, 4)
        End If
    Next
    If m_hWnd <> 0 Then
        Call SetWindowLong(m_hWnd, GWL_USERDATA, 0)
        sc_Terminate m_hSubclassThunk
        Call DestroyWindow(m_hWnd)
    End If
    If m_uWsaData.wVersion <> 0 Then
        Call WSACleanup
    End If
End Sub

'=========================================================================
' Self-subclassing
'=========================================================================

'Based on [Simple SelfSub code by Paul Caton](http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73803&lngWId=1)

Private Function sc_Subclass(ByVal hWnd As Long, Optional ByVal dwRefData As Long) As Long
    Dim lVtbl           As Long
    Dim lPtr            As Long
    Dim bytOrig         As Byte
    Dim bytJmp          As Byte
    Dim lIdx            As Long
    Dim hThunk          As Long

    GetMem4 ObjPtr(Me), lVtbl                       'get address of the usercontrol's vtable
'    lVtbl = UnsignedAdd(lVtbl, &H7A4)               'bump to the user part of the usercontrol's vtable
    lVtbl = UnsignedAdd(lVtbl, 7 * 4)               '--- skip IDispatch methods
    GetMem4 lVtbl, lPtr                             'read the address of the first entry point
    GetMem1 lPtr, bytOrig                           'read the jump opcode at the first entry point [&H33 for psuedo code, &HE9 for native code]
    For lIdx = 1 To 512                             'scan a reasonable number of vtable entries
        lVtbl = UnsignedAdd(lVtbl, 4)               'next entry address
        GetMem4 lVtbl, lPtr                         'read the address of the entry point
        If IsBadCodePtr(lPtr) Then                  'is the entry point address valid code?
            Exit For
        End If
        GetMem1 lPtr, bytJmp                        'read the jump opcode at the entry point
        If bytJmp <> bytOrig Then                   'does the jump opcode match that of the first vtable entry?
            Exit For
        End If
    Next
    If lIdx > 512 Then
        GoTo QH                                     'last vtable entry not found... increase the For limit?
    End If
    GetMem4 UnsignedAdd(lVtbl, -4), lPtr            'back one entry to the last private method
    hThunk = VirtualAlloc(0, 48, MEM_COMMIT, PAGE_EXECUTE_READWRITE) 'allocate executable memory
    If hThunk = 0 Then
        GoTo QH
    End If
    PutMem4 hThunk, hWnd                            'copy the subclass thunk to memory
    PutMem8 UnsignedAdd(hThunk, 4), -854782363258311.4703@
    PutMem8 UnsignedAdd(hThunk, 12), 205082594635713.8405@
    PutMem8 UnsignedAdd(hThunk, 20), 850253272047553.4847@
    PutMem8 UnsignedAdd(hThunk, 28), -518126163307069.4644@
    PutMem4 UnsignedAdd(hThunk, 36), lPtr           'call address
    PutMem8 UnsignedAdd(hThunk, 40), -802991802926118.8865@
    If SetWindowSubclass(hWnd, UnsignedAdd(hThunk, 4), ObjPtr(Me), dwRefData) = 0 Then
        Call VirtualFree(hThunk, 0, MEM_RELEASE)
        GoTo QH
    End If
    '--- success
    sc_Subclass = hThunk
QH:
End Function

Private Sub sc_Terminate(hThunk As Long)
    Dim hWnd            As Long
    
    If hThunk <> 0 Then
        GetMem4 hThunk, hWnd
        Call RemoveWindowSubclass(hWnd, UnsignedAdd(hThunk, 4), ObjPtr(Me))
        Call VirtualFree(hThunk, 0, MEM_RELEASE)
        hThunk = 0
    End If
End Sub

Private Function sc_WndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    #If hWnd And dwRefData Then  '--- touch args (silence MZ-Tools->Review Source Code)
    #End If
    sc_WndProc = pvHandleNotify(wMsg, wParam, lParam)
End Function
