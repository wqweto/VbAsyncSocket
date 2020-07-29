VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cRequestHandler"

#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)

'=========================================================================
' API
'=========================================================================

Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_sKey                  As String
Private m_lParentWeakRef        As Long
Private m_sRequest              As String

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Parent() As Form2
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oSocket As cTlsSocket, sKey As String, oParent As Form2) As Boolean
    Set m_oSocket = oSocket
    m_sKey = sKey
    m_lParentWeakRef = ObjPtr(oParent)
    '--- success
    Init = True
QH:
End Function

Private Function HandleRequest(sText As String) As Boolean
    Const FUNC_NAME     As String = "HandleRequest"
    Dim vSplit          As Variant
    Dim sRetVal         As String
    
    On Error GoTo EH
    m_sRequest = m_sRequest & sText
    If InStr(m_sRequest, vbCrLf) = 0 Then
        Exit Function
    End If
    vSplit = Split(m_sRequest, vbCrLf)
    vSplit = Split(vSplit(0), " ")
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, "Path=" & vSplit(1)
    #End If
    sRetVal = "<html><body>" & Now & "</body></html>"
    sRetVal = "HTTP/1.0 200 Ok" & vbCrLf & _
        "Content-Type: text/html; charset=UTF-8" & vbCrLf & _
        "Content-Length: " & Len(sRetVal) & vbCrLf & _
        "Connection: Close" & vbCrLf & vbCrLf & _
        sRetVal
    If Not m_oSocket.SendArray(ToUtf8Array(sRetVal)) Then
        GoTo QH
    End If
    HandleRequest = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnClose()
    Parent.frRemoveHandler m_sKey
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    With m_oSocket.LastError
        If .Number <> 0 Then
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME & ", " & Replace(.Source, vbCrLf, ", "), .Description & " &H" & Hex$(.Number), vbLogEventTypeError
            #Else
                Debug.Print "Error: " & .Description & " &H" & Hex$(.Number) & " [" & MODULE_NAME & "." & FUNC_NAME & ", " & Replace(.Source, vbCrLf, ", ") & "]"
            #End If
        End If
    End With
    Parent.frRemoveHandler m_sKey
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baRecv()        As Byte
    
    On Error GoTo EH
    If Not m_oSocket.ReceiveArray(baRecv) Then
        GoTo QH
    End If
    If Not HandleRequest(FromUtf8Array(baRecv)) Then
        GoTo QH
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub
