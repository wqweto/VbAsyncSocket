VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinSockRequest"
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
Private Const MODULE_NAME As String = "cWinSockRequest"

#Const ImplSync = Not (ASYNCSOCKET_NO_SYNC <> 0)
#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)

'=========================================================================
' Events
'=========================================================================

Event OnError(ByVal ErrorNumber As Long, ErrorDescription As String)
Event OnResponseDataAvailable(Data() As Byte)
Event OnReadyStateChange()

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsReadyStateEnum
    ucsRdsUninitialized = 0
    ucsRdsLoading = 1
    ucsRdsLoaded = 2
    ucsRdsInteractive = 3
    ucsRdsCompleted = 4
End Enum

'=========================================================================
' API
'=========================================================================

Private Const ERR_TIMEOUT                   As Long = &H800705B4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#If Not ImplUseShared Then
    Private Declare Function GetSystemTimeAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As Currency) As Long
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ERR_SYNC_NOT_IMPL As String = "Synchronous operations not implemented"

Private m_sHostAddress          As String
Private m_lHostPort             As Long
Private m_bAsync                As Boolean
Private m_eReadyState           As UcsReadyStateEnum
Private m_eNextState            As UcsReadyStateEnum
Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_lResolveTimeout       As Long
Private m_lConnectTimeout       As Long
Private m_lSendTimeout          As Long
Private m_lReceiveTimeout       As Long
Private m_lPacketTimeout        As Long
Private m_baRequest()           As Byte
Private m_lRequestBytes         As Long
Private m_baResponse()          As Byte
Private m_dblResponseTimer      As Double
Private m_dblPacketTimer        As Double
Private m_lStatus               As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical: " & Err.Description & " (" & MODULE_NAME & "." & sFunction & ")"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get ReadyState() As UcsReadyStateEnum
    ReadyState = m_eReadyState
End Property

Property Get ResponseBody() As Variant
    ResponseBody = m_baResponse
End Property

'--- ToDo: impl w/ CreateStreamOnHGlobal
'Property Get ResponseStream() As Variant
'
'End Property

Property Get ResponseText(Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    If UBound(m_baResponse) >= 0 Then
        ResponseText = m_oSocket.FromTextArray(m_baResponse, CodePage)
    End If
End Property

Property Get Status() As Long
    Status = m_lStatus
End Property

Property Get StatusText() As String
    If m_lStatus <> 0 Then
        StatusText = m_oSocket.GetErrorDescription(m_lStatus)
    End If
End Function

'=========================================================================
' Methods
'=========================================================================

''
' Paramter `Address` format is "[server:]port[/protocol]"
'
'    server    optional part for IP/Host/FQDN (default: localhost)
'    port      required numeric port
'    protocol  optional part for "tcp" or "udp" (default: "tcp")
'
' Samples:
'    "www.microsoft.com:80"  -> connect to www.microsoft.com on 80/tcp
'    "172.17.17.1:53/udp"    -> connect to 172.17.17.1 on 53/udp
'    "80"                    -> connect to localhost on 80/tcp
'
Public Sub Open_(Address As String, Optional ByVal Async As Boolean)
    Const FUNC_NAME     As String = "Open"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    m_lStatus = 0
    m_baRequest = vbNullString
    m_baResponse = vbNullString
    vSplit = Split(Address, ":")
    lIdx = IIf(UBound(vSplit) < 1, 0, 1)
    m_sHostAddress = At(vSplit, lIdx - 1)
    m_lHostPort = Val(At(vSplit, lIdx))
    #If ImplSync Then
        m_bAsync = Async
    #Else
        If Not Async Then
            On Error GoTo 0
            Err.Raise vbObjectError, , STR_ERR_SYNC_NOT_IMPL
        End If
        m_bAsync = True
    #End If
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=IIf(Right$(At(vSplit, lIdx), 4) = "/udp", ucsSckDatagram, ucsSckStream)) Then
        pvSetStatus m_oSocket.LastError
        GoTo QH
    End If
    Exit Sub
QH:
    If m_lStatus <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, , StatusText
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    GoTo QH
End Sub

Public Sub Send(Body As Variant)
    Const FUNC_NAME     As String = "Send"
    
    On Error GoTo EH
    m_lStatus = 0
    '-- ToDo: impl send from stream
    Select Case VarType(Body)
    Case vbArray + vbByte
        m_baRequest = Body
    Case Else
        m_baRequest = StrConv(Body, vbFromUnicode)
    End Select
    m_lRequestBytes = 0
    m_baResponse = vbNullString
    m_eNextState = ucsRdsLoading
    '--- note: Connect is noop for UDP, just used to async resolve m_sHostAddress
    If Not m_oSocket.Connect(m_sHostAddress, m_lHostPort) Then
        pvSetStatus m_oSocket.LastError
        GoTo QH
    End If
    #If ImplSync Then
        If Not m_bAsync Then
            If Not pvWaitForEvent([_ucsSfdResolve], m_lResolveTimeout, ucsRdsUninitialized) Then
                GoTo QH
            End If
            If Not pvWaitForEvent(ucsSfdConnect, m_lConnectTimeout, ucsRdsLoading) Then
                GoTo QH
            End If
            If Not pvWaitForEvent(ucsSfdWrite, m_lSendTimeout, ucsRdsLoaded) Then
                GoTo QH
            End If
            If Not pvWaitForEvent(ucsSfdAll, m_lReceiveTimeout, ucsRdsCompleted) Then
                GoTo QH
            End If
        End If
    #End If
    Exit Sub
QH:
    If m_lStatus <> 0 Then
        Abort
        On Error GoTo 0
        Err.Raise vbObjectError, , StatusText
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    GoTo QH
End Sub

Public Function WaitForResponse(Optional ByVal Timeout As Long) As Boolean
    #If ImplSync Then
        WaitForResponse = pvWaitForEvent(ucsSfdAll, Timeout, ucsRdsCompleted)
    #Else
        Err.Raise vbObjectError, , STR_ERR_SYNC_NOT_IMPL
    #End If
End Function

Public Sub Abort()
    #If ImplSync Then
        m_oSocket.SyncCancelWait
    #End If
    m_oSocket.Close_
End Sub

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long, ByVal PacketTimeout As Long)
    m_lResolveTimeout = ResolveTimeout
    m_lConnectTimeout = ConnectTimeout
    m_lSendTimeout = SendTimeout
    m_lReceiveTimeout = ReceiveTimeout
    m_lPacketTimeout = PacketTimeout
End Sub

'= private ===============================================================

#If ImplSync Then
Private Function pvWaitForEvent( _
        ByVal eEventMask As UcsAsyncSocketEventMaskEnum, _
        ByVal lTimeout As Long, _
        ByVal eReadyState As UcsReadyStateEnum) As Boolean
    Dim dblTimer        As Double
    Dim lRestTimeout    As Long
    
    dblTimer = DateTimer
    Do While ReadyState < eReadyState
        If lTimeout > 0 Then
            lRestTimeout = lTimeout - (DateTimer - dblTimer) * 1000
            If lRestTimeout < 0 Then
                pvSetStatus ERR_TIMEOUT
                GoTo QH
            End If
        End If
        If Not m_oSocket.SyncWaitForEvent(lRestTimeout, eEventMask) Then
            m_oSocket.SyncProcessMsgQueue
            If ReadyState >= eReadyState Then
                Exit Do
            End If
            pvSetStatus m_oSocket.LastError
            GoTo QH
        End If
        m_oSocket.SyncProcessMsgQueue
        If m_lStatus <> 0 Then
            GoTo QH
        End If
    Loop
    '--- success
    pvWaitForEvent = True
QH:
End Function
#End If

Private Sub pvSetReadyState(ByVal eValue As UcsReadyStateEnum)
    m_eNextState = ucsRdsUninitialized
    If m_eReadyState <> eValue Then
        m_eReadyState = eValue
        RaiseEvent OnReadyStateChange
    End If
End Sub

Private Sub pvSetStatus(ByVal lValue As Long)
    If m_lStatus <> lValue Then
        m_lStatus = lValue
        If lValue <> 0 Then
            RaiseEvent OnError(Status, StatusText)
        End If
    End If
End Sub

'= shared ================================================================

#If Not ImplUseShared Then
Private Property Get DateTimer() As Double
    Dim cDateTime       As Currency
    
    Call GetSystemTimeAsFileTime(cDateTime)
    DateTimer = CDbl(cDateTime - 9435304800000@) / 1000#
End Property

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function
#End If

'=========================================================================
' Events
'=========================================================================

Private Sub m_oSocket_OnResolve(Address As String)
    Const FUNC_NAME     As String = "m_oSocket_OnResolve"
    
    On Error GoTo EH
    If m_eNextState <> ucsRdsLoading Then
        Exit Sub
    End If
    m_sHostAddress = Address
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnConnect()
    Const FUNC_NAME     As String = "m_oSocket_OnConnect"
    
    On Error GoTo EH
    If m_eNextState <> ucsRdsLoading Then
        Exit Sub
    End If
    pvSetReadyState ucsRdsLoading
    m_eNextState = ucsRdsLoaded
    m_oSocket.PostEvent ucsSfdWrite
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    Dim lResult         As Long
    
    On Error GoTo EH
    If m_eNextState <> ucsRdsLoaded Then
        GoTo QH
    End If
    If m_lRequestBytes <= UBound(m_baRequest) Then
        '--- note: pass host/port to work w/ both TCP and UDP sockets
        lResult = m_oSocket.Send(VarPtr(m_baRequest(m_lRequestBytes)), UBound(m_baRequest) + 1 - m_lRequestBytes, _
            m_sHostAddress, m_lHostPort)
        If lResult < 0 Then
            pvSetStatus m_oSocket.LastError
            GoTo QH
        End If
        m_lRequestBytes = m_lRequestBytes + lResult
        m_oSocket.PostEvent ucsSfdWrite
    Else
        pvSetReadyState ucsRdsLoaded
        If m_lPacketTimeout > 0 Then
            m_eNextState = ucsRdsInteractive
            m_dblResponseTimer = DateTimer
            m_dblPacketTimer = m_dblResponseTimer
        Else
            pvSetReadyState ucsRdsCompleted
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eNextState <> ucsRdsInteractive And m_eNextState <> ucsRdsCompleted Then
        GoTo QH
    End If
    '--- note: pass host/port to work w/ both TCP and UDP sockets
    If Not m_oSocket.ReceiveArray(baBuffer, m_sHostAddress, m_lHostPort) Then
        pvSetStatus m_oSocket.LastError
        GoTo QH
    End If
    If UBound(baBuffer) >= 0 Then
        ReDim Preserve m_baResponse(0 To UBound(m_baResponse) + UBound(baBuffer) + 1) As Byte
        Call CopyMemory(m_baResponse(UBound(m_baResponse) - UBound(baBuffer)), baBuffer(0), UBound(baBuffer) + 1)
        pvSetReadyState ucsRdsInteractive
        m_eNextState = ucsRdsCompleted
        RaiseEvent OnResponseDataAvailable(baBuffer)
        m_dblPacketTimer = DateTimer
    ElseIf m_lPacketTimeout > 0 And m_lPacketTimeout < (DateTimer - m_dblPacketTimer) * 1000 Then
        pvSetReadyState ucsRdsCompleted
        GoTo QH
    ElseIf m_lReceiveTimeout > 0 And m_lReceiveTimeout < (DateTimer - m_dblResponseTimer) * 1000 Then
        pvSetStatus ERR_TIMEOUT
        GoTo QH
    End If
    m_oSocket.PostEvent [_ucsSfdForceRead]
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    
    On Error GoTo EH
    If m_eNextState <> ucsRdsCompleted Then
        Exit Sub
    End If
    pvSetReadyState ucsRdsCompleted
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    pvSetStatus ErrorCode
    Abort
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lResolveTimeout = 0
    m_lConnectTimeout = 60000
    m_lSendTimeout = 30000
    m_lReceiveTimeout = 30000
    m_lPacketTimeout = 50
    m_baResponse = vbNullString
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
End Sub

