VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTlsRemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cTlsRemaster"

'=========================================================================
' Public enums
'=========================================================================

Public Enum RemasterProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
    sckTLSProtocol = 2
End Enum

Public Enum RemasterSockState
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum

'=========================================================================
' Public events
'=========================================================================

Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

'=========================================================================
' API
'=========================================================================

Private Const DUPLICATE_SAME_ACCESS         As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'=========================================================================
' Contants and member variables
'=========================================================================

Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_eState                As RemasterSockState
Private m_lLocalPort            As Long
Private m_eProtocol             As RemasterProtocolConstants
Private m_sRemoteHost           As String
Private m_lRemotePort           As Long
Private m_lTimeout              As Long
Private m_baRecvBuffer()        As Byte
Private m_baSendBuffer()        As Byte
Private m_lSendPos              As Long
Private m_oRequestSocket        As cTlsSocket
Private m_sTag                  As String

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

Private Sub ErrRaise(ByVal Number As Long, Optional Source As Variant, Optional Description As Variant)
    Err.Raise Number, Source, Description
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get BytesReceived() As Long
    BytesReceived = UBound(m_baRecvBuffer) + 1
End Property

Public Property Get LocalHostName() As String
    pvSocket.GetLocalHost LocalHostName, vbNullString
End Property

Public Property Get LocalIP() As String
    pvSocket.GetLocalHost vbNullString, LocalIP
End Property

Public Property Get LocalPort() As Long
    If pvHasSocket Then
        pvSocket.GetSockName vbNullString, LocalPort
    Else
        LocalPort = m_lLocalPort
    End If
End Property

Public Property Let LocalPort(ByVal lValue As Long)
    If m_lLocalPort <> lValue Then
        CloseSck
        m_lLocalPort = lValue
    End If
End Property

Public Property Get Protocol() As RemasterProtocolConstants
    Protocol = m_eProtocol
End Property

Public Property Let Protocol(ByVal eValue As RemasterProtocolConstants)
    If m_eProtocol <> eValue Then
        CloseSck
        m_eProtocol = eValue
    End If
End Property

Public Property Get RemoteHost() As String
    RemoteHost = m_sRemoteHost
End Property

Public Property Let RemoteHost(ByVal sValue As String)
    If m_sRemoteHost <> sValue Then
        m_sRemoteHost = sValue
        m_baSendBuffer = vbNullString
    End If
End Property

Public Property Get RemoteHostIP() As String
    pvSocket.GetPeerName RemoteHostIP, 0
End Property

Public Property Get RemotePort() As Long
    If pvHasSocket Then
        pvSocket.GetPeerName vbNullString, RemotePort
    Else
        RemotePort = m_lRemotePort
    End If
End Property

Public Property Let RemotePort(ByVal lValue As Long)
    If m_lRemotePort <> lValue Then
        m_lRemotePort = lValue
        m_baSendBuffer = vbNullString
    End If
End Property

Public Property Get SocketHandle() As Long
    If pvHasSocket Then
        SocketHandle = pvSocket.SocketHandle
    End If
End Property

Public Property Get State() As RemasterSockState
    State = m_eState
End Property

Public Property Let Tag(ByVal sValue As String)
    m_sTag = sValue
End Property

Public Property Get Tag() As String
    Tag = m_sTag
End Property

'= private ===============================================================

Private Property Get pvHasSocket() As Boolean
    If Not m_oRequestSocket Is Nothing Or Not m_oSocket Is Nothing Then
        pvHasSocket = True
    End If
End Property

Private Property Get pvSocket() As cTlsSocket
    Const FUNC_NAME     As String = "pvSocket [get]"
    
    On Error GoTo EH
    If Not m_oRequestSocket Is Nothing Then
        Set pvSocket = m_oRequestSocket
    Else
        If m_oSocket Is Nothing Then
            Set m_oSocket = New cTlsSocket
            m_oSocket.Create SocketType:=(m_eProtocol And 1)
        End If
        Set pvSocket = m_oSocket
    End If
    Exit Property
EH:
    PrintError FUNC_NAME
End Property

Private Property Let pvState(ByVal eValue As RemasterSockState)
    m_eState = eValue
    If eValue <> sckClosing Then
        m_baRecvBuffer = vbNullString
        m_baSendBuffer = vbNullString
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Accept(requestID As Long)
    Dim hDuplicate      As Long
    
    On Error GoTo EH
    CloseSck
    If Not g_oRequestSocket Is Nothing Then
        If g_oRequestSocket.SocketHandle = requestID Then
            Set m_oSocket = g_oRequestSocket
        End If
    End If
    If m_oSocket Is Nothing Then
        If DuplicateHandle(GetCurrentProcess(), requestID, GetCurrentProcess(), hDuplicate, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then
            On Error GoTo 0
            pvSetError LastDllError:=Err.LastDllError, RaiseError:=True
        End If
        Set m_oSocket = New cTlsSocket
        If Not m_oSocket.Attach(hDuplicate) Then
            On Error GoTo 0
            pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
        End If
    End If
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub Bind(Optional LocalPort As Variant, Optional LocalIP As Variant)
    On Error GoTo EH
    CloseSck
    If LocalPort <> 0 Then
        m_lLocalPort = LocalPort
    End If
    If Not pvSocket.Bind(LocalIP & vbNullString, m_lLocalPort) Then
        On Error GoTo 0
        pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
    End If
    pvState = sckOpen
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub CloseSck()
    On Error GoTo EH
    If State <> sckClosed Then
        If Not m_oSocket Is Nothing Then
            pvState = sckClosing
            m_oSocket.Close_
            Set m_oSocket = Nothing
        End If
        pvState = sckClosed
    End If
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub Connect(Optional RemoteHost As Variant, Optional RemotePort As Variant)
    Const LocalFeatures As Long = 0
    On Error GoTo EH
    CloseSck
    If LenB(RemoteHost) <> 0 Then
        m_sRemoteHost = RemoteHost
    End If
    If RemotePort <> 0 Then
        m_lRemotePort = RemotePort
    End If
    pvState = sckResolvingHost
    If Not pvSocket.Connect(m_sRemoteHost, m_lRemotePort, UseTls:=(m_eProtocol = sckTLSProtocol), LocalFeatures:=LocalFeatures) Then
        On Error GoTo 0
        pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
    End If
    pvState = sckConnected
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub GetData(Data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If IsMissing(varType) Then
        varType = VBA.varType(Data)
    End If
    Select Case varType
    Case vbString, vbByte + vbArray
    Case Else
        ErrRaise vbObjectError, , "Unsupported data type: " & varType
    End Select
    baBuffer = vbNullString
    If UBound(m_baRecvBuffer) >= 0 Then
        If IsMissing(maxLen) Then
            baBuffer = m_baRecvBuffer
            m_baRecvBuffer = vbNullString
        ElseIf maxLen = 0 Then
            baBuffer = vbNullString
        Else
            baBuffer = m_baRecvBuffer
            lIdx = UBound(m_baRecvBuffer) + 1 - maxLen
            If lIdx > 0 Then
                ReDim m_baRecvBuffer(0 To lIdx - 1) As Byte
                Call CopyMemory(m_baRecvBuffer(0), baBuffer(maxLen), lIdx)
                ReDim Preserve baBuffer(0 To maxLen - 1)
            Else
                m_baRecvBuffer = vbNullString
            End If
        End If
    Else
        If IsMissing(maxLen) Then
            If pvSocket.AvailableBytes <= 0 Then
                baBuffer = vbNullString
            Else
                pvSocket.SyncReceiveArray baBuffer, Timeout:=m_lTimeout
            End If
        Else
            pvSocket.SyncReceiveArray baBuffer, maxLen, Timeout:=m_lTimeout
        End If
    End If
    Select Case varType
    Case vbString
        Data = pvSocket.FromTextArray(baBuffer, ucsScpAcp)
    Case vbByte + vbArray
        Data = baBuffer
    End Select
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub Listen( _
            Optional CertFile As String, _
            Optional Password As String, _
            Optional CertSubject As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection, _
            Optional AlpnProtocols As String, _
            Optional ByVal LocalFeatures As Long)
    On Error GoTo EH
    If m_eProtocol = sckTLSProtocol Then
        If Not pvSocket.InitServerTls(CertFile, Password, CertSubject, Certificates, PrivateKey, AlpnProtocols, LocalFeatures) Then
            On Error GoTo 0
            pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
        End If
    End If
    If Not pvSocket.Listen() Then
        On Error GoTo 0
        pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
    End If
    pvState = sckListening
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub PeekData(Data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo EH
    If IsMissing(varType) Then
        varType = VBA.varType(Data)
    End If
    Select Case varType
    Case vbString, vbByte + vbArray
    Case Else
        ErrRaise vbObjectError, , "Unsupported data type: " & varType
    End Select
    If IsMissing(maxLen) Then
        If pvSocket.AvailableBytes <= 0 Then
            baBuffer = vbNullString
        Else
            pvSocket.SyncReceiveArray baBuffer, Timeout:=m_lTimeout
        End If
    Else
        pvSocket.SyncReceiveArray baBuffer, maxLen, Timeout:=m_lTimeout
    End If
    Select Case varType
    Case vbString
        Data = pvSocket.FromTextArray(baBuffer, ucsScpAcp)
    Case vbByte + vbArray
        Data = baBuffer
    End Select
    If UBound(m_baRecvBuffer) >= 0 And UBound(baBuffer) >= 0 Then
        lIdx = UBound(m_baRecvBuffer) + 1
        ReDim Preserve m_baRecvBuffer(0 To lIdx + UBound(baBuffer))
        Call CopyMemory(m_baRecvBuffer(lIdx), baBuffer(0), UBound(baBuffer) + 1)
    Else
        m_baRecvBuffer = baBuffer
    End If
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub SendData(Data As Variant)
    Dim baAppend()      As Byte
    Dim lPos            As Long
    
    On Error GoTo EH
    If UBound(m_baSendBuffer) < 0 Then
        Select Case VBA.varType(Data)
        Case vbString
            m_baSendBuffer = pvSocket.ToTextArray(CStr(Data), ucsScpAcp)
        Case vbByte + vbArray
            m_baSendBuffer = Data
        Case Else
            ErrRaise vbObjectError, , "Unsupported data type: " & TypeName(Data)
        End Select
    Else
        Select Case VBA.varType(Data)
        Case vbString
            baAppend = pvSocket.ToTextArray(CStr(Data), ucsScpAcp)
        Case vbByte + vbArray
            baAppend = Data
        Case Else
            ErrRaise vbObjectError, , "Unsupported data type: " & TypeName(Data)
        End Select
        If UBound(baAppend) >= 0 Then
            lPos = UBound(m_baSendBuffer) + 1
            ReDim Preserve m_baSendBuffer(0 To lPos + UBound(baAppend)) As Byte
            Call CopyMemory(m_baSendBuffer(lPos), baAppend(0), UBound(baAppend) + 1)
        End If
    End If
    If UBound(m_baSendBuffer) >= 0 Then
        m_oSocket_OnSend
    End If
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Private Sub pvSetError(Optional ByVal LastDllError As Long, Optional LastError As VBA.ErrObject, Optional ByVal RaiseError As Boolean)
    Const LNG_FACILITY_WIN32 As Long = &H80070000
    Dim Number          As Long
    Dim Description     As String
    Dim Source          As String
    Dim bCancel         As Boolean
    
    pvState = sckError
    If LastDllError <> 0 Then
        Number = LastDllError Or IIf(LastDllError < 0, 0, LNG_FACILITY_WIN32)
        Description = pvSocket.GetErrorDescription(LastDllError)
    ElseIf Not LastError Is Nothing Then
        Number = LastError.Number
        Source = LastError.Source
        Description = LastError.Description
        LastDllError = LastError.Number And &HFFFF&
    End If
    RaiseEvent Error(Number, Description, LastDllError, Source, App.HelpFile, 0, bCancel)
    If Not bCancel And RaiseError Then
        ErrRaise Number, Source, Description
    End If
End Sub

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnConnect()
    pvState = sckConnected
    RaiseEvent Connect
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    
    On Error GoTo EH
    pvState = sckClosing
    RaiseEvent CloseSck
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnAccept()
    Const FUNC_NAME     As String = "m_oSocket_OnAccept"
    Dim oTemp           As cTlsSocket
    
    On Error GoTo EH
    pvState = sckConnectionPending
    Set oTemp = New cTlsSocket
    If Not m_oSocket.Accept(oTemp, UseTls:=(m_eProtocol = sckTLSProtocol)) Then
        pvSetError LastError:=m_oSocket.LastError
        GoTo QH
    End If
    Set m_oRequestSocket = oTemp
    Set g_oRequestSocket = oTemp
    RaiseEvent ConnectionRequest(oTemp.SocketHandle)
    Set m_oRequestSocket = Nothing
    Set g_oRequestSocket = Nothing
QH:
    pvState = sckListening
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnResolve(IpAddress As String)
    pvState = sckHostResolved
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eState = 0 Then
        m_oSocket_OnConnect
    End If
    RaiseEvent DataArrival(pvSocket.AvailableBytes)
    PeekData baBuffer
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    Dim lSent           As Long
    
    On Error GoTo EH
    If m_eState = 0 Then
        m_oSocket_OnConnect
    End If
    Do While m_lSendPos <= UBound(m_baSendBuffer)
        lSent = pvSocket.Send(VarPtr(m_baSendBuffer(m_lSendPos)), UBound(m_baSendBuffer) + 1 - m_lSendPos, m_sRemoteHost, m_lRemotePort)
        If lSent < 0 Then
            If Not pvSocket.HasPendingEvent Then
                pvSetError LastError:=m_oSocket.LastError
                GoTo QH
            End If
            Exit Do
        Else
            m_lSendPos = m_lSendPos + lSent
            RaiseEvent SendProgress(m_lSendPos, UBound(m_baSendBuffer) + 1 - m_lSendPos)
        End If
    Loop
    If m_lSendPos > UBound(m_baSendBuffer) And m_lSendPos > 0 Then
        m_lSendPos = 0
        m_baSendBuffer = vbNullString
        RaiseEvent SendComplete
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    If m_oSocket.LastError <> 0 Then
        pvSetError LastDllError:=ErrorCode
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    m_baRecvBuffer = vbNullString
    m_baSendBuffer = vbNullString
End Sub
