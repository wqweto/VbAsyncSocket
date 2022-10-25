VERSION 5.00
Begin VB.UserControl ctxWinsock 
   BackColor       =   &H80000018&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Label labLogo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   576
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ctxWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2022 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "ctxWinsock"

'=========================================================================
' Events
'=========================================================================

Event Connect()
Event CloseEvent()
Event ConnectionRequest(ByVal requestID As Long)
Event DataArrival(ByVal bytesTotal As Long)
Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Event SendComplete()
Event Error(ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
    sckTLSProtocol = 2
End Enum

Public Enum UcsStateConstants
    sckClosed = 0
    sckOpen = 1
    sckListening = 2
    sckConnectionPending = 3
    sckResolvingHost = 4
    sckHostResolved = 5
    sckConnecting = 6
    sckConnected = 7
    sckClosing = 8
    sckError = 9
End Enum

Public Enum UcsErrorConstants
    sckInvalidPropertyValue = 380
    sckGetNotSupported = 394
    sckSetNotSupported = 383
    sckOutOfMemory = 7
    sckBadState = 40006
    sckInvalidArg = 40014
    sckSuccess = 40017
    sckUnsupported = 40018
    sckInvalidOp = 40020
    sckOutOfRange = 40021
    sckWrongProtocol = 40026
    sckOpCanceled = 10004
    sckInvalidArgument = 10014
    sckWouldBlock = 10035
    sckInProgress = 10036
    sckAlreadyComplete = 10037
    sckNotSocket = 10038
    sckMsgTooBig = 10040
    sckPortNotSupported = 10043
    sckAddressInUse = 10048
    sckAddressNotAvailable = 10049
    sckNetworkSubsystemFailed = 10050
    sckNetworkUnreachable = 10051
    sckNetReset = 10052
    sckConnectAborted = 10053
    sckConnectionReset = 10054
    sckNoBufferSpace = 10055
    sckAlreadyConnected = 10056
    sckNotConnected = 10057
    sckSocketShutdown = 10058
    sckTimedout = 10060
    sckConnectionRefused = 10061
    sckNotInitialized = 10093
    sckHostNotFound = 11001
    sckHostNotFoundTryAgain = 11002
    sckNonRecoverableError = 11003
    sckNoData = 11004
End Enum

Public Enum UcsSckLocalFeaturesEnum '--- bitmask
    ucsSckSupportTls10 = 2 ^ 0
    ucsSckSupportTls11 = 2 ^ 1
    ucsSckSupportTls12 = 2 ^ 2
    ucsSckSupportTls13 = 2 ^ 3
    ucsSckIgnoreServerCertificateErrors = 2 ^ 4
    ucsSckIgnoreServerCertificateRevocation = 2 ^ 5
    ucsSckSupportAll = ucsSckSupportTls10 Or ucsSckSupportTls11 Or ucsSckSupportTls12 Or ucsSckSupportTls13
End Enum

Public Enum UcsSckOptionLevelEnum
    ucsSckIP = 0
    ucsSckICMP = 1
    ucsSckIGMP = 2
    ucsSckTCP = 6
    ucsSckUDP = 17
    ucsSckSocket = &HFFFF&                  ' SOL_SOCKET
End Enum

Public Enum UcsSckOptionNameEnum
    ucsSckDebug = &H1                       ' Debugging is enabled.
    ucsSckAcceptConnection = &H2            ' Socket is listening.
    ucsSckReuseAddress = &H4                ' The socket can be bound to an address which is already in use. Not applicable for ATM sockets.
    ucsSckKeepAlive = &H8                   ' Keep-alives are being sent. Not supported on ATM sockets.
    ucsSckDontRoute = &H10                  ' Routing is disabled. Not supported on ATM sockets.
    ucsSckBroadcast = &H20                  ' Socket is configured for the transmission of broadcast messages.
    ucsSckUseLoopback = &H40                ' Bypass hardware when possible.
    ucsSckLinger = &H80                     ' Linger on close if unsent data is present.
    ucsSckOutOfBandInline = &H100           ' Receives out-of-band data in the normal data stream.
    ucsSckDontLinger = Not ucsSckLinger     ' Close socket gracefully without lingering.
    ucsSckExclusiveAddressUse = Not ucsSckReuseAddress ' Enables a socket to be bound for exclusive access.
    ucsSckSendBuffer = &H1001               ' Buffer size for sends.
    ucsSckReceiveBuffer = &H1002            ' Buffer size for receives.
    ucsSckSendLowWater = &H1003             ' Specifies the total per-socket buffer space reserved for receives.
    ucsSckReceiveLowWater = &H1004          ' Receive low water mark.
    ucsSckSendTimeout = &H1005              ' Sends time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSckReceiveTimeout = &H1006           ' Receives time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSckError = &H1007                    ' Get error status and clear.
    ucsSckType = &H1008                     ' Get socket type.
'    ucsSckGroupId = &H2001                  ' Reserved.
'    ucsSckGroupPriority = &H2002            ' Reserved.
    ucsSckMaxMsgSize = &H2003               ' Maximum size of a message for message-oriented socket types (for example, SOCK_DGRAM). Has no meaning for stream oriented sockets.
    ucsSckProtocolInfo = &H2004             ' Description of protocol information for protocol that is bound to this socket.
    ucsSckReuseUnicastPort = &H3007         ' Defer ephemeral port allocation for outbound connections
    ucsSckMaxConnections = &H7FFFFFFF       ' Maximum queue length specifiable by listen.
    '-- IP
    ucsSckIPOptions = 1                     ' IP options.
    ucsSckHeaderIncluded = 2                ' Header is included with data.
    ucsSckTypeOfService = 3                 ' IP type of service and preced.
    ucsSckIpTimeToLive = 4                  ' IP time to live.
    ucsSckMulticastInterface = 9            ' IP multicast interface.
    ucsSckMulticastTimeToLive = 10          ' IP multicast time to live.
    ucsSckMulticastLoopback = 11            ' IP Multicast loopback.
    ucsSckAddMembership = 12                ' Add an IP group membership.
    ucsSckDropMembership = 13               ' Drop an IP group membership.
    ucsSckDontFragment = 14                 ' Don't fragment IP datagrams.
    ucsSckAddSourceMembership = 15          ' Join IP group/source.
    ucsSckDropSourceMembership = 16         ' Leave IP group/source.
    ucsSckBlockSource = 17                  ' Block IP group/source.
    ucsSckUnblockSource = 18                ' Unblock IP group/source.
    ucsSckPacketInformation = 19            ' Receive packet information for ipv4.
    '-- TCP
    ucsSckNoDelay = 1                       ' Disables the Nagle algorithm for send coalescing.
    ucsSckExpedited = 2
    '--- UDP
    ucsSckNoChecksum = 1
    ucsSckChecksumCoverage = 20             ' Udp-Lite checksum coverage.
    ucsSckUpdateAcceptContext = &H700B
    ucsSckUpdateConnectContext = &H7010
End Enum

'=========================================================================
' API
'=========================================================================

Private Const DUPLICATE_SAME_ACCESS         As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_LOGO              As String = "AS" & vbCrLf & "WS"
Private Const DEF_LOCALPORT         As Long = 0
Private Const DEF_PROTOCOL          As Long = 0
Private Const DEF_REMOTEHOST        As String = vbNullString
Private Const DEF_REMOTEPORT        As Long = 0
Private Const DEF_TIMEOUT           As Long = 5000

Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_eState                As UcsStateConstants
Private m_lLocalPort            As Long
Private m_eProtocol             As UcsProtocolConstants
Private m_sRemoteHost           As String
Private m_lRemotePort           As Long
Private m_lTimeout              As Long
Private m_baRecvBuffer()        As Byte
Private m_baSendBuffer()        As Byte
Private m_lSendPos              As Long
Private m_oRequestSocket        As cTlsSocket

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

Property Get LocalPort() As Long
    If pvHasSocket Then
        pvSocket.GetSockName vbNullString, LocalPort
    Else
        LocalPort = m_lLocalPort
    End If
End Property

Property Let LocalPort(ByVal lValue As Long)
    If m_lLocalPort <> lValue Then
        Close_
        pvState = sckClosed
        m_lLocalPort = lValue
        PropertyChanged
    End If
End Property

Property Get Protocol() As UcsProtocolConstants
    Protocol = m_eProtocol
End Property

Property Let Protocol(ByVal eValue As UcsProtocolConstants)
    If m_eProtocol <> eValue Then
        Close_
        pvState = sckClosed
        m_eProtocol = eValue
        PropertyChanged
    End If
End Property

Property Get RemoteHost() As String
    RemoteHost = m_sRemoteHost
End Property

Property Let RemoteHost(sValue As String)
    If m_sRemoteHost <> sValue Then
        m_sRemoteHost = sValue
        m_baSendBuffer = vbNullString
        PropertyChanged
    End If
End Property

Property Get RemotePort() As Long
    If pvHasSocket Then
        pvSocket.GetPeerName vbNullString, RemotePort
    Else
        RemotePort = m_lRemotePort
    End If
End Property

Property Let RemotePort(ByVal lValue As Long)
    If m_lRemotePort <> lValue Then
        m_lRemotePort = lValue
        m_baSendBuffer = vbNullString
        PropertyChanged
    End If
End Property

Property Get Timeout() As Long
    Timeout = m_lTimeout
End Property

Property Let Timeout(ByVal lValue As Long)
    If m_lTimeout <> lValue Then
        m_lTimeout = lValue
        PropertyChanged
    End If
End Property

'= run-time ==============================================================

Property Get SocketHandle() As Long
    SocketHandle = pvSocket.SocketHandle
End Property

Property Get State() As UcsStateConstants
    State = m_eState
End Property

Private Property Let pvState(ByVal eValue As UcsStateConstants)
    m_eState = eValue
    m_baRecvBuffer = vbNullString
    m_baSendBuffer = vbNullString
End Property

Property Get LocalHostName() As String
    pvSocket.GetLocalHost LocalHostName, vbNullString
End Property

Property Get LocalIP() As String
    pvSocket.GetLocalHost vbNullString, LocalIP
End Property

Property Get RemoteHostIP() As String
    pvSocket.GetPeerName RemoteHostIP, 0
End Property

Property Get SockOpt(ByVal OptionName As UcsSckOptionNameEnum, Optional ByVal Level As UcsSckOptionLevelEnum = ucsSckSocket) As Long
    SockOpt = pvSocket.Socket.SockOpt(OptionName, Level)
End Property

Property Let SockOpt(ByVal OptionName As UcsSckOptionNameEnum, Optional ByVal Level As UcsSckOptionLevelEnum = ucsSckSocket, ByVal Value As Long)
    pvSocket.Socket.SockOpt(OptionName, Level) = Value
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

'=========================================================================
' Methods
'=========================================================================

Public Sub Accept(ByVal requestID As Long)
    Dim hDuplicate      As Long
    
    On Error GoTo EH
    Close_
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

Public Sub Close_()
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

Public Sub Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String)
    On Error GoTo EH
    Close_
    If LocalPort <> 0 Then
        m_lLocalPort = LocalPort
    End If
    If Not pvSocket.Bind(LocalIP, m_lLocalPort) Then
        On Error GoTo 0
        pvSetError LastError:=m_oSocket.LastError, RaiseError:=True
    End If
    pvState = sckOpen
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub Connect(Optional RemoteHost As String, Optional ByVal RemotePort As Long, Optional ByVal LocalFeatures As UcsSckLocalFeaturesEnum)
    On Error GoTo EH
    Close_
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

Public Sub Listen( _
            Optional CertFile As String, _
            Optional Password As String, _
            Optional CertSubject As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection)
    On Error GoTo EH
    If m_eProtocol = sckTLSProtocol Then
        If Not pvSocket.InitServerTls(CertFile, Password, CertSubject, Certificates, PrivateKey) Then
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

Public Sub PeekData(data As Variant, Optional ByVal type_ As Long, Optional ByVal maxLen As Long = -1)
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo EH
    If type_ = 0 Then
        type_ = VarType(data)
    End If
    Select Case type_
    Case vbString, vbByte + vbArray
    Case Else
        ErrRaise vbObjectError, , "Unsupported data type: " & type_
    End Select
    If maxLen < 0 Then
        If pvSocket.AvailableBytes <= 0 Then
            baBuffer = vbNullString
        Else
            pvSocket.SyncReceiveArray baBuffer, Timeout:=m_lTimeout
        End If
    Else
        pvSocket.SyncReceiveArray baBuffer, maxLen, Timeout:=m_lTimeout
    End If
    Select Case type_
    Case vbString
        data = pvSocket.FromTextArray(baBuffer, ucsScpAcp)
    Case vbByte + vbArray
        data = baBuffer
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

Public Sub GetData(data As Variant, Optional ByVal type_ As Long, Optional ByVal maxLen As Long = -1)
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If type_ = 0 Then
        type_ = VarType(data)
    End If
    Select Case type_
    Case vbString, vbByte + vbArray
    Case Else
        ErrRaise vbObjectError, , "Unsupported data type: " & type_
    End Select
    baBuffer = vbNullString
    If UBound(m_baRecvBuffer) >= 0 Then
        If maxLen < 0 Then
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
        If maxLen < 0 Then
            If pvSocket.AvailableBytes <= 0 Then
                baBuffer = vbNullString
            Else
                pvSocket.SyncReceiveArray baBuffer, Timeout:=m_lTimeout
            End If
        Else
            pvSocket.SyncReceiveArray baBuffer, maxLen, Timeout:=m_lTimeout
        End If
    End If
    Select Case type_
    Case vbString
        data = pvSocket.FromTextArray(baBuffer, ucsScpAcp)
    Case vbByte + vbArray
        data = baBuffer
    End Select
    Exit Sub
EH:
    pvSetError LastError:=Err, RaiseError:=True
End Sub

Public Sub SendData(data As Variant)
    Dim baAppend()      As Byte
    Dim lPos            As Long
    
    On Error GoTo EH
    If UBound(m_baSendBuffer) < 0 Then
        Select Case VarType(data)
        Case vbString
            m_baSendBuffer = pvSocket.ToTextArray(CStr(data), ucsScpAcp)
        Case vbByte + vbArray
            m_baSendBuffer = data
        Case Else
            ErrRaise vbObjectError, , "Unsupported data type: " & TypeName(data)
        End Select
    Else
        Select Case VarType(data)
        Case vbString
            baAppend = pvSocket.ToTextArray(CStr(data), ucsScpAcp)
        Case vbByte + vbArray
            baAppend = data
        Case Else
            ErrRaise vbObjectError, , "Unsupported data type: " & TypeName(data)
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
    RaiseEvent CloseEvent
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If
    pvState = sckClosed
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
    pvState = sckListening
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnResolve(IpAddress As String)
    pvState = sckHostResolved
End Sub

Private Sub m_oSocket_OnReceive()
    If m_eState = 0 Then
        m_oSocket_OnConnect
    End If
    RaiseEvent DataArrival(pvSocket.AvailableBytes)
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
' Control events
'=========================================================================

Private Sub UserControl_Resize()
    '--- note: skip error handler not to clear Err object
    Width = ScaleX(32, vbPixels)
    Height = ScaleX(32, vbPixels)
    labLogo.Move 0, (ScaleHeight - labLogo.Height) / 2, ScaleWidth
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    labLogo.Caption = STR_LOGO
    LocalPort = DEF_LOCALPORT
    Protocol = DEF_PROTOCOL
    RemoteHost = DEF_REMOTEHOST
    RemotePort = DEF_REMOTEPORT
    Timeout = DEF_TIMEOUT
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    labLogo.Caption = STR_LOGO
    With PropBag
        LocalPort = .ReadProperty("LocalPort", DEF_LOCALPORT)
        Protocol = .ReadProperty("Protocol", DEF_PROTOCOL)
        RemoteHost = .ReadProperty("RemoteHost", DEF_REMOTEHOST)
        RemotePort = .ReadProperty("RemotePort", DEF_REMOTEPORT)
        Timeout = .ReadProperty("Timeout", DEF_TIMEOUT)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "LocalPort", LocalPort, DEF_LOCALPORT
        .WriteProperty "Protocol", Protocol, DEF_PROTOCOL
        .WriteProperty "RemoteHost", RemoteHost, DEF_REMOTEHOST
        .WriteProperty "RemotePort", RemotePort, DEF_REMOTEPORT
        .WriteProperty "Timeout", Timeout, DEF_TIMEOUT
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_Initialize()
    m_baRecvBuffer = vbNullString
    m_baSendBuffer = vbNullString
End Sub
