VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncSocket"
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
Private Const MODULE_NAME As String = "cAsyncSocket"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplSync = Not (ASYNCSOCKET_NO_SYNC <> 0)

'=========================================================================
' Public events
'=========================================================================

Event OnResolve(IpAddress As String)
Event OnAccept()
Event OnClose()
Event OnConnect()
Event OnReceive()
Event OnSend()
Event OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
Event OnMessagePending(Handled As Boolean)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsAsyncSocketTypeEnum
    ucsSckStream = 1
    ucsSckDatagram = 2
End Enum

Public Enum UcsAsyncSocketEventMaskEnum
    ucsSfdRead = 2 ^ 0
    ucsSfdWrite = 2 ^ 1
    ucsSfdOob = 2 ^ 2
    ucsSfdAccept = 2 ^ 3
    ucsSfdConnect = 2 ^ 4
    ucsSfdClose = 2 ^ 5
    ucsSfdAll = 2 ^ 6 - 1
    [_ucsSfdResolve] = 2 ^ 15
    [_ucsSfdForceRead] = 2 ^ 14
End Enum

Public Enum UcsAsyncSocketOptionLevelEnum
    ucsSolSocket = &HFFFF&                  ' SOL_SOCKET
End Enum

Public Enum UcsAsyncSocketOptionNameEnum
    ucsSsoDebug = &H1                       ' Debugging is enabled.
    ucsSsoAcceptConnection = &H2            ' Socket is listening.
    ucsSsoReuseAddress = &H4                ' The socket can be bound to an address which is already in use. Not applicable for ATM sockets.
    ucsSsoKeepAlive = &H8                   ' Keep-alives are being sent. Not supported on ATM sockets.
    ucsSsoDontRoute = &H10                  ' Routing is disabled. Not supported on ATM sockets.
    ucsSsoBroadcast = &H20                  ' Socket is configured for the transmission of broadcast messages.
    ucsSsoUseLoopback = &H40                ' Bypass hardware when possible.
    ucsSsoLinger = &H80                     ' Linger on close if unsent data is present.
    ucsSsoOutOfBandInline = &H100           ' Receives out-of-band data in the normal data stream.
    ucsSsoDontLinger = Not ucsSsoLinger     ' Close socket gracefully without lingering.
    ucsSsoExclusiveAddressUse = Not ucsSsoReuseAddress ' Enables a socket to be bound for exclusive access.
    ucsSsoSendBuffer = &H1001               ' Buffer size for sends.
    ucsSsoReceiveBuffer = &H1002            ' Buffer size for receives.
    ucsSsoSendLowWater = &H1003             ' Specifies the total per-socket buffer space reserved for receives.
    ucsSsoReceiveLowWater = &H1004          ' Receive low water mark.
    ucsSsoSendTimeout = &H1005              ' Sends time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSsoReceiveTimeout = &H1006           ' Receives time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSsoError = &H1007                    ' Get error status and clear.
    ucsSsoType = &H1008                     ' Get socket type.
'    ucsSsoGroupId = &H2001                  ' Reserved.
'    ucsSsoGroupPriority = &H2002            ' Reserved.
    ucsSsoMaxMsgSize = &H2003               ' Maximum size of a message for message-oriented socket types (for example, SOCK_DGRAM). Has no meaning for stream oriented sockets.
    ucsSsoProtocolInfo = &H2004             ' Description of protocol information for protocol that is bound to this socket.
    ucsSsoReuseUnicastPort = &H3007         ' Defer ephemeral port allocation for outbound connections
    ucsSsoMaxConnections = &H7FFFFFFF       ' Maximum queue length specifiable by listen.
End Enum

Public Enum UcsAsyncSocketCodePageEnum
    ucsScpAcp = 0
    ucsScpOem = 1
    ucsScpUtf8 = 65001
    ucsScpUnicode = -1                      '--- handled internally -> no conversion, keep text wide-char
    ucsScpNone = ucsScpUnicode              '--- and some more aliases
    ucsScpWideChar = ucsScpUnicode
End Enum

'=========================================================================
' API
'=========================================================================

Private Const SOCKET_ERROR                  As Long = -1
Private Const INVALID_SOCKET                As Long = -1
Private Const AF_INET                       As Long = 2
Private Const ERR_TIMEOUT                   As Long = &H800705B4
'--- for ioctlsocket
Private Const FIONREAD                      As Long = &H4004667F
'--- for WSAGetLastError
Private Const WSABASEERR                    As Long = 10000
Private Const WSAEINTR                      As Long = (WSABASEERR + 4)
Private Const WSAEINVAL                     As Long = (WSABASEERR + 22)
Private Const WSAEMFILE                     As Long = (WSABASEERR + 24)
Private Const WSAEWOULDBLOCK                As Long = (WSABASEERR + 35)
Private Const WSAEALREADY                   As Long = (WSABASEERR + 37)
Private Const WSAENOTCONN                   As Long = (WSABASEERR + 57)
Private Const WSANOTINITIALISED             As Long = (WSABASEERR + 93)
'--- for WSAAsyncGetHostByName
Private Const MAXGETHOSTSTRUCT              As Long = 1024
Private Const INADDR_ANY                    As Long = 0
Private Const INADDR_NONE                   As Long = -1
'--- Windows Messages
Private Const WM_PAINT                      As Long = &HF
Private Const WM_TIMER                      As Long = &H113
Private Const WM_USER                       As Long = &H400
Private Const WM_SOCKET_RESOLVE             As Long = WM_USER + 1
Private Const WM_SOCKET_NOTIFY              As Long = WM_USER + 2
'--- for Get/SetWindowLong
Private Const GWL_USERDATA                  As Long = (-21)
'--- for PeekMessage
Private Const PM_NOREMOVE                   As Long = 0
Private Const PM_REMOVE                     As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function ws_htonl Lib "ws2_32" Alias "htonl" (ByVal hostlong As Long) As Long
Private Declare Function ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function ws_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function ws_inet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_recvfrom Lib "ws2_32" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
Private Declare Function ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_sendto Lib "ws2_32" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef toaddr As SOCKADDR_IN, ByVal tolen As Long) As Long
Private Declare Function ws_getpeername Lib "ws2_32" Alias "getpeername" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function ws_getsockname Lib "ws2_32" Alias "getsockname" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function ws_shutdown Lib "ws2_32" Alias "shutdown" (ByVal s As Long, ByVal How As Long) As Long
Private Declare Function ws_listen Lib "ws2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function ws_accept Lib "ws2_32" Alias "accept" (ByVal s As Long, addr As Any, addrlen As Any) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_inet_addr Lib "ws2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function ws_gethostbyname Lib "ws2_32" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function ws_gethostname Lib "ws2_32" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function ws_WSAAsyncGetHostByName Lib "ws2_32" Alias "WSAAsyncGetHostByName" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal BufLen As Long) As Long
Private Declare Function ws_WSAAsyncSelect Lib "ws2_32" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function ws_WSACancelAsyncRequest Lib "ws2_32" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long
#If ImplSync Then
    Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As APIMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As APIMSG) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function WaitMessage Lib "user32" () As Long
#End If

Private Type SOCKADDR_IN
    sin_family          As Integer
    sin_port            As Integer
    sin_addr            As Long
    sin_zero(1 To 8)    As Byte
End Type

Private Type APIPOINT
    X                   As Long
    Y                   As Long
End Type

Private Type APIMSG
    hWnd                As Long
    lMessage            As Long
    wParam              As Long
    lParam              As Long
    lTime               As Long
    pt                  As APIPOINT
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_HELPER_CLASS      As String = "STATIC"
Private Const STR_HELPER_CAPTION    As String = "AsyncSocket Helper Window"
Private Const STR_CHR1              As String = "" '--- CHAR(1)
Private Const DEF_TIMEOUT           As Long = 5000

Private m_oHelperWindow         As cAsyncSocketHelperWindow
Private m_hSocket               As Long
Private m_lIndex                As Long
Private m_cMsgQueue             As Collection
Private m_uGetHostByName        As UcsAsyncGetHostByNameType
Private m_bBlocking             As Boolean
Private m_lLastError            As Long

Private Type UcsAsyncGetHostByNameType
    Handle          As Long
    Buffer()        As Byte
    Port            As Long
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

Property Get SocketHandle() As Long
    SocketHandle = m_hSocket
End Property

Property Get LastError() As Long
    LastError = m_lLastError
End Property

Property Get SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket) As Long
    Call ws_getsockopt(m_hSocket, Level, OptionName, SockOpt, 4)
    m_lLastError = Err.LastDllError
End Property

Property Let SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket, ByVal Value As Long)
    Call ws_setsockopt(m_hSocket, Level, OptionName, Value, 4)
    m_lLastError = Err.LastDllError
End Property

Property Get AvailableBytes() As Long
    If Not IOCtl(FIONREAD, AvailableBytes) Then
        AvailableBytes = SOCKET_ERROR
    End If
End Property

Property Get HasPendingEvent() As Boolean
    HasPendingEvent = (m_lLastError = WSAEWOULDBLOCK)
End Property

Property Get HasPendingResolve() As Boolean
    HasPendingResolve = (m_uGetHostByName.Handle <> 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Create( _
            Optional ByVal SocketPort As Long, _
            Optional ByVal SocketType As UcsAsyncSocketTypeEnum = ucsSckStream, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll, _
            Optional SocketAddress As String) As Boolean
    Const FUNC_NAME     As String = "Create"
    
    On Error GoTo EH
    If m_hSocket <> INVALID_SOCKET Then
        m_lLastError = WSAEALREADY
        GoTo QH
    End If
    If m_oHelperWindow.hWnd = 0 Then
        m_lLastError = WSANOTINITIALISED
        GoTo QH
    End If
    m_hSocket = ws_socket(AF_INET, SocketType, 0)
    If m_hSocket = INVALID_SOCKET Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    If Not m_oHelperWindow.frAddSocket(Me, m_lIndex) Then
        m_lLastError = WSAEMFILE
        Close_
        GoTo QH
    End If
    If Not AsyncSelect(EventMask) Then
        Close_
        GoTo QH
    End If
    If Not Bind(SocketAddress, SocketPort) Then
        Close_
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    Create = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Attach( _
            ByVal SocketHandle As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Const FUNC_NAME     As String = "Attach"

    On Error GoTo EH
    If m_hSocket <> INVALID_SOCKET Then
        m_lLastError = WSAEALREADY
        GoTo QH
    End If
    m_hSocket = SocketHandle
    If Not m_oHelperWindow.frAddSocket(Me, m_lIndex) Then
        m_hSocket = INVALID_SOCKET
        m_lLastError = WSAEMFILE
        GoTo QH
    End If
    If Not AsyncSelect(EventMask) Then
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    Attach = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Detach() As Long
    Const FUNC_NAME     As String = "Detach"

    On Error GoTo EH
    m_lLastError = 0
    Detach = m_hSocket
    m_oHelperWindow.frRemoveSocket Me, m_lIndex
    m_hSocket = INVALID_SOCKET
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetPeerName(PeerAddress As String, PeerPort As Long) As Boolean
    Const FUNC_NAME     As String = "GetPeerName"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If ws_getpeername(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    pvFromSockAddr uAddr, PeerAddress, PeerPort
    '--- success
    m_lLastError = 0
    GetPeerName = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetSockName(SocketAddress As String, SocketPort As Long) As Boolean
    Const FUNC_NAME     As String = "GetSockName"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If ws_getsockname(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    pvFromSockAddr uAddr, SocketAddress, SocketPort
    '--- success
    m_lLastError = 0
    GetSockName = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetLocalHost(HostName As String, HostAddress As String) As Boolean
    Const FUNC_NAME     As String = "GetLocalHost"
    Dim sBuffer         As String
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    sBuffer = String$(256, 0)
    If ws_gethostname(sBuffer, Len(sBuffer) - 1) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    HostName = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    pvToSockAddr HostName, 0, uAddr
    pvFromSockAddr uAddr, HostAddress, 0
    '--- success
    m_lLastError = 0
    GetLocalHost = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Accept(ConnectedSocket As cAsyncSocket, Optional SocketAddress As String, Optional SocketPort As Long) As Boolean
    Const FUNC_NAME     As String = "Accept"
    Dim hTemp           As Long
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If SocketPort = 0 Then
        hTemp = ws_accept(m_hSocket, ByVal 0, ByVal 0)
    Else
        pvToSockAddr SocketAddress, SocketPort, uAddr
        hTemp = ws_accept(m_hSocket, uAddr, LenB(uAddr))
    End If
    If hTemp = INVALID_SOCKET Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    If Not ConnectedSocket.Attach(hTemp) Then
        m_lLastError = ConnectedSocket.LastError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    Accept = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function AsyncSelect(Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Const FUNC_NAME     As String = "AsyncSelect"
    
    On Error GoTo EH
    If ws_WSAAsyncSelect(m_hSocket, m_oHelperWindow.hWnd, WM_SOCKET_NOTIFY + m_lIndex, EventMask) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    AsyncSelect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Bind(Optional SocketAddress As String, Optional ByVal SocketPort As Long) As Boolean
    Const FUNC_NAME     As String = "Bind"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If Not pvToSockAddr(SocketAddress, SocketPort, uAddr) Then
        GoTo QH
    End If
    If ws_bind(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    Bind = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub Close_()
    m_lLastError = 0
    '--- note: used in terminate -> no error handling
    If m_uGetHostByName.Handle <> 0 Then
        Call ws_WSACancelAsyncRequest(m_uGetHostByName.Handle)
        m_uGetHostByName.Handle = 0
        Erase m_uGetHostByName.Buffer
    End If
    If m_hSocket <> INVALID_SOCKET Then
        Call ws_WSAAsyncSelect(m_hSocket, m_oHelperWindow.hWnd, 0, 0)
        Call ws_shutdown(m_hSocket, 1)
        Call ws_closesocket(m_hSocket)
        m_oHelperWindow.frRemoveSocket Me, m_lIndex
        m_hSocket = INVALID_SOCKET
    End If
End Sub

Public Function Connect(HostAddress As String, ByVal HostPort As Long) As Boolean
    Const FUNC_NAME     As String = "Connect"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If m_hSocket = INVALID_SOCKET Then
        Create
    End If
    If Not pvToSockAddr(HostAddress, HostPort, uAddr, Resolve:=False) Then
        ReDim m_uGetHostByName.Buffer(0 To MAXGETHOSTSTRUCT) As Byte
        m_uGetHostByName.Port = HostPort
        '--- note: concat "" to prevent passing NULL for HostAddress
        m_uGetHostByName.Handle = ws_WSAAsyncGetHostByName(m_oHelperWindow.hWnd, WM_SOCKET_RESOLVE, HostAddress & "", m_uGetHostByName.Buffer(0), MAXGETHOSTSTRUCT)
        If m_uGetHostByName.Handle = 0 Then
            m_lLastError = Err.LastDllError
            Erase m_uGetHostByName.Buffer
            GoTo QH
        End If
        m_lLastError = WSAEWOULDBLOCK
    Else
        If SockOpt(ucsSsoType) = ucsSckStream Then
            If ws_connect(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
                m_lLastError = Err.LastDllError
                If m_lLastError <> WSAEWOULDBLOCK Then
                    GoTo QH
                End If
            Else
                PostEvent ucsSfdConnect
            End If
        Else
            PostEvent ucsSfdConnect
        End If
    End If
    '--- success
    m_lLastError = 0
    Connect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function IOCtl(ByVal Command As Long, Argument As Long) As Boolean
    Const FUNC_NAME     As String = "IOCtl"
    
    On Error GoTo EH
    If ws_ioctlsocket(m_hSocket, Command, Argument) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    IOCtl = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Listen(Optional ByVal ConnectionBacklog As Long = 5) As Boolean
    Const FUNC_NAME     As String = "Listen"
    
    On Error GoTo EH
    If ws_listen(m_hSocket, ConnectionBacklog) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    Listen = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ReceiveText( _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "ReceiveText"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If ReceiveArray(baBuffer, HostAddress, HostPort) Then
        ReceiveText = FromTextArray(baBuffer, CodePage)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ReceiveArray( _
            Buffer() As Byte, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "ReceiveArray"
    Dim uAddr           As SOCKADDR_IN
    Dim lBytes          As Long
    
    On Error GoTo EH
    lBytes = AvailableBytes
    If lBytes <= 0 Then
        Buffer = vbNullString
    Else
        ReDim Buffer(0 To lBytes - 1) As Byte
        If HostAddress <> STR_CHR1 Then
            lBytes = ws_recvfrom(m_hSocket, Buffer(0), lBytes, Flags, uAddr, LenB(uAddr))
        Else
            lBytes = ws_recv(m_hSocket, Buffer(0), lBytes, Flags)
        End If
        If lBytes = SOCKET_ERROR Then
            m_lLastError = Err.LastDllError
            GoTo QH
        End If
        If uAddr.sin_family <> 0 Then
            pvFromSockAddr uAddr, HostAddress, HostPort
        End If
    End If
    '--- success
    m_lLastError = 0
    ReceiveArray = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Receive( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Long
    Const FUNC_NAME     As String = "Receive"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If HostAddress <> STR_CHR1 Then
        Receive = ws_recvfrom(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
    Else
        Receive = ws_recv(m_hSocket, ByVal BufPtr, BufLen, Flags)
    End If
    If Receive = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    If uAddr.sin_family <> 0 Then
        pvFromSockAddr uAddr, HostAddress, HostPort
    End If
    '--- success
    m_lLastError = 0
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SendText( _
            Text As String, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Boolean
    Const FUNC_NAME     As String = "SendText"
    
    On Error GoTo EH
    SendText = SendArray(ToTextArray(Text, CodePage), HostAddress, HostPort)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SendArray( _
            Buffer() As Byte, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SendArray"
    Dim uAddr           As SOCKADDR_IN
    Dim lBytes          As Long
    Dim lResult         As Long
    
    On Error GoTo EH
    If LenB(HostAddress) <> 0 Then
        If Not pvToSockAddr(HostAddress, HostPort, uAddr) Then
            GoTo QH
        End If
    End If
    Do
        If uAddr.sin_family <> 0 Then
            lResult = ws_sendto(m_hSocket, Buffer(lBytes), UBound(Buffer) + 1 - lBytes, Flags, uAddr, LenB(uAddr))
        Else
            lResult = ws_send(m_hSocket, Buffer(lBytes), UBound(Buffer) + 1 - lBytes, Flags)
        End If
        If lResult < 0 Then
            m_lLastError = Err.LastDllError
            GoTo QH
        End If
        lBytes = lBytes + lResult
    Loop While lBytes <= UBound(Buffer)
    '--- success
    m_lLastError = 0
    SendArray = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Send( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Long
    Const FUNC_NAME     As String = "Send"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If LenB(HostAddress) <> 0 Then
        If Not pvToSockAddr(HostAddress, HostPort, uAddr) Then
            Send = SOCKET_ERROR
            GoTo QH
        End If
    End If
    If uAddr.sin_family <> 0 Then
        Send = ws_sendto(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
    Else
        Send = ws_send(m_hSocket, ByVal BufPtr, BufLen, Flags)
    End If
    If Send = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ShutDown(Optional ByVal How As Long = 1) As Boolean
    Const FUNC_NAME     As String = "ShutDown"
    
    On Error GoTo EH
    If ws_shutdown(m_hSocket, How) = SOCKET_ERROR Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    ShutDown = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetErrorDescription(ByVal ErrorNumber As Long) As String
    Const FUNC_NAME     As String = "GetErrorDescription"
    Dim lSize           As Long
    
    On Error GoTo EH
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, ErrorNumber, 0, GetErrorDescription, Len(GetErrorDescription) + 1, 0)
    If lSize > 2 Then
        If Mid$(GetErrorDescription, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetErrorDescription = Left$(GetErrorDescription, lSize)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub PostEvent(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "PostEvent"
    
    On Error GoTo EH
    Call PostMessage(m_oHelperWindow.hWnd, WM_SOCKET_NOTIFY + m_lIndex, m_hSocket, EventMask)
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= static helpers ========================================================

Public Function FromTextArray(baText() As Byte, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "FromTextArray"
    Dim lSize           As Long
    
    On Error GoTo EH
    If CodePage = ucsScpUnicode Then
        FromTextArray = baText
    ElseIf UBound(baText) >= 0 Then
        FromTextArray = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CodePage, 0, baText(0), UBound(baText) + 1, StrPtr(FromTextArray), Len(FromTextArray) + 1)
        If lSize <> Len(FromTextArray) Then
            FromTextArray = Left$(FromTextArray, lSize)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    Const FUNC_NAME     As String = "ToTextArray"
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    On Error GoTo EH
    If CodePage = ucsScpUnicode Then
        baRetVal = sText
    ElseIf LenB(sText) <> 0 Then
        ReDim baRetVal(0 To 2 * Len(sText)) As Byte
        lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
        If lSize > 0 Then
            ReDim Preserve baRetVal(0 To lSize - 1) As Byte
        Else
            baRetVal = vbNullString
        End If
    Else
        baRetVal = vbNullString
    End If
    ToTextArray = baRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= friend ================================================================

Friend Function frNotifyEvent(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    If m_hSocket = wParam Then
        #If ImplSync Then
            SyncProcessMsgQueue
        #End If
        pvDoNotify wParam, lParam
        '--- handled
        frNotifyEvent = True
    End If
End Function

Friend Function frNotifyGetHostByName(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    If m_uGetHostByName.Handle = wParam Then
        #If ImplSync Then
            SyncProcessMsgQueue
        #End If
        pvDoResolve wParam, lParam
        '--- handled
        frNotifyGetHostByName = True
    End If
End Function

Private Sub pvDoNotify(ByVal wParam As Long, ByVal lParam As Long)
    Dim eEvent          As UcsAsyncSocketEventMaskEnum
    Dim lErrorCode      As Long
    Dim lBytes          As Long

    If m_hSocket = wParam Then
        eEvent = lParam And &HFFFF&
        lErrorCode = lParam \ &H10000
        If lErrorCode <> 0 Then
            m_lLastError = lErrorCode
            RaiseEvent OnError(lErrorCode, eEvent)
        Else
            Select Case eEvent
            Case ucsSfdRead
                If Not IOCtl(FIONREAD, lBytes) Then
                    RaiseEvent OnError(m_lLastError, eEvent)
                    If m_hSocket = INVALID_SOCKET Then
                        GoTo QH
                    End If
                End If
                If lBytes <> 0 Then
                    m_lLastError = 0
                    RaiseEvent OnReceive
                End If
            Case [_ucsSfdForceRead]
                m_lLastError = 0
                RaiseEvent OnReceive
            Case ucsSfdWrite
                m_lLastError = 0
                RaiseEvent OnSend
            Case ucsSfdConnect
                m_lLastError = 0
                RaiseEvent OnConnect
            Case ucsSfdAccept
                m_lLastError = 0
                RaiseEvent OnAccept
            Case ucsSfdClose
                m_lLastError = 0
                RaiseEvent OnClose
            End Select
        End If
    End If
QH:
End Sub

Private Sub pvDoResolve(ByVal wParam As Long, ByVal lParam As Long)
    Dim lErrorCode      As Long
    Dim uAddr           As SOCKADDR_IN
    Dim lPtr            As Long
    
    If m_uGetHostByName.Handle = wParam Then
        lErrorCode = lParam \ &H10000
        If lErrorCode <> 0 Then
            m_uGetHostByName.Handle = 0
            Erase m_uGetHostByName.Buffer
            m_lLastError = lErrorCode
            RaiseEvent OnError(lErrorCode, [_ucsSfdResolve])
        Else
            '--- set sin_addr to first IP address resolved
            Call CopyMemory(lPtr, m_uGetHostByName.Buffer(12), 4)
            Call CopyMemory(lPtr, ByVal lPtr, 4)
            Call CopyMemory(uAddr.sin_addr, ByVal lPtr, 4)
            m_uGetHostByName.Handle = 0
            Erase m_uGetHostByName.Buffer
            m_lLastError = 0
            RaiseEvent OnResolve(pvToString(ws_inet_ntoa(uAddr.sin_addr)))
            If m_hSocket = INVALID_SOCKET Then
                m_lLastError = WSAEINTR
                GoTo QH
            End If
            If SockOpt(ucsSsoType) = ucsSckStream Then
                uAddr.sin_family = AF_INET
                uAddr.sin_port = ws_htons(m_uGetHostByName.Port)
                If ws_connect(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
                    m_lLastError = Err.LastDllError
                    If m_lLastError <> WSAEWOULDBLOCK Then
                        RaiseEvent OnError(m_lLastError, ucsSfdConnect)
                    End If
                Else
                    PostEvent ucsSfdConnect
                End If
            Else
                PostEvent ucsSfdConnect
            End If
        End If
    End If
QH:
End Sub

Private Function pvToSockAddr( _
            Address As String, _
            ByVal Port As Long, _
            uAddr As SOCKADDR_IN, _
            Optional ByVal Resolve As Boolean = True) As Boolean
    Dim lPtr            As Long
    
    uAddr.sin_family = AF_INET
    If LenB(Address) = 0 Then
        uAddr.sin_addr = ws_htonl(INADDR_ANY)
    Else
        uAddr.sin_addr = ws_inet_addr(Address)
        If uAddr.sin_addr = INADDR_NONE Then
            If Resolve Then
                lPtr = ws_gethostbyname(Address)
            End If
            If lPtr = 0 Then
                m_lLastError = WSAEINVAL
                GoTo QH
            End If
            Call CopyMemory(lPtr, ByVal UnsignedAdd(lPtr, 12), 4)
            Call CopyMemory(lPtr, ByVal lPtr, 4)
            Call CopyMemory(uAddr.sin_addr, ByVal lPtr, 4)
        End If
    End If
    uAddr.sin_port = ws_htons(Port)
    '--- success
    pvToSockAddr = True
QH:
End Function

Private Sub pvFromSockAddr(uAddr As SOCKADDR_IN, Address As String, Port As Long)
    Port = pvToLong(ws_ntohs(uAddr.sin_port))
    Address = pvToString(ws_inet_ntoa(uAddr.sin_addr))
End Sub

Private Function pvToLong(ByVal nLoWord As Integer) As Long
    Call CopyMemory(pvToLong, nLoWord, 2)
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call lstrcpy(pvToString, lPtr)
    End If
End Function

Private Function pvGetHelperWindow() As Long
    pvGetHelperWindow = FindWindow(STR_HELPER_CLASS, App.hInstance & ":" & App.ThreadID & ":" & STR_HELPER_CAPTION)
End Function

'= sync ==================================================================

#If ImplSync Then
Public Function SyncConnect(HostAddress As String, ByVal HostPort As Long, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "SyncConnect"
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = DEF_TIMEOUT
    End If
    If Not Connect(HostAddress, HostPort) Then
        GoTo QH
    End If
    If Not SyncWaitForEvent(Timeout, ucsSfdConnect) Then
        GoTo QH
    End If
    '--- success
    m_lLastError = 0
    SyncConnect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceiveText( _
            Optional ByVal NeedLen As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "SyncReceiveText"
    Dim baText()        As Byte
    Dim dblTimer        As Double
    Dim lElapsed        As Long
    Dim baBuffer()      As Byte
    Dim bResult         As Boolean
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = Znl(SockOpt(ucsSsoReceiveTimeout), DEF_TIMEOUT)
    End If
    baText = vbNullString
    dblTimer = TimerEx
    Do
        If Timeout > 0 Then
            lElapsed = Int((TimerEx - dblTimer) * 1000)
            If lElapsed >= Timeout Then
                m_lLastError = ERR_TIMEOUT
                GoTo QH
            End If
        End If
        bResult = SyncReceiveArray(baBuffer, NeedLen - Len(SyncReceiveText), Timeout - lElapsed, HostAddress, HostPort)
        ReDim Preserve baText(0 To UBound(baText) + UBound(baBuffer) + 1) As Byte
        Call CopyMemory(baText(UBound(baText) - UBound(baBuffer)), baBuffer(0), UBound(baBuffer) + 1)
        SyncReceiveText = FromTextArray(baText, CodePage)
        If Not bResult Then
            GoTo QH
        End If
    Loop While Len(SyncReceiveText) < NeedLen And NeedLen > 0
    '--- success
    m_lLastError = 0
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function


Public Function SyncReceiveArray( _
            Buffer() As Byte, _
            Optional ByVal NeedLen As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncReceiveArray"
    Dim lAvailable      As Long
    Dim dblTimer        As Double
    Dim lResult         As Long
    Dim uAddr           As SOCKADDR_IN
    Dim lBytes          As Long
    Dim lElapsed        As Long
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = Znl(SockOpt(ucsSsoReceiveTimeout), DEF_TIMEOUT)
    End If
    lAvailable = NeedLen
    If lAvailable <= 0 Then
        lAvailable = AvailableBytes
    End If
    If lAvailable <= 0 Then
        lAvailable = SockOpt(ucsSsoReceiveBuffer)
    End If
    If lAvailable <= 0 Then
        lAvailable = 4096
    End If
    ReDim Buffer(0 To lAvailable - 1) As Byte
    dblTimer = TimerEx
    Do
        If HostAddress <> STR_CHR1 Then
            lResult = ws_recvfrom(m_hSocket, Buffer(lBytes), lAvailable - lBytes, Flags, uAddr, LenB(uAddr))
        Else
            lResult = ws_recv(m_hSocket, Buffer(lBytes), lAvailable - lBytes, Flags)
        End If
        If lResult = SOCKET_ERROR Then
            m_lLastError = Err.LastDllError
            If m_lLastError <> WSAEWOULDBLOCK Then
                GoTo QH
            End If
        ElseIf lResult > 0 Then
            If uAddr.sin_family <> 0 Then
                pvFromSockAddr uAddr, HostAddress, HostPort
            End If
            lBytes = lBytes + lResult
            If lBytes >= NeedLen Or NeedLen = 0 Then
                Exit Do
            End If
        End If
        If Timeout > 0 Then
            lElapsed = Int((TimerEx - dblTimer) * 1000)
            If lElapsed >= Timeout Then
                m_lLastError = ERR_TIMEOUT
                GoTo QH
            End If
        End If
        If Not SyncWaitForEvent(Timeout - lElapsed, ucsSfdRead) Then
            GoTo QH
        End If
    Loop
    '--- success
    m_lLastError = 0
    SyncReceiveArray = True
QH:
    If UBound(Buffer) + 1 <> lBytes Then
        If lBytes > 0 Then
            ReDim Preserve Buffer(0 To lBytes - 1) As Byte
        Else
            Buffer = vbNullString
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceive( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Received As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncReceive"
    Dim lResult         As Long
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = Znl(SockOpt(ucsSsoReceiveTimeout), DEF_TIMEOUT)
    End If
    Do
        If HostAddress <> STR_CHR1 Then
            lResult = ws_recvfrom(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
        Else
            lResult = ws_recv(m_hSocket, ByVal BufPtr, BufLen, Flags)
        End If
        If lResult = SOCKET_ERROR Then
            m_lLastError = Err.LastDllError
            If m_lLastError <> WSAEWOULDBLOCK Then
                GoTo QH
            End If
            If Not SyncWaitForEvent(Timeout, ucsSfdRead) Then
                GoTo QH
            End If
        ElseIf lResult > 0 Then
            If uAddr.sin_family <> 0 Then
                pvFromSockAddr uAddr, HostAddress, HostPort
            End If
            Received = lResult
            Exit Do
        End If
    Loop
    '--- success
    m_lLastError = 0
    SyncReceive = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSendText( _
            Text As String, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Boolean
    Const FUNC_NAME     As String = "SendText"
    
    On Error GoTo EH
    SyncSendText = SyncSendArray(ToTextArray(Text, CodePage), Timeout, HostAddress, HostPort)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSendArray( _
            Buffer() As Byte, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncSendArray"
    
    On Error GoTo EH
    If UBound(Buffer) < 0 Then
        SyncSendArray = True
    Else
        SyncSendArray = SyncSend(VarPtr(Buffer(0)), UBound(Buffer) + 1, Timeout, HostAddress, HostPort, Flags)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSend( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncSend"
    Dim uAddr           As SOCKADDR_IN
    Dim lResult         As Long
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = Znl(SockOpt(ucsSsoSendTimeout), DEF_TIMEOUT)
    End If
    If LenB(HostAddress) <> 0 Then
        If Not pvToSockAddr(HostAddress, HostPort, uAddr) Then
            SyncSend = SOCKET_ERROR
            GoTo QH
        End If
    End If
    Do
        Do
            If uAddr.sin_family <> 0 Then
                lResult = ws_sendto(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
            Else
                lResult = ws_send(m_hSocket, ByVal BufPtr, BufLen, Flags)
            End If
            If lResult >= 0 Then
                Exit Do
            End If
            m_lLastError = Err.LastDllError
            If m_lLastError <> WSAEWOULDBLOCK Then
                GoTo QH
            End If
            If Not SyncWaitForEvent(Timeout, ucsSfdWrite) Then
                GoTo QH
            End If
        Loop
        If lResult = BufLen Then
            Exit Do
        End If
        BufPtr = UnsignedAdd(BufPtr, lResult)
        BufLen = BufLen - lResult
    Loop
    '--- success
    m_lLastError = 0
    SyncSend = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncWaitForEvent( _
            ByVal Timeout As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Const FUNC_NAME     As String = "SyncWaitForEvent"
    Dim hWnd            As Long
    Dim uMsg            As APIMSG
    Dim lTimerID        As Long
    Dim bBlocking       As Boolean
    Dim bHandled        As Boolean
    Dim bPeek           As Boolean
    
    On Error GoTo EH
    '--- check if no pending resolution
    If (EventMask And [_ucsSfdResolve]) <> 0 Then
        If m_uGetHostByName.Handle = 0 Then
            '--- success
            m_lLastError = 0
            SyncWaitForEvent = True
            GoTo QH
        End If
    End If
    If (EventMask And ucsSfdRead) <> 0 Then
        EventMask = EventMask Or [_ucsSfdForceRead]
    End If
    bBlocking = m_bBlocking
    m_bBlocking = True
    hWnd = m_oHelperWindow.hWnd
    If Timeout > 0 Then
        lTimerID = SetTimer(hWnd, 1, Timeout, 0)
    End If
    bPeek = True
    Do
        If PeekMessage(uMsg, hWnd, WM_SOCKET_NOTIFY + m_lIndex, WM_SOCKET_NOTIFY + m_lIndex, PM_REMOVE) <> 0 Then
            If uMsg.wParam <> INVALID_SOCKET And uMsg.lParam <> 0 Then
                If uMsg.wParam = m_hSocket Then
                    m_cMsgQueue.Add Array(WM_SOCKET_NOTIFY, uMsg.wParam, uMsg.lParam)
                    If (uMsg.lParam And EventMask) <> 0 Then
                        '--- success
                        m_lLastError = 0
                        SyncWaitForEvent = True
                        Exit Do
                    End If
                    If (uMsg.lParam And ucsSfdClose) <> 0 Then
                        m_lLastError = WSAENOTCONN
                        Exit Do
                    End If
                End If
                bPeek = True
            End If
        End If
        If PeekMessage(uMsg, hWnd, WM_SOCKET_RESOLVE, WM_SOCKET_RESOLVE, PM_REMOVE) <> 0 Then
            If uMsg.wParam = m_uGetHostByName.Handle Then
                pvDoResolve uMsg.wParam, uMsg.lParam
                If (EventMask And [_ucsSfdResolve]) <> 0 Then
                    '--- success
                    m_lLastError = 0
                    SyncWaitForEvent = True
                    Exit Do
                End If
            End If
            bPeek = True
        End If
        If PeekMessage(uMsg, hWnd, WM_TIMER, WM_TIMER, PM_REMOVE) <> 0 Then
            m_lLastError = ERR_TIMEOUT
            Exit Do
        End If
        If Not m_bBlocking Then
            m_lLastError = WSAEINTR
            Exit Do
        End If
        If bPeek Then
            If PeekMessage(uMsg, 0, 0, 0, PM_NOREMOVE) <> 0 Then
                bHandled = False
                RaiseEvent OnMessagePending(bHandled)
                If m_hSocket = INVALID_SOCKET Then
                    m_lLastError = WSAEINTR
                    GoTo QH
                End If
                If Not bHandled Then
                    If PeekMessage(uMsg, 0, WM_PAINT, WM_PAINT, PM_REMOVE) <> 0 Then
                        Call DispatchMessage(uMsg)
                        If m_hSocket = INVALID_SOCKET Then
                            m_lLastError = WSAEINTR
                            GoTo QH
                        End If
                    End If
                End If
                bPeek = False
            Else
                Call WaitMessage
                bPeek = True
            End If
        Else
            Call WaitMessage
            bPeek = True
        End If
    Loop
    If lTimerID <> 0 Then
        Call KillTimer(hWnd, lTimerID)
    End If
    If Not m_bBlocking Then
        m_lLastError = WSAEINTR
        GoTo QH
    End If
QH:
    m_bBlocking = bBlocking
    If m_cMsgQueue.Count > 0 Then
        Call PostMessage(hWnd, WM_SOCKET_NOTIFY + m_lIndex, INVALID_SOCKET, 0)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub SyncCancelWait()
    m_bBlocking = False
    m_lLastError = 0
End Sub

Public Sub SyncProcessMsgQueue()
    Const FUNC_NAME     As String = "SyncProcessMsgQueue"
    Dim vElem           As Variant
    
    On Error GoTo EH
    Do While m_cMsgQueue.Count > 0
        vElem = m_cMsgQueue.Item(1)
        m_cMsgQueue.Remove 1
        Select Case vElem(0)
        Case WM_SOCKET_NOTIFY
            pvDoNotify vElem(1), vElem(2)
        Case WM_SOCKET_RESOLVE
            pvDoResolve vElem(1), vElem(2)
        End Select
    Loop
    m_lLastError = 0
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
#End If

'= shared ================================================================

#If Not ImplUseShared Then
Private Function Znl(ByVal lValue As Long, Optional IfEmptyLong As Variant = Null, Optional ByVal EmptyLong As Long = 0) As Variant
    Znl = IIf(lValue = EmptyLong, IfEmptyLong, lValue)
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim lPtr            As Long
    Dim pUnk            As Object
    
    lPtr = GetWindowLong(pvGetHelperWindow(), GWL_USERDATA)
    If lPtr <> 0 Then
        Call CopyMemory(pUnk, lPtr, 4)
        Set m_oHelperWindow = pUnk
        Call CopyMemory(pUnk, 0&, 4)
    Else
        Set m_oHelperWindow = New cAsyncSocketHelperWindow
    End If
    m_hSocket = INVALID_SOCKET
    m_lIndex = -1
    Set m_cMsgQueue = New Collection
End Sub

Private Sub Class_Terminate()
    Close_
    Set m_oHelperWindow = Nothing
End Sub
