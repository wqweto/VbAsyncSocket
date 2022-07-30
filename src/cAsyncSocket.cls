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
Private Const MODULE_NAME As String = "cAsyncSocket"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplSync = Not (ASYNCSOCKET_NO_SYNC <> 0)
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)

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
Event BeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Cancel As Boolean)
Event AfterNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsAsyncSocketTypeEnum
    ucsSckStream = 1
    ucsSckDatagram = 2
    ucsSckRaw = 3
End Enum

Public Enum UcsAsyncSocketProtocolEnum
    ucsScpICMP = 1
    ucsScpIGMP = 2
    ucsScpTCP = 6
    ucsScpUDP = 17
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
    ucsSolIP = 0
    ucsSolICMP = 1
    ucsSolIGMP = 2
    ucsSolTCP = 6
    ucsSolUDP = 17
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
    '-- IP
    ucsSsoIPOptions = 1                     ' IP options.
    ucsSsoHeaderIncluded = 2                ' Header is included with data.
    ucsSsoTypeOfService = 3                 ' IP type of service and preced.
    ucsSsoIpTimeToLive = 4                  ' IP time to live.
    ucsSsoMulticastInterface = 9            ' IP multicast interface.
    ucsSsoMulticastTimeToLive = 10          ' IP multicast time to live.
    ucsSsoMulticastLoopback = 11            ' IP Multicast loopback.
    ucsSsoAddMembership = 12                ' Add an IP group membership.
    ucsSsoDropMembership = 13               ' Drop an IP group membership.
    ucsSsoDontFragment = 14                 ' Don't fragment IP datagrams.
    ucsSsoAddSourceMembership = 15          ' Join IP group/source.
    ucsSsoDropSourceMembership = 16         ' Leave IP group/source.
    ucsSsoBlockSource = 17                  ' Block IP group/source.
    ucsSsoUnblockSource = 18                ' Unblock IP group/source.
    ucsSsoPacketInformation = 19            ' Receive packet information for ipv4.
    '-- TCP
    ucsSsoNoDelay = 1                       ' Disables the Nagle algorithm for send coalescing.
    ucsSsoExpedited = 2
    '--- UDP
    ucsSsoNoChecksum = 1
    ucsSsoChecksumCoverage = 20             ' Udp-Lite checksum coverage.
    ucsSsoUpdateAcceptContext = &H700B
    ucsSsoUpdateConnectContext = &H7010
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
Private Const FIONBIO                       As Long = &H8004667E
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
Private Const HWND_MESSAGE                  As Long = -3
'--- for PeekMessage
Private Const PM_NOREMOVE                   As Long = 0
Private Const PM_REMOVE                     As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_FROM_HMODULE   As Long = &H800
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Args As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal lX As Long, ByVal lY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (AdapterInfo As Any, OutBufLen As Long) As Long
Private Declare Function GetIpForwardTable Lib "iphlpapi" (pIpForwardTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function WSAAsyncGetHostByName Lib "ws2_32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal BufLen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSACancelAsyncRequest Lib "ws2_32" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
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
Private Declare Function ws_accept Lib "ws2_32" Alias "accept" (ByVal s As Long, Addr As Any, addrlen As Any) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_inet_addr Lib "ws2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function ws_gethostbyname Lib "ws2_32" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function ws_gethostname Lib "ws2_32" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
#If ImplSync Then
    Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As APIMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As APIMSG) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'    Private Declare Function WaitMessage Lib "user32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
#If Not ImplUseShared Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    '--- for Modern Subclassing Thunk (MST)
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
    Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    '--- end MST
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

Private Type IP_ADDR_STRING
    Next                As Long
    IpAddress(0 To 15)  As Byte
    IpMask(0 To 15)     As Byte
    Context             As Long
End Type

Private Type IP_ADAPTER_INFO
    Next                As Long
    ComboIndex          As Long
    AdapterName         As String * 260
    Description         As String * 132
    AddressLength       As Long
    Address(0 To 7)     As Byte
    Index               As Long
    Type                As Long
    DhcpEnabled         As Long
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Byte
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As Long
    LeaseExpires        As Long
End Type

Private Type MIB_IPFORWARDROW
    dwForwardDest       As Long
    dwForwardMask       As Long
    dwForwardPolicy     As Long
    dwForwardNextHop    As Long
    dwForwardIfIndex    As Long
    dwForwardType       As Long
    dwForwardProto      As Long
    dwForwardAge        As Long
    dwForwardNextHopAS  As Long
    dwForwardMetric1    As Long
    dwForwardMetric2    As Long
    dwForwardMetric3    As Long
    dwForwardMetric4    As Long
    dwForwardMetric5    As Long
End Type
Private Const sizeof_MIB_IPFORWARDROW As Long = 56

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_HELPER_CLASS      As String = "STATIC"
Private Const STR_HELPER_NAME       As String = "AsyncSocket Helper Window"
Private Const STR_CHR1              As String = "" '--- CHAR(1)
Private Const DEF_TIMEOUT           As Long = 5000
'--- helper window
Private Const MAX_SOCKETS           As Long = &HC000& - WM_SOCKET_NOTIFY

Private m_hSocket               As Long
Private m_pCleanup              As IUnknown
Private m_oHelperWindow         As cAsyncSocket
Private m_lIndex                As Long
Private m_lLastError            As Long
Private m_lLastSendBytes        As Long
Private m_uGetHostByName()      As UcsAsyncGetHostByNameType
Private m_uWindowState()        As UcsHelperWindowStateType
#If ImplSync Then
    Private m_cMsgQueue         As VBA.Collection
    Private m_bBlocking         As Boolean
#End If

Private Type UcsAsyncGetHostByNameType
    Handle              As Long
    Port                As Long
    Buffer(0 To MAXGETHOSTSTRUCT) As Byte
End Type

Private Type UcsHelperWindowStateType
    hWnd                As Long
    Notify              As IUnknown
    Cleanup             As IUnknown
    SocketPtr()         As Long
    Pos                 As Long
    Count               As Long
End Type

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

Public Property Get SocketHandle() As Long
    SocketHandle = m_hSocket
End Property

Public Property Get LastError() As VBA.ErrObject
    Set LastError = Err
    With LastError
        .Description = GetErrorDescription(m_lLastError)
        .Number = m_lLastError
        .Source = "cAsyncSocket"
    End With
End Property

Public Property Get SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket) As Long
    Call ws_getsockopt(m_hSocket, Level, OptionName, SockOpt, 4)
    m_lLastError = Err.LastDllError
End Property

Public Property Let SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket, ByVal Value As Long)
    Call ws_setsockopt(m_hSocket, Level, OptionName, Value, 4)
    m_lLastError = Err.LastDllError
End Property

Public Property Get AvailableBytes() As Long
     If Not IOCtl(FIONREAD, AvailableBytes) Then
        AvailableBytes = SOCKET_ERROR
    End If
End Property

Public Property Get BlockingMode() As Long
    If Not IOCtl(FIONBIO, BlockingMode) Then
        BlockingMode = SOCKET_ERROR
    End If
End Property

Public Property Get LastSendBytes() As Long
    LastSendBytes = m_lLastSendBytes
End Property

Public Property Get HasPendingEvent() As Boolean
    HasPendingEvent = (m_lLastError = WSAEWOULDBLOCK)
End Property

Public Property Get HasPendingResolve() As Boolean
    HasPendingResolve = pvResolveHandle <> 0
End Property

Public Property Get IsClosed() As Boolean
    IsClosed = (m_hSocket = INVALID_SOCKET)
End Property

'= helper window =========================================================

Friend Property Get frMessageHWnd() As Long
    If Peek(ArrPtr(m_uWindowState)) <> 0 Then
        frMessageHWnd = m_uWindowState(0).hWnd
    End If
End Property

'= private ===============================================================

Private Property Get pvResolveHandle() As Long
    If Peek(ArrPtr(m_uGetHostByName)) <> 0 Then
        pvResolveHandle = m_uGetHostByName(0).Handle
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Create( _
            Optional ByVal SocketPort As Long, _
            Optional ByVal SocketType As UcsAsyncSocketTypeEnum = ucsSckStream, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll, _
            Optional SocketAddress As String, _
            Optional ByVal SocketProtocol As UcsAsyncSocketProtocolEnum) As Boolean
    Const FUNC_NAME     As String = "Create"
    
    On Error GoTo EH
    If m_hSocket <> INVALID_SOCKET Then
        m_lLastError = WSAEALREADY
        GoTo QH
    End If
    If m_oHelperWindow.frMessageHWnd = 0 Then
        m_lLastError = WSANOTINITIALISED
        GoTo QH
    End If
    m_hSocket = ws_socket(AF_INET, SocketType, SocketProtocol)
    If m_hSocket = INVALID_SOCKET Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    Set m_pCleanup = InitCleanupThunk(m_hSocket, "ws2_32", "closesocket")
    If Not m_oHelperWindow.frAddSocket(Me, m_lIndex) Then
        m_lLastError = WSAEMFILE
        pvClose
        GoTo QH
    End If
    If Not AsyncSelect(EventMask) Then
        pvClose
        GoTo QH
    End If
    If SocketPort <> 0 Or LenB(SocketAddress) <> 0 Then
        If Not Bind(SocketAddress, SocketPort) Then
            pvClose
            GoTo QH
        End If
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
    Set m_pCleanup = InitCleanupThunk(m_hSocket, "ws2_32", "closesocket")
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
    '--- note: prevent closesocket(m_hSocket) being called
    ThunkPrivateData(m_pCleanup) = INVALID_SOCKET
    m_oHelperWindow.frRemoveSocket Me, m_lIndex
    m_hSocket = INVALID_SOCKET
    Set m_pCleanup = Nothing
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

Public Function GetLocalHost(HostName As String, HostAddress As String, Optional AdaptersInfo As Variant, Optional ForwardTable As Variant) As Boolean
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
    If Not IsMissing(AdaptersInfo) Then
        pvGetAdaptersInfo AdaptersInfo
    End If
    If Not IsMissing(ForwardTable) Then
        pvGetForwardTable ForwardTable
    End If
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
    If ConnectedSocket Is Nothing Then
        Set ConnectedSocket = New cAsyncSocket
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
    If WSAAsyncSelect(m_hSocket, m_oHelperWindow.frMessageHWnd, WM_SOCKET_NOTIFY + m_lIndex, EventMask) = SOCKET_ERROR Then
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
    Const FUNC_NAME     As String = "Close_"
    
    On Error GoTo EH
    If pvResolveHandle <> 0 Then
        Call WSACancelAsyncRequest(m_uGetHostByName(0).Handle)
        Erase m_uGetHostByName
    End If
    If m_hSocket <> INVALID_SOCKET Then
        Call ws_shutdown(m_hSocket, 1)
        pvClose
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Function Connect(HostAddress As String, ByVal HostPort As Long) As Boolean
    Const FUNC_NAME     As String = "Connect"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If m_hSocket = INVALID_SOCKET Then
        Create
    End If
    If Not pvToSockAddr(HostAddress, HostPort, uAddr, Resolve:=False) Then
        ReDim m_uGetHostByName(0 To 0) As UcsAsyncGetHostByNameType
        m_uGetHostByName(0).Port = HostPort
        '--- note: concat "" to prevent passing NULL for HostAddress
        m_uGetHostByName(0).Handle = WSAAsyncGetHostByName(m_oHelperWindow.frMessageHWnd, WM_SOCKET_RESOLVE, HostAddress & "", m_uGetHostByName(0).Buffer(0), MAXGETHOSTSTRUCT)
        If m_uGetHostByName(0).Handle = 0 Then
            m_lLastError = Err.LastDllError
            Erase m_uGetHostByName
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
    Dim lPos            As Long
    Dim lSize           As Long
    
    On Error GoTo EH
    lBytes = AvailableBytes
    If lBytes <= 0 Then
        Buffer = vbNullString
    Else
        ReDim Buffer(0 To lBytes - 1) As Byte
        Do
            lPos = lPos + lSize
            If HostAddress <> STR_CHR1 Then
                lSize = ws_recvfrom(m_hSocket, Buffer(lPos), lBytes - lPos, Flags, uAddr, LenB(uAddr))
            Else
                lSize = ws_recv(m_hSocket, Buffer(lPos), lBytes - lPos, Flags)
            End If
            If lSize = SOCKET_ERROR Then
                m_lLastError = Err.LastDllError
                GoTo QH
            ElseIf lSize = 0 Or SockOpt(ucsSsoType) <> ucsSckStream Then
                If lPos + lSize = 0 Then
                    Buffer = vbNullString
                Else
                    ReDim Preserve Buffer(0 To lPos + lSize - 1) As Byte
                End If
                Exit Do
            End If
        Loop While lPos + lSize < lBytes
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
    Dim lPtr            As Long
    
    On Error GoTo EH
    If LenB(HostAddress) <> 0 Then
        If Not pvToSockAddr(HostAddress, HostPort, uAddr) Then
            GoTo QH
        End If
    End If
    Do
        If lBytes <= UBound(Buffer) Then
            lPtr = VarPtr(Buffer(lBytes))
        End If
        If uAddr.sin_family <> 0 Then
            lResult = ws_sendto(m_hSocket, ByVal lPtr, UBound(Buffer) + 1 - lBytes, Flags, uAddr, LenB(uAddr))
        Else
            lResult = ws_send(m_hSocket, ByVal lPtr, UBound(Buffer) + 1 - lBytes, Flags)
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
    m_lLastSendBytes = BufLen
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

Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
    Const FUNC_NAME     As String = "GetErrorDescription"
    Dim lSize           As Long
    
    On Error GoTo EH
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrorCode, 0, GetErrorDescription, Len(GetErrorDescription), 0)
    If lSize = 0 Then
        lSize = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, LoadLibrary("mswsock"), ErrorCode And &HFFFF&, 0, GetErrorDescription, Len(GetErrorDescription), 0)
    End If
    If lSize = 0 Then
        lSize = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, LoadLibrary("winhttp"), ErrorCode And &HFFFF&, 0, GetErrorDescription, Len(GetErrorDescription), 0)
    End If
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

Public Sub PostEvent(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Optional ByVal Immediate As Boolean)
    Const FUNC_NAME     As String = "PostEvent"
    
    On Error GoTo EH
    If Immediate Then
        pvDoNotify m_hSocket, EventMask
    Else
        Call PostMessage(m_oHelperWindow.frMessageHWnd, WM_SOCKET_NOTIFY + m_lIndex, m_hSocket, EventMask)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Sub SetLastError(ByVal lValue As Long)
    m_lLastError = lValue
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
            lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(CodePage, 0, StrPtr(sText), Len(sText), baRetVal(0), UBound(baRetVal) + 1, 0, 0)
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
    If pvResolveHandle = wParam Then
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
    Dim bCancel         As Boolean
    Dim lBytes          As Long

    If m_hSocket <> wParam Then
        GoTo QH
    End If
    eEvent = lParam And &HFFFF&
    m_lLastError = lParam \ &H10000
    If m_lLastError <> 0 Then
        RaiseEvent OnError(m_lLastError, eEvent)
        GoTo QH
    End If
    RaiseEvent BeforeNotify(eEvent, bCancel)
    If bCancel Then
        GoTo QH
    End If
    Select Case eEvent
    Case ucsSfdRead
        If Not IOCtl(FIONREAD, lBytes) Then
            RaiseEvent OnError(m_lLastError, eEvent)
            If m_hSocket = INVALID_SOCKET Then
                GoTo QH
            End If
        End If
        If lBytes <> 0 Then
            RaiseEvent OnReceive
        End If
    Case [_ucsSfdForceRead]
        RaiseEvent OnReceive
    Case ucsSfdWrite
        RaiseEvent OnSend
    Case ucsSfdConnect
        RaiseEvent OnConnect
    Case ucsSfdAccept
        RaiseEvent OnAccept
    Case ucsSfdClose
        RaiseEvent OnClose
    End Select
    RaiseEvent AfterNotify(eEvent)
QH:
End Sub

Private Sub pvDoResolve(ByVal wParam As Long, ByVal lParam As Long)
    Dim uAddr           As SOCKADDR_IN
    Dim lPtr            As Long
    
    If pvResolveHandle <> wParam Then
        GoTo QH
    End If
    m_lLastError = lParam \ &H10000
    If m_lLastError <> 0 Then
        Erase m_uGetHostByName
        RaiseEvent OnError(m_lLastError, [_ucsSfdResolve])
        GoTo QH
    End If
    uAddr.sin_family = AF_INET
    uAddr.sin_port = ws_htons(m_uGetHostByName(0).Port)
    '--- set sin_addr to first IP address resolved
    Call CopyMemory(lPtr, m_uGetHostByName(0).Buffer(12), 4)
    Call CopyMemory(lPtr, ByVal lPtr, 4)
    Call CopyMemory(uAddr.sin_addr, ByVal lPtr, 4)
    Erase m_uGetHostByName
    m_lLastError = 0
    RaiseEvent OnResolve(pvToString(ws_inet_ntoa(uAddr.sin_addr)))
    If m_hSocket = INVALID_SOCKET Then
        m_lLastError = WSAEINTR
        GoTo QH
    End If
    If SockOpt(ucsSsoType) = ucsSckStream Then
        '--- note uAddr is prepared above, before m_uGetHostByName being erased
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
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

Private Sub pvClose()
    '--- note: used in terminate -> no error handling
    If m_hSocket <> INVALID_SOCKET Then
        Call WSAAsyncSelect(m_hSocket, m_oHelperWindow.frMessageHWnd, 0, 0)
        m_oHelperWindow.frRemoveSocket Me, m_lIndex
        m_hSocket = INVALID_SOCKET
        Set m_pCleanup = Nothing
    End If
End Sub

Private Sub pvGetAdaptersInfo(vOutput As Variant)
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    Dim lPtr            As Long
    Dim uInfo           As IP_ADAPTER_INFO
    Dim vGateway        As Variant
    Dim vIpAddress      As Variant
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    vOutput = Array()
    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize = 0 Then
        GoTo QH
    End If
    ReDim baBuffer(0 To lSize - 1) As Byte
    If GetAdaptersInfo(baBuffer(0), lSize) <> 0 Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    lPtr = VarPtr(baBuffer(0))
    Do While lPtr <> 0
        Call CopyMemory(uInfo, ByVal lPtr, Len(uInfo))
        pvGetAddrList VarPtr(uInfo.GatewayList), vGateway
        For lIdx = 0 To UBound(vGateway)
            vGateway(lIdx) = vGateway(lIdx)(0)
        Next
        pvGetAddrList VarPtr(uInfo.IpAddressList), vIpAddress
        For Each vElem In vIpAddress
            ReDim Preserve vOutput(UBound(vOutput) + 1)
            vOutput(UBound(vOutput)) = Array(vElem(0), vElem(1), vGateway, Replace(uInfo.AdapterName, vbNullChar, vbNullString), Replace(uInfo.Description, vbNullChar, vbNullString), uInfo.Index, uInfo.Type)
        Next
        lPtr = uInfo.Next
    Loop
QH:
End Sub

Private Sub pvGetAddrList(ByVal lPtr As Long, vOutput As Variant)
    Dim uString         As IP_ADDR_STRING
    Dim sAddr           As String
    
    vOutput = Array()
    Do While lPtr <> 0
        Call CopyMemory(uString, ByVal lPtr, Len(uString))
        sAddr = pvToString(VarPtr(uString.IpAddress(0)))
        If sAddr <> "0.0.0.0" Then
            ReDim Preserve vOutput(UBound(vOutput) + 1)
            vOutput(UBound(vOutput)) = Array(sAddr, pvToString(VarPtr(uString.IpMask(0))))
        End If
        lPtr = uString.Next
    Loop
End Sub

Private Sub pvGetForwardTable(vOutput As Variant)
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim uInfo()         As MIB_IPFORWARDROW
    Dim lIdx            As Long
    Dim sDest           As String
    Dim sMask           As String
    Dim sNextHop        As String
    
    vOutput = Array()
    Call GetIpForwardTable(ByVal 0, lSize, 0)
    If lSize = 0 Then
        GoTo QH
    End If
    ReDim baBuffer(0 To lSize - 1) As Byte
    If GetIpForwardTable(baBuffer(0), lSize, 0) <> 0 Then
        m_lLastError = Err.LastDllError
        GoTo QH
    End If
    Call CopyMemory(lSize, baBuffer(0), 4)
    If lSize = 0 Then
        GoTo QH
    End If
    ReDim uInfo(0 To lSize - 1) As MIB_IPFORWARDROW
    Call CopyMemory(uInfo(0), baBuffer(4), lSize * sizeof_MIB_IPFORWARDROW)
    ReDim vOutput(0 To lSize - 1) As Variant
    For lIdx = 0 To lSize - 1
        With uInfo(lIdx)
            sDest = pvToString(ws_inet_ntoa(.dwForwardDest))
            sMask = pvToString(ws_inet_ntoa(.dwForwardMask))
            sNextHop = pvToString(ws_inet_ntoa(.dwForwardNextHop))
            vOutput(lIdx) = Array(sDest, sMask, sNextHop, .dwForwardIfIndex, .dwForwardType, .dwForwardMetric1)
        End With
    Next
QH:
End Sub

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
            Optional ByVal NeedLen As Long = 1, _
            Optional ByVal Timeout As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "SyncReceiveText"
    Dim baText()        As Byte
    Dim dblTimerEx      As Double
    Dim lElapsed        As Long
    Dim baBuffer()      As Byte
    Dim bResult         As Boolean
    
    On Error GoTo EH
    If Timeout = 0 Then
        Timeout = Znl(SockOpt(ucsSsoReceiveTimeout), DEF_TIMEOUT)
    End If
    baText = vbNullString
    dblTimerEx = TimerEx
    Do
        If Timeout > 0 Then
            lElapsed = Int((TimerEx - dblTimerEx) * 1000)
            If lElapsed >= Timeout Then
                m_lLastError = ERR_TIMEOUT
                GoTo QH
            End If
        End If
        bResult = SyncReceiveArray(baBuffer, 0, Timeout - lElapsed, HostAddress, HostPort)
        If UBound(baBuffer) >= 0 Then
            ReDim Preserve baText(0 To UBound(baText) + UBound(baBuffer) + 1) As Byte
            Call CopyMemory(baText(UBound(baText) - UBound(baBuffer)), baBuffer(0), UBound(baBuffer) + 1)
            SyncReceiveText = FromTextArray(baText, CodePage)
        End If
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
    Dim dblTimerEx      As Double
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
    dblTimerEx = TimerEx
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
            lElapsed = Int((TimerEx - dblTimerEx) * 1000)
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
        ElseIf lResult = 0 Then
            If Not SyncWaitForEvent(Timeout, ucsSfdRead) Then
                GoTo QH
            End If
        Else
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
    Const FUNC_NAME     As String = "SyncSendText"
    
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
        If pvResolveHandle = 0 Then
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
    hWnd = m_oHelperWindow.frMessageHWnd
    If Timeout > 0 Then
        lTimerID = SetTimer(hWnd, 1, Timeout, 0)
    End If
    bPeek = True
    Do
        If PeekMessage(uMsg, hWnd, WM_SOCKET_NOTIFY + m_lIndex, WM_SOCKET_NOTIFY + m_lIndex, PM_REMOVE) <> 0 Then
            If uMsg.wParam <> INVALID_SOCKET And uMsg.lParam <> 0 Then
                If uMsg.wParam = m_hSocket Then
                    If m_cMsgQueue Is Nothing Then
                        Set m_cMsgQueue = New VBA.Collection
                    End If
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
            If uMsg.wParam = pvResolveHandle Then
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
'                Call WaitMessage
                Call Sleep(1)
                bPeek = True
            End If
        Else
'            Call WaitMessage
            Call Sleep(1)
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
    If Not m_cMsgQueue Is Nothing Then
        If m_cMsgQueue.Count > 0 Then
            Call PostMessage(hWnd, WM_SOCKET_NOTIFY + m_lIndex, INVALID_SOCKET, 0)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub SyncCancelWait()
    m_bBlocking = False
End Sub

Public Sub SyncProcessMsgQueue()
    Const FUNC_NAME     As String = "SyncProcessMsgQueue"
    Dim vElem           As Variant
    
    On Error GoTo EH
    If Not m_cMsgQueue Is Nothing Then
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
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
#End If

'= helper window =========================================================

Friend Function frAddSocket(Socket As cAsyncSocket, Index As Long) As Boolean
    Const FUNC_NAME     As String = "frAddSocket"
    Dim lSize           As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    If frMessageHWnd = 0 Then
        GoTo QH
    End If
    With m_uWindowState(0)
        If Index <> -1 Then
            Debug.Assert .SocketPtr(Index) = ObjPtr(Socket)
            '--- success
            frAddSocket = True
            Exit Function
        End If
        lSize = UBound(.SocketPtr) + 1
        If .Count >= lSize Then
            lSize = 2 * lSize
            If lSize > MAX_SOCKETS Then
                lSize = MAX_SOCKETS
            End If
            ReDim Preserve .SocketPtr(0 To lSize - 1) As Long
        End If
        For lIdx = 0 To UBound(.SocketPtr)
            Index = (.Pos + lIdx) Mod lSize
            If .SocketPtr(Index) = 0 Then
                '--- note: weak reference
                .SocketPtr(Index) = ObjPtr(Socket)
                .Pos = Index + 1
                .Count = .Count + 1
                '--- success
                frAddSocket = True
                Exit Function
            End If
        Next
    End With
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
    If frMessageHWnd = 0 Then
        GoTo QH
    End If
    With m_uWindowState(0)
        If Index <> -1 Then
            Debug.Assert .SocketPtr(Index) = ObjPtr(Socket)
            .SocketPtr(Index) = 0
            Index = -1
            .Count = .Count - 1
        End If
        '--- success
        frRemoveSocket = True
    End With
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
    If frMessageHWnd = 0 Then
        GoTo QH
    End If
    With m_uWindowState(0)
        Select Case wMsg
        Case WM_SOCKET_NOTIFY - 1
            For lIdx = 0 To UBound(.SocketPtr)
                If .SocketPtr(lIdx) <> 0 Then
                    If pvToSocket(.SocketPtr(lIdx)).frNotifyGetHostByName(wParam, lParam) Then
                        pvHandleNotify = 1
                        GoTo QH
                    End If
                End If
            Next
        Case WM_SOCKET_NOTIFY To WM_SOCKET_NOTIFY + UBound(.SocketPtr)
            If wParam <> INVALID_SOCKET Then
                lIdx = wMsg - WM_SOCKET_NOTIFY
                If .SocketPtr(lIdx) <> 0 Then
                    pvToSocket(.SocketPtr(lIdx)).frNotifyEvent wParam, lParam
                    pvHandleNotify = 1
                End If
            End If
        End Select
    End With
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvToSocket(ByVal lPtr As Long) As cAsyncSocket
    Call vbaObjSetAddref(pvToSocket, lPtr)
End Function

Private Property Get pvAddressOfNotifyProc() As cAsyncSocket
    Set pvAddressOfNotifyProc = InitAddressOfMethod(Me, 4)
End Property

Public Function NotifyProc(ByVal lTablePtr As Long, ByVal lTableSize As Long, ByVal lStartIdx As Long, ByVal lEndIdx As Long) As Long
Attribute NotifyProc.VB_MemberFlags = "40"
    Dim uMsg        As APIMSG
    
    Do While lStartIdx <> lEndIdx
        Call CopyMemory(uMsg.lMessage, ByVal UnsignedAdd(lTablePtr, 12 * lStartIdx), 12)
        pvHandleNotify uMsg.lMessage, uMsg.wParam, uMsg.lParam
        lStartIdx = (lStartIdx + 1) And (lTableSize - 1)
    Loop
End Function

'= shared ================================================================

#If Not ImplUseShared Then
Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function Znl(ByVal lValue As Long, Optional IfEmptyLong As Variant = Null, Optional ByVal EmptyLong As Long = 0) As Variant
    Znl = IIf(lValue = EmptyLong, IfEmptyLong, lValue)
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitAsyncSelectNotifyThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABageplEEQAV1aLdCQUg8YIgz4AdDeL+oHHBBNEAIvCBQgRRACri8IFRBFEAKuLwgVUEUQAq4vCBYARRACri8IFgBJEAKu5CQAAAPOlgcIEE0QAi/qDxziLwqu4AQAAAKszwKuLRCQMqzPAq4t0JBSlpTPAq6urq4PvLGoAV/9yDP93DP9SFItEJBiJOF5fuGgTRAAtYBBEAAUAMAAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GlKLClL/cQz/cgz/URhaiwr/chBqAP9RJDPAwgQADx8AVYvsi1UY/0IE/0IoUosK/3UU/3UQ/3UM/3UI/1EcWlCLQiBAJf8DAAA7Qhx0GlBWV416LItCII08x408h411DKWlpV9ej0Igi0IkhcAPhZ4AAACLQhCFwA+FkwAAAIsKi0EwhcB0MFL/0FqJQgiD+AF3a4sKi0E0hcB0CFL/0FqFwHVaiwpSavD/cSj/USxaqQAAAAh1R4tCKDtCBHRSi0IUhcB0S8dCJAEAAAD/ciBSM8BQVP9yIP9yHGgABAAAjUIsUP9yFP9SGFhax0IkAAAAAFiJQhw7QiB0Feu5UosK/3EQagBqAGoA/1EgWolCEP9KKFLo2v7//1hdwhgAkOgAAAAAWoHqhRJEAIHCPBNEAP9CBP9CKIsKi0EwhcB0DFL/0FqJQgiD+AF3SlKLCv9yEGoA/1EkWotCKDtCBHQvi0IUhcB0KP9yIFIzwFBU/3Ig/3IcaAAEAACNQixQ/3IU/1IYWFpYiUIcO0IgdcnHQhAAAAAA/0ooUuhU/v//whAAkI1EQAGNREAB" ' 27.6.2020 20:49:42
    Const THUNK_SIZE    As Long = 13064
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long

    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    If hThunk = 0 Then
        hThunk = pvThunkGlobalData("InitAsyncSelectNotifyThunk" & App.ThreadID)
    End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(2) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(3) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(6) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        pvThunkGlobalData("InitAsyncSelectNotifyThunk" & App.ThreadID) = hThunk
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitAsyncSelectNotifyThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitCleanupThunk(ByVal hHandle As Long, sModuleName As String, sProcName As String) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepQEDwBV1aLdCQUgz4AdCeL+oHHPBE8AYvCBcwQPAGri8IFCBE8AauLwgUYETwBq7kCAAAA86WBwjwRPAFSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBE8AS1QEDwBwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRL/cgj/UgyLVCQEiwpS/1EQM8DCBAAPHwA=" ' 25.3.2019 14:03:56
    Const THUNK_SIZE    As Long = 256
    Static hThunk       As Long
    Dim aParams(0 To 1) As Long
    Dim pfnCleanup      As Long
    Dim lSize           As Long
    
    If hThunk = 0 Then
        hThunk = pvThunkGlobalData("InitCleanupThunk")
    End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        aParams(0) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(1) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        pvThunkGlobalData("InitCleanupThunk") = hThunk
    End If
    pfnCleanup = GetProcAddress(GetModuleHandle(sModuleName), sProcName)
    If pfnCleanup <> 0 Then
        lSize = CallWindowProc(hThunk, hHandle, pfnCleanup, VarPtr(aParams(0)), VarPtr(InitCleanupThunk))
        Debug.Assert lSize = THUNK_SIZE
    End If
End Function

Private Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, 4)
    End If
End Property

Private Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, 4)
    End If
End Property

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Const ERR_INVALID_WINDOW_HANDLE As Long = 1400
    Dim sHelperName     As String
    Dim hWndParent      As Long
    Dim hWndHelper      As Long
    Dim lPtr            As Long
    Dim baWSAData()     As Byte
    
    sHelperName = GetCurrentProcessId() & ":" & App.ThreadID & ":" & STR_HELPER_NAME
    hWndParent = HWND_MESSAGE
    Do
        hWndHelper = FindWindowEx(hWndParent, 0, STR_HELPER_CLASS, sHelperName)
        If hWndHelper = 0 And Err.LastDllError = ERR_INVALID_WINDOW_HANDLE Then
            hWndParent = 0
        Else
            Exit Do
        End If
    Loop
    lPtr = GetWindowLong(hWndHelper, GWL_USERDATA)
    If hWndHelper = 0 Or lPtr <> 0 Then
        If hWndHelper = 0 Then
            hWndHelper = CreateWindowEx(0, STR_HELPER_CLASS, sHelperName, 0, 0, 0, 0, 0, hWndParent, 0, App.hInstance, ByVal 0)
        End If
        If lPtr <> 0 Then
            Set m_oHelperWindow = pvToSocket(lPtr)
        Else
            Set m_oHelperWindow = New cAsyncSocket
        End If
        m_hSocket = INVALID_SOCKET
        m_lIndex = -1
    Else
        '-- init helper window
        ReDim baWSAData(0 To 1000) As Byte
        ReDim m_uWindowState(0 To 0) As UcsHelperWindowStateType
        With m_uWindowState(0)
            If WSAStartup(&H101, baWSAData(0)) = 0 Then
                .hWnd = hWndHelper
                Call SetWindowLong(hWndHelper, GWL_USERDATA, ObjPtr(Me))
                Set .Notify = InitAsyncSelectNotifyThunk(hWndHelper, Me, pvAddressOfNotifyProc.NotifyProc(0, 0, 0, 0))
                Set .Cleanup = InitCleanupThunk(hWndHelper, "user32", "DestroyWindow")
            End If
            ReDim .SocketPtr(0 To 511) As Long
        End With
    End If
End Sub

Private Sub Class_Terminate()
    If frMessageHWnd = 0 Then
        pvClose
        Set m_oHelperWindow = Nothing
    Else
        With m_uWindowState(0)
            '--- terminate notifications
            Debug.Assert ThunkPrivateData(.Notify, 3) = ObjPtr(Me)
            ThunkPrivateData(.Notify, 3) = 0
            Set .Notify = Nothing
            '--- terminate helper window
            Call SetWindowLong(.hWnd, GWL_USERDATA, 0)
            Set .Cleanup = Nothing
            '--- terminate winsock
            Call WSACleanup
        End With
        Erase m_uWindowState
    End If
End Sub
