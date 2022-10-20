VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpRequest"
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
Private Const MODULE_NAME As String = "cHttpRequest"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplUseTls = True
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)

'=========================================================================
' Public enums
'=========================================================================

Public Enum WinHttpRequestOption
    WinHttpRequestOption_UserAgentString = 0
    WinHttpRequestOption_URL = 1
    WinHttpRequestOption_URLCodePage = 2
    WinHttpRequestOption_EscapePercentInURL = 3
    WinHttpRequestOption_SslErrorIgnoreFlags = 4
    WinHttpRequestOption_SelectCertificate = 5
    WinHttpRequestOption_EnableRedirects = 6
    WinHttpRequestOption_UrlEscapeDisable = 7
    WinHttpRequestOption_UrlEscapeDisableQuery = 8
    WinHttpRequestOption_SecureProtocols = 9
    WinHttpRequestOption_EnableTracing = 10
    WinHttpRequestOption_RevertImpersonationOverSsl = 11
    WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
    WinHttpRequestOption_EnablePassportAuthentication = 13
    WinHttpRequestOption_MaxAutomaticRedirects = 14
    WinHttpRequestOption_MaxResponseHeaderSize = 15
    WinHttpRequestOption_MaxResponseDrainSize = 16
    WinHttpRequestOption_EnableHttp1_1 = 17
    WinHttpRequestOption_EnableCertificateRevocationCheck = 18
    WinHttpRequestOption_RejectUserpwd = 19
    WinHttpRequestOption_RootCA = 20
End Enum
Private Const sizeof_WinHttpRequestOption As Long = WinHttpRequestOption_RootCA

Public Enum WinHttpRequestAutoLogonPolicy
    AutoLogonPolicy_Always = 0
    AutoLogonPolicy_OnlyIfBypassProxy = 1
    AutoLogonPolicy_Never = 2
End Enum

Public Enum WinHttpRequestSslErrorFlags '--- bitmask
    SslErrorFlag_UnknownCA = 2 ^ 8
    SslErrorFlag_CertWrongUsage = 2 ^ 9
    SslErrorFlag_CertCNInvalid = 2 ^ 12
    SslErrorFlag_CertDateInvalid = 2 ^ 13
    SslErrorFlag_Ignore_All = SslErrorFlag_UnknownCA Or SslErrorFlag_CertWrongUsage Or SslErrorFlag_CertCNInvalid Or SslErrorFlag_CertDateInvalid
End Enum

Public Enum WinHttpRequestSecureProtocols '--- bitmask
    SecureProtocol_SSL2 = 2 ^ 3
    SecureProtocol_SSL3 = 2 ^ 5
    SecureProtocol_TLS1 = 2 ^ 7
    SecureProtocol_TLS1_1 = 2 ^ 9
    SecureProtocol_TLS1_2 = 2 ^ 11
    SecureProtocol_TLS1_3 = 2 ^ 13
    SecureProtocol_ALL = SecureProtocol_SSL2 Or SecureProtocol_SSL3 Or SecureProtocol_TLS1 '-- nothing past TLS1
End Enum

'=========================================================================
' Events
'=========================================================================

Event OnResponseStart(Status As Long, ContentType As String)
Event OnResponseDataAvailable(Data() As Byte)
Event OnResponseFinished()
Event OnError(ErrorNumber As Long, ErrorDescription As String)

'=========================================================================
' API
'=========================================================================

'--- for IStream_Seek
Private Const STREAM_SEEK_SET               As Long = 0
Private Const STREAM_SEEK_END               As Long = 2
'--- errors
Private Const WINHTTP_ERROR_BASE                    As Long = &H80070000 + 12000
Private Const ERROR_WINHTTP_TIMEOUT                 As Long = (WINHTTP_ERROR_BASE + 2)
Private Const ERROR_WINHTTP_INVALID_URL             As Long = (WINHTTP_ERROR_BASE + 5)
Private Const ERROR_WINHTTP_UNRECOGNIZED_SCHEME     As Long = (WINHTTP_ERROR_BASE + 6)
Private Const ERROR_WINHTTP_NAME_NOT_RESOLVED       As Long = (WINHTTP_ERROR_BASE + 7)
Private Const ERROR_WINHTTP_INVALID_SERVER_RESPONSE As Long = (WINHTTP_ERROR_BASE + 152)
Private Const ERROR_WINHTTP_SECURE_INVALID_CERT     As Long = (WINHTTP_ERROR_BASE + 169)
Private Const ERROR_WINHTTP_HEADER_SIZE_OVERFLOW    As Long = (WINHTTP_ERROR_BASE + 182)
Private Const ERROR_INVALID_PARAMETER               As Long = &H80070057
Private Const E_NOTIMPL                             As Long = &H80004001
Private Const E_POINTER                             As Long = &H80004003
'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function WinHttpGetDefaultProxyConfiguration Lib "winhttp" (pProxyInfo As Any) As Long
'--- shlwapi
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeW" (ByVal pszURL As Long, ByVal pszEscaped As Long, pcchEscaped As Long, ByVal dwFlags As Long) As Long
'--- libarchive
Private Declare Function archive_read_new Lib "archiveint" Alias "_archive_read_new@0" () As Long
Private Declare Function archive_read_free Lib "archiveint" Alias "_archive_read_free@4" (ByVal hArchive As Long) As Long
Private Declare Function archive_read_support_filter_gzip Lib "archiveint" Alias "_archive_read_support_filter_gzip@4" (ByVal hArchive As Long) As Long
Private Declare Function archive_read_support_format_raw Lib "archiveint" Alias "_archive_read_support_format_raw@4" (ByVal hArchive As Long) As Long
Private Declare Function archive_read_open Lib "archiveint" Alias "_archive_read_open@20" (ByVal hArchive As Long, ByVal pClientData As Long, ByVal lOpener As Long, ByVal lReader As Long, ByVal lCloser As Long) As Long
Private Declare Function archive_read_next_header Lib "archiveint" Alias "_archive_read_next_header@8" (ByVal hArchive As Long, pHeader As Long) As Long
Private Declare Function archive_read_data Lib "archiveint" Alias "_archive_read_data@12" (ByVal hArchive As Long, pBuffer As Any, ByVal lSize As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    '--- for Modern Subclassing Thunk (MST)
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    '--- end MST
#End If

Private Type WINHTTP_PROXY_INFO
    dwAccessType        As Long
    lpszProxy           As Long
    lpszProxyBypass     As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_HOST              As String = "Host"
Private Const HDR_USER_AGENT        As String = "User-Agent"
Private Const HDR_CONTENT_TYPE      As String = "Content-Type"
Private Const HDR_CONTENT_LENGTH    As String = "Content-Length"
Private Const HDR_LOCATION          As String = "Location"
Private Const HDR_CONNECTION        As String = "Connection"
Private Const HDR_PROXY_AUTH        As String = "Proxy-Authorization"
Private Const HDR_PROXY_CONNECTION  As String = "Proxy-Connection"
Private Const HDR_SET_COOKIE        As String = "Set-Cookie"
Private Const HDR_COOKIE            As String = "Cookie"
Private Const HDR_DELIM             As String = vbCrLf & vbCrLf
Private Const DEF_BUFFER_SIZE       As Long = 256& * 1024
Private Const DEF_USER_AGENT        As String = "Mozilla/4.0 (compatible; Win32; VbAsyncSocket)"
Private Const DEF_CONTENT_TYPE      As String = "text/plain"
Private Const DEF_CHARSET           As String = "charset=utf-8"
Private Const STR_KEEP_ALIVE        As String = "Keep-Alive"

#If ImplUseTls Then
    Private WithEvents m_oSocket As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
#Else
    Private WithEvents m_oSocket As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
#End If
Private m_vLastError            As Variant
Private m_eState                As UcsStateEnum
Private m_vOptions(0 To sizeof_WinHttpRequestOption) As Variant
Private m_uRequest              As UcsRequestType
Private m_uResponse             As UcsResponseType
Private m_oCookies              As Object '--- keyed on domains, contains collections of path + cookie per domain

Private Enum UcsStateEnum
    ucsIdle
    ucsWaitResolve
    ucsWaitConnect
    ucsWaitSendProxy
    ucsWaitRecvProxy
    ucsWaitSendBody
    ucsWaitRecvHeaders
    ucsWaitRecvBody
End Enum

Private Type UcsParsedUrl
    Protocol            As String
    Host                As String
    UserPort            As String
    Port                As Long
    Path                As String
    QueryString         As String
    Anchor              As String
    User                As String
    Pass                As String
End Type

Private Type UcsBuffer
    Data()              As Byte
    Pos                 As Long
    Size                As Long
End Type

Private Type UcsRequestType
    ResolveTimeout      As Long
    ConnectTimeout      As Long
    SendTimeout         As Long
    ReceiveTimeout      As Long
    Method              As String
    Remote              As UcsParsedUrl
    ProxyList()         As UcsParsedUrl
    BypassList          As Variant
    CurrentProxy        As UcsParsedUrl
    Async               As Boolean
    Headers             As Object
    NumOfRedirects      As Long
    Stream              As IUnknown
    SendBuffer          As UcsBuffer
End Type

Private Type UcsResponseType
    Status              As Long
    StatusText          As String
    AllHeaders          As String
    Headers             As Object
    Stream              As IUnknown
    Text                As String
    BytesProgress       As Double
    ContentLength       As Double
    CurrentRedirect     As Long
    TransferEncoding    As String
    TransferChunks()    As UcsBuffer
    ContentEncoding     As String
    ContentData()       As Byte
    RecvBuffer          As UcsBuffer
    ChunksBuffer        As UcsBuffer
End Type

Private Type UcsRedirectReaderThunk
    Code(0 To 8)        As Long
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

Private Sub RaiseError(sFunction As String)
    Dim vErr                As Variant
    
    vErr = Array(Err.Number, Err.Source, Err.Description)
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
    Err.Raise vErr(0), MODULE_NAME & "." & sFunction & vbCrLf & vErr(1), vErr(2)
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get ResponseStream() As Variant
    Set ResponseStream = m_uResponse.Stream
End Property

Public Property Get ResponseBody() As Byte()
    Const FUNC_NAME     As String = "ResponseBody"
    Dim lSize           As Long
    Dim baData()        As Byte
    Dim hResult         As Long
    
    On Error GoTo EH
    m_vLastError = Empty
    lSize = IStream_GetSize(m_uResponse.Stream)
    If lSize > 0 Then
        ReDim baData(0 To lSize - 1) As Byte
        hResult = IStream_Read(m_uResponse.Stream, baData)
        If hResult < 0 Then
            pvSetError hResult, "IStream_Read"
            GoTo QH
        End If
    Else
        baData = vbNullString
    End If
    ResponseBody = baData
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Public Property Get ResponseText() As String
    Dim eCodePage       As UcsAsyncSocketCodePageEnum
    Dim sCharset        As String
    
    With m_uResponse
        If .Text = vbNullChar Then
            eCodePage = ucsScpUtf8
            If .Headers.Exists(HDR_CONTENT_TYPE) Then
                sCharset = pvParseHeaderAttribs(.Headers.Item(HDR_CONTENT_TYPE)).Item("charset")
                If LenB(sCharset) <> 0 Then
                    eCodePage = pvMapCharset(sCharset)
                End If
            End If
            .Text = pvFromTextArray(ResponseBody, eCodePage)
        End If
        ResponseText = .Text
    End With
End Property

Public Property Get Status() As Long
    Status = m_uResponse.Status
End Property

Public Property Get StatusText() As String
    StatusText = m_uResponse.StatusText
End Property

Public Property Get Option_(ByVal eOption As WinHttpRequestOption) As Variant
    If IsObject(m_vOptions(eOption)) Then
        Set Option_ = m_vOptions(eOption)
    Else
        Option_ = m_vOptions(eOption)
    End If
End Property

Public Property Let Option_(ByVal eOption As WinHttpRequestOption, ByVal RHS As Variant)
    If IsObject(m_vOptions(eOption)) Then
        Set m_vOptions(eOption) = RHS
    Else
        Select Case VarType(m_vOptions(eOption))
        Case vbBoolean
            RHS = CBool(RHS)
        Case vbLong
            RHS = CLng(RHS)
        Case vbString
            RHS = CStr(RHS)
        End Select
        m_vOptions(eOption) = RHS
    End If
End Property

Private Property Get pvAddressOfReaderCallback() As cHttpRequest
    Set pvAddressOfReaderCallback = InitAddressOfMethod(Me, 3)
End Property

#If ImplUseTls Then
Private Property Get pvLocalFeatures() As UcsTlsLocalFeaturesEnum
    Dim eProt           As WinHttpRequestSecureProtocols
        
    eProt = m_vOptions(WinHttpRequestOption_SecureProtocols)
    If eProt = SecureProtocol_ALL Then
        pvLocalFeatures = ucsTlsSupportAll
    Else
        pvLocalFeatures = ucsTlsSupportTls13 * -((eProt And SecureProtocol_TLS1_3) <> 0) _
            Or ucsTlsSupportTls12 * -((eProt And SecureProtocol_TLS1_2) <> 0) _
            Or ucsTlsSupportTls11 * -((eProt And SecureProtocol_TLS1_1) <> 0) _
            Or ucsTlsSupportTls10 * -((eProt And SecureProtocol_TLS1) <> 0)
    End If
    If m_vOptions(WinHttpRequestOption_SslErrorIgnoreFlags) <> 0 Then
        pvLocalFeatures = pvLocalFeatures Or ucsTlsIgnoreServerCertificateErrors
    End If
    If Not m_vOptions(WinHttpRequestOption_EnableCertificateRevocationCheck) Then
        pvLocalFeatures = pvLocalFeatures Or ucsTlsIgnoreServerCertificateRevocation
    End If
End Property
#End If

'=========================================================================
' Methods
'=========================================================================

Public Sub Open_(ByVal Method As String, ByVal Url As String, Optional ByVal Async As Variant)
    Const FUNC_NAME     As String = "Open"
    Dim sEscaped        As String
    Dim uParsed         As UcsParsedUrl
    
    On Error GoTo EH
    m_vLastError = Empty
    sEscaped = pvEscapeUrl(Url)
    m_vOptions(WinHttpRequestOption_URL) = sEscaped
    If Not pvParseUrl(sEscaped, uParsed, "http") Or LenB(uParsed.Host) = 0 Then
        pvSetError ERROR_WINHTTP_INVALID_URL
        GoTo QH
    End If
    If Not pvIsProtocolSupported(uParsed.Protocol) Then
        pvSetError ERROR_WINHTTP_UNRECOGNIZED_SCHEME
        GoTo QH
    End If
    If m_vOptions(WinHttpRequestOption_RejectUserpwd) Then
        uParsed.User = vbNullString
        uParsed.Pass = vbNullString
    End If
    With m_uRequest
        If LenB(.CurrentProxy.Host) <> 0 And pvIsEqual(GetResponseHeader(HDR_PROXY_CONNECTION), STR_KEEP_ALIVE) And pvIsEqual(.Remote.Protocol, "http") Then
            '--- keep proxy connection alive
        ElseIf LenB(.CurrentProxy.Host) = 0 And pvIsEqual(GetResponseHeader(HDR_CONNECTION), STR_KEEP_ALIVE) And pvIsEqual(pvFormatHost(.Remote), pvFormatHost(uParsed)) Then
            '--- keep host connection alive
        Else
            m_oSocket.Close_
            Set m_oSocket = pvCreateNewSocket
        End If
        .Method = UCase$(Method)
        .Remote = uParsed
        .Async = IIf(Not IsMissing(Async), Async, False)
        .NumOfRedirects = 0
        '--- init headers
        SetRequestHeader HDR_HOST, .Remote.Host
        SetRequestHeader HDR_COOKIE, pvGetAllCookies(.Remote.Host, .Remote.Path)
        SetRequestHeader HDR_USER_AGENT, m_vOptions(WinHttpRequestOption_UserAgentString)
        If m_vOptions(WinHttpRequestOption_EnableHttp1_1) Then
            SetRequestHeader HDR_CONNECTION, STR_KEEP_ALIVE
        End If
        SetRequestHeader "Accept", "*/*"
    End With
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub Send(Optional ByVal Body As Variant)
    Const FUNC_NAME     As String = "Send"
    Dim baData()        As Byte
    Dim eCodePage       As UcsAsyncSocketCodePageEnum
    Dim sCharset        As String
    
    On Error GoTo EH
    m_vLastError = Empty
    With m_uRequest
        '--- preapare request stream from body
        If IsObject(Body) Then
            Set .Stream = AsIStream(Body)
        ElseIf IsMissing(Body) Or IsEmpty(Body) Or IsNull(Body) Then
            Set .Stream = Nothing
        ElseIf VarType(Body) = vbByte + vbArray Then
            If UBound(Body) >= 0 Then
                Set .Stream = SHCreateMemStream(Body(0), UBound(Body) + 1)
            Else
                Set .Stream = Nothing
            End If
        ElseIf VarType(Body) = vbString Then
            If LenB(Body) <> 0 Then
                eCodePage = ucsScpUtf8
                If Not .Headers.Exists(HDR_CONTENT_TYPE) Then
                    SetRequestHeader HDR_CONTENT_TYPE, DEF_CONTENT_TYPE & "; " & DEF_CHARSET
                Else
                    sCharset = pvParseHeaderAttribs(.Headers.Item(HDR_CONTENT_TYPE)).Item("charset")
                    If LenB(sCharset) <> 0 Then
                        eCodePage = pvMapCharset(sCharset)
                    Else
                        SetRequestHeader HDR_CONTENT_TYPE, .Headers.Item(HDR_CONTENT_TYPE) & "; " & DEF_CHARSET
                    End If
                End If
                baData = pvToTextArray(CStr(Body), eCodePage)
                Set .Stream = SHCreateMemStream(baData(0), UBound(baData) + 1)
            Else
                Set .Stream = Nothing
            End If
        Else
            pvSetError ERROR_INVALID_PARAMETER
            GoTo QH
        End If
        '--- prepare headers
        If Not .Stream Is Nothing Or .Method <> "GET" Then
            SetRequestHeader HDR_CONTENT_LENGTH, IStream_GetSize(.Stream)
        End If
        If LenB(.Remote.User) <> 0 Or LenB(.Remote.Pass) <> 0 Then
            SetRequestHeader "Authorization", "Basic " & Replace(ToBase64Array( _
                pvToTextArray(.Remote.User & ":" & .Remote.Pass, ucsScpAcp)), vbCrLf, vbNullString)
        End If
        SetRequestHeader HDR_HOST, .Remote.Host
        SetRequestHeader HDR_COOKIE, pvGetAllCookies(.Remote.Host, .Remote.Path)
        pvResetResponse m_uResponse
        '--- try to reuse socket if not closed (keep-alive)
        pvSetState ucsWaitResolve
        If m_oSocket.IsClosed Then
            If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
                pvSetError m_oSocket.LastError, "m_oSocket.Create"
                GoTo QH
            End If
            If Not pvConnectRemote(m_uRequest) Then
                pvSetError m_oSocket.LastError, "m_oSocket.Connect"
                GoTo QH
            End If
        Else
            pvDrainKeepAlive
            m_oSocket.PostEvent ucsSfdWrite
        End If
        '--- handle synchronous mode
        If Not .Async Then
            If Not pvWaitForState(ucsWaitResolve, .ResolveTimeout) Then
                pvSetError ERROR_WINHTTP_TIMEOUT
                GoTo QH
            End If
            If Not pvWaitForState(ucsWaitConnect, .ConnectTimeout) Then
                pvSetError ERROR_WINHTTP_TIMEOUT
                GoTo QH
            End If
            If Not pvWaitForState(ucsWaitSendBody, .SendTimeout) Then
                pvSetError ERROR_WINHTTP_TIMEOUT
                GoTo QH
            End If
            If Not pvWaitForState(ucsWaitRecvBody, .ReceiveTimeout) Then
                pvSetError ERROR_WINHTTP_TIMEOUT
                GoTo QH
            End If
        End If
    End With
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function WaitForResponse(Optional ByVal Timeout As Variant) As Boolean
    Const FUNC_NAME     As String = "WaitForResponse"
    
    On Error GoTo EH
    m_vLastError = Empty
    If Not pvWaitForState(ucsWaitRecvBody, IIf(Not IsMissing(Timeout), Timeout, -1)) Then
        GoTo QH
    End If
    '--- success
    WaitForResponse = True
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
    With m_uRequest
        .ResolveTimeout = ResolveTimeout
        .ConnectTimeout = ConnectTimeout
        .SendTimeout = SendTimeout
        .ReceiveTimeout = ReceiveTimeout
    End With
End Sub

Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)
    With m_uRequest
        If LenB(Value) <> 0 Then
            .Headers.Item(Header) = Value
        ElseIf .Headers.Exists(Header) Then
            .Headers.Remove Header
        End If
    End With
End Sub

Public Sub Abort()
    pvSetState ucsIdle
    Set m_oSocket = pvCreateNewSocket
    Set m_uRequest.Stream = Nothing
    m_uRequest.Headers.RemoveAll
    Set m_uResponse.Stream = Nothing
    m_uResponse.Text = vbNullChar
End Sub

Public Function GetAllResponseHeaders() As String
    Dim lPos            As Long
    
    lPos = InStr(m_uResponse.AllHeaders, vbCrLf)
    If lPos > 0 Then
        GetAllResponseHeaders = Mid$(m_uResponse.AllHeaders, lPos + 2)
    End If
End Function

Public Function GetResponseHeader(ByVal Header As String) As String
    With m_uResponse
        If .Headers.Exists(Header) Then
            GetResponseHeader = .Headers.Item(Header)
        End If
    End With
End Function

Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As Long)
    Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER As Long = 0
    Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY As Long = 1
    
    Select Case Flags
    Case HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
        With m_uRequest.Remote
            .User = UserName
            .Pass = Password
        End With
    Case HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
        With m_uRequest.CurrentProxy
            .User = UserName
            .Pass = Password
        End With
    Case Else
        Err.Raise E_NOTIMPL, MODULE_NAME & ".SetCredentials", m_oSocket.GetErrorDescription(E_NOTIMPL)
    End Select
End Sub

Public Sub SetProxy(ByVal ProxySetting As Long, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
    Const FUNC_NAME     As String = "SetProxy"
    Const HTTPREQUEST_PROXYSETTING_PRECONFIG As Long = 0
    Const HTTPREQUEST_PROXYSETTING_DIRECT As Long = 1
    Const HTTPREQUEST_PROXYSETTING_PROXY As Long = 2
    Const WINHTTP_ACCESS_TYPE_NAMED_PROXY As Long = 3
    Dim uProxyList()    As UcsParsedUrl
    Dim vBypassList     As Variant
    Dim uParsed         As UcsParsedUrl
    Dim uInfo           As WINHTTP_PROXY_INFO
    Dim vElem           As Variant
    Dim lSize           As Long
    
    On Error GoTo EH
    m_vLastError = Empty
    If ProxySetting = HTTPREQUEST_PROXYSETTING_PRECONFIG Then
         If WinHttpGetDefaultProxyConfiguration(uInfo) = 0 Then
            pvSetError Err.LastDllError
            GoTo QH
        End If
        ProxyServer = pvToStringAndFree(uInfo.lpszProxy)
        BypassList = pvToStringAndFree(uInfo.lpszProxyBypass)
        If uInfo.dwAccessType = WINHTTP_ACCESS_TYPE_NAMED_PROXY Then
            ProxySetting = HTTPREQUEST_PROXYSETTING_PROXY
        Else
            ProxySetting = HTTPREQUEST_PROXYSETTING_DIRECT
        End If
    End If
    Select Case ProxySetting
    Case HTTPREQUEST_PROXYSETTING_DIRECT
        '--- will clear proxy settings
        ReDim uProxyList(-1 To -1) As UcsParsedUrl
    Case HTTPREQUEST_PROXYSETTING_PROXY
        '--- comma/semicolon/whitespace separated list of [<scheme>=][<scheme>"://"]<server>[":"<port>]
        For Each vElem In Split(Replace(Replace(IIf(IsMissing(ProxyServer), vbNullString, ProxyServer), ";", " "), ",", " "))
            vElem = Split(vElem, "=", Limit:=2)
            If Not pvParseUrl(Trim$(vElem(UBound(vElem))), uParsed, "http") Or LenB(uParsed.Host) = 0 Then
                pvSetError ERROR_WINHTTP_INVALID_URL
                GoTo QH
            End If
            If Not pvIsProtocolSupported(uParsed.Protocol) Then
                pvSetError ERROR_WINHTTP_UNRECOGNIZED_SCHEME
                GoTo QH
            End If
            '--- keep proxy's target remote protocol in unused anchor member
            uParsed.Anchor = IIf(UBound(vElem) > 0, Trim$(vElem(0)), vbNullString)
            If lSize = 0 Then
                ReDim uProxyList(0 To lSize) As UcsParsedUrl
            Else
                ReDim Preserve uProxyList(0 To lSize) As UcsParsedUrl
            End If
            uProxyList(lSize) = uParsed
            lSize = lSize + 1
        Next
        vBypassList = Split(Replace(Replace(IIf(IsMissing(BypassList) Or IsEmpty(BypassList), "<local>", BypassList), ";", " "), ",", " "))
    Case Else
        pvSetError E_NOTIMPL
        GoTo QH
    End Select
    m_uRequest.ProxyList = uProxyList
    m_uRequest.BypassList = vBypassList
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub SetClientCertificate(ByVal ClientCertificate As String)
#If ImplUseTls Then
    Const FUNC_NAME     As String = "SetClientCertificate"
    
    On Error GoTo EH
    m_vLastError = Empty
    If Not m_oSocket.ImportSystemStoreCertificates(ClientCertificate) Then
        pvSetError ERROR_WINHTTP_SECURE_INVALID_CERT
        GoTo QH
    End If
    m_vOptions(WinHttpRequestOption_SelectCertificate) = ClientCertificate
QH:
    If IsArray(m_vLastError) Then
        On Error GoTo 0
        Err.Raise m_vLastError(0), MODULE_NAME & "." & FUNC_NAME & vbCrLf & m_vLastError(1), m_vLastError(2)
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
#Else
    Err.Raise E_NOTIMPL, MODULE_NAME & ".SetClientCertificate", m_oSocket.GetErrorDescription(E_NOTIMPL)
#End If
End Sub

'= not implemented =======================================================

Public Sub SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)
    #If AutoLogonPolicy Then '--- touch args for MZ-Tools
    #End If
    Err.Raise E_NOTIMPL, MODULE_NAME & ".SetAutoLogonPolicy", m_oSocket.GetErrorDescription(E_NOTIMPL)
End Sub

'= private ===============================================================

Private Function pvSendHeaders(uRequest As UcsRequestType) As Boolean
    Const FUNC_NAME     As String = "pvSendHeaders"
    Dim sHttpVersion    As String
    Dim sProxyAuth      As String
    Dim sRequestLine    As String
    
    On Error GoTo EH
    With uRequest
        .SendBuffer.Size = 0
        .SendBuffer.Pos = 0
        sHttpVersion = "HTTP/1." & -m_vOptions(WinHttpRequestOption_EnableHttp1_1)
        If LenB(.CurrentProxy.User) <> 0 Or LenB(.CurrentProxy.Pass) <> 0 Then
            sProxyAuth = "Basic " & Replace(ToBase64Array(pvToTextArray(.CurrentProxy.User & ":" & .CurrentProxy.Pass, ucsScpAcp)), vbCrLf, vbNullString)
        End If
        If m_eState < ucsWaitSendProxy And pvIsEqual(.Remote.Protocol, "https") And LenB(.CurrentProxy.Host) <> 0 Then
            pvBufferWriteArray .SendBuffer, pvToTextArray( _
                "CONNECT " & .Remote.Host & ":" & .Remote.Port & " " & sHttpVersion & vbCrLf & _
                HDR_HOST & ": " & .Remote.Host & ":" & .Remote.Port & IIf(LenB(sProxyAuth) <> 0, vbCrLf & _
                HDR_PROXY_AUTH & ": " & sProxyAuth, vbNullString) & HDR_DELIM, ucsScpAcp)
            pvSetState ucsWaitSendProxy
        Else
            sRequestLine = .Remote.Path & .Remote.QueryString & .Remote.Anchor & " " & sHttpVersion
            If pvIsEqual(.Remote.Protocol, "http") And LenB(.CurrentProxy.Host) <> 0 Then
                sRequestLine = pvFormatHost(.Remote) & sRequestLine
                SetRequestHeader HDR_PROXY_AUTH, sProxyAuth
                SetRequestHeader HDR_PROXY_CONNECTION, STR_KEEP_ALIVE
            End If
            pvBufferWriteArray .SendBuffer, pvToTextArray( _
                .Method & " " & sRequestLine & vbCrLf & _
                pvJoinHeaders(.Headers) & HDR_DELIM, ucsScpAcp)
            pvSetState ucsWaitSendBody
        End If
    End With
    '--- success
    pvSendHeaders = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvSendBody(uRequest As UcsRequestType) As Boolean
    Const FUNC_NAME     As String = "pvSendBody"
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hResult         As Long
    
    On Error GoTo EH
    With uRequest
        Do
            lSize = .SendBuffer.Size - .SendBuffer.Pos
            If lSize <= 0 Then
                If m_eState < ucsWaitRecvProxy Then
                    pvSetState ucsWaitRecvProxy
                    pvSendBody = True
                    GoTo QH
                Else
                    If pvArraySize(.SendBuffer.Data) <> DEF_BUFFER_SIZE Then
                        ReDim .SendBuffer.Data(0 To DEF_BUFFER_SIZE - 1) As Byte
                    End If
                    hResult = IStream_Read(.Stream, .SendBuffer.Data, lSize)
                    If hResult < 0 Then
                        pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Read"
                        GoTo QH
                    End If
                    If lSize = 0 Then
                        m_uResponse.RecvBuffer.Size = 0
                        m_uResponse.ChunksBuffer.Size = 0
                        pvSetState ucsWaitRecvHeaders
                        pvSendBody = True
                        GoTo QH
                    End If
                    .SendBuffer.Size = lSize
                    .SendBuffer.Pos = 0
                End If
            End If
            If lSize <= 0 Then
                Exit Do
            End If
            lResult = m_oSocket.Send(VarPtr(.SendBuffer.Data(.SendBuffer.Pos)), lSize)
            If lResult < 0 Then
                If Not m_oSocket.HasPendingEvent Then
                    pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Send"
                    GoTo QH
                End If
            Else
                .SendBuffer.Pos = .SendBuffer.Pos + lResult
            End If
            If m_oSocket.HasPendingEvent Then
                Exit Do
            End If
        Loop
    End With
    '--- success
    pvSendBody = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvRecvHeaders(baData() As Byte, uResponse As UcsResponseType, uRequest As UcsRequestType) As Boolean
    Const FUNC_NAME     As String = "pvRecvHeaders"
    Const sizeof_HDR_DELIM As Long = 4   '-- vbCrLf & vbCrLf
    Const sizeof_HTTP_VER As Long = 8    '-- "HTTP/1.1"
    Const sizeof_HTTP_STATUS As Long = 3 '-- "200"
    Dim lPos            As Long
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim uRedirect       As UcsParsedUrl
    Dim lIdx            As Long
    Dim sEscaped        As String
    
    On Error GoTo EH
    With uResponse
        pvBufferWriteArray .RecvBuffer, baData
        lPos = InStrB(1, .RecvBuffer.Data, StrConv(HDR_DELIM, vbFromUnicode))
        If lPos > 0 Then
            .AllHeaders = Left$(pvFromTextArray(.RecvBuffer.Data, ucsScpAcp), lPos - 1)
            '--- unfold headers
            If InStr(.AllHeaders, vbCrLf & " ") > 0 Or InStr(.AllHeaders, vbCrLf & vbTab) > 0 Then
                .AllHeaders = Replace(.AllHeaders, vbCrLf & " ", " ")
                .AllHeaders = Replace(.AllHeaders, vbCrLf & vbTab, vbTab)
            End If
            vSplit = Split(.AllHeaders, vbCrLf)
            lIdx = 0
            For Each vElem In vSplit
                vElem = Split(vElem, ":", Limit:=2)
                vElem(0) = Trim$(vElem(0))
                If UBound(vElem) > 0 Then
                    vElem(1) = Trim$(vElem(1))
                    If Not .Headers.Exists(vElem(0)) Then
                        .Headers.Item(vElem(0)) = vElem(1)
                    End If
                    If pvIsEqual(CStr(vElem(0)), HDR_SET_COOKIE) Then
                        With pvParseHeaderAttribs(vElem(1))
                            pvSetCookie .Item("domain"), .Item("path"), .Item(0)
                        End With
                    End If
                Else
                    .Headers.Item(lIdx) = vElem(0)
                End If
                lIdx = lIdx + 1
            Next
            .Status = Val(Mid$(vSplit(0), 1 + sizeof_HTTP_VER + 1, sizeof_HTTP_STATUS))
            .StatusText = Mid$(vSplit(0), 1 + sizeof_HTTP_VER + 1 + sizeof_HTTP_STATUS + 1)
            Select Case .Status
            Case 0
                pvSetError ERROR_WINHTTP_INVALID_SERVER_RESPONSE, MODULE_NAME & "." & FUNC_NAME
                GoTo QH
            Case 300 To 399
                If Not m_vOptions(WinHttpRequestOption_EnableRedirects) Or Not .Headers.Exists(HDR_LOCATION) Then
                    GoTo TryDownload
                End If
                sEscaped = pvEscapeUrl(.Headers.Item(HDR_LOCATION))
                m_vOptions(WinHttpRequestOption_URL) = sEscaped
                #If ImplUseDebugLog Then
                    DebugLog MODULE_NAME, FUNC_NAME, "Redirecting to " & sEscaped
                #End If
                With uRequest
                    .NumOfRedirects = .NumOfRedirects + 1
                    If .NumOfRedirects > m_vOptions(WinHttpRequestOption_MaxAutomaticRedirects) Then
                        GoTo TryDownload
                    End If
                    If Not pvParseUrl(sEscaped, uRedirect, .Remote.Protocol) Then
                        pvSetError ERROR_WINHTTP_INVALID_URL, MODULE_NAME & "." & FUNC_NAME
                        GoTo QH
                    End If
                    If LenB(uRedirect.Host) = 0 Then
                        uRedirect.Protocol = .Remote.Protocol
                        uRedirect.Host = .Remote.Host
                        uRedirect.Port = .Remote.Port
                    End If
                    If Not pvIsProtocolSupported(uRedirect.Protocol) Then
                        pvSetError ERROR_WINHTTP_UNRECOGNIZED_SCHEME, MODULE_NAME & "." & FUNC_NAME
                        GoTo QH
                    End If
                    If Not m_vOptions(WinHttpRequestOption_EnableHttpsToHttpRedirects) Then
                        If pvIsProtocolSecure(.Remote.Protocol) And Not pvIsProtocolSecure(uRedirect.Protocol) Then
                            GoTo TryDownload
                        End If
                    End If
                    If LenB(.CurrentProxy.Host) <> 0 And pvIsEqual(GetResponseHeader(HDR_PROXY_CONNECTION), STR_KEEP_ALIVE) And pvIsEqual(.Remote.Protocol, "http") Then
                        '--- keep proxy connection alive
                    ElseIf LenB(.CurrentProxy.Host) = 0 And pvIsEqual(GetResponseHeader(HDR_CONNECTION), STR_KEEP_ALIVE) And pvIsEqual(pvFormatHost(.Remote), pvFormatHost(uRedirect)) Then
                        '--- keep host connection alive
                    Else
                        m_oSocket.Close_
                        Set m_oSocket = pvCreateNewSocket
                    End If
                    .Remote = uRedirect
                    '--- prepare headers
                    SetRequestHeader HDR_HOST, .Remote.Host
                    SetRequestHeader HDR_COOKIE, pvGetAllCookies(.Remote.Host, .Remote.Path)
                    pvResetResponse uResponse
                    '--- try to reuse socket if not closed (keep-alive)
                    pvSetState ucsWaitResolve
                    If m_oSocket.IsClosed Then
                        If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
                            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Create"
                            GoTo QH
                        End If
                        If Not pvConnectRemote(uRequest) Then
                            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Connect"
                            GoTo QH
                        End If
                    Else
                        pvDrainKeepAlive
                        m_oSocket.PostEvent ucsSfdWrite
                    End If
                End With
            Case Else
TryDownload:
                If m_eState = ucsWaitRecvProxy Then
                    If .Status >= 200 And .Status <= 299 Then
                        #If ImplUseTls Then
                            If pvIsProtocolSecure(uRequest.Remote.Protocol) Then
                                If pvIsProtocolSecure(uRequest.CurrentProxy.Protocol) Then
                                    Dim oBaseSocket As cTlsSocket
                                    Set oBaseSocket = m_oSocket
                                    Set m_oSocket = pvCreateNewSocket
                                    Set m_oSocket.Socket = oBaseSocket
                                    Set oBaseSocket.CallbackWeakRef = m_oSocket
                                End If
                                m_oSocket.StartTls uRequest.Remote.Host, pvLocalFeatures
                            End If
                        #End If
                        pvResetResponse uResponse
                        pvSendHeaders uRequest
                    End If
                Else
                    Set uRequest.Stream = Nothing
                    uRequest.Headers.RemoveAll
                    If .Headers.Exists(HDR_CONTENT_LENGTH) Then
                        .ContentLength = Val(.Headers.Item(HDR_CONTENT_LENGTH))
                    End If
                    .TransferEncoding = GetResponseHeader("Transfer-Encoding")
                    .ContentEncoding = GetResponseHeader("Content-Encoding")
                    pvSetState ucsWaitRecvBody
                    lPos = lPos + sizeof_HDR_DELIM - 1
                    If .RecvBuffer.Size <= lPos Then
                        baData = vbNullString
                    Else
                        ReDim baData(0 To .RecvBuffer.Size - lPos - 1) As Byte
                        Call CopyMemory(baData(0), .RecvBuffer.Data(lPos), UBound(baData) + 1)
                    End If
                    RaiseEvent OnResponseStart(.Status, GetResponseHeader(HDR_CONTENT_TYPE))
                    .RecvBuffer.Size = 0
                    Set .Stream = SHCreateMemStream(ByVal 0, 0)
                    .Text = vbNullChar
                End If
            End Select
        ElseIf .RecvBuffer.Size > m_vOptions(WinHttpRequestOption_MaxResponseHeaderSize) Then
            pvSetError ERROR_WINHTTP_HEADER_SIZE_OVERFLOW, MODULE_NAME & "." & FUNC_NAME
            GoTo QH
        End If
    End With
    '--- success
    pvRecvHeaders = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvRecvBody(baData() As Byte, uResponse As UcsResponseType, Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvRecvBody"
    Dim hResult         As Long
    Dim baTemp()        As Byte
    Dim lPos            As Long
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim uChunk          As UcsBuffer
    Dim pOutput         As IUnknown
    
    On Error GoTo EH
    With uResponse
        If UBound(baData) >= 0 Then
            If pvIsEqual(.TransferEncoding, "chunked") Then
                pvBufferWriteArray .ChunksBuffer, baData
                pvParseChunks .ChunksBuffer, .TransferChunks, lSize
                For lIdx = 0 To lSize - 1
                    uChunk = .TransferChunks(lIdx)
                    lPos = uChunk.Pos + uChunk.Size
                    If uChunk.Size = 0 Then
                        pvSetState ucsIdle
                        Flush = True
                        Exit For
                    End If
                    ReDim baTemp(0 To uChunk.Size - 1) As Byte
                    Call CopyMemory(baTemp(0), .ChunksBuffer.Data(uChunk.Pos), UBound(baTemp) + 1)
                    .BytesProgress = .BytesProgress + UBound(baTemp) + 1
                    RaiseEvent OnResponseDataAvailable(baTemp)
                    pvBufferWriteArray .RecvBuffer, baTemp
                Next
                If lPos > 0 Then
                    lSize = .ChunksBuffer.Size - lPos
                    If lSize > 0 Then
                        Call CopyMemory(.ChunksBuffer.Data(0), .ChunksBuffer.Data(lPos), lSize)
                        .ChunksBuffer.Size = lSize
                    Else
                        .ChunksBuffer.Size = 0
                    End If
                End If
            Else
                .BytesProgress = .BytesProgress + UBound(baData) + 1
                RaiseEvent OnResponseDataAvailable(baData)
                pvBufferWriteArray .RecvBuffer, baData
                If .BytesProgress >= .ContentLength And .ContentLength >= 0 Then
                    pvSetState ucsIdle
                    Flush = True
                End If
            End If
        ElseIf .BytesProgress >= .ContentLength And .ContentLength >= 0 Then
            pvSetState ucsIdle
            Flush = True
        End If
        If .RecvBuffer.Size > 0 Then
            If .RecvBuffer.Size >= DEF_BUFFER_SIZE Or Flush Then
                If .RecvBuffer.Size < UBound(.RecvBuffer.Data) + 1 Then
                    ReDim Preserve .RecvBuffer.Data(0 To .RecvBuffer.Size - 1) As Byte
                End If
                hResult = IStream_Write(.Stream, .RecvBuffer.Data)
                .RecvBuffer.Size = 0
                If hResult < 0 Then
                    pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Write"
                    GoTo QH
                End If
            End If
        End If
        If Flush Then
            If pvIsEqual(.ContentEncoding, "gzip") And Not .Stream Is Nothing Then
                hResult = IStream_Seek(.Stream, 0, STREAM_SEEK_SET)
                If hResult < 0 Then
                    pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Seek(STREAM_SEEK_SET)"
                End If
                If pvUngzipStream(.Stream, pOutput) Then
                    Set .Stream = pOutput
                    .ContentEncoding = vbNullString
                End If
            End If
            RaiseEvent OnResponseFinished
        End If
    End With
    '--- success
    pvRecvBody = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvEscapeUrl(sUrl As String) As String
    Const URL_ESCAPE_PERCENT As Long = &H1000
    Dim sEncoded        As String
    Dim lFlags          As Long
    Dim vSplit          As Variant
    
    sEncoded = StrConv(pvToTextArray(sUrl, m_vOptions(WinHttpRequestOption_URLCodePage)), vbUnicode)
    If Not m_vOptions(WinHttpRequestOption_UrlEscapeDisable) Then
        If m_vOptions(WinHttpRequestOption_EscapePercentInURL) Then
            lFlags = lFlags Or URL_ESCAPE_PERCENT
        End If
        vSplit = Split(sEncoded, "?", Limit:=2)
        pvEscapeUrl = pvEscapePart(vSplit(0), lFlags)
        '--- append query string and anchor
        If UBound(vSplit) > 0 Then
            If Not m_vOptions(WinHttpRequestOption_UrlEscapeDisableQuery) Then
                pvEscapeUrl = pvEscapeUrl & "?" & pvEscapePart(vSplit(1), lFlags)
            Else
                pvEscapeUrl = pvEscapeUrl & "?" & vSplit(1)
            End If
        End If
    Else
        pvEscapeUrl = sEncoded
    End If
End Function

Private Function pvEscapePart(ByVal sPart As String, ByVal lFlags As Long) As String
    Dim sOutput         As String
    Dim lSize           As Long
    Dim hResult         As Long
    Dim baTemp()        As Byte
    
    lSize = Len(sPart)
    sOutput = String$(lSize, 0)
    hResult = UrlEscape(StrPtr(sPart), StrPtr(sOutput), lSize, lFlags)
    If hResult = E_POINTER Then
        sOutput = String$(lSize, 0)
        hResult = UrlEscape(StrPtr(sPart), StrPtr(sOutput), lSize, lFlags)
    End If
    If hResult < 0 Then
        Err.Raise hResult, "UrlEscape", m_oSocket.GetErrorDescription(hResult)
    End If
    baTemp = pvToTextArray(Left$(sOutput, lSize), ucsScpAcp)
    For lSize = 0 To UBound(baTemp)
        If baTemp(lSize) < &H80 Then
            pvEscapePart = pvEscapePart & Chr$(baTemp(lSize))
        Else
            pvEscapePart = pvEscapePart & "%" & Right$("0" & Hex$(baTemp(lSize)), 2)
        End If
    Next
End Function

Private Function pvCreateNewSocket() As Object
    #If ImplUseTls Then
        Set pvCreateNewSocket = New cTlsSocket
    #Else
        Set pvCreateNewSocket = New cAsyncSocket
    #End If
End Function

Private Function pvGetBaseSocket() As Object
    #If ImplUseTls Then
        Set pvGetBaseSocket = m_oSocket.Socket
    #Else
        Set pvGetBaseSocket = m_oSocket
    #End If
End Function

Private Function pvConnectRemote(uRequest As UcsRequestType) As Boolean
    Const FUNC_NAME     As String = "pvConnectRemote"
    Dim lIdx            As Long
    Dim uRemote         As UcsParsedUrl

    With uRequest
        .CurrentProxy.Host = vbNullString
        For lIdx = 0 To UBound(.ProxyList)
            If LenB(.ProxyList(lIdx).Host) <> 0 Then
                If (pvIsEqual(.ProxyList(lIdx).Anchor, .Remote.Protocol) Or LenB(.ProxyList(lIdx).Anchor) = 0) Then
                    If Not pvMatchBypassList(.Remote, .BypassList) Then
                        .CurrentProxy.Protocol = .ProxyList(lIdx).Protocol
                        .CurrentProxy.Host = .ProxyList(lIdx).Host
                        .CurrentProxy.Port = .ProxyList(lIdx).Port
                        If LenB(.ProxyList(lIdx).User) <> 0 Or LenB(.ProxyList(lIdx).Pass) <> 0 Then
                            .CurrentProxy.User = .ProxyList(lIdx).User
                            .CurrentProxy.Pass = .ProxyList(lIdx).Pass
                        End If
                    Else
                        #If ImplUseDebugLog Then
                            DebugLog MODULE_NAME, FUNC_NAME, "Bypass proxy for " & pvFormatHost(.Remote)
                        #End If
                    End If
                    Exit For
                End If
            End If
        Next
        If LenB(.CurrentProxy.Host) <> 0 Then
            uRemote = .CurrentProxy
        Else
            uRemote = .Remote
        End If
        #If ImplUseTls Then
            Dim oRootCa         As cTlsSocket
            
            Set oRootCa = m_vOptions(WinHttpRequestOption_RootCA)
            pvConnectRemote = m_oSocket.Connect(uRemote.Host, uRemote.Port, _
                UseTls:=pvIsProtocolSecure(uRemote.Protocol), LocalFeatures:=pvLocalFeatures, _
                RootCa:=oRootCa)
        #Else
            pvConnectRemote = m_oSocket.Connect(uRemote.Host, uRemote.Port)
        #End If
    End With
End Function

Private Function pvMatchBypassList(uRemote As UcsParsedUrl, vBypassList As Variant) As Boolean
    Dim vElem           As Variant
    
    If IsArray(vBypassList) Then
        For Each vElem In vBypassList
            If pvMatchRemote(uRemote, LCase$(vElem)) Then
                pvMatchBypassList = True
                Exit For
            End If
        Next
    End If
End Function

Private Function pvMatchRemote(uRemote As UcsParsedUrl, sBypass As String) As Boolean
    Dim sHost           As String
    
    If sBypass = "<local>" Then
        If InStr(uRemote.Host, ".") = 0 Then
            pvMatchRemote = True
        ElseIf uRemote.Host Like "192.168.*.*" Then
            pvMatchRemote = True
        ElseIf uRemote.Host Like "172.*.*.*" Then
            pvMatchRemote = True
        ElseIf uRemote.Host Like "10.*.*.*" Then
            pvMatchRemote = True
        End If
    Else
        sHost = LCase$(uRemote.Host)
        If sHost Like sBypass Then
            pvMatchRemote = True
        ElseIf "." & sHost Like sBypass Then
            pvMatchRemote = True
        ElseIf sHost & ":" & uRemote.Port Like sBypass Then
            pvMatchRemote = True
        ElseIf "." & sHost & ":" & uRemote.Port Like sBypass Then
            pvMatchRemote = True
        End If
    End If
End Function

Private Function pvIsProtocolSupported(sProtocol As String) As Boolean
    #If ImplUseTls Then
        pvIsProtocolSupported = (pvIsEqual(sProtocol, "http") Or pvIsEqual(sProtocol, "https"))
    #Else
        pvIsProtocolSupported = pvIsEqual(sProtocol, "http")
    #End If
End Function

Private Function pvIsProtocolSecure(sProtocol As String) As Boolean
    pvIsProtocolSecure = pvIsEqual(sProtocol, "https")
End Function

Private Function pvFormatHost(uRemote As UcsParsedUrl) As String
    If LenB(uRemote.Host) <> 0 Then
        pvFormatHost = uRemote.Protocol & "://" & uRemote.Host & IIf(LenB(uRemote.UserPort) <> 0, ":" & uRemote.UserPort, vbNullString)
    End If
End Function

Private Sub pvDrainKeepAlive()
    Dim baData()        As Byte
    Dim lTotal          As Long
    
    With pvGetBaseSocket
        .ReceiveArray baData
        lTotal = pvArraySize(baData)
        If lTotal > 0 Then
            Do While lTotal < m_vOptions(WinHttpRequestOption_MaxResponseDrainSize)
                If Not .SyncWaitForEvent(100, ucsSfdRead) Then
                    Exit Do
                End If
                .ReceiveArray baData
                lTotal = lTotal + pvArraySize(baData)
            Loop
        End If
    End With
End Sub

Private Function pvWaitForState(ByVal eState As UcsStateEnum, ByVal Timeout As Long) As Boolean
    Dim dblTimerEx      As Double
    Dim lRemaining      As Long
    
    dblTimerEx = TimerEx
    Do While m_eState <> ucsIdle And m_eState <= eState
        If Timeout >= 0 Then
            lRemaining = Timeout - Int((TimerEx - dblTimerEx) * 1000)
            If lRemaining <= 0 Then
                GoTo QH
            End If
        End If
        '--- note: call pvGetBaseSocket on each iteration because socket instance might change on redirect
        With pvGetBaseSocket
            If .HasPendingResolve Then
                If Not .SyncWaitForEvent(lRemaining, [_ucsSfdResolve]) Then
                    GoTo QH
                End If
            Else
                If Not .SyncWaitForEvent(lRemaining) Then
                    GoTo QH
                End If
            End If
            .SyncProcessMsgQueue
        End With
    Loop
    '--- success
    pvWaitForState = True
QH:
End Function

Private Sub pvResetResponse(uResponse As UcsResponseType)
    With uResponse
        .Status = 0
        .StatusText = vbNullString
        .AllHeaders = vbNullString
        .Headers.RemoveAll
        .BytesProgress = 0
        .ContentLength = -1
        .TransferEncoding = vbNullString
        .ContentEncoding = vbNullString
        .RecvBuffer.Size = 0
        .ChunksBuffer.Size = 0
        Set .Stream = Nothing
        .Text = vbNullChar
    End With
End Sub

Private Sub pvSetState(ByVal eValue As UcsStateEnum)
    m_eState = eValue
End Sub

Private Sub pvSetError(ByVal ErrNumber As Long, Optional ErrSource As String, Optional ErrDescription As String)
    Const FUNC_NAME     As String = "pvSetError"
    
    On Error GoTo EH
    If LenB(ErrDescription) = 0 Then
        If Not m_oSocket Is Nothing Then
            ErrDescription = m_oSocket.GetErrorDescription(ErrNumber)
        Else
            ErrDescription = "Error " & ErrNumber
        End If
    End If
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, "pvSetError", ErrDescription & " &H" & Hex$(ErrNumber) & " [" & Replace(ErrSource, vbCrLf, "; ") & "]"
    #End If
    m_vLastError = Array(ErrNumber, ErrSource, ErrDescription)
    RaiseEvent OnError(ErrNumber, ErrDescription & IIf(LenB(ErrSource) <> 0, " [" & Replace(ErrSource, vbCrLf, "; ") & "]", vbNullString))
    Abort
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvJoinHeaders(oDict As Object) As String
    Const FUNC_NAME     As String = "pvJoinHeaders"
    Dim vSplit          As Variant
    Dim vKey            As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = oDict.Keys
    For Each vKey In vSplit
        vSplit(lIdx) = vSplit(lIdx) & ": " & oDict.Item(vKey)
        lIdx = lIdx + 1
    Next
    pvJoinHeaders = Join(vSplit, vbCrLf)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvParseChunks(uBuffer As UcsBuffer, aOutput() As UcsBuffer, lOutSize As Long)
    Const FUNC_NAME     As String = "pvParseChunks"
    Const sizeof_CrLf   As Long = 2
    Dim lPtr            As Long
    Dim lStart          As Long
    Dim lPos            As Long
    Dim sTemp           As String
    Dim lSize           As Long
    
    On Error GoTo EH
    Call CopyMemory(lPtr, ByVal ArrPtr(aOutput), 4)
    If lPtr = 0 Then
        ReDim aOutput(0 To 7) As UcsBuffer
    End If
    lStart = 1
    lPos = InStrB(lStart, uBuffer.Data, StrConv(vbCrLf, vbFromUnicode))
    If lPos = 1 Then
        lStart = 3
        lPos = InStrB(lStart, uBuffer.Data, StrConv(vbCrLf, vbFromUnicode))
    End If
    Do While lPos > 0 And lPos < uBuffer.Size
        sTemp = String$(lPos - lStart, 0)
        Call CopyMemory(ByVal sTemp, uBuffer.Data(lStart - 1), Len(sTemp))
        lSize = Val("&H" & sTemp & "&")
        If lPos + sizeof_CrLf + lSize > uBuffer.Size Then
            Exit Do
        End If
        If UBound(aOutput) < lOutSize Then
            ReDim Preserve aOutput(0 To 2 * UBound(aOutput) + 1) As UcsBuffer
        End If
        With aOutput(lOutSize)
            .Pos = lPos + sizeof_CrLf - 1
            .Size = lSize
        End With
        lOutSize = lOutSize + 1
        If lSize = 0 Then
            Exit Do
        End If
        lStart = lPos + sizeof_CrLf + lSize + sizeof_CrLf
        lPos = InStrB(lStart, uBuffer.Data, StrConv(vbCrLf, vbFromUnicode))
    Loop
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Pattern = "^(?:(?:(.+?):)?//)?(?:(?:([^:]*?):)?([^@]*)@)?([A-Za-z0-9\-\.]+)?(:[0-9]+)?(/[^?#]*?)?(\?[^#]*?)?(#.*)?$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = IIf(LenB(.Item(0)) = 0, DefProtocol, .Item(0))
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.UserPort = Mid$(.Item(4), 2)
                    uParsed.Port = Val(uParsed.UserPort)
                    If uParsed.Port = 0 Then
                        Select Case LCase$(uParsed.Protocol)
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                    uParsed.QueryString = .Item(6)
                    uParsed.Anchor = .Item(7)
                End With
                pvParseUrl = True
            End If
        End With
    End With
End Function

Private Function pvParseHeaderAttribs(ByVal sHeader As String) As Object
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim sKey            As String
    Dim sValue          As String
    Dim lIdx            As Long
    
    '--- TODO: need regex parser for semi-colons in quoted values
    Set pvParseHeaderAttribs = CreateObject("Scripting.Dictionary")
    m_uRequest.Headers.CompareMode = vbTextCompare
    vSplit = Split(sHeader, ";")
    If UBound(vSplit) > 0 Then
        For Each vElem In vSplit
            pvParseHeaderAttribs.Item(lIdx) = Trim$(vElem)
            vElem = Split(Trim$(vElem), "=", Limit:=2)
            sKey = Trim$(vElem(0))
            If Left$(sKey, 1) = """" And Right$(sKey, 1) = """" And Len(sKey) > 1 Then
                sKey = Trim$(Mid$(sKey, 2, Len(sKey) - 2))
            End If
            If UBound(vElem) > 0 Then
                sValue = Trim$(vElem(1))
                If Left$(sValue, 1) = """" And Right$(sValue, 1) = """" And Len(sValue) > 1 Then
                    sValue = Trim$(Mid$(sValue, 2, Len(sValue) - 2))
                End If
            Else
                sValue = vbNullString
            End If
            If Not pvParseHeaderAttribs.Exists(sKey) Then
                pvParseHeaderAttribs.Item(sKey) = sValue
            End If
            lIdx = lIdx + 1
        Next
    End If
End Function

Private Function pvMapCharset(sCharset As String) As UcsAsyncSocketCodePageEnum
    Dim vElem           As Variant
    
    ' https://www.iana.org/assignments/character-sets/character-sets.xhtml
    ' https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
    Select Case LCase$(sCharset)
    Case "utf-8"
        pvMapCharset = ucsScpUtf8
    Case "utf-7"
        pvMapCharset = 65000
    Case "utf-16"
        pvMapCharset = 1200
    Case "utf-32"
        pvMapCharset = 12000
    Case "iso-8859-1", "latin1"
        pvMapCharset = 28591
    Case "iso-8859-2", "latin2"
        pvMapCharset = 28592
    Case "iso-8859-3", "latin3"
        pvMapCharset = 28593
    Case "iso-8859-4", "latin4"
        pvMapCharset = 28594
    Case "iso-8859-5", "cyrillic"
        pvMapCharset = 28595
    Case "iso-8859-6", "arabic"
        pvMapCharset = 28596
    Case "iso-8859-7", "greek"
        pvMapCharset = 28597
    Case "iso-8859-8", "hebrew"
        pvMapCharset = 28598
    Case "iso-8859-9", "latin5"
        pvMapCharset = 28599
    Case "iso-8859-13"
        pvMapCharset = 28603
    Case "iso-8859-15", "latin-9"
        pvMapCharset = 28603
    Case "koi8-r"
        pvMapCharset = 28605
    Case "koi8-u"
        pvMapCharset = 21866
    Case "us-ascii"
        pvMapCharset = 20127
    Case "shift_jis"
        pvMapCharset = 932
    Case "gb2312"
        pvMapCharset = 936
    Case "ks_c_5601-1987"
        pvMapCharset = 949
    Case "big5"
        pvMapCharset = 950
    Case "cp1025"
        pvMapCharset = 21025
    Case Else
        For Each vElem In Split("windows- dos- cp x-cp ibm")
            If pvIsEqual(Left$(sCharset, Len(vElem)), CStr(vElem)) Then
                pvMapCharset = Val(Mid$(sCharset, Len(vElem) + 1))
                If pvMapCharset <> 0 Then
                    Exit Function
                End If
            End If
        Next
        '--- fallback
        pvMapCharset = ucsScpAcp
    End Select
End Function

Private Function pvUngzipStream(pInput As IUnknown, pOutput As IUnknown) As Boolean
    Const FUNC_NAME     As String = "pvUngzipStream"
    Const BUFF_SIZE     As Long = 65536
    Dim baData()        As Byte
    Dim hArchive        As Long
    Dim lSize           As Long
    Dim lResult         As Long
    Dim lReaderPfn      As Long
    Dim uReaderThunk    As UcsRedirectReaderThunk
    Dim hResult         As Long
        
    On Error GoTo EH
    ReDim m_uResponse.ContentData(0 To BUFF_SIZE - 1) As Byte
    hArchive = archive_read_new()
    If hArchive = 0 Then
        GoTo QH
    End If
    lResult = archive_read_support_filter_gzip(hArchive)
    If lResult <> 0 Then
        GoTo QH
    End If
    lResult = archive_read_support_format_raw(hArchive)
    If lResult <> 0 Then
        GoTo QH
    End If
    lReaderPfn = pvAddressOfReaderCallback.ArchiveReaderCallback(0, Nothing, 0)
    lReaderPfn = InitRedirectReaderThunk(uReaderThunk, ObjPtr(Me), lReaderPfn)
    lResult = archive_read_open(hArchive, ObjPtr(pInput), 0, lReaderPfn, 0)
    If lResult <> 0 Then
        GoTo QH
    End If
    lResult = archive_read_next_header(hArchive, 0)
    If lResult <> 0 Then
        GoTo QH
    End If
    ReDim baData(0 To BUFF_SIZE - 1) As Byte
    Do
        lSize = archive_read_data(hArchive, baData(0), UBound(baData) + 1)
        If lSize < 0 Then
            pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "archive_read_data", "Decompression error (" & lSize & ")"
            GoTo QH
        End If
        If lSize = 0 Then
            Exit Do
        End If
        If UBound(baData) + 1 <> lSize Then
            ReDim Preserve baData(0 To lSize - 1) As Byte
        End If
        If pOutput Is Nothing Then
            Set pOutput = SHCreateMemStream(baData(0), UBound(baData) + 1)
            If pOutput Is Nothing Then
                pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "SHCreateMemStream", "Out of memory"
                GoTo QH
            End If
            hResult = IStream_Seek(pOutput, 0, STREAM_SEEK_END, 0)
            If hResult < 0 Then
                pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Seek"
                GoTo QH
            End If
        Else
            hResult = IStream_Write(pOutput, baData)
            If hResult < 0 Then
                pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Write"
                GoTo QH
            End If
        End If
    Loop
    '--- success
    pvUngzipStream = True
QH:
    If hArchive <> 0 Then
        Call archive_read_free(hArchive)
    End If
    Erase m_uResponse.ContentData
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ArchiveReaderCallback(ByVal hArchive As Long, ByVal pInput As IUnknown, lBufferPtr As Long) As Long
Attribute ArchiveReaderCallback.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "ArchiveReaderCallback"
    Dim hResult         As Long
    Dim lSize           As Long

    On Error GoTo EH
    #If hArchive Then '--- touch args
    #End If
    hResult = IStream_Read(pInput, m_uResponse.ContentData, lSize)
    If hResult < 0 Then
        GoTo QH
    End If
    If lSize > 0 Then
        lBufferPtr = VarPtr(m_uResponse.ContentData(0))
        '--- success
        ArchiveReaderCallback = lSize
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function InitRedirectReaderThunk(uThunk As UcsRedirectReaderThunk, ByVal ParamValue As Long, ByVal pfnDest As Long) As Long
    ' 54                   push        esp
    ' 54                   push        esp
    ' 33 C0                xor         eax,eax
    ' B0 14                mov         al,14h
    ' FF 34 04             push        dword ptr [esp+eax]
    ' FF 34 04             push        dword ptr [esp+eax]
    ' FF 34 04             push        dword ptr [esp+eax]
    ' B8 XX XX XX XX       mov         eax,XXXXXXXXXX
    ' 50                   push        eax
    ' 90                   nop
    ' 90                   nop
    ' B8 YY YY YY YY       mov         eax,YYYYYYYYYY
    ' FF D0                call        eax
    ' 58                   pop         eax
    ' C2 0C 00             ret         0Ch
    ' 90                   nop
    ' 90                   nop
    With uThunk
        .Code(0) = &HC0335454
        .Code(1) = &H34FF14B0
        .Code(2) = &H434FF04
        .Code(3) = &HB80434FF
        .Code(4) = ParamValue
        .Code(5) = &HB8909050
        .Code(6) = pfnDest
        .Code(7) = &HC258D0FF
        .Code(8) = &H9090000C
        Call VirtualProtect(.Code(0), Len(uThunk), PAGE_EXECUTE_READWRITE, 0)
    End With
    InitRedirectReaderThunk = VarPtr(uThunk.Code(0))
End Function

Private Function pvToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    pvToTextArray = m_oSocket.ToTextArray(sText, CodePage)
End Function

Private Function pvFromTextArray(baText() As Byte, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    pvFromTextArray = m_oSocket.FromTextArray(baText, CodePage)
End Function

Private Function pvToStringAndFree(ByVal lPtr As Long) As Variant
    If lPtr Then
        pvToStringAndFree = String$(lstrlenW(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvToStringAndFree), ByVal lPtr, LenB(pvToStringAndFree))
        Call GlobalFree(lPtr)
    End If
End Function

Private Function pvSetCookie(ByVal sDomain As String, ByVal sPath As String, ByVal sValue As String)
    Dim oCol            As Collection
    
    If Left$(sDomain, 1) = "." Then
        sDomain = Mid$(sDomain, 2)
    End If
    If Right$(sDomain, 1) = "." Then
        sDomain = Left$(sDomain, Len(sDomain) - 1)
    End If
    If LenB(sDomain) = 0 Then
        sDomain = m_uRequest.Remote.Host
    End If
    If Left$(sPath, 1) <> "/" Then
        sPath = m_uRequest.Remote.Path
    End If
    If Right$(sPath, 1) <> "/" Then
        sPath = sPath & "/"
    End If
    If m_oCookies.Exists(sDomain) Then
        Set oCol = m_oCookies.Item(sDomain)
    Else
        Set oCol = New Collection
        Set m_oCookies.Item(sDomain) = oCol
    End If
    oCol.Add sPath & Chr$(1) & sValue
End Function

Private Function pvGetAllCookies(ByVal sDomain As String, ByVal sPath As String) As String
    Dim vRetVal         As Variant
    Dim oCol            As Collection
    Dim vKey            As Variant
    
    vRetVal = Array()
    If Right$(sDomain, 1) = "." Then
        sDomain = Left$(sDomain, Len(sDomain) - 1)
    End If
    If Left$(sPath, 1) <> "/" Then
        sPath = "/"
    End If
    Do While InStr(sDomain, ".") <> 0
        If m_oCookies.Exists(sDomain) Then
            Set oCol = m_oCookies.Item(sDomain)
            For Each vKey In oCol
                vKey = Split(vKey, Chr$(1), Limit:=2)
                If Left$(sPath, Len(vKey(0))) = vKey(0) Then
                    ReDim Preserve vRetVal(0 To UBound(vRetVal) + 1) As Variant
                    vRetVal(UBound(vRetVal)) = vKey(1)
                End If
            Next
        End If
        sDomain = Mid$(sDomain, InStr(sDomain, ".") + 1)
    Loop
    pvGetAllCookies = Join(vRetVal, "; ")
End Function

Private Function pvIsEqual(sFirst As String, sSecond As String) As Boolean
    If Len(sFirst) = Len(sSecond) Then
        pvIsEqual = (StrComp(sFirst, sSecond, vbTextCompare) = 0)
    End If
End Function

'= buffers ===============================================================

Private Function AsIStream(ByVal pUnk As IUnknown) As IUnknown
    Const IDX_QueryInterface As Long = 0
    Static IID_IStream(0 To 3) As Long
    Dim hResult         As Long
    
    If IID_IStream(0) = 0 Then
        IID_IStream(0) = &HC: IID_IStream(1) = &H0
        IID_IStream(2) = &HC0: IID_IStream(3) = &H46000000
    End If
    If Not pUnk Is Nothing Then
        hResult = DispCallByVtbl(pUnk, IDX_QueryInterface, VarPtr(IID_IStream(0)), VarPtr(AsIStream))
        If hResult < 0 Then
            Err.Raise hResult, "IUnknown.QueryInterface(IID_IStream)"
        End If
    End If
End Function

Private Function IStream_Read(pStm As IUnknown, baData() As Byte, Optional BytesRead As Long) As Long
    If Not pStm Is Nothing And UBound(baData) >= 0 Then
        IStream_Read = DispCallByVtbl(pStm, 3, VarPtr(baData(0)), UBound(baData) + 1, VarPtr(BytesRead))
    End If
End Function

Private Function IStream_Write(pStm As IUnknown, baData() As Byte, Optional BytesWritten As Long) As Long
    If Not pStm Is Nothing And UBound(baData) >= 0 Then
        IStream_Write = DispCallByVtbl(pStm, 4, VarPtr(baData(0)), UBound(baData) + 1, VarPtr(BytesWritten))
    End If
End Function

Private Function IStream_Seek(pStm As stdole.IUnknown, ByVal cMove As Currency, ByVal dwOrigin As Long, Optional NewPosition As Currency) As Long
    If Not pStm Is Nothing Then
        IStream_Seek = DispCallByVtbl(pStm, 5, cMove, dwOrigin, VarPtr(NewPosition))
    Else
        IStream_Seek = -1
    End If
End Function

Private Function IStream_GetSize(pStream As IUnknown) As Double
    Dim hResult         As Long
    Dim cPos            As Currency
    
    If Not pStream Is Nothing Then
        hResult = IStream_Seek(pStream, 0, STREAM_SEEK_END, cPos)
        If hResult < 0 Then
            Err.Raise hResult, "IStream_Seek(STREAM_SEEK_END)"
        End If
        hResult = IStream_Seek(pStream, 0, STREAM_SEEK_SET)
        If hResult < 0 Then
            Err.Raise hResult, "IStream_Seek(STREAM_SEEK_SET)"
        End If
    End If
    IStream_GetSize = CDbl(cPos) * 10000#
End Function

Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Private Sub pvBufferWriteArray(uOutput As UcsBuffer, baSrc() As Byte)
    Dim lSize       As Long
    
    With uOutput
        lSize = pvArraySize(baSrc)
        If lSize > 0 Then
            .Size = pvArrayWriteBlob(.Data, .Size, VarPtr(baSrc(0)), lSize)
        End If
    End With
End Sub

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long

    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

Private Function pvArrayWriteBlob(baArray() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baArray), 4)
    If lBufPtr = 0 Then
        ReDim baArray(0 To lPos + lSize - 1) As Byte
    ElseIf UBound(baArray) < lPos + lSize - 1 Then
        ReDim Preserve baArray(0 To lPos + lSize - 1) As Byte
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baArray(lPos), ByVal lPtr, lSize)
    End If
    pvArrayWriteBlob = lPos + lSize
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

Private Function ToBase64Array(baData() As Byte) As String
    With CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .NodeTypedValue = baData
        ToBase64Array = .Text
    End With
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
#End If

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnResolve(IpAddress As String)
    If m_eState = ucsWaitResolve Then
        pvSetState ucsWaitConnect
    End If
End Sub

Private Sub m_oSocket_OnConnect()
    m_oSocket_OnSend
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    
    On Error GoTo EH
    If m_eState = ucsWaitResolve Or m_eState = ucsWaitConnect Then
        If Not pvSendHeaders(m_uRequest) Then
            GoTo QH
        End If
    End If
    If m_eState = ucsWaitSendProxy Or m_eState = ucsWaitSendBody Then
        If Not pvSendBody(m_uRequest) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baData()        As Byte
    
    On Error GoTo EH
    If Not m_oSocket.ReceiveArray(baData) Then
        If m_uResponse.BytesProgress >= m_uResponse.ContentLength And m_uResponse.ContentLength >= 0 Then
            GoTo QH
        End If
        If m_oSocket.LastError <> 0 Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
        End If
        GoTo QH
    End If
    If UBound(baData) < 0 Then
        GoTo QH
    End If
    If m_eState = ucsWaitRecvProxy Or m_eState = ucsWaitRecvHeaders Then
        If Not pvRecvHeaders(baData, m_uResponse, m_uRequest) Then
            GoTo QH
        End If
    End If
    If m_eState = ucsWaitRecvBody Then
        If Not pvRecvBody(baData, m_uResponse) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    Dim baData()        As Byte
    Dim dblBytes        As Double
    
    On Error GoTo EH
    If m_eState = ucsWaitRecvBody Then
        Do
            dblBytes = m_uResponse.BytesProgress
            m_oSocket_OnReceive
        Loop While m_uResponse.BytesProgress > dblBytes
        baData = vbNullString
        pvRecvBody baData, m_uResponse, Flush:=True
    End If
    pvSetState ucsIdle
    m_oSocket.Close_
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    Const WSAHOST_NOT_FOUND As Long = 11001
    
    On Error GoTo EH
    With m_oSocket.LastError
        Select Case .Number
        Case WSAHOST_NOT_FOUND
            pvSetError ERROR_WINHTTP_NAME_NOT_RESOLVED
        Case Is <> 0
            pvSetError .Number, .Source, .Description
        End Select
    End With
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

#If ImplUseTls Then
Private Sub m_oSocket_OnCertificate(Issuers As Object, Confirmed As Boolean)
    Const FUNC_NAME     As String = "m_oSocket_OnCertificate"
    
    On Error GoTo EH
    '--- try to match a client certificate in CURRENT_USER\MY certificate store by Issuer's RDN
    If Issuers.Count > 0 Then
        If m_oSocket.ImportSystemStoreCertificates(Issuers) Then
            Confirmed = True
            GoTo QH
        End If
    End If
    If LenB(m_vOptions(WinHttpRequestOption_SelectCertificate)) <> 0 Then
        If m_oSocket.ImportSystemStoreCertificates(m_vOptions(WinHttpRequestOption_SelectCertificate)) Then
            Confirmed = True
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_oSocket = pvCreateNewSocket
    m_vOptions(WinHttpRequestOption_UserAgentString) = DEF_USER_AGENT
    m_vOptions(WinHttpRequestOption_URLCodePage) = CLng(65001)
    m_vOptions(WinHttpRequestOption_EscapePercentInURL) = False
    m_vOptions(WinHttpRequestOption_SslErrorIgnoreFlags) = 0
    m_vOptions(WinHttpRequestOption_SelectCertificate) = vbNullString
    m_vOptions(WinHttpRequestOption_EnableRedirects) = True
    m_vOptions(WinHttpRequestOption_UrlEscapeDisable) = False
    m_vOptions(WinHttpRequestOption_UrlEscapeDisableQuery) = True
    m_vOptions(WinHttpRequestOption_SecureProtocols) = SecureProtocol_ALL
    m_vOptions(WinHttpRequestOption_EnableTracing) = False                     '--- ignored
    m_vOptions(WinHttpRequestOption_RevertImpersonationOverSsl) = True         '--- ignored
    m_vOptions(WinHttpRequestOption_EnableHttpsToHttpRedirects) = False
    m_vOptions(WinHttpRequestOption_EnablePassportAuthentication) = False      '--- ignored
    m_vOptions(WinHttpRequestOption_MaxAutomaticRedirects) = CLng(10)
    m_vOptions(WinHttpRequestOption_MaxResponseHeaderSize) = CLng(65536)
    m_vOptions(WinHttpRequestOption_MaxResponseDrainSize) = CLng(1024000)
    m_vOptions(WinHttpRequestOption_EnableHttp1_1) = True
    m_vOptions(WinHttpRequestOption_EnableCertificateRevocationCheck) = False
    m_vOptions(WinHttpRequestOption_RejectUserpwd) = False
    Set m_vOptions(WinHttpRequestOption_RootCA) = Nothing
    SetTimeouts 5000, 5000, 15000, 15000
    Set m_uRequest.Headers = CreateObject("Scripting.Dictionary")
    m_uRequest.Headers.CompareMode = vbTextCompare
    Set m_uResponse.Headers = CreateObject("Scripting.Dictionary")
    m_uResponse.Headers.CompareMode = vbTextCompare
    Set m_oCookies = CreateObject("Scripting.Dictionary")
    m_oCookies.CompareMode = vbTextCompare
    ReDim m_uRequest.ProxyList(-1 To -1) As UcsParsedUrl
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
End Sub
