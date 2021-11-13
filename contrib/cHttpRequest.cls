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
' VbAsyncSocket Project (c) 2018-2021 by wqweto@gmail.com
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
End Enum

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
'--- oleaut32
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
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
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    '--- end MST
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_HOST              As String = "Host"
Private Const HDR_USER_AGENT        As String = "User-Agent"
Private Const HDR_CONTENT_TYPE      As String = "Content-Type"
Private Const HDR_CONTENT_LENGTH    As String = "Content-Length"
Private Const HDR_LOCATION          As String = "Location"
Private Const HDR_DELIM             As String = vbCrLf & vbCrLf
Private Const DEF_BUFFER_SIZE       As Long = 256& * 1024
Private Const DEF_USER_AGENT        As String = "Mozilla/4.0 (compatible; Win32; VbAsyncSocket)"
Private Const DEF_CONTENT_TYPE      As String = "text/plain"
Private Const DEF_CHARSET           As String = "charset=utf-8"

#If ImplUseTls Then
    Private WithEvents m_oSocket As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
#Else
    Private WithEvents m_oSocket As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
#End If
Private m_vLastError            As Variant
Private m_lResolveTimeout       As Long
Private m_lConnectTimeout       As Long
Private m_lSendTimeout          As Long
Private m_lReceiveTimeout       As Long
Private m_eState                As UcsStateEnum
Private m_sMethod               As String
Private m_uRemote               As UcsParsedUrl
Private m_bAsync                As Boolean
Private m_oRequestHeaders       As Object
Private m_pRequestStream        As IUnknown
Private m_vOptions(0 To WinHttpRequestOption_RejectUserpwd) As Variant
Private m_lStatus               As Long
Private m_sStatusText           As String
Private m_sAllResponseHeaders   As String
Private m_oResponseHeaders      As Object
Private m_pResponseStream       As IUnknown
Private m_sResponseText         As String
Private m_dblContentLength      As Double
Private m_dblBytesProgress      As Double
Private m_lCurrentRedirect      As Long
Private m_sTransferEncoding     As String
Private m_aTransferChunks()     As UcsBuffer
Private m_sContentEncoding      As String
'--- I/O buffers
Private m_uSendBuffer           As UcsBuffer
Private m_uRecvBuffer           As UcsBuffer
Private m_uChunkBuffer          As UcsBuffer
Private m_baGzipBuffer()        As Byte

Private Enum UcsStateEnum
    ucsIdle
    ucsWaitResolve
    ucsWaitConnect
    ucsWaitSendBody
    ucsWaitRecvHeaders
    ucsWaitRecvBody
End Enum

Private Type UcsParsedUrl
    Protocol            As String
    Host                As String
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
    Set ResponseStream = m_pResponseStream
End Property

Public Property Get ResponseBody() As Byte()
    Const FUNC_NAME     As String = "ResponseBody"
    Dim lSize           As Long
    Dim baRetVal()      As Byte
    Dim hResult         As Long
    
    On Error GoTo EH
    m_vLastError = Empty
    lSize = pvGetStreamSize(m_pResponseStream)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        hResult = IStream_Read(m_pResponseStream, baRetVal)
        If hResult < 0 Then
            pvSetError hResult, "IStream_Read"
            GoTo QH
        End If
    Else
        baRetVal = vbNullString
    End If
    ResponseBody = baRetVal
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
    
    If m_sResponseText = vbNullChar Then
        eCodePage = ucsScpUtf8
        If m_oResponseHeaders.Exists(HDR_CONTENT_TYPE) Then
            sCharset = pvGetHeaderAttrib(m_oResponseHeaders.Item(HDR_CONTENT_TYPE), "charset")
            If LenB(sCharset) <> 0 Then
                eCodePage = pvMapCharset(sCharset)
            End If
        End If
        m_sResponseText = m_oSocket.FromTextArray(ResponseBody, eCodePage)
    End If
    ResponseText = m_sResponseText
End Property

Public Property Get Status() As Long
    Status = m_lStatus
End Property

Public Property Get StatusText() As String
    StatusText = m_sStatusText
End Property

Public Property Get Option_(ByVal eOption As WinHttpRequestOption) As Variant
    Option_ = m_vOptions(eOption)
End Property

Public Property Let Option_(ByVal eOption As WinHttpRequestOption, ByVal RHS As Variant)
    Select Case VarType(m_vOptions(eOption))
    Case vbBoolean
        RHS = CBool(RHS)
    Case vbLong
        RHS = CLng(RHS)
    Case vbString
        RHS = CStr(RHS)
    End Select
    m_vOptions(eOption) = RHS
End Property

Private Property Get pvAddressOfReaderCallback() As cHttpRequest
    Set pvAddressOfReaderCallback = InitAddressOfMethod(Me, 3)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Open_(ByVal Method As String, ByVal Url As String, Optional ByVal Async As Variant)
    Const FUNC_NAME     As String = "Open"
    Dim sEscaped        As String
    Dim uRemote         As UcsParsedUrl
    
    On Error GoTo EH
    m_vLastError = Empty
    sEscaped = pvEscapeUrl(Url)
    Option_(WinHttpRequestOption_URL) = sEscaped
    If Not pvParseUrl(sEscaped, uRemote, "http") Then
        pvSetError ERROR_WINHTTP_INVALID_URL
        GoTo QH
    End If
    If Not pvIsProtocolSupported(uRemote.Protocol) Then
        pvSetError ERROR_WINHTTP_UNRECOGNIZED_SCHEME
        GoTo QH
    End If
    If Option_(WinHttpRequestOption_RejectUserpwd) Then
        uRemote.User = vbNullString
        uRemote.Pass = vbNullString
    End If
    m_sMethod = Method
    If LenB(m_uRemote.Host) <> 0 Then
        If m_uRemote.Host <> uRemote.Host Or m_uRemote.Port <> uRemote.Port Or m_uRemote.Protocol <> uRemote.Protocol Then
            Set m_oSocket = pvCreateNewSocket
        End If
    End If
    m_uRemote = uRemote
    m_bAsync = IIf(Not IsMissing(Async), Async, False)
    Set m_oRequestHeaders = CreateObject("Scripting.Dictionary")
    m_oRequestHeaders.CompareMode = vbTextCompare
    Set m_pRequestStream = Nothing
    m_lStatus = 0
    m_sStatusText = vbNullString
    m_sAllResponseHeaders = vbNullString
    Set m_oResponseHeaders = CreateObject("Scripting.Dictionary")
    m_oResponseHeaders.CompareMode = vbTextCompare
    m_dblBytesProgress = 0
    m_dblContentLength = -1
    m_lCurrentRedirect = 0
    m_sTransferEncoding = vbNullString
    m_sContentEncoding = vbNullString
    '--- init headers
    SetRequestHeader HDR_HOST, m_uRemote.Host
    SetRequestHeader HDR_USER_AGENT, Option_(WinHttpRequestOption_UserAgentString)
    If Option_(WinHttpRequestOption_EnableHttp1_1) Then
        SetRequestHeader "Connection", "Keep-Alive"
    End If
    SetRequestHeader "Accept", "*/*"
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
    Dim baBuffer()      As Byte
    Dim eCodePage       As UcsAsyncSocketCodePageEnum
    Dim sCharset        As String
    
    On Error GoTo EH
    m_vLastError = Empty
    '--- preapare request stream from body
    If IsObject(Body) Then
        Set m_pRequestStream = AsIStream(Body)
    ElseIf IsMissing(Body) Or IsEmpty(Body) Or IsNull(Body) Then
        Set m_pRequestStream = Nothing
    ElseIf VarType(Body) = vbByte + vbArray Then
        If UBound(Body) >= 0 Then
            Set m_pRequestStream = SHCreateMemStream(Body(0), UBound(Body) + 1)
        Else
            Set m_pRequestStream = Nothing
        End If
    ElseIf VarType(Body) = vbString Then
        If LenB(Body) <> 0 Then
            eCodePage = ucsScpUtf8
            If Not m_oRequestHeaders.Exists(HDR_CONTENT_TYPE) Then
                SetRequestHeader HDR_CONTENT_TYPE, DEF_CONTENT_TYPE & "; " & DEF_CHARSET
            Else
                sCharset = pvGetHeaderAttrib(m_oRequestHeaders.Item(HDR_CONTENT_TYPE), "charset")
                If LenB(sCharset) <> 0 Then
                    eCodePage = pvMapCharset(sCharset)
                Else
                    SetRequestHeader HDR_CONTENT_TYPE, m_oRequestHeaders.Item(HDR_CONTENT_TYPE) & "; " & DEF_CHARSET
                End If
            End If
            baBuffer = pvToTextArray(CStr(Body), eCodePage)
            Set m_pRequestStream = SHCreateMemStream(baBuffer(0), UBound(baBuffer) + 1)
        Else
            Set m_pRequestStream = Nothing
        End If
    Else
        pvSetError ERROR_INVALID_PARAMETER
        GoTo QH
    End If
    '--- prepare headers
    If Not m_pRequestStream Is Nothing Then
        SetRequestHeader HDR_CONTENT_LENGTH, pvGetStreamSize(m_pRequestStream)
    End If
    If LenB(m_uRemote.User) <> 0 Or LenB(m_uRemote.Pass) <> 0 Then
        SetRequestHeader "Authorization", "Basic " & Replace(ToBase64Array( _
            pvToTextArray(m_uRemote.User & ":" & m_uRemote.Pass, ucsScpAcp)), vbCrLf, vbNullString)
    End If
    SetRequestHeader HDR_HOST, m_uRemote.Host
    '--- try to reuse socket if not closed (keep-alive)
    m_eState = ucsWaitResolve
    If m_oSocket.IsClosed Then
        If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
            pvSetError m_oSocket.LastError, "m_oSocket.Create"
            GoTo QH
        End If
        If Not pvConnectRemote(m_uRemote) Then
            pvSetError m_oSocket.LastError, "m_oSocket.Connect"
            GoTo QH
        End If
    Else
        pvDrainKeepAlive
        m_oSocket.PostEvent ucsSfdWrite
    End If
    '--- handle synchronous mode
    If Not m_bAsync Then
        If Not pvWaitForState(ucsWaitResolve, m_lResolveTimeout) Then
            pvSetError ERROR_WINHTTP_TIMEOUT
            GoTo QH
        End If
        If Not pvWaitForState(ucsWaitConnect, m_lConnectTimeout) Then
            pvSetError ERROR_WINHTTP_TIMEOUT
            GoTo QH
        End If
        If Not pvWaitForState(ucsWaitSendBody, m_lSendTimeout) Then
            pvSetError ERROR_WINHTTP_TIMEOUT
            GoTo QH
        End If
        If Not pvWaitForState(ucsWaitRecvBody, m_lReceiveTimeout) Then
            pvSetError ERROR_WINHTTP_TIMEOUT
            GoTo QH
        End If
    End If
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
    m_lResolveTimeout = ResolveTimeout
    m_lConnectTimeout = ConnectTimeout
    m_lSendTimeout = SendTimeout
    m_lReceiveTimeout = ReceiveTimeout
End Sub

Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)
    If LenB(Value) <> 0 Then
        m_oRequestHeaders.Item(Header) = Value
    Else
        m_oRequestHeaders.Remove Header
    End If
End Sub

Public Sub Abort()
    m_eState = ucsIdle
    Set m_oSocket = pvCreateNewSocket
    Set m_pRequestStream = Nothing
    Set m_pResponseStream = Nothing
    m_sResponseText = vbNullChar
End Sub

Public Function GetAllResponseHeaders() As String
    Dim lPos            As Long
    
    lPos = InStr(m_sAllResponseHeaders, vbCrLf)
    If lPos > 0 Then
        GetAllResponseHeaders = Mid$(m_sAllResponseHeaders, lPos + 2)
    End If
End Function

Public Function GetResponseHeader(ByVal Header As String) As String
    If m_oResponseHeaders.Exists(Header) Then
        GetResponseHeader = m_oResponseHeaders.Item(Header)
    End If
End Function

Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As Long)
    Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER As Long = 0
    
    If Flags = HTTPREQUEST_SETCREDENTIALS_FOR_SERVER Then
        m_uRemote.User = UserName
        m_uRemote.Pass = Password
    End If
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

Public Sub SetProxy(ByVal ProxySetting As Long, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
    #If ProxySetting And ProxyServer And BypassList Then
    #End If
    Err.Raise E_NOTIMPL, MODULE_NAME & ".SetProxy", m_oSocket.GetErrorDescription(E_NOTIMPL)
End Sub

'= private ===============================================================

Private Function pvSendBody(uSend As UcsBuffer) As Boolean
    Const FUNC_NAME     As String = "pvSendBody"
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hResult         As Long
    
    On Error GoTo EH
    Do
        lSize = uSend.Size - uSend.Pos
        If lSize <= 0 Then
            If pvArraySize(uSend.Data) <> DEF_BUFFER_SIZE Then
                ReDim uSend.Data(0 To DEF_BUFFER_SIZE - 1) As Byte
            End If
            hResult = IStream_Read(m_pRequestStream, uSend.Data, lSize)
            If hResult < 0 Then
                pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Read"
                GoTo QH
            End If
            If lSize = 0 Then
                m_uRecvBuffer.Size = 0
                m_uChunkBuffer.Size = 0
                m_eState = ucsWaitRecvHeaders
                GoTo QH
            End If
            uSend.Size = lSize
            uSend.Pos = 0
        End If
        If lSize <= 0 Then
            Exit Do
        End If
        lResult = m_oSocket.Send(VarPtr(uSend.Data(uSend.Pos)), lSize)
        If lResult < 0 Then
            If Not m_oSocket.HasPendingEvent Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Send"
                GoTo QH
            End If
        Else
            uSend.Pos = uSend.Pos + lResult
        End If
        If m_oSocket.HasPendingEvent Then
            Exit Do
        End If
    Loop
    '--- success
    pvSendBody = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvRecvHeaders(baBuffer() As Byte, uRecv As UcsBuffer) As Boolean
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
    pvBufferWriteArray uRecv, baBuffer
    lPos = InStrB(1, uRecv.Data, StrConv(HDR_DELIM, vbFromUnicode))
    If lPos > 0 Then
        m_sAllResponseHeaders = Left$(m_oSocket.FromTextArray(uRecv.Data, ucsScpAcp), lPos - 1)
        vSplit = Split(m_sAllResponseHeaders, vbCrLf)
        lIdx = 0
        For Each vElem In vSplit
            vElem = Split(vElem, ":", Limit:=2)
            If UBound(vElem) > 0 Then
                m_oResponseHeaders.Item(vElem(0)) = Trim$(vElem(1))
            Else
                m_oResponseHeaders.Item(lIdx) = Trim$(vElem(0))
            End If
            lIdx = lIdx + 1
        Next
        m_lStatus = Val(Mid$(vSplit(0), 1 + sizeof_HTTP_VER + 1, sizeof_HTTP_STATUS))
        m_sStatusText = Mid$(vSplit(0), 1 + sizeof_HTTP_VER + 1 + sizeof_HTTP_STATUS + 1)
        Select Case m_lStatus
        Case 0
            pvSetError ERROR_WINHTTP_INVALID_SERVER_RESPONSE, MODULE_NAME & "." & FUNC_NAME
            GoTo QH
        Case 300 To 399
            If Not Option_(WinHttpRequestOption_EnableRedirects) Or Not m_oResponseHeaders.Exists(HDR_LOCATION) Then
                GoTo TryDownload
            End If
            m_lCurrentRedirect = m_lCurrentRedirect + 1
            If m_lCurrentRedirect > Option_(WinHttpRequestOption_MaxAutomaticRedirects) Then
                GoTo TryDownload
            End If
            sEscaped = pvEscapeUrl(m_oResponseHeaders.Item(HDR_LOCATION))
            Option_(WinHttpRequestOption_URL) = sEscaped
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, "Redirecting to " & sEscaped
            #End If
            If Not pvParseUrl(sEscaped, uRedirect, m_uRemote.Protocol) Then
                pvSetError ERROR_WINHTTP_INVALID_URL, MODULE_NAME & "." & FUNC_NAME
                GoTo QH
            End If
            uRedirect.User = m_uRemote.User
            uRedirect.Pass = m_uRemote.Pass
            If Not pvIsProtocolSupported(uRedirect.Protocol) Then
                pvSetError ERROR_WINHTTP_UNRECOGNIZED_SCHEME, MODULE_NAME & "." & FUNC_NAME
                GoTo QH
            End If
            If Not Option_(WinHttpRequestOption_EnableHttpsToHttpRedirects) Then
                If pvIsProtocolSecure(m_uRemote.Protocol) And Not pvIsProtocolSecure(uRedirect.Protocol) Then
                    GoTo TryDownload
                End If
            End If
            If m_uRemote.Host <> uRedirect.Host Or m_uRemote.Port <> uRedirect.Port Or m_uRemote.Protocol <> uRedirect.Protocol Then
                Set m_oSocket = pvCreateNewSocket
            End If
            m_uRemote = uRedirect
            m_oResponseHeaders.RemoveAll
            SetRequestHeader HDR_HOST, m_uRemote.Host
            m_eState = ucsWaitResolve
            If m_oSocket.IsClosed Then
                If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
                    pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Create"
                    GoTo QH
                End If
                If Not pvConnectRemote(m_uRemote) Then
                    pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.Connect"
                    GoTo QH
                End If
            Else
                pvDrainKeepAlive
                m_oSocket.PostEvent ucsSfdWrite
            End If
        Case Else
TryDownload:
            If m_oResponseHeaders.Exists(HDR_CONTENT_LENGTH) Then
                m_dblContentLength = Val(m_oResponseHeaders.Item(HDR_CONTENT_LENGTH))
            End If
            m_sTransferEncoding = LCase$(GetResponseHeader("Transfer-Encoding"))
            m_sContentEncoding = LCase$(GetResponseHeader("Content-Encoding"))
            m_eState = ucsWaitRecvBody
            lPos = lPos + sizeof_HDR_DELIM - 1
            If uRecv.Size <= lPos Then
                baBuffer = vbNullString
            Else
                ReDim baBuffer(0 To uRecv.Size - lPos - 1) As Byte
                Call CopyMemory(baBuffer(0), uRecv.Data(lPos), UBound(baBuffer) + 1)
            End If
            RaiseEvent OnResponseStart(m_lStatus, GetResponseHeader(HDR_CONTENT_TYPE))
            uRecv.Size = 0
            Set m_pResponseStream = SHCreateMemStream(ByVal 0, 0)
            m_sResponseText = vbNullChar
            pvRecvBody baBuffer, uRecv
        End Select
    ElseIf uRecv.Size > Option_(WinHttpRequestOption_MaxResponseHeaderSize) Then
        pvSetError ERROR_WINHTTP_HEADER_SIZE_OVERFLOW, MODULE_NAME & "." & FUNC_NAME
        GoTo QH
    End If
    '--- success
    pvRecvHeaders = True
QH:
    Exit Function
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvRecvBody(baBuffer() As Byte, uRecv As UcsBuffer, Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvRecvBody"
    Const sizeof_CrLf   As Long = 2
    Dim hResult         As Long
    Dim baTemp()        As Byte
    Dim lPos            As Long
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim uChunk          As UcsBuffer
    Dim pOutput         As IUnknown
    
    On Error GoTo EH
    If UBound(baBuffer) >= 0 Then
        If m_sTransferEncoding = "chunked" Then
            pvBufferWriteArray m_uChunkBuffer, baBuffer
            pvParseChunks m_uChunkBuffer, m_aTransferChunks, lSize
            For lIdx = 0 To lSize - 1
                uChunk = m_aTransferChunks(lIdx)
                lPos = uChunk.Pos + uChunk.Size + sizeof_CrLf
                If uChunk.Size = 0 Then
                    m_eState = ucsIdle
                    Flush = True
                    Exit For
                End If
                ReDim baTemp(0 To uChunk.Size - 1) As Byte
                Call CopyMemory(baTemp(0), m_uChunkBuffer.Data(uChunk.Pos), UBound(baTemp) + 1)
                m_dblBytesProgress = m_dblBytesProgress + UBound(baTemp) + 1
                RaiseEvent OnResponseDataAvailable(baTemp)
                pvBufferWriteArray uRecv, baTemp
            Next
            If lPos > 0 Then
                lSize = m_uChunkBuffer.Size - lPos
                If lSize > 0 Then
                    Call CopyMemory(m_uChunkBuffer.Data(0), m_uChunkBuffer.Data(lPos), lSize)
                End If
                m_uChunkBuffer.Size = lSize
            End If
        Else
            m_dblBytesProgress = m_dblBytesProgress + UBound(baBuffer) + 1
            RaiseEvent OnResponseDataAvailable(baBuffer)
            pvBufferWriteArray uRecv, baBuffer
            If m_dblBytesProgress >= m_dblContentLength And m_dblContentLength >= 0 Then
                m_eState = ucsIdle
                Flush = True
            End If
        End If
    End If
    If uRecv.Size > 0 Then
        If uRecv.Size >= DEF_BUFFER_SIZE Or Flush Then
            If uRecv.Size < UBound(uRecv.Data) + 1 Then
                ReDim Preserve uRecv.Data(0 To uRecv.Size - 1) As Byte
            End If
            hResult = IStream_Write(m_pResponseStream, uRecv.Data)
            uRecv.Size = 0
            If hResult < 0 Then
                pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Write"
                GoTo QH
            End If
        End If
    End If
    If Flush Then
        If m_sContentEncoding = "gzip" Then
            hResult = IStream_Seek(m_pResponseStream, 0, STREAM_SEEK_SET)
            If hResult < 0 Then
                pvSetError hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "IStream_Seek(STREAM_SEEK_SET)"
            End If
            If pvUngzipStream(m_pResponseStream, pOutput) Then
                Set m_pResponseStream = pOutput
                m_sContentEncoding = vbNullString
            End If
        End If
        RaiseEvent OnResponseFinished
    End If
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
    
    sEncoded = StrConv(pvToTextArray(sUrl, Option_(WinHttpRequestOption_URLCodePage)), vbUnicode)
    If Not Option_(WinHttpRequestOption_UrlEscapeDisable) Then
        If Option_(WinHttpRequestOption_EscapePercentInURL) Then
            lFlags = lFlags Or URL_ESCAPE_PERCENT
        End If
        vSplit = Split(sEncoded, "?", Limit:=2)
        pvEscapeUrl = pvEscapePart(vSplit(0), lFlags)
        '--- append query string and anchor
        If UBound(vSplit) > 0 Then
            If Not Option_(WinHttpRequestOption_UrlEscapeDisableQuery) Then
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

Private Function pvGetAsyncSocket() As cAsyncSocket
    #If ImplUseTls Then
        Set pvGetAsyncSocket = m_oSocket.Socket
    #Else
        Set pvGetAsyncSocket = m_oSocket
    #End If
End Function

Private Function pvConnectRemote(uRemote As UcsParsedUrl) As Boolean
    #If ImplUseTls Then
        Dim eProt           As WinHttpRequestSecureProtocols
        Dim eFeatures       As UcsTlsLocalFeaturesEnum
        
        eProt = Option_(WinHttpRequestOption_SecureProtocols)
        If eProt = SecureProtocol_ALL Then
            eFeatures = ucsTlsSupportAll
        Else
            eFeatures = ucsTlsSupportTls13 * -((eProt And SecureProtocol_TLS1_3) <> 0) _
                Or ucsTlsSupportTls12 * -((eProt And SecureProtocol_TLS1_2) <> 0) _
                Or ucsTlsSupportTls11 * -((eProt And SecureProtocol_TLS1_1) <> 0) _
                Or ucsTlsSupportTls10 * -((eProt And SecureProtocol_TLS1) <> 0)
        End If
        If Option_(WinHttpRequestOption_SslErrorIgnoreFlags) <> 0 Then
            eFeatures = eFeatures Or ucsTlsIgnoreServerCertificateErrors
        End If
        If Not Option_(WinHttpRequestOption_EnableCertificateRevocationCheck) Then
            eFeatures = eFeatures Or ucsTlsIgnoreServerCertificateRevocation
        End If
        pvConnectRemote = m_oSocket.Connect(uRemote.Host, uRemote.Port, UseTls:=pvIsProtocolSecure(uRemote.Protocol), LocalFeatures:=eFeatures)
    #Else
        pvConnectRemote = m_oSocket.Connect(uRemote.Host, uRemote.Port)
    #End If
End Function

Private Function pvIsProtocolSupported(sProtocol As String) As Boolean
    #If ImplUseTls Then
        pvIsProtocolSupported = (LCase$(sProtocol) = "http" Or LCase$(sProtocol) = "https")
    #Else
        pvIsProtocolSupported = (LCase$(sProtocol) = "http")
    #End If
End Function

Private Function pvIsProtocolSecure(sProtocol As String) As Boolean
    pvIsProtocolSecure = (LCase$(sProtocol) = "https")
End Function

Private Sub pvDrainKeepAlive()
    Dim baBuffer()      As Byte
    Dim lTotal          As Long
    
    With pvGetAsyncSocket
        .ReceiveArray baBuffer
        lTotal = pvArraySize(baBuffer)
        If lTotal > 0 Then
            Do While lTotal < Option_(WinHttpRequestOption_MaxResponseDrainSize)
                If Not .SyncWaitForEvent(100, ucsSfdRead) Then
                    Exit Do
                End If
                .ReceiveArray baBuffer
                lTotal = lTotal + pvArraySize(baBuffer)
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
        '-- note: call pvGetAsyncSocket on each iteration because socket instance might change on redirect
        With pvGetAsyncSocket
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

Private Sub pvParseChunks(uChunk As UcsBuffer, aOutput() As UcsBuffer, lOutSize As Long)
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
    lPos = InStrB(lStart, uChunk.Data, StrConv(vbCrLf, vbFromUnicode))
    Do While lPos > 0 And lPos < uChunk.Size
        sTemp = String$(lPos - lStart, 0)
        Call CopyMemory(ByVal sTemp, uChunk.Data(lStart - 1), Len(sTemp))
        lSize = Val("&H" & sTemp & "&")
        If lPos + sizeof_CrLf + lSize > m_uChunkBuffer.Size Then
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
        lPos = InStrB(lStart, uChunk.Data, StrConv(vbCrLf, vbFromUnicode))
    Loop
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(?:(.+):)?//)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(/[^?#]*)?(\?[^#]*)?(#.*)?$"
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
                    uParsed.Port = Val(Mid$(.Item(4), 2))
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

Private Function pvGetStreamSize(pStream As IUnknown) As Double
    Dim hResult         As Long
    Dim cPos            As Currency
    
    If Not pStream Is Nothing Then
        hResult = IStream_Seek(pStream, 0, STREAM_SEEK_END, cPos)
        If hResult < 0 Then
            Err.Raise hResult, "IStream_Seek(STREAM_SEEK_END)", m_oSocket.GetErrorDescription(hResult)
        End If
        hResult = IStream_Seek(pStream, 0, STREAM_SEEK_SET)
        If hResult < 0 Then
            Err.Raise hResult, "IStream_Seek(STREAM_SEEK_SET)", m_oSocket.GetErrorDescription(hResult)
        End If
    End If
    pvGetStreamSize = CDbl(cPos) * 10000#
End Function

Private Function pvGetHeaderAttrib(sHeader As String, sAttrib As String) As String
    Dim vSplit          As Variant
    Dim vElem           As Variant
    
    '--- TODO: need regex parser for semi-colons in quoted values
    vSplit = Split(sHeader, ";")
    If UBound(vSplit) > 0 Then
        For Each vElem In vSplit
            vElem = Split(vElem, "=", Limit:=2)
            If UBound(vElem) > 0 Then
                If LCase$(Trim$(vElem(0))) = LCase$(sAttrib) Then
                    pvGetHeaderAttrib = Trim$(vElem(1))
                    Exit For
                End If
            End If
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
            If LCase$(Left$(sCharset, Len(vElem))) = vElem Then
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

Private Function pvToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    pvToTextArray = m_oSocket.ToTextArray(sText, CodePage)
End Function

Private Function pvUngzipStream(pInput As IUnknown, pOutput As IUnknown) As Boolean
    Const FUNC_NAME     As String = "pvUngzipStream"
    Const BUFF_SIZE     As Long = 65536
    Dim baBuffer()      As Byte
    Dim hArchive        As Long
    Dim lSize           As Long
    Dim lResult         As Long
    Dim lReaderPfn      As Long
    Dim uReaderThunk    As UcsRedirectReaderThunk
    Dim hResult         As Long
        
    On Error GoTo EH
    ReDim m_baGzipBuffer(0 To BUFF_SIZE - 1) As Byte
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
    ReDim baBuffer(0 To BUFF_SIZE - 1) As Byte
    Do
        lSize = archive_read_data(hArchive, baBuffer(0), UBound(baBuffer) + 1)
        If lSize = 0 Then
            Exit Do
        End If
        If UBound(baBuffer) + 1 <> lSize Then
            ReDim Preserve baBuffer(0 To lSize - 1) As Byte
        End If
        If pOutput Is Nothing Then
            Set pOutput = SHCreateMemStream(baBuffer(0), UBound(baBuffer) + 1)
        Else
            hResult = IStream_Write(pOutput, baBuffer)
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
    Erase m_baGzipBuffer
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
    hResult = IStream_Read(pInput, m_baGzipBuffer, lSize)
    If hResult < 0 Then
        GoTo QH
    End If
    If lSize > 0 Then
        lBufferPtr = VarPtr(m_baGzipBuffer(0))
        '--- success
        ArchiveReaderCallback = lSize
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function InitRedirectReaderThunk(uThunk As UcsRedirectReaderThunk, ByVal ParamValue As Long, ByVal pfnDest As Long) As Long
    Const PAGE_EXECUTE_READWRITE As Long = &H40
    
    With uThunk
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

Private Function IStream_Read(pStm As IUnknown, baBuffer() As Byte, Optional BytesRead As Long) As Long
    If Not pStm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Read = DispCallByVtbl(pStm, 3, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesRead))
    End If
End Function

Private Function IStream_Write(pStm As IUnknown, baBuffer() As Byte, Optional BytesWritten As Long) As Long
    If Not pStm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Write = DispCallByVtbl(pStm, 4, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesWritten))
    End If
End Function

Private Function IStream_Seek(pStm As stdole.IUnknown, ByVal cMove As Currency, ByVal dwOrigin As Long, Optional NewPosition As Currency) As Long
    If Not pStm Is Nothing Then
        IStream_Seek = DispCallByVtbl(pStm, 5, cMove, dwOrigin, VarPtr(NewPosition))
    Else
        IStream_Seek = -1
    End If
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
' Socket events
'=========================================================================

Private Sub m_oSocket_OnResolve(IpAddress As String)
    If m_eState = ucsWaitResolve Then
        m_eState = ucsWaitConnect
    End If
End Sub

Private Sub m_oSocket_OnConnect()
    m_oSocket_OnSend
End Sub

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    
    On Error GoTo EH
    If m_eState = ucsWaitResolve Or m_eState = ucsWaitConnect Then
        m_uSendBuffer.Size = 0
        m_uSendBuffer.Pos = 0
        pvBufferWriteArray m_uSendBuffer, pvToTextArray( _
            m_sMethod & " " & m_uRemote.Path & m_uRemote.QueryString & m_uRemote.Anchor & " HTTP/1." & IIf(Option_(WinHttpRequestOption_EnableHttp1_1), "1", "0") & vbCrLf & _
            pvJoinHeaders(m_oRequestHeaders) & HDR_DELIM, ucsScpAcp)
        m_eState = ucsWaitSendBody
    End If
    If m_eState = ucsWaitSendBody Then
        If Not pvSendBody(m_uSendBuffer) Then
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
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eState = ucsWaitRecvHeaders Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvHeaders(baBuffer, m_uRecvBuffer) Then
            GoTo QH
        End If
    End If
    If m_eState = ucsWaitRecvBody Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            If m_dblBytesProgress < m_dblContentLength Then
                pvSetError m_oSocket.LastError, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "m_oSocket.ReceiveArray"
            End If
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvBody(baBuffer, m_uRecvBuffer) Then
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
    Dim baBuffer()      As Byte
    Dim dblBytes        As Double
    
    On Error GoTo EH
    If m_eState = ucsWaitRecvBody Then
        Do
            dblBytes = m_dblBytesProgress
            m_oSocket_OnReceive
        Loop While m_dblBytesProgress > dblBytes
        baBuffer = vbNullString
        pvRecvBody baBuffer, m_uRecvBuffer, Flush:=True
    End If
    m_eState = ucsIdle
    m_oSocket.Close_
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    #If ImplUseTls Then
        With m_oSocket.LastError
            If .Number <> 0 Then
                pvSetError .Number, .Source, .Description
            End If
        End With
    #Else
        With m_oSocket
            If .LastError <> 0 Then
                pvSetError .LastError, MODULE_NAME & "." & FUNC_NAME, .GetErrorDescription(.LastError)
            End If
        End With
    #End If
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

#If ImplUseTls Then
Private Sub m_oSocket_OnClientCertificate(Issuers As Object, Confirmed As Boolean)
    Const FUNC_NAME     As String = "m_oSocket_OnClientCertificate"
    Dim vElem           As Variant
    
    On Error GoTo EH
    '--- try to match a client certificate in CURRENT_USER\MY certificate store by Issuer's RDN
    For Each vElem In Issuers
        If m_oSocket.ImportSystemStoreCertificates(vElem) Then
            Confirmed = True
            GoTo QH
        End If
    Next
    If LenB(Option_(WinHttpRequestOption_SelectCertificate)) <> 0 Then
        If m_oSocket.ImportSystemStoreCertificates(Option_(WinHttpRequestOption_SelectCertificate)) Then
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
    SetTimeouts 5000, 5000, 15000, 15000
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
End Sub
