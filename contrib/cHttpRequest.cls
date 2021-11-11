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
Private Const ERROR_WINHTTP_HEADER_SIZE_OVERFLOW    As Long = (WINHTTP_ERROR_BASE + 182)
Private Const ERROR_INVALID_PARAMETER               As Long = &H80070057
Private Const E_NOTIMPL                             As Long = &H80004001
Private Const E_POINTER                             As Long = &H80004003

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeW" (ByVal pszURL As Long, ByVal pszEscaped As Long, pcchEscaped As Long, ByVal dwFlags As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

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
'--- i/o buffers
Private m_uSendBuffer           As UcsBuffer
Private m_uRecvBuffer           As UcsBuffer
Private m_uChunkBuffer          As UcsBuffer

Private Enum UcsStateEnum
    ucsIdle
    ucsConnecting
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
    Dim lSize               As Long
    Dim baRetVal()          As Byte
    Dim hResult             As Long
    
    lSize = pvGetStreamSize(m_pResponseStream)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        hResult = IStream_Read(m_pResponseStream, baRetVal)
        If hResult < 0 Then
            Err.Raise vbObjectError, "IStream_Read", m_oSocket.GetErrorDescription(hResult)
        End If
    Else
        baRetVal = vbNullString
    End If
    ResponseBody = baRetVal
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

'=========================================================================
' Methods
'=========================================================================

Public Sub Open_(ByVal Method As String, ByVal Url As String, Optional ByVal Async As Variant)
    Const FUNC_NAME     As String = "Open"
    Dim sEscaped        As String
    Dim uRemote         As UcsParsedUrl
    
    On Error GoTo EH
    sEscaped = pvEscapeUrl(Url)
    Option_(WinHttpRequestOption_URL) = sEscaped
    If Not pvParseUrl(sEscaped, uRemote, "http") Then
        On Error GoTo 0
        Err.Raise ERROR_WINHTTP_INVALID_URL, , m_oSocket.GetErrorDescription(ERROR_WINHTTP_INVALID_URL)
    End If
    If Not pvIsProtocolSupported(uRemote.Protocol) Then
        On Error GoTo 0
        Err.Raise ERROR_WINHTTP_UNRECOGNIZED_SCHEME, , m_oSocket.GetErrorDescription(ERROR_WINHTTP_UNRECOGNIZED_SCHEME)
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
    '--- init headers
    SetRequestHeader "Host", m_uRemote.Host
    SetRequestHeader "User-Agent", Option_(WinHttpRequestOption_UserAgentString)
    If Option_(WinHttpRequestOption_EnableHttp1_1) Then
        SetRequestHeader "Connection", "Keep-Alive"
    End If
    SetRequestHeader "Accept", "*.*"
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
    '--- restore headers
    SetRequestHeader "Host", m_uRemote.Host
    SetRequestHeader "User-Agent", Option_(WinHttpRequestOption_UserAgentString)
    '--- prepare auth header
    If LenB(m_uRemote.User) <> 0 Or LenB(m_uRemote.Pass) <> 0 Then
        SetRequestHeader "Authorization", "Basic " & Replace(ToBase64Array( _
            pvToTextArray(m_uRemote.User & ":" & m_uRemote.Pass, ucsScpAcp)), vbCrLf, vbNullString)
    End If
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
        On Error GoTo 0
        Err.Raise ERROR_INVALID_PARAMETER, , m_oSocket.GetErrorDescription(ERROR_INVALID_PARAMETER)
    End If
    If Not m_pRequestStream Is Nothing Then
        SetRequestHeader HDR_CONTENT_LENGTH, pvGetStreamSize(m_pRequestStream)
    End If
    '--- try to reuse socket if not closed (keep-alive)
    m_eState = ucsConnecting
    If m_oSocket.IsClosed Then
        If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
            On Error GoTo 0
            Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
        End If
        If Not pvConnectRemote(m_uRemote) Then
            On Error GoTo 0
            Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
        End If
    Else
        pvDrainKeepAlive
        m_oSocket.PostEvent ucsSfdWrite
    End If
    '--- handle synchronous mode
    If Not m_bAsync Then
        If Not WaitForResponse(m_lSendTimeout) Then
            If m_eState < ucsWaitRecvHeaders Then
                Abort
                On Error GoTo 0
                Err.Raise ERROR_WINHTTP_TIMEOUT, , m_oSocket.GetErrorDescription(ERROR_WINHTTP_TIMEOUT)
            End If
        End If
        If Not WaitForResponse(m_lReceiveTimeout) Then
            Abort
            On Error GoTo 0
            Err.Raise ERROR_WINHTTP_TIMEOUT, , m_oSocket.GetErrorDescription(ERROR_WINHTTP_TIMEOUT)
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
    m_lResolveTimeout = ResolveTimeout
    m_lConnectTimeout = ConnectTimeout
    m_lSendTimeout = SendTimeout
    m_lReceiveTimeout = ReceiveTimeout
End Sub

Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)
    Const FUNC_NAME     As String = "SetRequestHeader"
    
    On Error GoTo EH
    If LenB(Value) <> 0 Then
        m_oRequestHeaders.Item(Header) = Value
    Else
        m_oRequestHeaders.Remove Header
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub Abort()
    Const FUNC_NAME     As String = "Abort"
    
    On Error GoTo EH
    m_eState = ucsIdle
    Set m_oSocket = pvCreateNewSocket
    Set m_pRequestStream = Nothing
    Set m_pResponseStream = Nothing
    m_sResponseText = vbNullChar
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function WaitForResponse(Optional ByVal Timeout As Variant) As Boolean
    Const FUNC_NAME     As String = "WaitForResponse"
    Dim dblTimerEx      As Double
    Dim lTimeout        As Long
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    If IsMissing(Timeout) Then
        Timeout = -1
    End If
    Do While m_eState <> ucsIdle
        If Timeout >= 0 Then
            lTimeout = Timeout - Int((TimerEx - dblTimerEx) * 1000)
            If lTimeout <= 0 Then
                GoTo QH
            End If
        End If
        With pvGetAsyncSocket
            If Not .SyncWaitForEvent(lTimeout) Then
                GoTo QH
            End If
            .SyncProcessMsgQueue
        End With
    Loop
    '--- success
    WaitForResponse = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

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

'= not implemented =======================================================

Public Sub SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)
    #If AutoLogonPolicy Then '--- touch args for MZ-Tools
    #End If
    Err.Raise E_NOTIMPL, , m_oSocket.GetErrorDescription(E_NOTIMPL)
End Sub

Public Sub SetClientCertificate(ByVal ClientCertificate As String)
    #If ClientCertificate Then
    #End If
    Err.Raise E_NOTIMPL, , m_oSocket.GetErrorDescription(E_NOTIMPL)
End Sub

Public Sub SetProxy(ByVal ProxySetting As Long, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
    #If ProxySetting And ProxyServer And BypassList Then
    #End If
    Err.Raise E_NOTIMPL, , m_oSocket.GetErrorDescription(E_NOTIMPL)
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
            SetRequestHeader "Host", m_uRemote.Host
            m_eState = ucsConnecting
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
        DebugLog MODULE_NAME, "pvSetError", Hex$(ErrNumber) & " " & ErrDescription
    #End If
    RaiseEvent OnError(ErrNumber, ErrDescription & " [" & Replace(ErrSource, vbCrLf, "; ") & "]")
    Abort
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvJoinHeaders(oDict As Object) As String
    Const FUNC_NAME     As String = "pvJoinHeaders"
    Dim vTemp           As Variant
    Dim vKey            As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vTemp = oDict.Keys
    For Each vKey In vTemp
        vTemp(lIdx) = vTemp(lIdx) & ": " & oDict.Item(vKey)
        lIdx = lIdx + 1
    Next
    pvJoinHeaders = Join(vTemp, vbCrLf)
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
        lSize = Val("&H" & sTemp)
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
            Err.Raise vbObjectError, "IStream_Seek(STREAM_SEEK_END)", m_oSocket.GetErrorDescription(hResult)
        End If
        hResult = IStream_Seek(pStream, 0, STREAM_SEEK_SET)
        If hResult < 0 Then
            Err.Raise vbObjectError, "IStream_Seek(STREAM_SEEK_SET)", m_oSocket.GetErrorDescription(hResult)
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
    Select Case LCase$(sCharset)
    Case "iso-8859-1", "windows-1252"
        pvMapCharset = 1252
    Case "windows-1251"
        pvMapCharset = 1251
    Case "utf-8"
        pvMapCharset = ucsScpUtf8
    Case Else
        pvMapCharset = ucsScpAcp
    End Select
End Function

Private Function pvToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    pvToTextArray = m_oSocket.ToTextArray(sText, CodePage)
End Function

Private Function AsIStream(ByVal pUnk As IUnknown) As IUnknown
    Const IDX_QueryInterface As Long = 0
    Dim IID_IStream(0 To 3) As Long
    Dim hResult         As Long
    
    IID_IStream(0) = &HC: IID_IStream(1) = &H0
    IID_IStream(2) = &HC0: IID_IStream(3) = &H46000000
    hResult = DispCallByVtbl(pUnk, IDX_QueryInterface, VarPtr(IID_IStream(0)), VarPtr(AsIStream))
    If hResult < 0 Then
        Err.Raise hResult, "IUnknown.QueryInterface(IID_IStream)"
    End If
End Function

Private Function IStream_Read(pstm As IUnknown, baBuffer() As Byte, Optional BytesRead As Long) As Long
    If Not pstm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Read = DispCallByVtbl(pstm, 3, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesRead))
    End If
End Function

Private Function IStream_Write(pstm As IUnknown, baBuffer() As Byte, Optional BytesWritten As Long) As Long
    If Not pstm Is Nothing And UBound(baBuffer) >= 0 Then
        IStream_Write = DispCallByVtbl(pstm, 4, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, VarPtr(BytesWritten))
    End If
End Function

Private Function IStream_Seek(pstm As stdole.IUnknown, ByVal cMove As Currency, ByVal dwOrigin As Long, Optional NewPosition As Currency) As Long
    If Not pstm Is Nothing Then
        IStream_Seek = DispCallByVtbl(pstm, 5, cMove, dwOrigin, VarPtr(NewPosition))
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
#End If

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnSend()
    Const FUNC_NAME     As String = "m_oSocket_OnSend"
    
    On Error GoTo EH
    If m_eState = ucsConnecting Then
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
                pvSetError .LastError, , .GetErrorDescription(.LastError)
            End If
        End With
    #End If
    Exit Sub
EH:
    pvSetError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Err.Description
End Sub

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
