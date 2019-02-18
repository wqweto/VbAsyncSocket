VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTlsClient"
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
Private Const MODULE_NAME As String = "cTlsClient"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplDebug = False

'=========================================================================
' API
'=========================================================================

'--- for AcquireCredentialsHandle
Private Const UNISP_NAME                    As String = "Microsoft Unified Security Protocol Provider"
Private Const SECPKG_CRED_OUTBOUND          As Long = 2
Private Const SCHANNEL_CRED_VERSION         As Long = 4
Private Const SCH_CRED_MANUAL_CRED_VALIDATION As Long = 8
Private Const SCH_CRED_NO_DEFAULT_CREDS     As Long = &H10
'-- for InitializeSecurityContext
'Private Const ISC_REQ_MUTUAL_AUTH           As Long = &H2
Private Const ISC_REQ_REPLAY_DETECT         As Long = &H4
Private Const ISC_REQ_SEQUENCE_DETECT       As Long = &H8
Private Const ISC_REQ_CONFIDENTIALITY       As Long = &H10
Private Const ISC_REQ_USE_SUPPLIED_CREDS    As Long = &H80
Private Const ISC_REQ_ALLOCATE_MEMORY       As Long = &H100
'Private Const ISC_REQ_CONNECTION            As Long = &H800
Private Const ISC_REQ_EXTENDED_ERROR        As Long = &H4000
Private Const ISC_REQ_STREAM                As Long = &H8000&
'Private Const ISC_REQ_INTEGRITY             As Long = &H10000
'Private Const ISC_REQ_MANUAL_CRED_VALIDATION As Long = &H80000
Private Const SECURITY_NATIVE_DREP          As Long = &H10
'--- for ApiSecBuffer.BufferType
Private Const SECBUFFER_EMPTY               As Long = 0   ' Undefined, replaced by provider
Private Const SECBUFFER_DATA                As Long = 1   ' Packet data
Private Const SECBUFFER_TOKEN               As Long = 2   ' Security token
Private Const SECBUFFER_EXTRA               As Long = 5   ' Extra data
Private Const SECBUFFER_STREAM_TRAILER      As Long = 6   ' Security Trailer
Private Const SECBUFFER_STREAM_HEADER       As Long = 7   ' Security Header
Private Const SECBUFFER_ALERT               As Long = 17
Private Const SECBUFFER_VERSION             As Long = 0
'--- SSPI/Schannel retvals
Private Const SEC_I_CONTINUE_NEEDED         As Long = &H90312
Private Const SEC_I_COMPLETE_NEEDED         As Long = &H90313
Private Const SEC_I_COMPLETE_AND_CONTINUE   As Long = &H90314
Private Const SEC_I_CONTEXT_EXPIRED         As Long = &H90317
Private Const SEC_I_INCOMPLETE_CREDENTIALS  As Long = &H90320
Private Const SEC_I_RENEGOTIATE             As Long = &H90321
Private Const SEC_E_INCOMPLETE_MESSAGE      As Long = &H80090318
Private Const SEC_E_OK                      As Long = 0
'--- for QueryContextAttributes
Private Const SECPKG_ATTR_STREAM_SIZES      As Long = 4
'--- for ApplyControlToken
Private Const SCHANNEL_SHUTDOWN             As Long = 1   ' gracefully close down a connection

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AcquireCredentialsHandle Lib "security" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonId As Long, pAuthData As Any, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, phCredential As Currency, ByVal ptsExpiry As Long) As Long
Private Declare Function FreeCredentialsHandle Lib "security" (phContext As Currency) As Long
Private Declare Function InitializeSecurityContext Lib "security" Alias "InitializeSecurityContextA" (phCredential As Currency, phContext As Any, pszTargetName As Any, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, pInput As Any, ByVal Reserved2 As Long, phNewContext As Currency, pOutput As Any, pfContextAttr As Long, ByVal ptsExpiry As Long) As Long
Private Declare Function DeleteSecurityContext Lib "security" (phContext As Currency) As Long
Private Declare Function FreeContextBuffer Lib "security" (ByVal pvContextBuffer As Long) As Long
Private Declare Function CompleteAuthToken Lib "security" (phContext As Currency, pToken As Any) As Long
Private Declare Function QueryContextAttributes Lib "security" Alias "QueryContextAttributesA" (phContext As Currency, ByVal ulAttribute As Long, pBuffer As Any) As Long
Private Declare Function DecryptMessage Lib "secur32" (phContext As Currency, pMessage As Any, ByVal MessageSeqNo As Long, ByVal pfQOP As Long) As Long
Private Declare Function EncryptMessage Lib "secur32" (phContext As Currency, ByVal fQOP As Long, pMessage As Any, ByVal MessageSeqNo As Long) As Long
Private Declare Function ApplyControlToken Lib "secur32" (phContext As Currency, pInput As Any) As Long
#If ImplDebug Then
    Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
#End If

Private Type SCHANNEL_CRED
    dwVersion               As Long
    cCreds                  As Long
    paCred                  As Long
    hRootStore              As Long
    cMappers                As Long
    aphMappers              As Long
    cSupportedAlgs          As Long
    palgSupportedAlgs       As Long
    grbitEnabledProtocols   As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan       As Long
    dwFlags                 As Long
    dwCredFormat            As Long
End Type

Private Type ApiSecBuffer
    cbBuffer                As Long
    BufferType              As Long
    pvBuffer                As Long
End Type

Private Type ApiSecBufferDesc
    ulVersion               As Long
    cBuffers                As Long
    pBuffers                As Long
End Type

Private Type ApiSecPkgContext_StreamSizes
    cbHeader                As Long
    cbTrailer               As Long
    cbMaximumMessage        As Long
    cBuffers                As Long
    cbBlockSize             As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ERR_CONTEXT_MISMATCH As String = "Context attributes do not match requested: &H%1<>&H%2"
Private Const STR_ERR_CLIENT_CERT_REQUIRED As String = "Client certificate required"
Private Const STR_ERR_UNEXPECTED_RESULT1 As String = "Unexpected result from InitializeSecurityContext: &H%1"
Private Const STR_ERR_UNEXPECTED_RESULT2 As String = "Unexpected result from DecryptMessage: &H%1"
Private Const STR_ERR_UNEXPECTED_RESULT3 As String = "Unexpected result from EncryptMessage: &H%1"

Private m_oSocket               As cAsyncSocket
Private m_lLastErrNumber        As Long
Private m_sLastErrDescription   As String
Private m_sLastErrSource        As String
Private m_lResolveTimeout       As Long
Private m_lConnectTimeout       As Long
Private m_lWriteTimeout         As Long
Private m_lReadTimeout          As Long
'--- for TLS
Private m_sTlsHostAddress       As String
Private m_hTlsCredentials       As Currency
Private m_hTlsContext           As Currency
Private m_uTlsSizes             As ApiSecPkgContext_StreamSizes
Private m_baTlsExtra()          As Byte

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

Property Get LastError() As VBA.ErrObject
    Err.Number = m_lLastErrNumber
    Err.Source = m_sLastErrSource
    Err.Description = m_sLastErrDescription
    Set LastError = Err
End Property

Property Get AvailableBytes() As Long
    AvailableBytes = m_oSocket.AvailableBytes
    If m_hTlsContext <> 0 And AvailableBytes >= 0 Then
        AvailableBytes = AvailableBytes + UBound(m_baTlsExtra) + 1
    End If
End Property

Property Get TlsHostAddress() As String
    TlsHostAddress = m_sTlsHostAddress
End Property

'=========================================================================
' Methods
'=========================================================================

''
' Connects to a TCP host with or w/o negotiating TLS channel
'
Public Function Connect( _
            HostAddress As String, _
            ByVal HostPort As Long, _
            Optional ByVal UseTls As Boolean) As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    Close_
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.SyncConnect(HostAddress, HostPort, Timeout:=m_lConnectTimeout) Then
        pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncConnect"
        GoTo QH
    End If
    If UseTls Then
        If Not StartTls(HostAddress) Then
            GoTo QH
        End If
    End If
    '--- success
    Connect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

''
' Disconects from remote host, shutting down TLS channel if present
'
Public Sub Close_()
    '--- note: used in terminate -> no error handling
    pvTlsClose
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If
End Sub

''
' Starts or re-negotiates TLS channel. No method provided to revert to plain socket
'
Public Function StartTls(HostAddress As String) As Boolean
    Const FUNC_NAME     As String = "StartTls"
    Dim uCred           As SCHANNEL_CRED
    
    On Error GoTo EH
    pvTlsClose
    uCred.dwVersion = SCHANNEL_CRED_VERSION
    uCred.dwFlags = uCred.dwFlags Or SCH_CRED_MANUAL_CRED_VALIDATION    ' Prevent Schannel from validating the received server certificate chain.
    uCred.dwFlags = uCred.dwFlags Or SCH_CRED_NO_DEFAULT_CREDS          ' Prevent Schannel from attempting to automatically supply a certificate chain for client authentication.
    If AcquireCredentialsHandle(0, UNISP_NAME, SECPKG_CRED_OUTBOUND, 0, uCred, 0, 0, m_hTlsCredentials, 0) <> 0 Then
        GoTo QH
    End If
    m_sTlsHostAddress = HostAddress
    If Not pvTlsHandshake(m_sTlsHostAddress, m_baTlsExtra) Then
        GoTo QH
    End If
    If QueryContextAttributes(m_hTlsContext, SECPKG_ATTR_STREAM_SIZES, m_uTlsSizes) <> 0 Then
        GoTo QH
    End If
    '--- success
    StartTls = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

''
' Receives text from remote host, optionally over TLS channel. Blocks until remote host responds
' or read timeout is reached. On error returns empty string and sets LastError property.
'
Public Function ReadText(Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Dim baBuffer()      As Byte
    
    If ReadArray(baBuffer) Then
        ReadText = m_oSocket.FromTextArray(baBuffer, CodePage)
    End If
End Function

''
' Receives data from remote host, optionally over TLS channel. Blocks until remote host responds
' or read timeout is reached. On error returns False and sets LastError property.
'
Public Function ReadArray(Buffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "ReadArray"
    Dim lAvailable      As Long
    Dim lBytes          As Long
    Dim lReceived       As Long
    
    On Error GoTo EH
    If m_hTlsContext = 0 Then
        lAvailable = m_oSocket.AvailableBytes
        If lAvailable <= 0 Then
            lAvailable = m_oSocket.SockOpt(ucsSsoReceiveBuffer)
        End If
        If lAvailable <= 0 Then
            lAvailable = 4096
        End If
        ReDim Buffer(0 To lBytes + lAvailable - 1) As Byte
        Do
            If Not m_oSocket.SyncReceive(VarPtr(Buffer(lBytes)), lAvailable, lReceived, Timeout:=m_lReadTimeout) Then
                pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncReceive"
                GoTo QH
            End If
        Loop While lReceived = 0
#If ImplDebug Then
        Debug.Print "Read recv " & lReceived
        Debug.Print DesignDumpMemory(VarPtr(Buffer(lBytes)), IIf(lReceived > 32, 32, lReceived))
#End If
        If lBytes + lReceived - 1 <> UBound(Buffer) Then
            ReDim Preserve Buffer(0 To lBytes + lReceived - 1) As Byte
        End If
    Else
        If UBound(m_baTlsExtra) >= 0 Then
            Buffer = m_baTlsExtra
            GoTo InLoop
        End If
        Do
            lAvailable = m_oSocket.AvailableBytes
            If lAvailable <= 0 Then
                lAvailable = m_oSocket.SockOpt(ucsSsoReceiveBuffer)
            End If
            If lAvailable <= 0 Then
                lAvailable = 4096
            End If
            lBytes = UBound(m_baTlsExtra) + 1
            ReDim Buffer(0 To lBytes + lAvailable - 1) As Byte
            If UBound(m_baTlsExtra) >= 0 Then
                CopyMemory Buffer(0), m_baTlsExtra(0), UBound(m_baTlsExtra) + 1
            End If
            Do
                If Not m_oSocket.SyncReceive(VarPtr(Buffer(lBytes)), lAvailable, lReceived, Timeout:=m_lReadTimeout) Then
                    pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncReceive"
                    GoTo QH
                End If
            Loop While lReceived = 0
#If ImplDebug Then
            Debug.Print "Read recv " & lReceived
            Debug.Print DesignDumpMemory(VarPtr(Buffer(lBytes)), IIf(lReceived > 32, 32, lReceived))
#End If
            If lBytes + lReceived - 1 <> UBound(Buffer) Then
                ReDim Preserve Buffer(0 To lBytes + lReceived - 1) As Byte
            End If
InLoop:
            m_baTlsExtra = vbNullString
            If Not pvTlsDecrypt(Buffer, m_sTlsHostAddress, m_baTlsExtra) Then
                GoTo QH
            End If
            If UBound(Buffer) >= 0 Then
                Exit Do
            End If
        Loop
    End If
    '--- succees
    ReadArray = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

''
' Next two sends text/data to remote host, optionally over TLS channel. Block until all data is
' transmitted or write timeout is reached. On error return False and set LastError property.
'
Public Function WriteText(Text As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Boolean
    WriteText = WriteArray(m_oSocket.ToTextArray(Text, CodePage))
End Function

Public Function WriteArray(baBuffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "WriteArray"
    Dim baEncr()        As Byte
    
    On Error GoTo EH
    If m_hTlsContext = 0 Then
        If Not m_oSocket.SyncSend(VarPtr(baBuffer(0)), UBound(baBuffer) + 1, Timeout:=m_lWriteTimeout) Then
            pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncSend"
        End If
    Else
        If Not pvTlsEncrypt(baBuffer, baEncr) Then
            GoTo QH
        End If
#If ImplDebug Then
        Debug.Print "Write encr " & UBound(baEncr) + 1
        Debug.Print DesignDumpMemory(VarPtr(baEncr(0)), IIf(UBound(baEncr) + 1 > 32, 32, UBound(baEncr) + 1))
#End If
        If Not m_oSocket.SyncSend(VarPtr(baEncr(0)), UBound(baEncr) + 1, Timeout:=m_lWriteTimeout) Then
            pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncSend"
        End If
    End If
    '--- success
    WriteArray = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal WriteTimeout As Long, ByVal ReadTimeout As Long)
    m_lResolveTimeout = ResolveTimeout
    m_lConnectTimeout = ConnectTimeout
    m_lWriteTimeout = WriteTimeout
    m_lReadTimeout = ReadTimeout
End Sub

'= private ===============================================================

''
' Create m_hTlsContext (if empty) and perform TLS handshake
'
Private Function pvTlsHandshake(sHostAddress As String, baExtra() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvTlsHandshake"
    Dim lContextReq     As Long
    Dim lContextAttr    As Long
    Dim uInDesc         As ApiSecBufferDesc
    Dim uInBuffers()    As ApiSecBuffer
    Dim uOutDesc        As ApiSecBufferDesc
    Dim uOutBuffers()   As ApiSecBuffer
    Dim lResult         As Long
    Dim baBuffer()      As Byte
    Dim lBytes          As Long
    Dim bNeedRecv       As Boolean
    Dim lRepeat         As Long
    Dim lIdx            As Long

    On Error GoTo EH
    ReDim baBuffer(0 To m_oSocket.SockOpt(ucsSsoReceiveBuffer) - 1) As Byte
'    lContextReq = lContextReq Or ISC_REQ_MUTUAL_AUTH               ' The mutual authentication policy of the service will be satisfied.
    lContextReq = lContextReq Or ISC_REQ_REPLAY_DETECT              ' Detect replayed messages that have been encoded by using the EncryptMessage or MakeSignature functions.
    lContextReq = lContextReq Or ISC_REQ_SEQUENCE_DETECT            ' Detect messages received out of sequence.
    lContextReq = lContextReq Or ISC_REQ_CONFIDENTIALITY            ' Encrypt messages by using the EncryptMessage function.
'    lContextReq = lContextReq Or ISC_REQ_USE_SUPPLIED_CREDS         ' Schannel must not attempt to supply credentials for the client automatically.
    lContextReq = lContextReq Or ISC_REQ_ALLOCATE_MEMORY            ' The security package allocates output buffers for you. When you have finished using the output buffers, free them by calling the FreeContextBuffer function.
'    lContextReq = lContextReq Or ISC_REQ_CONNECTION                ' The security context will not handle formatting messages.
    lContextReq = lContextReq Or ISC_REQ_EXTENDED_ERROR             ' When errors occur, the remote party will be notified.
    lContextReq = lContextReq Or ISC_REQ_STREAM                     ' Support a stream-oriented connection.
'    lContextReq = lContextReq Or ISC_REQ_INTEGRITY                 ' Sign messages and verify signatures by using the EncryptMessage and MakeSignature functions.
'    lContextReq = lContextReq Or ISC_REQ_MANUAL_CRED_VALIDATION    ' Schannel must not authenticate the server automatically.
    pvInitSecDesc uInDesc, 2, uInBuffers
    pvInitSecDesc uOutDesc, 3, uOutBuffers
    For lRepeat = 0 To 1000
        pvInitSecBuffer uOutBuffers(0), SECBUFFER_TOKEN
        pvInitSecBuffer uOutBuffers(1), SECBUFFER_ALERT
        For lIdx = 2 To uOutDesc.cBuffers - 1
            pvInitSecBuffer uOutBuffers(lIdx), SECBUFFER_EMPTY
        Next
        lContextAttr = 0
        If m_hTlsContext = 0 Then
            '--- note: on initial loop don't pass [in] phContext and [in] pInput
            lResult = InitializeSecurityContext(m_hTlsCredentials, ByVal 0, ByVal sHostAddress, lContextReq, 0, _
                    SECURITY_NATIVE_DREP, ByVal 0, 0, m_hTlsContext, uOutDesc, lContextAttr, 0)
        Else
            If UBound(baExtra) >= 0 Then
                If lBytes + UBound(baExtra) > UBound(baBuffer) Then
                    ReDim Preserve baBuffer(0 To lBytes + UBound(baExtra) + 1024) As Byte
                End If
                Call CopyMemory(baBuffer(lBytes), baExtra(0), UBound(baExtra) + 1)
                lBytes = lBytes + UBound(baExtra) + 1
                baExtra = vbNullString
            End If
            If bNeedRecv Then
                If Not m_oSocket.SyncReceive(VarPtr(baBuffer(lBytes)), UBound(baBuffer) + 1 - lBytes, lResult, Timeout:=m_lReadTimeout) Then
                    pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncReceive"
                    GoTo QH
                End If
#If ImplDebug Then
                Debug.Print "Handshake recv " & lResult
                Debug.Print DesignDumpMemory(VarPtr(baBuffer(lBytes)), IIf(lResult > 32, 32, lResult))
#End If
                lBytes = lBytes + lResult
            End If
            pvInitSecBuffer uInBuffers(0), SECBUFFER_TOKEN, lBytes, VarPtr(baBuffer(0))
            For lIdx = 1 To uInDesc.cBuffers - 1
                pvInitSecBuffer uInBuffers(lIdx), SECBUFFER_EMPTY
            Next
            '--- note: pass m_hTlsContext for [in] phContext and uInDesc for [in] pInput
            lResult = InitializeSecurityContext(m_hTlsCredentials, m_hTlsContext, ByVal sHostAddress, lContextReq, 0, _
                SECURITY_NATIVE_DREP, uInDesc, 0, m_hTlsContext, uOutDesc, lContextAttr, 0)
        End If
        If lContextReq <> lContextAttr Then
            pvSetError vbObjectError, FUNC_NAME, Replace(Replace(STR_ERR_CONTEXT_MISMATCH, "%1", Hex(lContextReq)), "%2", Hex(lContextAttr))
            GoTo QH
        End If
        '--- note: these success codes are generally not returned in TLS/SSL case
        Select Case lResult
        Case SEC_I_COMPLETE_AND_CONTINUE, SEC_I_COMPLETE_NEEDED
            If CompleteAuthToken(m_hTlsContext, uOutDesc) < 0 Then
                pvSetError Err.LastDllError, FUNC_NAME & vbCrLf & "CompleteAuthToken"
                GoTo QH
            End If
        End Select
        For lIdx = 0 To UBound(uOutBuffers)
            With uOutBuffers(lIdx)
                If .BufferType = SECBUFFER_TOKEN And .cbBuffer > 0 Then
#If ImplDebug Then
                    Debug.Print "Handshake send " & .cbBuffer
                    Debug.Print DesignDumpMemory(.pvBuffer, .cbBuffer)
#End If
                    If Not m_oSocket.SyncSend(.pvBuffer, .cbBuffer, Timeout:=m_lWriteTimeout) Then
                        pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncSend"
                        GoTo QH
                    End If
                ElseIf .BufferType = SECBUFFER_ALERT And .cbBuffer > 0 Then
#If ImplDebug Then
                    Debug.Print "Handshake alert " & .cbBuffer
                    Debug.Print DesignDumpMemory(.pvBuffer, .cbBuffer)
#End If
                End If
                If .pvBuffer <> 0 Then
                    Call FreeContextBuffer(.pvBuffer)
                    .pvBuffer = 0
                End If
            End With
        Next
        lBytes = 0
        For lIdx = 0 To UBound(uInBuffers)
            With uInBuffers(lIdx)
                If .BufferType = SECBUFFER_EXTRA And .cbBuffer > 0 Then
                    '-- note: if input buffers are not allocated by Schannel -> pvBuffer is NULL not to be freed w/ FreeContextBuffer
                    If .pvBuffer = 0 Then
                        .pvBuffer = VarPtr(baBuffer(uInBuffers(0).cbBuffer - .cbBuffer))
                    End If
                    Call CopyMemory(baBuffer(lBytes), ByVal .pvBuffer, .cbBuffer)
                    lBytes = lBytes + .cbBuffer
                End If
            End With
        Next
        bNeedRecv = True
        Select Case lResult
        Case SEC_E_OK, SEC_I_COMPLETE_NEEDED
            '--- success
            Exit For
        Case SEC_I_COMPLETE_AND_CONTINUE, SEC_I_CONTINUE_NEEDED, SEC_E_INCOMPLETE_MESSAGE
            '--- continue
        Case SEC_I_INCOMPLETE_CREDENTIALS
            '--- ToDo: GetNewClientCredentials
#If ImplDebug Then
            Debug.Print "Schannel will attempt to locate a client certificate and send it to the server", Timer
#End If
            If (lContextReq And ISC_REQ_USE_SUPPLIED_CREDS) = 0 Then
                lContextReq = lContextReq Or ISC_REQ_USE_SUPPLIED_CREDS
            Else
                pvSetError vbObjectError, FUNC_NAME, STR_ERR_CLIENT_CERT_REQUIRED
                GoTo QH
            End If
            bNeedRecv = False
            '--- continue
        Case Is < 0
            '--- failed
            pvSetError lResult, FUNC_NAME & vbCrLf & "InitializeSecurityContext"
            GoTo QH
        Case Else
            pvSetError vbObjectError, FUNC_NAME, Replace(STR_ERR_UNEXPECTED_RESULT1, "%1", Hex(lResult))
            GoTo QH
        End Select
    Next
    If lBytes > 0 Then
        '--- note: store remaining baBuffer as a look-ahead read buffer
        ReDim baExtra(0 To lBytes - 1) As Byte
        Call CopyMemory(baExtra(0), baBuffer(0), lBytes)
    Else
        baExtra = vbNullString
    End If
    '--- success
    pvTlsHandshake = True
    Exit Function
QH:
    pvFreeSecBuffers uOutBuffers
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

''
' Gracefully shutdown TLS channel
'
Private Function pvTlsShutdown() As Boolean
    Const FUNC_NAME     As String = "pvTlsShutdown"
    Dim lType           As Long
    Dim uCtlDesc        As ApiSecBufferDesc
    Dim uCtlBuffers()   As ApiSecBuffer
    Dim uOutDesc        As ApiSecBufferDesc
    Dim uOutBuffers()   As ApiSecBuffer
    Dim lResult         As Long
    Dim lContextReq     As Long
    Dim lContextAttr    As Long
    Dim lIdx            As Long
    
    On Error GoTo QH
    lType = SCHANNEL_SHUTDOWN
    pvInitSecDesc uCtlDesc, 1, uCtlBuffers
    pvInitSecBuffer uCtlBuffers(0), SECBUFFER_TOKEN, 4, VarPtr(lType)
    lResult = ApplyControlToken(m_hTlsContext, uCtlDesc)
    If lResult < 0 Then
        pvSetError lResult, FUNC_NAME & vbCrLf & "ApplyControlToken"
        GoTo QH
    End If
    '--- note: same as in pvTlsHandshake
    lContextReq = ISC_REQ_REPLAY_DETECT Or ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_CONFIDENTIALITY _
        Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_EXTENDED_ERROR Or ISC_REQ_STREAM
    pvInitSecDesc uOutDesc, 1, uOutBuffers
    pvInitSecBuffer uOutBuffers(0), SECBUFFER_TOKEN
    lResult = InitializeSecurityContext(m_hTlsCredentials, m_hTlsContext, ByVal 0, lContextReq, 0, _
        0, ByVal 0, 0, m_hTlsContext, uOutDesc, lContextAttr, 0)
    If lResult < 0 Then
        pvSetError lResult, FUNC_NAME & vbCrLf & "InitializeSecurityContext"
        GoTo QH
    End If
    For lIdx = 0 To UBound(uOutBuffers)
        With uOutBuffers(lIdx)
            If .BufferType = SECBUFFER_TOKEN And .cbBuffer > 0 Then
#If ImplDebug Then
                Debug.Print "Shutdown send " & .cbBuffer
                Debug.Print DesignDumpMemory(.pvBuffer, .cbBuffer)
#End If
                If Not m_oSocket.SyncSend(.pvBuffer, .cbBuffer, Timeout:=m_lWriteTimeout) Then
                    pvSetError m_oSocket.LastError, FUNC_NAME & vbCrLf & "cAsyncSocket.SyncSend"
                    GoTo QH
                End If
            End If
        End With
    Next
    '--- success
    pvTlsShutdown = True
QH:
    pvFreeSecBuffers uOutBuffers
    Exit Function
EH:
    PrintError FUNC_NAME
    GoTo QH
End Function

''
' Tries to decrypt [in] baBuffer and leave excess data in [out] baExtra
' Param [in] sHostAddress used only to re-negotiate TLS channel
'
Private Function pvTlsDecrypt(baBuffer() As Byte, sHostAddress As String, baExtra() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvTlsDecrypt"
    Dim uMessage        As ApiSecBufferDesc
    Dim uMsgBuffers()   As ApiSecBuffer
    Dim lResult         As Long
    Dim lIdx            As Long
    Dim baData()        As Byte
    
    On Error GoTo EH
    pvInitSecDesc uMessage, m_uTlsSizes.cBuffers, uMsgBuffers
    pvInitSecBuffer uMsgBuffers(0), SECBUFFER_DATA, UBound(baBuffer) + 1, VarPtr(baBuffer(0))
    lResult = DecryptMessage(m_hTlsContext, uMessage, 0, 0)
    Select Case lResult
    Case SEC_E_INCOMPLETE_MESSAGE
        baExtra = baBuffer
        baBuffer = vbNullString
    Case SEC_I_CONTEXT_EXPIRED
        baBuffer = vbNullString
    Case SEC_E_OK, SEC_I_RENEGOTIATE
        baData = vbNullString
        For lIdx = 0 To UBound(uMsgBuffers)
            With uMsgBuffers(lIdx)
                If .cbBuffer > 0 Then
                    Select Case .BufferType
                    Case SECBUFFER_DATA
                        ReDim Preserve baData(0 To UBound(baData) + .cbBuffer) As Byte
                        Call CopyMemory(baData(UBound(baData) + 1 - .cbBuffer), ByVal .pvBuffer, .cbBuffer)
                    Case SECBUFFER_EXTRA
                        ReDim Preserve baExtra(0 To UBound(baExtra) + .cbBuffer) As Byte
                        Call CopyMemory(baExtra(UBound(baExtra) + 1 - .cbBuffer), ByVal .pvBuffer, .cbBuffer)
                    End Select
                End If
            End With
        Next
        baBuffer = baData
        If lResult = SEC_I_RENEGOTIATE Then
            If Not pvTlsHandshake(sHostAddress, baExtra) Then
                GoTo QH
            End If
        End If
    Case Is < 0
        pvSetError lResult, FUNC_NAME & vbCrLf & "DecryptMessage"
        GoTo QH
    Case Else
        pvSetError vbObjectError, FUNC_NAME, Replace(STR_ERR_UNEXPECTED_RESULT2, "%1", Hex(lResult))
        GoTo QH
    End Select
    '--- success
    pvTlsDecrypt = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

''
' Tries to encrypt [in] baBuffer and stores result in [out] baEncr
'
Private Function pvTlsEncrypt(baBuffer() As Byte, baEncr() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvTlsEncrypt"
    Dim lBufSize        As Long
    Dim uMessage        As ApiSecBufferDesc
    Dim uMsgBuffers()   As ApiSecBuffer
    Dim lResult         As Long
    
    On Error GoTo EH
    lBufSize = UBound(baBuffer) + 1
    ReDim baEncr(0 To m_uTlsSizes.cbHeader + lBufSize + m_uTlsSizes.cbTrailer - 1) As Byte
    Call CopyMemory(baEncr(m_uTlsSizes.cbHeader), baBuffer(0), lBufSize)
    pvInitSecDesc uMessage, m_uTlsSizes.cBuffers, uMsgBuffers
    pvInitSecBuffer uMsgBuffers(0), SECBUFFER_STREAM_HEADER, m_uTlsSizes.cbHeader, VarPtr(baEncr(0))
    pvInitSecBuffer uMsgBuffers(1), SECBUFFER_DATA, lBufSize, VarPtr(baEncr(m_uTlsSizes.cbHeader))
    pvInitSecBuffer uMsgBuffers(2), SECBUFFER_STREAM_TRAILER, m_uTlsSizes.cbTrailer, VarPtr(baEncr(m_uTlsSizes.cbHeader + lBufSize))
    lResult = EncryptMessage(m_hTlsContext, 0, uMessage, 0)
    Select Case lResult
    Case SEC_E_OK
        '--- do nothing
    Case Is < 0
        pvSetError lResult, FUNC_NAME & vbCrLf & "EncryptMessage"
        GoTo QH
    Case Else
        pvSetError vbObjectError, FUNC_NAME, Replace(STR_ERR_UNEXPECTED_RESULT3, "%1", Hex(lResult))
        GoTo QH
    End Select
    '--- success
    pvTlsEncrypt = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvTlsClose()
    '--- note: used in terminate -> no error handling
    If m_hTlsCredentials <> 0 And m_hTlsContext <> 0 Then
        pvTlsShutdown
    End If
    If m_hTlsCredentials <> 0 Then
        Call FreeCredentialsHandle(m_hTlsCredentials)
        m_hTlsCredentials = 0
    End If
    If m_hTlsContext <> 0 Then
        Call DeleteSecurityContext(m_hTlsContext)
    End If
    m_baTlsExtra = vbNullString
End Sub

Private Sub pvSetError( _
            ByVal ErrorNumber As Long, _
            ErrorSource As String, _
            Optional ErrorDescription As String)
    m_lLastErrNumber = ErrorNumber
    m_sLastErrSource = MODULE_NAME & "." & ErrorSource
    If LenB(ErrorDescription) = 0 And Not m_oSocket Is Nothing Then
        m_sLastErrDescription = m_oSocket.GetErrorDescription(ErrorNumber)
    Else
        m_sLastErrDescription = ErrorDescription
    End If
End Sub

'= Schannel buffers helpers ==============================================

Private Sub pvInitSecDesc(uDesc As ApiSecBufferDesc, ByVal lCount As Long, uBuffers() As ApiSecBuffer)
    ReDim uBuffers(0 To lCount - 1)
    With uDesc
        .ulVersion = SECBUFFER_VERSION
        .cBuffers = lCount
        .pBuffers = VarPtr(uBuffers(0))
    End With
End Sub

Private Sub pvInitSecBuffer(uBuffer As ApiSecBuffer, ByVal lType As Long, Optional ByVal lSize As Long, Optional ByVal lPtr As Long)
    With uBuffer
        .BufferType = lType
        .cbBuffer = lSize
        .pvBuffer = lPtr
    End With
End Sub

Private Sub pvFreeSecBuffers(uBuffers() As ApiSecBuffer)
    Dim lIdx            As Long
    
    For lIdx = 0 To UBound(uBuffers)
        With uBuffers(lIdx)
            If .pvBuffer <> 0 Then
                Call FreeContextBuffer(.pvBuffer)
                .pvBuffer = 0
            End If
        End With
    Next
End Sub

#If ImplDebug Then
Private Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(UnsignedAdd(lPtr, lIdx), 1) = 0 Then
                Call CopyMemory(lValue, ByVal UnsignedAdd(lPtr, lIdx), 1)
                sHex = sHex & Right$("00" & Hex$(lValue), 2) & " "
                If lValue >= 32 Then
                    sChar = sChar & Chr$(lValue)
                Else
                    sChar = sChar & "."
                End If
            Else
                sHex = sHex & "?? "
                sChar = sChar & "."
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            DesignDumpMemory = DesignDumpMemory & Right$("0000" & Hex$(lIdx - 15), 4) & ": " & sHex & " " & sChar & vbCrLf
            sHex = vbNullString
            sChar = vbNullString
        End If
    Next
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
    m_lResolveTimeout = 0
    m_lConnectTimeout = 60000
    m_lWriteTimeout = 30000
    m_lReadTimeout = 30000
End Sub

Private Sub Class_Terminate()
    Close_
End Sub
