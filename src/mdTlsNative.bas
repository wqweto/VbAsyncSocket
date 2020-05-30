Attribute VB_Name = "mdTlsNative"
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2020 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdTlsNative"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)

'=========================================================================
' API
'=========================================================================

'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40
'--- for AcquireCredentialsHandle
Private Const UNISP_NAME                                As String = "Microsoft Unified Security Protocol Provider"
Private Const SECPKG_CRED_INBOUND                       As Long = 1
Private Const SECPKG_CRED_OUTBOUND                      As Long = 2
Private Const SCHANNEL_CRED_VERSION                     As Long = 4
Private Const SCH_CRED_MANUAL_CRED_VALIDATION           As Long = 8
Private Const SCH_CRED_NO_DEFAULT_CREDS                 As Long = &H10
'-- for InitializeSecurityContext
Private Const ISC_REQ_REPLAY_DETECT                     As Long = &H4
Private Const ISC_REQ_SEQUENCE_DETECT                   As Long = &H8
Private Const ISC_REQ_CONFIDENTIALITY                   As Long = &H10
Private Const ISC_REQ_USE_SUPPLIED_CREDS                As Long = &H80
Private Const ISC_REQ_ALLOCATE_MEMORY                   As Long = &H100
Private Const ISC_REQ_EXTENDED_ERROR                    As Long = &H4000
Private Const ISC_REQ_STREAM                            As Long = &H8000&
Private Const SECURITY_NATIVE_DREP                      As Long = &H10
'--- for ApiSecBuffer.BufferType
Private Const SECBUFFER_EMPTY                           As Long = 0   ' Undefined, replaced by provider
Private Const SECBUFFER_DATA                            As Long = 1   ' Packet data
Private Const SECBUFFER_TOKEN                           As Long = 2   ' Security token
Private Const SECBUFFER_EXTRA                           As Long = 5   ' Extra data
Private Const SECBUFFER_STREAM_TRAILER                  As Long = 6   ' Security Trailer
Private Const SECBUFFER_STREAM_HEADER                   As Long = 7   ' Security Header
Private Const SECBUFFER_ALERT                           As Long = 17
Private Const SECBUFFER_VERSION                         As Long = 0
'--- SSPI/Schannel retvals
Private Const SEC_E_OK                                  As Long = 0
Private Const SEC_I_CONTINUE_NEEDED                     As Long = &H90312
Private Const SEC_I_CONTEXT_EXPIRED                     As Long = &H90317
Private Const SEC_I_INCOMPLETE_CREDENTIALS              As Long = &H90320
Private Const SEC_I_RENEGOTIATE                         As Long = &H90321
Private Const SEC_E_INCOMPLETE_MESSAGE                  As Long = &H80090318
'--- for QueryContextAttributes
Private Const SECPKG_ATTR_STREAM_SIZES                  As Long = 4
Private Const SECPKG_ATTR_REMOTE_CERT_CONTEXT           As Long = &H53
Private Const SECPKG_ATTR_ISSUER_LIST_EX                As Long = &H59
'--- for ApplyControlToken
Private Const SCHANNEL_SHUTDOWN                         As Long = 1   ' gracefully close down a connection
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
'--- for CertOpenStore
Private Const CERT_STORE_PROV_MEMORY                    As Long = 2
Private Const CERT_STORE_CREATE_NEW_FLAG                As Long = &H2000
'--- for CertAddEncodedCertificateToStore
Private Const CERT_STORE_ADD_USE_EXISTING               As Long = 2
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_NEWKEYSET                           As Long = &H8
Private Const CRYPT_MACHINE_KEYSET                      As Long = &H20
Private Const AT_KEYEXCHANGE                            As Long = 1
'--- for CertGetCertificateContextProperty
Private Const CERT_KEY_PROV_INFO_PROP_ID                As Long = 2
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_ECC_PUBLIC_KEY                      As String = "1.2.840.10045.2.1"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"
'--- NCrypt
Private Const BCRYPT_ECDSA_PRIVATE_P256_MAGIC           As Long = &H32534345
Private Const BCRYPT_ECDSA_PRIVATE_P384_MAGIC           As Long = &H34534345
Private Const BCRYPT_ECDSA_PRIVATE_P521_MAGIC           As Long = &H36534345
Private Const MS_KEY_STORAGE_PROVIDER                   As String = "Microsoft Software Key Storage Provider"
Private Const NCRYPTBUFFER_PKCS_KEY_NAME                As Long = 45
Private Const NCRYPT_OVERWRITE_KEY_FLAG                 As Long = &H80

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
'--- security
Private Declare Function AcquireCredentialsHandle Lib "security" Alias "AcquireCredentialsHandleA" (ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, ByVal pvLogonId As Long, pAuthData As Any, ByVal pGetKeyFn As Long, ByVal pvGetKeyArgument As Long, phCredential As Currency, ByVal ptsExpiry As Long) As Long
Private Declare Function FreeCredentialsHandle Lib "security" (phContext As Currency) As Long
Private Declare Function AcceptSecurityContext Lib "security" (phCredential As Currency, ByVal phContext As Long, pInput As Any, ByVal fContextReq As Long, ByVal TargetDataRep As Long, phNewContext As Currency, pOutput As Any, pfContextAttr As Long, ByVal ptsExpiry As Long) As Long
Private Declare Function InitializeSecurityContext Lib "security" Alias "InitializeSecurityContextA" (phCredential As Currency, ByVal phContext As Long, pszTargetName As Any, ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, pInput As Any, ByVal Reserved2 As Long, phNewContext As Currency, pOutput As Any, pfContextAttr As Long, ByVal ptsExpiry As Long) As Long
Private Declare Function DeleteSecurityContext Lib "security" (phContext As Currency) As Long
Private Declare Function FreeContextBuffer Lib "security" (ByVal pvContextBuffer As Long) As Long
Private Declare Function QueryContextAttributes Lib "security" Alias "QueryContextAttributesA" (phContext As Currency, ByVal ulAttribute As Long, pBuffer As Any) As Long
Private Declare Function DecryptMessage Lib "secur32" (phContext As Currency, pMessage As Any, ByVal MessageSeqNo As Long, ByVal pfQOP As Long) As Long
Private Declare Function EncryptMessage Lib "secur32" (phContext As Currency, ByVal fQOP As Long, pMessage As Any, ByVal MessageSeqNo As Long) As Long
Private Declare Function ApplyControlToken Lib "secur32" (phContext As Currency, pInput As Any) As Long
'--- crypt32
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CertOpenStore Lib "crypt32" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
Private Declare Function CertCloseStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertAddEncodedCertificateToStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long, ByVal dwAddDisposition As Long, ByVal ppCertContext As Long) As Long
Private Declare Function CertSetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, ByVal dwFlags As Long, pvData As Any) As Long
Private Declare Function CertDuplicateCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
'--- ncrypt
Private Declare Function NCryptOpenStorageProvider Lib "ncrypt" (phProvider As Long, ByVal pszProviderName As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptImportKey Lib "ncrypt" (ByVal hProvider As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, phKey As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFreeObject Lib "ncrypt" (ByVal hObject As Long) As Long

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

Private Type CRYPT_KEY_PROV_INFO
    pwszContainerName   As Long
    pwszProvName        As Long
    dwProvType          As Long
    dwFlags             As Long
    cProvParam          As Long
    rgProvParam         As Long
    dwKeySpec           As Long
End Type

Private Type BCRYPT_ECCKEY_BLOB
    dwMagic             As Long
    cbKey               As Long
    Buffer(0 To 1000)   As Byte
End Type

Private Type CRYPT_BLOB_DATA
    cbData              As Long
    pbData              As Long
End Type

Private Type CRYPT_BIT_BLOB
    cbData              As Long
    pbData              As Long
    cUnusedBits         As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId            As Long
    Parameters          As CRYPT_BLOB_DATA
End Type

Private Type CERT_PUBLIC_KEY_INFO
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PublicKey           As CRYPT_BIT_BLOB
End Type

Private Type CRYPT_ECC_PRIVATE_KEY_INFO
    dwVersion           As Long
    PrivateKey          As CRYPT_BLOB_DATA
    szCurveOid          As Long
    PublicKey           As CRYPT_BLOB_DATA
End Type

Private Type CRYPT_PRIVATE_KEY_INFO
    Version             As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_BLOB_DATA
    pAttributes         As Long
End Type

Private Type CERT_CONTEXT
    dwCertEncodingType  As Long
    pbCertEncoded       As Long
    cbCertEncoded       As Long
    pCertInfo           As Long
    hCertStore          As Long
End Type

Private Type SecPkgContext_IssuerListInfoEx
    aIssuers            As Long
    cIssuers            As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS                 As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown certificate authority|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_UNKNOWN                   As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT              As String = "%1."
'--- errors
Private Const ERR_UNEXPECTED_RESULT         As String = "Unexpected result from %1 (%2)"
Private Const ERR_CONNECTION_CLOSED         As String = "Connection closed"
Private Const ERR_UNKNOWN_ECC_PRIVKEY       As String = "Unknown ECC private key (%1)"
Private Const ERR_UNKNOWN_PUBKEY            As String = "Unknown public key (%1)"
'--- numeric
Private Const TLS_CONTENT_TYPE_ALERT        As Long = 21
Private Const LNG_FACILITY_WIN32            As Long = &H80070000

Public Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 2
    ucsTlsSupportAll = ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum

Public Enum UcsTlsStatesEnum
    ucsTlsStateNew = 0
    ucsTlsStateClosed = 1
    ucsTlsStateHandshakeStart = 2
    ucsTlsStatePostHandshake = 8
    ucsTlsStateShutdown = 9
End Enum

Public Type UcsTlsContext
    '--- config
    IsServer            As Boolean
    RemoteHostName      As String
    LocalFeatures       As UcsTlsLocalFeaturesEnum
    OnClientCertificate As Long
    '--- state
    State               As UcsTlsStatesEnum
    LastErrNumber       As Long
    LastError           As String
    LastErrSource       As String
    LastAlertCode       As Long
    '--- handshake
    LocalCertificates   As Collection
    LocalPrivateKey     As Collection
    RemoteCertificates  As Collection
    '--- SSPI
    ContextReq          As Long
    hTlsCredentials     As Currency
    hTlsContext         As Currency
    TlsSizes            As ApiSecPkgContext_StreamSizes
    InDesc              As ApiSecBufferDesc
    InBuffers()         As ApiSecBuffer
    OutDesc             As ApiSecBufferDesc
    OutBuffers()        As ApiSecBuffer
    '--- I/O buffers
    RecvBuffer()        As Byte
    RecvPos             As Long
    SendBuffer()        As Byte
    SendPos             As Long
End Type

Private Type UcsKeyInfo
    AlgoObjId           As String
    KeyBlob()           As Byte
    BitLen              As Long
End Type

Public Enum UcsTlsAlertDescriptionsEnum
    uscTlsAlertCloseNotify = 0
    uscTlsAlertUnexpectedMessage = 10
    uscTlsAlertBadRecordMac = 20
    uscTlsAlertHandshakeFailure = 40
    uscTlsAlertBadCertificate = 42
    uscTlsAlertCertificateRevoked = 44
    uscTlsAlertCertificateExpired = 45
    uscTlsAlertCertificateUnknown = 46
    uscTlsAlertIllegalParameter = 47
    uscTlsAlertUnknownCa = 48
    uscTlsAlertDecodeError = 50
    uscTlsAlertDecryptError = 51
    uscTlsAlertProtocolVersion = 70
    uscTlsAlertInternalError = 80
    uscTlsAlertUserCanceled = 90
    uscTlsAlertMissingExtension = 109
    uscTlsAlertUnrecognizedName = 112
    uscTlsAlertCertificateRequired = 116
    uscTlsAlertNoApplicationProtocol = 120
End Enum

Public g_oRequestSocket             As cTlsSocket

'=========================================================================
' Properties
'=========================================================================

Public Property Get TlsIsClosed(uCtx As UcsTlsContext) As Boolean
    TlsIsClosed = (uCtx.State = ucsTlsStateClosed)
End Property

Public Property Get TlsIsStarted(uCtx As UcsTlsContext) As Boolean
    TlsIsStarted = (uCtx.State > ucsTlsStateClosed)
End Property

Public Property Get TlsIsReady(uCtx As UcsTlsContext) As Boolean
    TlsIsReady = (uCtx.State >= ucsTlsStatePostHandshake)
End Property

Public Property Get TlsIsShutdown(uCtx As UcsTlsContext) As Boolean
    TlsIsShutdown = (uCtx.State = ucsTlsStateShutdown)
End Property

'=========================================================================
' TLS support
'=========================================================================

Public Function TlsInitClient( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum = ucsTlsSupportAll, _
            Optional OnClientCertificate As Object) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    With uEmpty
        pvTlsSetLastError uEmpty
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        .OnClientCertificate = ObjPtr(OnClientCertificate)
    End With
    uCtx = uEmpty
    '--- success
    TlsInitClient = True
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
End Function

Public Function TlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    With uEmpty
        pvTlsSetLastError uEmpty
        .IsServer = True
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = ucsTlsSupportAll
        Set .LocalCertificates = Certificates
        Set .LocalPrivateKey = PrivateKey
    End With
    uCtx = uEmpty
    '--- success
    TlsInitServer = True
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
End Function

Public Function TlsTerminate(uCtx As UcsTlsContext)
    With uCtx
        .State = ucsTlsStateClosed
        If .hTlsContext <> 0 Then
            Call DeleteSecurityContext(.hTlsContext)
            .hTlsContext = 0
        End If
        If .hTlsCredentials <> 0 Then
            Call FreeCredentialsHandle(.hTlsCredentials)
            .hTlsCredentials = 0
        End If
    End With
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsHandshake"
    Dim uCred           As SCHANNEL_CRED
    Dim lContextAttr    As Long
    Dim hResult         As Long
    Dim lIdx            As Long
    Dim lPtr            As Long
    Dim oCallback       As Object
    Dim hMemStore       As Long
    Dim pCertContext    As Long
    Dim aCred(0 To 100) As Long
    Dim uIssuerInfo     As SecPkgContext_IssuerListInfoEx
    Dim uIssuerList()   As CRYPT_BLOB_DATA
    Dim cIssuers        As Collection
    Dim baCaDn()        As Byte
    Dim uCertContext    As CERT_CONTEXT
    Dim sApiSource      As String
    
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            GoTo QH
        End If
        pvTlsSetLastError uCtx
        If .ContextReq = 0 Then
            .ContextReq = .ContextReq Or ISC_REQ_REPLAY_DETECT              ' Detect replayed messages that have been encoded by using the EncryptMessage or MakeSignature functions.
            .ContextReq = .ContextReq Or ISC_REQ_SEQUENCE_DETECT            ' Detect messages received out of sequence.
            .ContextReq = .ContextReq Or ISC_REQ_CONFIDENTIALITY            ' Encrypt messages by using the EncryptMessage function.
            .ContextReq = .ContextReq Or ISC_REQ_ALLOCATE_MEMORY            ' The security package allocates output buffers for you. When you have finished using the output buffers, free them by calling the FreeContextBuffer function.
            .ContextReq = .ContextReq Or ISC_REQ_EXTENDED_ERROR             ' When errors occur, the remote party will be notified.
            .ContextReq = .ContextReq Or ISC_REQ_STREAM                     ' Support a stream-oriented connection.
        End If
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If lSize > 0 Then
            .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        End If
        If lSize = 7 Then
            If baInput(0) = TLS_CONTENT_TYPE_ALERT Then
                .LastAlertCode = baInput(6)
            End If
        End If
RetryCredentials:
        If .hTlsCredentials = 0 Then
            uCred.dwVersion = SCHANNEL_CRED_VERSION
            uCred.dwFlags = uCred.dwFlags Or SCH_CRED_MANUAL_CRED_VALIDATION    ' Prevent Schannel from validating the received server certificate chain.
            uCred.dwFlags = uCred.dwFlags Or SCH_CRED_NO_DEFAULT_CREDS          ' Prevent Schannel from attempting to automatically supply a certificate chain for client authentication.
            If pvCollectionCount(.LocalCertificates) > 0 Then
                If pvTlsImportToCertStore(.LocalCertificates, .LocalPrivateKey, hMemStore) Then
                    Do
                        pCertContext = CertEnumCertificatesInStore(hMemStore, pCertContext)
                        If pCertContext = 0 Then
                            Exit Do
                        End If
                        aCred(uCred.cCreds) = CertDuplicateCertificateContext(pCertContext)
                        uCred.cCreds = uCred.cCreds + 1
                        If Not .IsServer Then
                            Call CertFreeCertificateContext(pCertContext)
                            Exit Do
                        End If
                    Loop
                    Call CertCloseStore(hMemStore, 0)
                    uCred.paCred = VarPtr(aCred(0))
                    .ContextReq = .ContextReq Or ISC_REQ_USE_SUPPLIED_CREDS     ' Schannel must not attempt to supply credentials for the client automatically.
                End If
            End If
            hResult = AcquireCredentialsHandle(0, UNISP_NAME, IIf(.IsServer, SECPKG_CRED_INBOUND, SECPKG_CRED_OUTBOUND), 0, uCred, 0, 0, .hTlsCredentials, 0)
            If hResult < 0 Then
                pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "AcquireCredentialsHandle", , .LastAlertCode
                GoTo QH
            End If
            For lIdx = 0 To uCred.cCreds - 1
                Call CertFreeCertificateContext(aCred(lIdx))
            Next
        End If
        If .hTlsContext = 0 Then
            pvInitSecDesc .InDesc, 3, .InBuffers
            pvInitSecDesc .OutDesc, 3, .OutBuffers
        End If
        If .RecvPos > 0 Then
            lPtr = VarPtr(.RecvBuffer(0))
        Else
            lPtr = VarPtr(.RecvPos)
        End If
        pvInitSecBuffer .InBuffers(0), SECBUFFER_TOKEN, lPtr, .RecvPos
        If .IsServer Then
            hResult = AcceptSecurityContext(.hTlsCredentials, IIf(.hTlsContext <> 0, VarPtr(.hTlsContext), 0), .InDesc, .ContextReq, _
                SECURITY_NATIVE_DREP, .hTlsContext, .OutDesc, lContextAttr, 0)
            sApiSource = "AcceptSecurityContext"
        Else
            hResult = InitializeSecurityContext(.hTlsCredentials, IIf(.hTlsContext <> 0, VarPtr(.hTlsContext), 0), ByVal .RemoteHostName, .ContextReq, 0, _
                SECURITY_NATIVE_DREP, .InDesc, 0, .hTlsContext, .OutDesc, lContextAttr, 0)
            sApiSource = "InitializeSecurityContext"
        End If
        If hResult = SEC_E_INCOMPLETE_MESSAGE Then
            pvInitSecBuffer .InBuffers(1), SECBUFFER_EMPTY
        ElseIf hResult < 0 Then
            pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & sApiSource, , .LastAlertCode
            GoTo QH
        Else
            .RecvPos = 0
            For lIdx = 1 To UBound(.InBuffers)
                With .InBuffers(lIdx)
                    If .cbBuffer > 0 Then
                        Select Case .BufferType
                        Case SECBUFFER_EXTRA
                            lPtr = .pvBuffer
                            If lPtr = 0 Then
                                lPtr = VarPtr(uCtx.RecvBuffer(uCtx.InBuffers(0).cbBuffer - .cbBuffer))
                            End If
                            uCtx.RecvPos = pvWriteBuffer(uCtx.RecvBuffer, uCtx.RecvPos, lPtr, .cbBuffer)
                        Case SECBUFFER_ALERT
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, "InBuffers, SECBUFFER_ALERT:" & vbCrLf & DesignDumpMemory(.pvBuffer, .cbBuffer), vbLogEventTypeWarning
                            #End If
                        Case Else
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, ".BufferType(" & lIdx & ")=" & .BufferType
                            #End If
                        End Select
                    End If
                End With
                pvInitSecBuffer .InBuffers(lIdx), SECBUFFER_EMPTY
            Next
            For lIdx = 0 To UBound(.OutBuffers)
                With .OutBuffers(lIdx)
                    If .cbBuffer > 0 Then
                        Select Case .BufferType
                        Case SECBUFFER_TOKEN
                            lOutputPos = pvWriteBuffer(baOutput, lOutputPos, .pvBuffer, .cbBuffer)
                        Case SECBUFFER_ALERT
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, "OutBuffers, SECBUFFER_ALERT:" & vbCrLf & DesignDumpMemory(.pvBuffer, .cbBuffer), vbLogEventTypeWarning
                            #End If
                        End Select
                        If .pvBuffer <> 0 Then
                            Call FreeContextBuffer(.pvBuffer)
                            Debug.Assert Err.LastDllError = 0
                        End If
                    End If
                End With
                pvInitSecBuffer .OutBuffers(lIdx), SECBUFFER_EMPTY
            Next
            Select Case hResult
            Case SEC_E_OK
                If QueryContextAttributes(.hTlsContext, SECPKG_ATTR_STREAM_SIZES, .TlsSizes) <> 0 Then
                    GoTo QH
                End If
                pvInitSecDesc .InDesc, .TlsSizes.cBuffers, .InBuffers
                pvInitSecDesc .OutDesc, .TlsSizes.cBuffers, .OutBuffers
                If QueryContextAttributes(.hTlsContext, SECPKG_ATTR_REMOTE_CERT_CONTEXT, pCertContext) = 0 Then
                    Call CopyMemory(uCertContext, ByVal pCertContext, Len(uCertContext))
                    If Not pvTlsExportFromCertStore(uCertContext.hCertStore, .RemoteCertificates) Then
                        GoTo QH
                    End If
                End If
                .State = ucsTlsStatePostHandshake
            Case SEC_I_CONTINUE_NEEDED
                '--- do nothing
            Case SEC_I_INCOMPLETE_CREDENTIALS
                If .OnClientCertificate <> 0 Then
                    If QueryContextAttributes(.hTlsContext, SECPKG_ATTR_ISSUER_LIST_EX, uIssuerInfo) <> 0 Then
                        GoTo QH
                    End If
                    If uIssuerInfo.cIssuers > 0 Then
                        ReDim uIssuerList(0 To uIssuerInfo.cIssuers - 1) As CRYPT_BLOB_DATA
                        Call CopyMemory(uIssuerList(0), ByVal uIssuerInfo.aIssuers, uIssuerInfo.cIssuers * Len(uIssuerList(0)))
                        Set cIssuers = New Collection
                        For lIdx = 0 To UBound(uIssuerList)
                            pvWriteBuffer baCaDn, 0, uIssuerList(lIdx).pbData, uIssuerList(lIdx).cbData
                            pvArrayReallocate baCaDn, uIssuerList(lIdx).cbData, FUNC_NAME & ".baCaDn"
                            cIssuers.Add baCaDn
                        Next
                    End If
                    Call vbaObjSetAddref(oCallback, .OnClientCertificate)
                    If oCallback.FireOnClientCertificate(cIssuers) Then
                        Call FreeCredentialsHandle(.hTlsCredentials)
                        .hTlsCredentials = 0
                        GoTo RetryCredentials
                    End If
                ElseIf (.ContextReq And ISC_REQ_USE_SUPPLIED_CREDS) = 0 Then
                    .ContextReq = .ContextReq Or ISC_REQ_USE_SUPPLIED_CREDS
                    GoTo RetryCredentials
                End If
                pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME, , .LastAlertCode
                GoTo QH
            Case Else
                pvTlsSetLastError uCtx, vbObjectError, FUNC_NAME, Replace(Replace(ERR_UNEXPECTED_RESULT, "%1", sApiSource), "%2", "&H" & Hex$(hResult)), .LastAlertCode
                GoTo QH
            End Select
        End If
    End With
    '--- success
    TlsHandshake = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsReceive"
    Dim hResult         As Long
    Dim lIdx            As Long
    Dim lPtr            As Long
    
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            GoTo QH
        End If
        pvTlsSetLastError uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If lSize > 0 Then
            .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        End If
        Do
            If .RecvPos > 0 Then
                lPtr = VarPtr(.RecvBuffer(0))
            Else
                lPtr = VarPtr(.RecvPos)
            End If
            pvInitSecBuffer .InBuffers(0), SECBUFFER_DATA, lPtr, .RecvPos
            hResult = DecryptMessage(.hTlsContext, .InDesc, 0, 0)
            If hResult = SEC_E_INCOMPLETE_MESSAGE Then
                pvInitSecBuffer .InBuffers(1), SECBUFFER_EMPTY
                Exit Do
            ElseIf hResult < 0 Then
                pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "DecryptMessage"
                GoTo QH
            End If
            .RecvPos = 0
            For lIdx = 1 To UBound(.InBuffers)
                With .InBuffers(lIdx)
                    If .cbBuffer > 0 Then
                        Select Case .BufferType
                        Case SECBUFFER_DATA
                            lPos = pvWriteBuffer(baPlainText, lPos, .pvBuffer, .cbBuffer)
                        Case SECBUFFER_EXTRA
                            lPtr = .pvBuffer
                            If lPtr = 0 Then
                                lPtr = VarPtr(uCtx.RecvBuffer(uCtx.InBuffers(0).cbBuffer - .cbBuffer))
                            End If
                            uCtx.RecvPos = pvWriteBuffer(uCtx.RecvBuffer, uCtx.RecvPos, lPtr, .cbBuffer)
                        Case SECBUFFER_ALERT
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, "InBuffers, SECBUFFER_ALERT:" & vbCrLf & DesignDumpMemory(.pvBuffer, .cbBuffer), vbLogEventTypeWarning
                            #End If
                        Case SECBUFFER_STREAM_HEADER, SECBUFFER_STREAM_TRAILER
                            '--- do nothing
                        Case Else
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, ".BufferType(" & lIdx & ")=" & .BufferType
                            #End If
                        End Select
                    End If
                End With
                pvInitSecBuffer .InBuffers(lIdx), SECBUFFER_EMPTY
            Next
            Select Case hResult
            Case SEC_E_OK
                '--- do nothing
            Case SEC_I_RENEGOTIATE
                .State = ucsTlsStateHandshakeStart
                TlsHandshake uCtx, .RecvBuffer, .RecvPos, .SendBuffer, .SendPos
                Exit Do
            Case SEC_I_CONTEXT_EXPIRED
                .State = ucsTlsStateShutdown
                Exit Do
            Case Else
                pvTlsSetLastError uCtx, vbObjectError, FUNC_NAME, Replace(Replace(ERR_UNEXPECTED_RESULT, "%1", "DecryptMessage"), "%2", "&H" & Hex$(hResult))
                GoTo QH
            End Select
        Loop
    End With
    '--- success
    TlsReceive = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsSend"
    Dim hResult         As Long
    Dim lBufPos         As Long
    Dim lBufSize        As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            GoTo QH
        End If
        pvTlsSetLastError uCtx
        If lSize = 0 Then
            '--- flush
            If .SendPos > 0 Then
                lOutputPos = pvWriteBuffer(baOutput, lOutputPos, VarPtr(.SendBuffer(0)), .SendPos)
            End If
            '--- success
            TlsSend = True
            Exit Function
        End If
        '--- figure out upper bound of total output and reserve space in baOutput
        lIdx = (lSize + .TlsSizes.cbMaximumMessage - 1) \ .TlsSizes.cbMaximumMessage
        pvWriteReserved baOutput, lOutputPos, .TlsSizes.cbHeader * lIdx + lSize + .TlsSizes.cbTrailer * lIdx
        For lPos = 0 To lSize - 1 Step .TlsSizes.cbMaximumMessage
            lBufPos = lOutputPos
            lBufSize = lSize - lPos
            If lBufSize > .TlsSizes.cbMaximumMessage Then
                lBufSize = .TlsSizes.cbMaximumMessage
            End If
            pvWriteReserved baOutput, lOutputPos, .TlsSizes.cbHeader + lBufSize + .TlsSizes.cbTrailer
            pvInitSecBuffer .InBuffers(0), SECBUFFER_STREAM_HEADER, VarPtr(baOutput(lBufPos)), .TlsSizes.cbHeader
            lBufPos = lBufPos + .TlsSizes.cbHeader
            Call CopyMemory(baOutput(lBufPos), baPlainText(lPos), lBufSize)
            pvInitSecBuffer .InBuffers(1), SECBUFFER_DATA, VarPtr(baOutput(lBufPos)), lBufSize
            lBufPos = lBufPos + lBufSize
            pvInitSecBuffer .InBuffers(2), SECBUFFER_STREAM_TRAILER, VarPtr(baOutput(lBufPos)), .TlsSizes.cbTrailer
            For lIdx = 3 To UBound(.InBuffers)
                pvInitSecBuffer .InBuffers(lIdx), SECBUFFER_EMPTY
            Next
            hResult = EncryptMessage(.hTlsContext, 0, .InDesc, 0)
            If hResult < 0 Then
                pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "EncryptMessage"
                GoTo QH
            End If
            Debug.Assert .InBuffers(0).cbBuffer = .TlsSizes.cbHeader
            Debug.Assert .InBuffers(1).cbBuffer = lBufSize
            '--- note: trailing MAC might be returned by EncryptMessage shorter than initial .TlsSizes.cbTrailer
'            Debug.Assert .InBuffers(2).cbBuffer = .TlsSizes.cbTrailer
            lOutputPos = lOutputPos + .InBuffers(0).cbBuffer + .InBuffers(1).cbBuffer + .InBuffers(2).cbBuffer
            For lIdx = 1 To UBound(.InBuffers)
                With .InBuffers(lIdx)
                    If .cbBuffer > 0 Then
                        Select Case .BufferType
                        Case SECBUFFER_ALERT
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, "InBuffers, SECBUFFER_ALERT:" & vbCrLf & DesignDumpMemory(.pvBuffer, .cbBuffer), vbLogEventTypeWarning
                            #End If
                        Case SECBUFFER_DATA, SECBUFFER_STREAM_HEADER, SECBUFFER_STREAM_TRAILER
                            '--- do nothing
                        Case Else
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, ".BufferType(" & lIdx & ")=" & .BufferType
                            #End If
                        End Select
                    End If
                End With
                pvInitSecBuffer .InBuffers(lIdx), SECBUFFER_EMPTY
            Next
            Select Case hResult
            Case SEC_E_OK
                '--- do nothing
            Case Else
                pvTlsSetLastError uCtx, vbObjectError, FUNC_NAME, Replace(Replace(ERR_UNEXPECTED_RESULT, "%1", "EncryptMessage"), "%2", "&H" & Hex$(hResult))
                GoTo QH
            End Select
        Next
    End With
    '--- success
    TlsSend = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
    Const FUNC_NAME     As String = "pvTlsShutdown"
    Dim lType           As Long
    Dim hResult         As Long
    Dim lIdx            As Long
    Dim sApiSource      As String
    Dim lContextAttr    As Long
    
    On Error GoTo QH
    With uCtx
        If .State = ucsTlsStateClosed Or .State = ucsTlsStateShutdown Then
            '--- success
            TlsShutdown = True
            GoTo QH
        End If
        lType = SCHANNEL_SHUTDOWN
        pvInitSecBuffer .InBuffers(0), SECBUFFER_TOKEN, VarPtr(lType), 4
        hResult = ApplyControlToken(.hTlsContext, .InDesc)
        If hResult < 0 Then
            hResult = hResult
'            pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & "ApplyControlToken"
'            GoTo QH
        End If
        pvInitSecBuffer .OutBuffers(0), SECBUFFER_TOKEN
        For lIdx = 1 To UBound(.OutBuffers)
            pvInitSecBuffer .OutBuffers(lIdx), SECBUFFER_EMPTY
        Next
        If .IsServer Then
            hResult = AcceptSecurityContext(.hTlsCredentials, VarPtr(.hTlsContext), ByVal 0, .ContextReq, _
                SECURITY_NATIVE_DREP, .hTlsContext, .OutDesc, lContextAttr, 0)
            sApiSource = "AcceptSecurityContext"
        Else
            hResult = InitializeSecurityContext(.hTlsCredentials, VarPtr(.hTlsContext), ByVal .RemoteHostName, .ContextReq, 0, _
                SECURITY_NATIVE_DREP, ByVal 0, 0, .hTlsContext, .OutDesc, lContextAttr, 0)
            sApiSource = "InitializeSecurityContext"
        End If
        If hResult < 0 Then
            pvTlsSetLastError uCtx, hResult, MODULE_NAME & "." & FUNC_NAME & vbCrLf & sApiSource
            GoTo QH
        End If
        For lIdx = 0 To UBound(.OutBuffers)
            With .OutBuffers(lIdx)
                If .BufferType = SECBUFFER_TOKEN And .cbBuffer > 0 Then
                    lPos = pvWriteBuffer(baOutput, lPos, .pvBuffer, .cbBuffer)
                End If
                If .pvBuffer <> 0 Then
                    Call FreeContextBuffer(.pvBuffer)
                    Debug.Assert Err.LastDllError = 0
                    .pvBuffer = 0
                End If
            End With
        Next
        .State = ucsTlsStateShutdown
    End With
    '--- success
    TlsShutdown = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext, Optional LastErrNumber As Long) As String
    LastErrNumber = uCtx.LastErrNumber
    TlsGetLastError = uCtx.LastError
    If uCtx.LastAlertCode <> -1 Then
        TlsGetLastError = IIf(LenB(TlsGetLastError) <> 0, TlsGetLastError & ". ", vbNullString) & Replace(STR_FORMAT_ALERT, "%1", TlsGetLastAlert(uCtx))
    End If
End Function

Private Sub pvTlsSetLastError(uCtx As UcsTlsContext, Optional ByVal lNumber As Long, Optional sSource As String, Optional sDescription As String, Optional ByVal AlertDesc As UcsTlsAlertDescriptionsEnum = -1)
    With uCtx
        .LastErrNumber = lNumber
        .LastErrSource = sSource
        .LastAlertCode = AlertDesc
        If lNumber <> 0 And LenB(sDescription) = 0 Then
            With New cAsyncSocket
                uCtx.LastError = .GetErrorDescription(lNumber)
            End With
        Else
            .LastError = sDescription
        End If
        If Right$(.LastError, 2) = vbCrLf Then
            .LastError = Left$(.LastError, Len(.LastError) - 2)
        End If
        If Right$(.LastError, 1) = "." Then
            .LastError = Left$(.LastError, Len(.LastError) - 1)
        End If
        If .LastErrNumber <> 0 Then
            .State = ucsTlsStateClosed
        End If
    End With
End Sub

Private Function TlsGetLastAlert(uCtx As UcsTlsContext, Optional AlertCode As UcsTlsAlertDescriptionsEnum) As String
    Static vTexts       As Variant
    
    AlertCode = uCtx.LastAlertCode
    If AlertCode >= 0 Then
        If IsEmpty(vTexts) Then
            vTexts = SplitOrReindex(STR_VL_ALERTS, "|")
        End If
        If AlertCode <= UBound(vTexts) Then
            TlsGetLastAlert = vTexts(AlertCode)
        End If
        If LenB(TlsGetLastAlert) = 0 Then
            TlsGetLastAlert = Replace(STR_UNKNOWN, "%1", AlertCode)
        End If
    End If
End Function

Private Function pvTlsImportToCertStore(cCerts As Collection, cPrivKey As Collection, hMemStore As Long) As Boolean
    Const FUNC_NAME     As String = "pvTlsImportToCertStore"
    Const DEF_KEY_NAME  As String = "VbAsyncSocketKey"
    Dim hCertStore      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim pCertContext    As Long
    Dim baPrivKey()     As Byte
    Dim hProv           As Long
    Dim hKey            As Long
    Dim lPtr            As Long
    Dim uPrivKeyInfo    As UcsKeyInfo
    Dim uPublicKeyInfo  As CERT_PUBLIC_KEY_INFO
    Dim uProvInfo       As CRYPT_KEY_PROV_INFO
    Dim uEccBlob        As BCRYPT_ECCKEY_BLOB
    Dim lBlobSize       As Long
    Dim sKeyName        As String
    Dim sProvName       As String
    Dim hNProv          As Long
    Dim hNKey           As Long
    Dim uDesc           As ApiSecBufferDesc
    Dim uBuffers()      As ApiSecBuffer
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- load server X.509 certificates to an in-memory certificate store
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For lIdx = pvCollectionCount(cCerts) To 1 Step -1
        baCert = cCerts.Item(lIdx)
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, IIf(lIdx = 1, VarPtr(pCertContext), 0)) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    If pCertContext <> 0 And SearchCollection(cPrivKey, 1, RetVal:=baPrivKey) Then
        sKeyName = DEF_KEY_NAME
        If Not pvAsn1DecodePrivateKey(baPrivKey, uPrivKeyInfo) Then
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
        lPtr = UnsignedAdd(lPtr, 56)                                        '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
        Call CopyMemory(uPublicKeyInfo, ByVal lPtr, Len(uPublicKeyInfo))
        Select Case pvToString(uPublicKeyInfo.Algorithm.pszObjId)
        Case szOID_RSA_RSA
            If CryptAcquireContext(hProv, StrPtr(sKeyName), 0, PROV_RSA_FULL, CRYPT_MACHINE_KEYSET) = 0 Then
                If CryptAcquireContext(hProv, StrPtr(sKeyName), 0, PROV_RSA_FULL, CRYPT_NEWKEYSET Or CRYPT_MACHINE_KEYSET) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CryptAcquireContext"
                    GoTo QH
                End If
            End If
            If CryptImportKey(hProv, uPrivKeyInfo.KeyBlob(0), UBound(uPrivKeyInfo.KeyBlob) + 1, 0, 0, hKey) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptImportKey"
                GoTo QH
            End If
            uProvInfo.pwszContainerName = StrPtr(sKeyName)
            uProvInfo.dwProvType = PROV_RSA_FULL
            uProvInfo.dwFlags = CRYPT_MACHINE_KEYSET
            uProvInfo.dwKeySpec = AT_KEYEXCHANGE
        Case szOID_ECC_PUBLIC_KEY
            Select Case uPrivKeyInfo.AlgoObjId
            Case szOID_ECC_CURVE_P256
                uEccBlob.dwMagic = BCRYPT_ECDSA_PRIVATE_P256_MAGIC
            Case szOID_ECC_CURVE_P384
                uEccBlob.dwMagic = BCRYPT_ECDSA_PRIVATE_P384_MAGIC
            Case szOID_ECC_CURVE_P521
                uEccBlob.dwMagic = BCRYPT_ECDSA_PRIVATE_P521_MAGIC
            Case Else
                Err.Raise vbObjectError, , Replace(ERR_UNKNOWN_ECC_PRIVKEY, "%1", uPrivKeyInfo.AlgoObjId)
            End Select
            lBlobSize = uPublicKeyInfo.PublicKey.cbData - 1
            uEccBlob.cbKey = UBound(uPrivKeyInfo.KeyBlob) + 1
            Call CopyMemory(uEccBlob.Buffer(0), ByVal UnsignedAdd(uPublicKeyInfo.PublicKey.pbData, 1), lBlobSize)
            Call CopyMemory(uEccBlob.Buffer(lBlobSize), uPrivKeyInfo.KeyBlob(0), uEccBlob.cbKey)
            lBlobSize = 8 + lBlobSize + uEccBlob.cbKey
            '--- import key
            sProvName = MS_KEY_STORAGE_PROVIDER
            hResult = NCryptOpenStorageProvider(hNProv, StrPtr(sProvName), 0)
            If hResult < 0 Then
                sApiSource = "NCryptOpenStorageProvider"
                GoTo QH
            End If
            pvInitSecDesc uDesc, 1, uBuffers
            pvInitSecBuffer uBuffers(0), NCRYPTBUFFER_PKCS_KEY_NAME, StrPtr(sKeyName), LenB(sKeyName) + 2
            hResult = NCryptImportKey(hNProv, 0, StrPtr("ECCPRIVATEBLOB"), uDesc, hNKey, uEccBlob, lBlobSize, NCRYPT_OVERWRITE_KEY_FLAG)
            If hResult < 0 Then
                sApiSource = "NCryptImportKey"
                GoTo QH
            End If
            uProvInfo.pwszContainerName = StrPtr(sKeyName)
            uProvInfo.pwszProvName = StrPtr(sProvName)
        Case Else
            Err.Raise vbObjectError, , Replace(ERR_UNKNOWN_PUBKEY, "%1", pvToString(uPublicKeyInfo.Algorithm.pszObjId))
        End Select
        If CertSetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, 0, uProvInfo) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertSetCertificateContextProperty"
            GoTo QH
        End If
    End If
    hMemStore = hCertStore
    hCertStore = 0
    '--- success
    pvTlsImportToCertStore = True
QH:
    If hNKey <> 0 Then
        Call NCryptFreeObject(hNKey)
    End If
    If hNProv <> 0 Then
        Call NCryptFreeObject(hNProv)
    End If
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvTlsExportFromCertStore(ByVal hCertStore As Long, cCerts As Collection) As Boolean
    Const FUNC_NAME     As String = "pvTlsExportFromCertStore"
    Dim uCertContext        As CERT_CONTEXT
    Dim baCert()        As Byte
    Dim pCertContext    As Long

    '--- export server X.509 certificates from certificate store
    Set cCerts = New Collection
    Do
        pCertContext = CertEnumCertificatesInStore(hCertStore, pCertContext)
        If pCertContext = 0 Then
            Exit Do
        End If
        Call CopyMemory(uCertContext, ByVal pCertContext, Len(uCertContext))
        pvWriteBuffer baCert, 0, uCertContext.pbCertEncoded, uCertContext.cbCertEncoded
        pvArrayReallocate baCert, uCertContext.cbCertEncoded, FUNC_NAME & ".baCert"
        cCerts.Add baCert
    Loop
    '--- success
    pvTlsExportFromCertStore = True
End Function

Private Function pvAsn1DecodePrivateKey(baPrivKey() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "pvAsn1DecodePrivateKey"
    Dim lPkiPtr         As Long
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim lSize           As Long
    Dim uEccKeyInfo     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) <> 0 Then
        Call CopyMemory(uPrivKey, ByVal lPkiPtr, Len(uPrivKey))
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uPrivKey.PrivateKey.pbData, uPrivKey.PrivateKey.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)"
            GoTo QH
        End If
        uRetVal.AlgoObjId = pvToString(uPrivKey.Algorithm.pszObjId)
        pvArrayAllocate uRetVal.KeyBlob, lKeySize, FUNC_NAME & ".uRetVal.KeyBlob"
        Call CopyMemory(uRetVal.KeyBlob(0), ByVal lKeyPtr, lKeySize)
        Debug.Assert UBound(uRetVal.KeyBlob) >= 16
        Call CopyMemory(uRetVal.BitLen, uRetVal.KeyBlob(12), 4)
    ElseIf CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, 0) <> 0 Then
        Call CopyMemory(uEccKeyInfo, ByVal lKeyPtr, Len(uEccKeyInfo))
        uRetVal.AlgoObjId = pvToString(uEccKeyInfo.szCurveOid)
        pvArrayAllocate uRetVal.KeyBlob, uEccKeyInfo.PrivateKey.cbData, FUNC_NAME & ".uRetVal.KeyBlob"
        Call CopyMemory(uRetVal.KeyBlob(0), ByVal uEccKeyInfo.PrivateKey.pbData, uEccKeyInfo.PrivateKey.cbData)
    ElseIf Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        '--- no X509_ECC_PRIVATE_KEY struct type on NT4 -> decode in a wildly speculative way
        Call CopyMemory(lSize, baPrivKey(6), 1)
        If 7 + lSize <= UBound(baPrivKey) Then
            uRetVal.AlgoObjId = szOID_ECC_CURVE_P256
            pvArrayAllocate uRetVal.KeyBlob, lSize, FUNC_NAME & ".uRetVal.KeyBlob"
            Call CopyMemory(uRetVal.KeyBlob(0), baPrivKey(7), lSize)
        Else
            hResult = ERROR_FILE_NOT_FOUND
            sApiSource = "CryptDecodeObjectEx(X509_ECC_PRIVATE_KEY)"
            GoTo QH
        End If
    Else
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(X509_ECC_PRIVATE_KEY)"
        GoTo QH
    End If
    '--- success
    pvAsn1DecodePrivateKey = True
QH:
    If lKeyPtr <> 0 Then
        Call LocalFree(lKeyPtr)
    End If
    If lPkiPtr <> 0 Then
        Call LocalFree(lPkiPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    Debug.Assert RedimStats(MODULE_NAME & "." & sFuncName, lSize)
End Sub

Private Sub pvArrayReallocate(baArray() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim Preserve baArray(0 To lSize - 1) As Byte
    Else
        baArray = vbNullString
    End If
    Debug.Assert RedimStats(MODULE_NAME & "." & sFuncName, lSize)
End Sub

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long

    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

Private Function pvWriteBuffer(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvWriteBuffer"
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baBuffer)
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), 4)
    If lBufPtr = 0 Then
        pvArrayAllocate baBuffer, lPos + lSize, FUNC_NAME & ".baBuffer"
    ElseIf UBound(baBuffer) < lPos + lSize - 1 Then
        pvArrayReallocate baBuffer, lPos + lSize, FUNC_NAME & ".baRetVal"
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baBuffer(lPos), ByVal lPtr, lSize)
    End If
    pvWriteBuffer = lPos + lSize
End Function

Private Function pvWriteReserved(baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    pvWriteReserved = pvWriteBuffer(baBuffer, lPos, 0, lSize)
End Function

'= Schannel buffers helpers ==============================================

Private Sub pvInitSecDesc(uDesc As ApiSecBufferDesc, ByVal lCount As Long, uBuffers() As ApiSecBuffer)
    ReDim uBuffers(0 To lCount - 1)
    With uDesc
        .ulVersion = SECBUFFER_VERSION
        .cBuffers = lCount
        .pBuffers = VarPtr(uBuffers(0))
    End With
End Sub

Private Sub pvInitSecBuffer(uBuffer As ApiSecBuffer, ByVal lType As Long, Optional ByVal lPtr As Long, Optional ByVal lSize As Long)
    With uBuffer
        .BufferType = lType
        .pvBuffer = lPtr
        .cbBuffer = lSize
    End With
End Sub

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

Private Function pvCollectionCount(oCol As Collection) As Long
    If Not oCol Is Nothing Then
        pvCollectionCount = oCol.Count
    End If
End Function

#If Not ImplUseShared Then
Public Function RedimStats(sFuncName As String, ByVal lSize As Long) As Boolean
    #If sFuncName And lSize Then
    #End If
    RedimStats = True
End Function

Public Sub RemoveCollection(ByVal oCol As Collection, Index As Variant)
    If Not oCol Is Nothing Then
        pvCallCollectionRemove oCol, Index
    End If
End Sub

Public Function SearchCollection(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Boolean
    Dim vItem           As Variant
    
    If oCol Is Nothing Then
        GoTo QH
    ElseIf pvCallCollectionItem(oCol, Index, vItem) < 0 Then
        GoTo QH
    End If
    If IsObject(vItem) Then
        Set RetVal = vItem
    Else
        RetVal = vItem
    End If
    '--- success
    SearchCollection = True
QH:
End Function

Private Function pvCallCollectionItem(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM As Long = 7
    
    pvPatchMethodTrampoline AddressOf mdTlsNative.pvCallCollectionItem, IDX_COLLECTION_ITEM
    pvCallCollectionItem = pvCallCollectionItem(oCol, Index, RetVal)
End Function

Private Function pvCallCollectionRemove(ByVal oCol As Collection, Index As Variant) As Long
    Const IDX_COLLECTION_REMOVE As Long = 10
    
    pvPatchMethodTrampoline AddressOf mdTlsNative.pvCallCollectionRemove, IDX_COLLECTION_REMOVE
    pvCallCollectionRemove = pvCallCollectionRemove(oCol, Index)
End Function

Private Function pvPatchMethodTrampoline(ByVal Pfn As Long, ByVal lMethodIdx As Long) As Boolean
    Dim bInIDE          As Boolean

    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(Pfn, ByVal Pfn + &H16, 4)
    Else
        Call VirtualProtect(Pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal Pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (Pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * 4, 4)
    '--- success
    pvPatchMethodTrampoline = True
End Function

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function FromBase64Array(sText As String) As Byte()
    With VBA.CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .Text = sText
        FromBase64Array = .NodeTypedValue
    End With
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function SplitOrReindex(Expression As String, Delimiter As String) As Variant
    Dim vResult         As Variant
    Dim vTemp           As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    
    vResult = Split(Expression, Delimiter)
    '--- check if reindex needed
    If IsNumeric(vResult(0)) Then
        vTemp = vResult
        For lIdx = 0 To UBound(vTemp) Step 2
            If lSize < vTemp(lIdx) Then
                lSize = vTemp(lIdx)
            End If
        Next
        ReDim vResult(0 To lSize) As Variant
        Debug.Assert RedimStats(MODULE_NAME & ".SplitOrReindex.vResult", 0)
        For lIdx = 0 To UBound(vTemp) Step 2
            vResult(vTemp(lIdx)) = vTemp(lIdx + 1)
        Next
        SplitOrReindex = vResult
    End If
End Function
#End If
