VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTlsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'
' Based on RFC 8446 at https://tools.ietf.org/html/rfc8446
'   and illustrated traffic-dump at https://tls13.ulfheim.net/
'
' More TLS 1.3 implementations at https://github.com/h2o/picotls
'   and https://github.com/openssl/openssl
'
' Additional links with TLS 1.3 resources
'   https://github.com/tlswg/tls13-spec/wiki/Implementations
'   https://sans-io.readthedocs.io/how-to-sans-io.html
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cTlsSocket"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplSync = Not (ASYNCSOCKET_NO_SYNC <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)
#Const ImplUseLibSodium = (ASYNCSOCKET_USE_LIBSODIUM <> 0)

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

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 2
    ucsTlsSupportAll = ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_NEWKEYSET                           As Long = &H8
Private Const CRYPT_DELETEKEYSET                        As Long = &H10
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CNG_RSA_PRIVATE_KEY_BLOB                  As Long = 83
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
'--- for CryptSignHash
Private Const AT_SIGNATURE                              As Long = 2
Private Const RSA1024BIT_KEY                            As Long = &H4000000
Private Const NTE_BAD_ALGID                             As Long = &H80090008
Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
'--- for CertGetCertificateContextProperty
Private Const CERT_KEY_PROV_INFO_PROP_ID                As Long = 2
'--- for PFXImportCertStore
Private Const CRYPT_EXPORTABLE                          As Long = &H1
'--- for CryptExportKey
Private Const PUBLICKEYBLOB                             As Long = 6
Private Const PRIVATEKEYBLOB                            As Long = 7
'--- for CryptAcquireCertificatePrivateKey
Private Const CRYPT_ACQUIRE_CACHE_FLAG                  As Long = &H1
Private Const CRYPT_ACQUIRE_SILENT_FLAG                 As Long = &H40
Private Const CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG       As Long = &H10000
'Private Const CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG      As Long = &H20000
'--- for NCryptImportKey
Private Const NCRYPT_OVERWRITE_KEY_FLAG                 As Long = &H80
Private Const NCRYPT_DO_NOT_FINALIZE_FLAG               As Long = &H400
'--- for NCryptSetProperty
Private Const NCRYPT_PERSIST_FLAG                       As Long = &H80000000
'--- for CertStrToName
Private Const CERT_OID_NAME_STR                         As Long = 2
'--- for CertOpenStore
Private Const CERT_STORE_PROV_MEMORY                    As Long = 2
Private Const CERT_STORE_CREATE_NEW_FLAG                As Long = &H2000
'--- for CertAddEncodedCertificateToStore
Private Const CERT_STORE_ADD_USE_EXISTING               As Long = 2
'--- for CertGetCertificateChain
Private Const CERT_TRUST_IS_NOT_TIME_VALID              As Long = &H1
Private Const CERT_TRUST_IS_NOT_TIME_NESTED             As Long = &H2
Private Const CERT_TRUST_IS_REVOKED                     As Long = &H4
Private Const CERT_TRUST_IS_NOT_SIGNATURE_VALID         As Long = &H8
Private Const CERT_TRUST_IS_UNTRUSTED_ROOT              As Long = &H20
Private Const CERT_TRUST_REVOCATION_STATUS_UNKNOWN      As Long = &H40
Private Const CERT_TRUST_IS_PARTIAL_CHAIN               As Long = &H10000
'--- for CertFindCertificateInStore
Private Const CERT_FIND_EXISTING                        As Long = &HD0000
'--- for CERT_ALT_NAME_ENTRY
Private Const CERT_ALT_NAME_DNS_NAME                    As Long = 3
'--- for WSAGetLastError
Private Const WSABASEERR                                As Long = 10000
Private Const WSAEWOULDBLOCK                            As Long = (WSABASEERR + 35)
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_RSA_SSA_PSS                         As String = "1.2.840.113549.1.1.10"
Private Const szOID_ECC_PUBLIC_KEY                      As String = "1.2.840.10045.2.1"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"
Private Const szOID_PKCS_12_pbeWithSHA1And3KeyTripleDES As String = "1.2.840.113549.1.12.1.3"
Private Const szOID_SUBJECT_ALT_NAME2                   As String = "2.5.29.17"
'--- BLOBs magic
Private Const BCRYPT_RSAPRIVATE_MAGIC                   As Long = &H32415352
Private Const BCRYPT_ECDH_PRIVATE_P256_MAGIC            As Long = &H324B4345
Private Const BCRYPT_ECDH_PRIVATE_P384_MAGIC            As Long = &H344B4345
Private Const BCRYPT_ECDH_PRIVATE_P521_MAGIC            As Long = &H364B4345
'--- buffer types
Private Const NCRYPTBUFFER_PKCS_ALG_OID                 As Long = 41
Private Const NCRYPTBUFFER_PKCS_ALG_PARAM               As Long = 42
Private Const NCRYPTBUFFER_PKCS_KEY_NAME                As Long = 45
Private Const NCRYPTBUFFER_PKCS_SECRET                  As Long = 46
'--- export policy flags
Private Const NCRYPT_ALLOW_EXPORT_FLAG                  As Long = &H1
Private Const NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG        As Long = &H2
Private Const ERR_TIMEOUT                               As Long = &H800705B4
Private Const INVALID_SOCKET                            As Long = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
'--- Crypt32
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvParameters As Long, phCryptProvOrNCryptKey As Long, pdwKeySpec As Long, pfCallerFreeProvOrNCryptKey As Long) As Long
Private Declare Function PFXImportCertStore Lib "crypt32" (pPFX As Any, ByVal szPassword As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertStrToName Lib "crypt32" Alias "CertStrToNameW" (ByVal dwCertEncodingType As Long, ByVal pszX500 As Long, ByVal dwStrType As Long, ByVal pvReserved As Long, pbEncoded As Any, pcbEncoded As Long, ByVal ppszError As Long) As Long
Private Declare Function CertCreateSelfSignCertificate Lib "crypt32" (ByVal hCryptProvOrNCryptKey As Long, pSubjectIssuerBlob As Any, ByVal dwFlags As Long, pKeyProvInfo As Any, ByVal pSignatureAlgorithm As Long, pStartTime As Any, pEndTime As Any, ByVal pExtensions As Long) As Long
Private Declare Function CertOpenStore Lib "crypt32" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
Private Declare Function CertCloseStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertAddEncodedCertificateToStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long, ByVal dwAddDisposition As Long, ByVal ppCertContext As Long) As Long
Private Declare Function CertCreateCertificateChainEngine Lib "crypt32" (pConfig As Any, phChainEngine As Long) As Long
Private Declare Function CertFreeCertificateChainEngine Lib "crypt32" (ByVal hChainEngine As Long) As Long
Private Declare Function CertGetCertificateChain Lib "crypt32" (ByVal hChainEngine As Long, ByVal pCertContext As Long, ByVal pTime As Long, ByVal hAdditionalStore As Long, pChainPara As Any, ByVal dwFlags As Long, ByVal pvReserved As Long, ppChainContext As Long) As Long
Private Declare Function CertFreeCertificateChain Lib "crypt32" (ByVal pChainContext As Long) As Long
Private Declare Function CertFindExtension Lib "crypt32" (ByVal pszObjId As String, ByVal cExtensions As Long, ByVal rgExtensions As Long) As Long
Private Declare Function CertFindCertificateInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, pvFindPara As Any, ByVal pPrevCertContext As Long) As Long
'--- NCrypt
Private Declare Function NCryptImportKey Lib "ncrypt" (ByVal hProvider As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, phKey As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptExportKey Lib "ncrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Any, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFreeObject Lib "ncrypt" (ByVal hKey As Long) As Long
Private Declare Function NCryptGetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptSetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFinalizeKey Lib "ncrypt" (ByVal hKey As Long, ByVal dwFlags As Long) As Long
#If ImplUseLibSodium Then
    Private Declare Function crypto_aead_aes256gcm_is_available Lib "libsodium" () As Long
#End If
#If Not ImplUseShared Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

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

Private Type CRYPT_KEY_PROV_INFO
    pwszContainerName   As Long
    pwszProvName        As Long
    dwProvType          As Long
    dwFlags             As Long
    cProvParam          As Long
    rgProvParam         As Long
    dwKeySpec           As Long
End Type

Private Type CERT_CONTEXT
    dwCertEncodingType  As Long
    pbCertEncoded       As Long
    cbCertEncoded       As Long
    pCertInfo           As Long
    hCertStore          As Long
End Type

Private Type CRYPT_PRIVATE_KEY_INFO
    Version             As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_BLOB_DATA
    pAttributes         As Long
End Type

Private Type CRYPT_PKCS12_PBE_PARAMS
    iIterations         As Long
    cbSalt              As Long
    SaltBuffer(0 To 31) As Byte
End Type

Private Type CERT_ALT_NAME_ENTRY
    dwAltNameChoice     As Long
    pwszDNSName         As Long
    Padding             As Long
End Type

Private Type CERT_ALT_NAME_INFO
    cAltEntry           As Long
    rgAltEntry          As Long
End Type

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type CERT_INFO
    dwVersion           As Long
    SerialNumber        As CRYPT_BLOB_DATA
    SignatureAlgorithm  As CRYPT_ALGORITHM_IDENTIFIER
    Issuer              As CRYPT_BLOB_DATA
    NotBefore           As FILETIME
    NotAfter            As FILETIME
    Subject             As CRYPT_BLOB_DATA
    SubjectPublicKeyInfo As CERT_PUBLIC_KEY_INFO
    IssuerUniqueId      As CRYPT_BIT_BLOB
    SubjectUniqueId     As CRYPT_BIT_BLOB
    cExtension          As Long
    rgExtension         As Long
End Type

Private Type CERT_TRUST_STATUS
    dwErrorStatus       As Long
    dwInfoStatus        As Long
End Type

Private Type CERT_CHAIN_CONTEXT
    cbSize              As Long
    TrustStatus         As CERT_TRUST_STATUS
    cElems              As Long
    rgElem              As Long
    '--- more here
End Type

Private Type CTL_USAGE
    cUsageIdentifier    As Long
    rgpszUsageIdentifier As Long
End Type

Private Type CERT_USAGE_MATCH
    dwType              As Long
    Usage               As CTL_USAGE
End Type

Private Type CERT_CHAIN_PARA
    cbSize              As Long
    RequestedUsage      As CERT_USAGE_MATCH
End Type

Private Type CERT_EXTENSION
    pszObjId            As Long
    fCritical           As Long
    Value               As CRYPT_BLOB_DATA
End Type

Private Type NCryptBuffer
    cbBuffer            As Long
    BufferType          As Long
    pvBuffer            As Long
End Type

Private Type NCryptBufferDesc
    ulVersion           As Long
    cBuffers            As Long
    pBuffers            As Long
    Buffers()           As NCryptBuffer
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type CERT_CHAIN_ENGINE_CONFIG
    cbSize              As Long
    hRestrictedRoot     As Long
    hRestrictedTrust    As Long
    hRestrictedOther    As Long
    cAdditionalStore    As Long
    rghAdditionalStore  As Long
    dwFlags             As Long
    dwUrlRetrievalTimeout As Long
    MaximumCachedCertificates As Long
    CycleDetectionModulus As Long
    '--- following are Win7+ only
    hExclusiveRoot      As Long
    hExclusiveTrustedPeople As Long
    dwExclusiveFlags    As Long
End Type

Private Type CERT_CHAIN_ELEMENT
    cbSize              As Long
    pCertContext        As Long
    TrustStatus         As CERT_TRUST_STATUS
    pRevocationInfo     As Long
    pIssuanceUsage      As Long
    pApplicationUsage   As Long
    pwszExtendedErrorInfo As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS                             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown certificate authority|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_UNKNOWN                               As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT                          As String = "%1."
'--- numeric
Private Const LNG_AAD_SIZE                              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const LNG_LEGACY_AAD_SIZE                       As Long = 13    '--- for TLS 1.2
Private Const LNG_ANS1_TYPE_SEQUENCE                    As Long = &H30
Private Const LNG_ANS1_TYPE_INTEGER                     As Long = &H2
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000
'--- TLS
Private Const TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC       As Long = 20
Private Const TLS_CONTENT_TYPE_ALERT                    As Long = 21
Private Const TLS_CONTENT_TYPE_HANDSHAKE                As Long = 22
Private Const TLS_CONTENT_TYPE_APPDATA                  As Long = 23
Private Const TLS_HANDSHAKE_TYPE_CLIENT_HELLO           As Long = 1
Private Const TLS_HANDSHAKE_TYPE_SERVER_HELLO           As Long = 2
Private Const TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET     As Long = 4
'Private Const TLS_HANDSHAKE_TYPE_END_OF_EARLY_DATA      As Long = 5
Private Const TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS   As Long = 8
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE            As Long = 11
Private Const TLS_HANDSHAKE_TYPE_SERVER_KEY_EXCHANGE    As Long = 12
'Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST    As Long = 13
Private Const TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE      As Long = 14
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY     As Long = 15
Private Const TLS_HANDSHAKE_TYPE_CLIENT_KEY_EXCHANGE    As Long = 16
Private Const TLS_HANDSHAKE_TYPE_FINISHED               As Long = 20
Private Const TLS_HANDSHAKE_TYPE_KEY_UPDATE             As Long = 24
'Private Const TLS_HANDSHAKE_TYPE_COMPRESSED_CERTIFICATE As Long = 25
Private Const TLS_HANDSHAKE_TYPE_MESSAGE_HASH           As Long = 254
Private Const TLS_EXTENSION_TYPE_SERVER_NAME            As Long = 0
'Private Const TLS_EXTENSION_TYPE_STATUS_REQUEST         As Long = 5
Private Const TLS_EXTENSION_TYPE_SUPPORTED_GROUPS       As Long = 10
Private Const TLS_EXTENSION_TYPE_EC_POINT_FORMAT        As Long = 11
Private Const TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS   As Long = 13
'Private Const TLS_EXTENSION_TYPE_ALPN                   As Long = 16
'Private Const TLS_EXTENSION_TYPE_COMPRESS_CERTIFICATE   As Long = 27
'Private Const TLS_EXTENSION_TYPE_PRE_SHARED_KEY         As Long = 41
'Private Const TLS_EXTENSION_TYPE_EARLY_DATA             As Long = 42
Private Const TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS     As Long = 43
Private Const TLS_EXTENSION_TYPE_COOKIE                 As Long = 44
'Private Const TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES As Long = 45
Private Const TLS_EXTENSION_TYPE_KEY_SHARE              As Long = 51
Private Const TLS_CS_AES_128_GCM_SHA256                 As Long = &H1301
Private Const TLS_CS_AES_256_GCM_SHA384                 As Long = &H1302
Private Const TLS_CS_CHACHA20_POLY1305_SHA256           As Long = &H1303
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256 As Long = &HC02B&
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384 As Long = &HC02C&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256  As Long = &HC02F&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384  As Long = &HC030&
Private Const TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA8&
Private Const TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA9&
Private Const TLS_CS_RSA_WITH_AES_128_GCM_SHA256        As Long = &H9C
Private Const TLS_CS_RSA_WITH_AES_256_GCM_SHA384        As Long = &H9D
Private Const TLS_GROUP_SECP256R1                       As Long = 23
Private Const TLS_GROUP_SECP384R1                       As Long = 24
'Private Const TLS_GROUP_SECP521R1                       As Long = 25
Private Const TLS_GROUP_X25519                          As Long = 29
'Private Const TLS_GROUP_X448                            As Long = 30
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA1              As Long = &H201
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA256            As Long = &H401
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA384            As Long = &H501
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA512            As Long = &H601
Private Const TLS_SIGNATURE_ECDSA_SECP256R1_SHA256      As Long = &H403
Private Const TLS_SIGNATURE_ECDSA_SECP384R1_SHA384      As Long = &H503
Private Const TLS_SIGNATURE_ECDSA_SECP521R1_SHA512      As Long = &H603
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA512         As Long = &H806
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA256          As Long = &H809
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA384          As Long = &H80A
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA512          As Long = &H80B
'Private Const TLS_PSK_KE_MODE_PSK_DHE                   As Long = 1
Private Const TLS_PROTOCOL_VERSION_TLS12                As Long = &H303
Private Const TLS_PROTOCOL_VERSION_TLS13                As Long = &H304
Private Const TLS_CHACHA20_KEY_SIZE                     As Long = 32
Private Const TLS_CHACHA20POLY1305_IV_SIZE              As Long = 12
Private Const TLS_CHACHA20POLY1305_TAG_SIZE             As Long = 16
Private Const TLS_AES128_KEY_SIZE                       As Long = 16
Private Const TLS_AES256_KEY_SIZE                       As Long = 32
Private Const TLS_AESGCM_IV_SIZE                        As Long = 12
Private Const TLS_AESGCM_TAG_SIZE                       As Long = 16
Private Const TLS_COMPRESS_NULL                         As Long = 0
Private Const TLS_SERVER_NAME_TYPE_HOSTNAME             As Long = 0
Private Const TLS_ALERT_LEVEL_WARNING                   As Long = 1
Private Const TLS_ALERT_LEVEL_FATAL                     As Long = 2
Private Const TLS_SHA256_DIGEST_SIZE                    As Long = 32
Private Const TLS_SHA384_DIGEST_SIZE                    As Long = 48
Private Const TLS_X25519_KEY_SIZE                       As Long = 32
Private Const TLS_SECP256R1_KEY_SIZE                    As Long = 32
Private Const TLS_SECP384R1_KEY_SIZE                    As Long = 48
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = (16384 + 256)
Private Const TLS_RECORD_VERSION                        As Long = TLS_PROTOCOL_VERSION_TLS12 '--- always legacy version
Private Const TLS_LOCAL_LEGACY_VERSION                  As Long = &H303
Private Const TLS_HELLO_RANDOM_SIZE                     As Long = 32
'--- errors
Private Const ERR_CONNECTION_CLOSED                     As String = "Connection closed"
Private Const ERR_GENER_KEYPAIR_FAILED                  As String = "Failed generating key pair (%1)"
Private Const ERR_UNSUPPORTED_EXCH_GROUP                As String = "Unsupported exchange group (%1)"
Private Const ERR_UNSUPPORTED_CIPHER_SUITE              As String = "Unsupported cipher suite (%1)"
Private Const ERR_UNSUPPORTED_SIGNATURE_TYPE            As String = "Unsupported signature type (%1)"
Private Const ERR_UNSUPPORTED_CERTIFICATE               As String = "Unsupported certificate"
Private Const ERR_UNSUPPORTED_PUBLIC_KEY                As String = "Unsupported public key OID (%1)"
Private Const ERR_UNSUPPORTED_PRIVATE_KEY               As String = "Unsupported private key"
Private Const ERR_UNSUPPORTED_CURVE_SIZE                As String = "Unsupported curve size (%1)"
Private Const ERR_UNSUPPORTED_PROTOCOL                  As String = "Invalid protocol version"
Private Const ERR_ENCRYPTION_FAILED                     As String = "Encryption failed"
Private Const ERR_SIGNATURE_FAILED                      As String = "Certificate signature failed (%1)"
Private Const ERR_DECRYPTION_FAILED                     As String = "Decryption failed"
Private Const ERR_SERVER_HANDSHAKE_FAILED               As String = "Handshake verification failed"
Private Const ERR_NEGOTIATE_SIGNATURE_FAILED            As String = "Negotiate signature type failed"
Private Const ERR_CALL_FAILED                           As String = "Call failed (%1)"
Private Const ERR_RECORD_TOO_BIG                        As String = "Record size too big"
Private Const ERR_FATAL_ALERT                           As String = "Received fatal alert"
Private Const ERR_UNEXPECTED_RECORD_TYPE                As String = "Unexpected record type (%1)"
Private Const ERR_UNEXPECTED_MSG_TYPE                   As String = "Unexpected message type for %1 (%2)"
Private Const ERR_UNEXPECTED_PROTOCOL                   As String = "Unexpected protocol for %1 (%2)"
Private Const ERR_INVALID_STATE_HANDSHAKE               As String = "Invalid state for handshake content (%1)"
Private Const ERR_INVALID_SIZE_KEY_SHARE                As String = "Invalid data size for key share"
Private Const ERR_INVALID_REMOTE_KEY                    As String = "Invalid remote key size"
Private Const ERR_INVALID_SIZE_REMOTE_KEY               As String = "Invalid data size for remote key"
Private Const ERR_INVALID_SIZE_VERSIONS                 As String = "Invalid data size for supported versions"
Private Const ERR_INVALID_SIGNATURE                     As String = "Invalid certificate signature"
Private Const ERR_INVALID_HASH_SIZE                     As String = "Invalid hash size (%1)"
Private Const ERR_COOKIE_NOT_ALLOWED                    As String = "Cookie not allowed outside HelloRetryRequest"
Private Const ERR_NO_HANDSHAKE_MESSAGES                 As String = "Missing handshake messages"
Private Const ERR_NO_PREVIOUS_SECRET                    As String = "Missing previous secret (%1)"
Private Const ERR_NO_REMOTE_RANDOM                      As String = "Missing remote random"
Private Const ERR_NO_SERVER_CERTIFICATE                 As String = "Missing server certificate"
Private Const ERR_NO_SUPPORTED_CIPHER_SUITE             As String = "Missing supported ciphersuite"
Private Const ERR_TRUST_IS_REVOKED                      As String = "Trust for this certificate or one of the certificates in the certificate chain has been revoked"
Private Const ERR_TRUST_IS_PARTIAL_CHAIN                As String = "The certificate chain is not complete"
Private Const ERR_TRUST_IS_UNTRUSTED_ROOT               As String = "The certificate or certificate chain is based on an untrusted root"
Private Const ERR_TRUST_IS_NOT_TIME_VALID               As String = "The certificate has expired"
Private Const ERR_TRUST_REVOCATION_STATUS_UNKNOWN       As String = "The revocation status of the certificate or one of the certificates in the certificate chain is unknown"
Private Const ERR_NO_MATCHING_ALT_NAME                  As String = "No certificate subject name matches target host name"
Private Const STR_CHR1                                  As String = "" '--- CHAR(1)
Private Const DEF_TIMEOUT                               As Long = 5000

Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_bUseTls               As Boolean
Private m_sRemoteHostName       As String
Private m_eLocalFeatures        As UcsTlsLocalFeaturesEnum
Private m_uCtx                  As UcsTlsContext
Private m_hRootStore            As Long
Private m_oRootStoreFwd         As cTlsSocket
'--- socket buffers
Private m_baRecv()              As Byte
Private m_lRecvSize             As Long
Private m_baSend()              As Byte
Private m_lSendPos              As Long
Private m_lSendSize             As Long
Private m_lSendBytes            As Long
Private m_lLastSendBytes        As Long

Private Enum UcsTlsStatesEnum
    ucsTlsStateNew
    ucsTlsStateClosed
    ucsTlsStateHandshakeStart
    ucsTlsStateExpectServerHello
    ucsTlsStateExpectExtensions
    ucsTlsStateExpectServerFinished     '--- not used in TLS 1.3
    '--- server states
    ucsTlsStateExpectClientHello
    ucsTlsStateExpectClientFinished
    ucsTlsStatePostHandshake
    ucsTlsStateShutdown
End Enum

Private Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoExchX25519 = 1
    ucsTlsAlgoExchSecp256r1 = 2
    ucsTlsAlgoExchSecp384r1 = 3
    ucsTlsAlgoExchCertificate = 4
    '--- authenticated encryption w/ additional data
    ucsTlsAlgoAeadChacha20Poly1305 = 11
    ucsTlsAlgoAeadAes128 = 12
    ucsTlsAlgoAeadAes256 = 13
    '--- hash
    ucsTlsAlgoDigestSha256 = 21
    ucsTlsAlgoDigestSha384 = 22
    ucsTlsAlgoDigestSha512 = 23
    '--- verify signature
    ucsTlsAlgoSignaturePkcsSha1 = 31
    ucsTlsAlgoSignaturePkcsSha2 = 32
    ucsTlsAlgoSignaturePss = 33
End Enum

Private Enum UcsTlsAlertDescriptionsEnum
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

#If Not ImplUseShared Then
Private Enum UcsOsVersionEnum
    ucsOsvNt4 = 400
    ucsOsvWin98 = 410
    ucsOsvWin2000 = 500
    ucsOsvXp = 501
    ucsOsvVista = 600
    ucsOsvWin7 = 601
    ucsOsvWin8 = 602
    [ucsOsvWin8.1] = 603
    ucsOsvWin10 = 1000
End Enum
#End If

Private Type UcsTlsContext
    '--- config
    IsServer            As Boolean
    RemoteHostName      As String
    LocalFeatures       As UcsTlsLocalFeaturesEnum
    '--- state
    State               As UcsTlsStatesEnum
    LastError           As String
    LastAlertCode       As UcsTlsAlertDescriptionsEnum
    BlocksStack         As Collection
    '--- handshake
    LocalSessionID()    As Byte
    LocalExchRandom()   As Byte
    LocalExchPrivate()  As Byte
    LocalExchPublic()   As Byte
    LocalExchRsaEncrPriv() As Byte
    LocalCertificates   As Collection
    LocalPrivateKey()   As Byte
    LocalSignatureType  As Long
    RemoteSessionID()   As Byte
    RemoteExchRandom()  As Byte
    RemoteExchPublic()  As Byte
    RemoteCertificates  As Collection
    RemoteExtensions    As Collection
    RemoteTickets       As Collection
    '--- crypto settings
    ProtocolVersion     As Long
    ExchGroup           As Long
    ExchAlgo            As UcsTlsCryptoAlgorithmsEnum
    CipherSuite         As Long
    AeadAlgo            As UcsTlsCryptoAlgorithmsEnum
    MacSize             As Long '--- always 0 (not used w/ AEAD ciphers)
    KeySize             As Long
    IvSize              As Long
    IvDynamicSize       As Long '--- only for AES in TLS 1.2
    TagSize             As Long
    DigestAlgo          As UcsTlsCryptoAlgorithmsEnum
    DigestSize          As Long
    '--- bulk secrets
    HandshakeMessages() As Byte '--- ToDo: reduce to HandshakeHash only
    HandshakeSecret()   As Byte
    MasterSecret()      As Byte
    RemoteTrafficSecret() As Byte
    RemoteTrafficKey()  As Byte
    RemoteTrafficIV()   As Byte
    RemoteTrafficSeqNo  As Long
    LocalTrafficSecret() As Byte
    LocalTrafficKey()   As Byte
    LocalTrafficIV()    As Byte
    LocalTrafficSeqNo   As Long
    '--- hello retry request
    HelloRetryRequest   As Boolean
    HelloRetryCipherSuite As Long
    HelloRetryExchGroup As Long
    HelloRetryCookie()  As Byte
    '--- I/O buffers
    RecvBuffer()        As Byte
    RecvPos             As Long
    DecrBuffer()        As Byte
    DecrPos             As Long
    SendBuffer()        As Byte
    SendPos             As Long
    MessBuffer()        As Byte
    MessPos             As Long
    MessSize            As Long
End Type

Private Type UcsKeyInfo
    AlgoObjId           As String
    KeyBlob()           As Byte
    BitLen              As Long
    Modulus()           As Byte
    PubExp()            As Byte
    PrivExp()           As Byte
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

Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

Property Get SocketHandle() As Long
    SocketHandle = m_oSocket.SocketHandle
End Property

Property Get LastErrorText() As String
    If m_oSocket.LastError <> 0 Then
        LastErrorText = m_oSocket.GetErrorDescription(m_oSocket.LastError)
    Else
        LastErrorText = pvTlsGetLastError(m_uCtx)
    End If
End Property

Property Get LastError() As Long
    LastError = m_oSocket.LastError
    If LastError = 0 Then
        If LenB(LastErrorText) <> 0 Then
            LastError = -1
        End If
    End If
End Property

Property Get IsClosed() As Boolean
    If m_bUseTls Then
        IsClosed = pvTlsIsClosed(m_uCtx)
    Else
        IsClosed = (m_oSocket.SocketHandle = INVALID_SOCKET)
    End If
End Property

Property Get AvailableBytes() As Long
    If m_bUseTls Then
        AvailableBytes = m_lRecvSize
    Else
        AvailableBytes = m_oSocket.AvailableBytes
    End If
End Property

Property Get LastSendBytes() As Long
    If m_bUseTls Then
        LastSendBytes = m_lLastSendBytes
    Else
        LastSendBytes = m_oSocket.LastSendBytes
    End If
End Property

Property Get LocalCertificates() As Collection
    Set LocalCertificates = m_uCtx.LocalCertificates
End Property

Property Get HasPendingEvent() As Boolean
    HasPendingEvent = m_oSocket.HasPendingEvent
End Property

Friend Property Get frRootStore() As Long
    If Not m_oRootStoreFwd Is Nothing Then
        frRootStore = m_oRootStoreFwd.frRootStore
    Else
        frRootStore = m_hRootStore
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Listen(Optional ByVal ConnectionBacklog As Long = 5) As Boolean
    Const FUNC_NAME     As String = "Listen"
    
    On Error GoTo EH
    If m_bUseTls Then
        If m_uCtx.LocalCertificates Is Nothing Then
            If Not pvPkiGenerSelfSignedCertificate(m_uCtx.LocalCertificates, m_uCtx.LocalPrivateKey) Then
                GoTo QH
            End If
        End If
    End If
    If Not m_oSocket.Listen(ConnectionBacklog) Then
        GoTo QH
    End If
    '--- success
    Listen = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Accept( _
            ConnectedSocket As cTlsSocket, _
            Optional SocketAddress As String, _
            Optional SocketPort As Long, _
            Optional ByVal UseTls As Boolean = True) As Boolean
    Const FUNC_NAME     As String = "Accept"
    
    On Error GoTo EH
    If ConnectedSocket Is Nothing Then
        Set ConnectedSocket = New cTlsSocket
    End If
    If Not m_oSocket.Accept(ConnectedSocket.Socket, SocketAddress, SocketPort) Then
        GoTo QH
    End If
    If UseTls Then
        If Not ConnectedSocket.StartServerTls(m_uCtx.LocalCertificates, m_uCtx.LocalPrivateKey) Then
            GoTo QH
        End If
    End If
    '--- success
    Accept = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function StartServerTls( _
            Optional Certificates As Collection, _
            Optional PrivateKey As Variant) As Boolean
    m_bUseTls = True
    m_oSocket.GetPeerName m_sRemoteHostName, 0
    m_eLocalFeatures = ucsTlsSupportTls13
    If Not pvTlsInitServer(m_uCtx, m_sRemoteHostName, Certificates, PrivateKey) Then
        GoTo QH
    End If
    '--- success
    StartServerTls = True
QH:
End Function

Public Function Connect( _
            HostAddress As String, _
            ByVal HostPort As Long, _
            Optional ByVal UseTls As Boolean = True, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum) As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    m_bUseTls = UseTls
    m_sRemoteHostName = HostAddress
    m_eLocalFeatures = LocalFeatures
    If Not m_oSocket.Connect(HostAddress, HostPort) Then
        GoTo QH
    End If
    '-- success
    Connect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function StartTls( _
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum) As Boolean
    Const FUNC_NAME     As String = "StartTls"
    Dim baOutput()      As Byte
    Dim lSize           As Long
    Dim bResult         As Boolean
    Dim baEmpty()       As Byte
    
    On Error GoTo EH
    If m_uCtx.State > ucsTlsStateClosed Then
        GoTo QH
    End If
    If LenB(RemoteHostName) <> 0 Then
        m_sRemoteHostName = RemoteHostName
    End If
    If LocalFeatures = 0 Then
        m_eLocalFeatures = ucsTlsSupportAll Or IIf(OsVersion < ucsOsvXp, ucsTlsIgnoreServerCertificateErrors, 0)
    Else
        m_eLocalFeatures = LocalFeatures
    End If
    If Not pvTlsInitClient(m_uCtx, m_sRemoteHostName, m_eLocalFeatures) Then
        GoTo QH
    End If
    bResult = pvTlsHandshake(m_uCtx, baEmpty, 0, baOutput, lSize)
    If Not m_oSocket.SendArray(baOutput) Then
        GoTo QH
    End If
    If Not bResult Then
        GoTo QH
    End If
    '--- success
    StartTls = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function ReceiveText( _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "ReceiveText"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_bUseTls Then
        If ReceiveArray(baBuffer) Then
            ReceiveText = m_oSocket.FromTextArray(baBuffer, CodePage)
        End If
    Else
        ReceiveText = m_oSocket.ReceiveText(HostAddress, HostPort, CodePage)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ReceiveArray( _
            Buffer() As Byte, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long) As Boolean
    Const FUNC_NAME     As String = "ReceiveArray"
    
    On Error GoTo EH
    If m_bUseTls Then
        Buffer = vbNullString
        If m_lRecvSize = 0 Then
            If pvTlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Not pvOnReceive() Then
                GoTo QH
            End If
        End If
        If m_lRecvSize > 0 Then
            pvArrayAllocate Buffer, m_lRecvSize, FUNC_NAME & ".Buffer"
            Call CopyMemory(Buffer(0), m_baRecv(0), m_lRecvSize)
            m_lRecvSize = 0
        End If
        m_oSocket.Receive 0, 0 '--- enable FD_READ
        '--- success
        ReceiveArray = True
    Else
        ReceiveArray = m_oSocket.ReceiveArray(Buffer, HostAddress, HostPort)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function SendText( _
            Text As String, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Boolean
    Const FUNC_NAME     As String = "SendText"
    
    On Error GoTo EH
    If m_bUseTls Then
        SendText = SendArray(m_oSocket.ToTextArray(Text, CodePage), HostAddress, HostPort)
    Else
        SendText = m_oSocket.SendText(Text, HostAddress, HostPort, CodePage)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SendArray( _
            Buffer() As Byte, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long) As Boolean
    Const FUNC_NAME     As String = "SendArray"
    Dim lInputSize      As Long
    Dim baTemp()        As Byte
    Dim bResult         As Boolean
    Dim lPos            As Long
    Dim lSize           As Long
    
    On Error GoTo EH
    If m_bUseTls Then
        If pvTlsIsClosed(m_uCtx) Then
            GoTo QH
        End If
        lInputSize = pvArraySize(Buffer)
        m_lSendBytes = m_lSendBytes + lInputSize
        If lInputSize <= TLS_MAX_PLAINTEXT_RECORD_SIZE Then
            bResult = pvTlsSend(m_uCtx, Buffer, lInputSize, m_baSend, m_lSendSize)
        Else
            '--- chunk Buffer on TLS_MAX_PLAINTEXT_RECORD_SIZE application data records
            pvArrayAllocate baTemp, TLS_MAX_PLAINTEXT_RECORD_SIZE, FUNC_NAME & ".baTemp"
            Do
                lSize = IIf(lInputSize - lPos < TLS_MAX_PLAINTEXT_RECORD_SIZE, lInputSize - lPos, TLS_MAX_PLAINTEXT_RECORD_SIZE)
                Call CopyMemory(baTemp(0), Buffer(lPos), lSize)
                lPos = lPos + lSize
                bResult = pvTlsSend(m_uCtx, baTemp, lSize, m_baSend, m_lSendSize)
            Loop While lPos < lInputSize And bResult
        End If
        If Not bResult Then
            RaiseEvent OnError(-1, ucsSfdWrite)
            GoTo QH
        End If
        If Not pvOnSend() Then
            GoTo QH
        End If
        '--- success
        SendArray = True
    Else
        SendArray = m_oSocket.SendArray(Buffer, HostAddress, HostPort)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function Send( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long) As Long
    Const FUNC_NAME     As String = "Send"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If BufLen = 0 Then
        '--- do nothing
    ElseIf m_bUseTls Then
        pvWriteBuffer baBuffer, 0, BufPtr, BufLen
        If Not SendArray(baBuffer) Then
            Send = -1
        Else
            Send = BufLen
        End If
    Else
        Send = m_oSocket.Send(BufPtr, BufLen, HostAddress, HostPort)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function ShutDown() As Boolean
    Const FUNC_NAME     As String = "ShutDown"
    
    On Error GoTo EH
    If m_bUseTls Then
        If pvTlsIsReady(m_uCtx) And m_uCtx.State <> ucsTlsStateShutdown Then
            If Not pvOnShutdown() Then
                GoTo QH
            End If
        End If
        '--- success
        ShutDown = True
    Else
        ShutDown = m_oSocket.ShutDown
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
    If ErrorCode = -1 Then
        GetErrorDescription = pvTlsGetLastError(m_uCtx)
    Else
        GetErrorDescription = m_oSocket.GetErrorDescription(ErrorCode)
    End If
End Function

'= forwarded =============================================================

Public Function Create( _
            Optional ByVal SocketPort As Long, _
            Optional ByVal SocketType As UcsAsyncSocketTypeEnum = ucsSckStream, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll, _
            Optional SocketAddress As String) As Boolean
    Create = m_oSocket.Create(SocketPort, SocketType, EventMask, SocketAddress)
End Function

Public Function GetSockName(SocketAddress As String, SocketPort As Long) As Boolean
    GetSockName = m_oSocket.GetSockName(SocketAddress, SocketPort)
End Function

Public Function GetPeerName(PeerAddress As String, PeerPort As Long) As Boolean
    GetPeerName = m_oSocket.GetPeerName(PeerAddress, PeerPort)
End Function

Public Function GetLocalHost(HostName As String, HostAddress As String) As Boolean
    GetLocalHost = m_oSocket.GetLocalHost(HostName, HostAddress)
End Function

Public Function Attach( _
            ByVal SocketHandle As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Attach = m_oSocket.Attach(SocketHandle, EventMask)
End Function

Public Function Detach() As Long
    Detach = m_oSocket.Detach()
End Function

Public Sub Close_()
    If m_bUseTls Then
        If pvTlsIsReady(m_uCtx) And m_uCtx.State <> ucsTlsStateShutdown Then
            If Not pvOnShutdown() Then
                GoTo QH
            End If
        End If
    Else
        m_oSocket.Close_
    End If
QH:
End Sub

Public Sub PostEvent(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Optional ByVal Immediate As Boolean)
    If Immediate Then
        Select Case EventMask
        Case ucsSfdRead
            RaiseEvent OnReceive
        Case ucsSfdWrite
            RaiseEvent OnSend
        Case ucsSfdConnect
            RaiseEvent OnConnect
        Case ucsSfdAccept
            RaiseEvent OnAccept
        Case ucsSfdClose
            RaiseEvent OnClose
            m_uCtx.State = ucsTlsStateClosed
            m_bUseTls = False
        End Select
    Else
        m_oSocket.PostEvent EventMask, Immediate
    End If
End Sub

Public Function Bind(Optional SocketAddress As String, Optional ByVal SocketPort As Long) As Boolean
    Bind = m_oSocket.Bind(SocketAddress, SocketPort)
End Function

Public Function FromTextArray(baText() As Byte, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    FromTextArray = m_oSocket.FromTextArray(baText, CodePage)
End Function

Public Function ToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    ToTextArray = m_oSocket.ToTextArray(sText, CodePage)
End Function

'= sync ==================================================================

#If ImplSync Then
Public Function SyncConnect(HostAddress As String, ByVal HostPort As Long, Optional ByVal Timeout As Long, Optional ByVal UseTls As Boolean = True) As Boolean
    Const FUNC_NAME     As String = "SyncConnect"
    
    On Error GoTo EH
    If UseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        If Not Connect(HostAddress, HostPort, UseTls:=UseTls) Then
            GoTo QH
        End If
        If Not SyncWaitForEvent(Timeout, ucsSfdConnect) Then
            GoTo QH
        End If
        '--- success
        SyncConnect = True
    Else
        SyncConnect = m_oSocket.SyncConnect(HostAddress, HostPort, Timeout:=Timeout)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSendText( _
            Text As String, _
            Optional ByVal Timeout As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Boolean
    Const FUNC_NAME     As String = "SyncSendText"
    
    On Error GoTo EH
    SyncSendText = SyncSendArray(m_oSocket.ToTextArray(Text, CodePage), Timeout:=Timeout)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSendArray(Buffer() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "SyncSendArray"
    
    On Error GoTo EH
    If UBound(Buffer) < 0 Then
        SyncSendArray = True
    ElseIf m_bUseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        If Not SendArray(Buffer) Then
            GoTo QH
        End If
        If Not SyncWaitForEvent(Timeout, ucsSfdWrite) Then
            GoTo QH
        End If
        '--- success
        SyncSendArray = True
    Else
        SyncSendArray = m_oSocket.SyncSendArray(Buffer, Timeout:=Timeout)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceiveText( _
            Optional ByVal NeedLen As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "SyncReceiveText"
    Dim lElapsed        As Long
    Dim dblTimer        As Double
    Dim baRecv()        As Byte
    Dim lPos            As Long
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_bUseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        baBuffer = vbNullString
        dblTimer = TimerEx
        Do
            If NeedLen = 0 Then
                If Not ReceiveArray(baBuffer) Then
                    GoTo QH
                End If
            Else
                If Not ReceiveArray(baRecv) Then
                    GoTo QH
                End If
                lPos = pvWriteArray(baBuffer, lPos, baRecv)
            End If
            SyncReceiveText = m_oSocket.FromTextArray(baBuffer, CodePage)
            If Len(SyncReceiveText) >= NeedLen Then
                Exit Do
            ElseIf pvTlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.frLastError = ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not SyncWaitForEvent(Timeout - lElapsed, ucsSfdRead) Then
                GoTo QH
            End If
        Loop
    Else
        SyncReceiveText = m_oSocket.SyncReceiveText(NeedLen:=NeedLen, Timeout:=Timeout, CodePage:=CodePage)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceiveArray( _
            Buffer() As Byte, _
            Optional ByVal NeedLen As Long, _
            Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "SyncReceiveArray"
    Dim lElapsed        As Long
    Dim dblTimer        As Double
    Dim baRecv()        As Byte
    Dim lPos            As Long
    
    On Error GoTo EH
    If m_bUseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        Buffer = vbNullString
        dblTimer = TimerEx
        Do
            If NeedLen = 0 Then
                If Not ReceiveArray(Buffer) Then
                    GoTo QH
                End If
            Else
                If Not ReceiveArray(baRecv) Then
                    GoTo QH
                End If
                lPos = pvWriteArray(Buffer, lPos, baRecv)
            End If
            If UBound(Buffer) >= NeedLen Then
                Exit Do
            ElseIf pvTlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.frLastError = ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not SyncWaitForEvent(Timeout - lElapsed, ucsSfdRead) Then
                GoTo QH
            End If
        Loop
        '--- success
        SyncReceiveArray = True
    Else
        SyncReceiveArray = m_oSocket.SyncReceiveArray(Buffer, NeedLen:=NeedLen, Timeout:=Timeout)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncWaitForEvent( _
            ByVal Timeout As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Const FUNC_NAME     As String = "SyncWaitForEvent"
    Const LNG_TLS_MASK  As Long = ucsSfdConnect Or ucsSfdRead Or ucsSfdWrite
    Dim dblTimer        As Double
    Dim lElapsed        As Long
    
    On Error GoTo EH
    If m_bUseTls And (EventMask And LNG_TLS_MASK) <> 0 Then
        dblTimer = TimerEx
        m_oSocket.SyncProcessMsgQueue
        Do
            If (EventMask And ucsSfdConnect) <> 0 Then
                If pvTlsIsReady(m_uCtx) Then
                    Exit Do
                End If
            End If
            If (EventMask And ucsSfdRead) <> 0 Then
                If m_lRecvSize > 0 Then
                    Exit Do
                End If
            End If
            If (EventMask And ucsSfdWrite) <> 0 Then
                If m_lSendPos = 0 Then
                    Exit Do
                End If
            End If
            If pvTlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.frLastError = ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not m_oSocket.SyncWaitForEvent(Timeout - lElapsed, ucsSfdAll) Then
                GoTo QH
            End If
            m_oSocket.SyncProcessMsgQueue
            If m_oSocket.LastError <> 0 Then
                GoTo QH
            End If
        Loop
        '--- success
        SyncWaitForEvent = True
    Else
        SyncWaitForEvent = m_oSocket.SyncWaitForEvent(Timeout, EventMask)
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function
#End If

'= private ===============================================================

Private Function pvOnConnect() As Boolean
    Const FUNC_NAME     As String = "pvOnConnect"
    
    On Error GoTo EH
    If Not StartTls(m_sRemoteHostName, m_eLocalFeatures) Then
        GoTo QH
    End If
    '--- success
    pvOnConnect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOnReceive(Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvOnReceive"
    Dim baRecv()        As Byte
    Dim baOutput()      As Byte
    Dim lSize           As Long
    Dim bResult         As Boolean
    Dim lPrevSize       As Long
    
    On Error GoTo EH
    lPrevSize = m_lRecvSize
    Do
        If pvTlsIsClosed(m_uCtx) Then
            Exit Do
        End If
        If Not m_oSocket.ReceiveArray(baRecv) Then
            GoTo QH
        End If
        If pvArraySize(baRecv) <= 0 Then
            Exit Do
        End If
        If Not pvTlsIsReady(m_uCtx) Then
            bResult = pvTlsHandshake(m_uCtx, baRecv, -1, baOutput, lSize)
            baRecv = vbNullString
            If pvArraySize(baOutput) > 0 Then
                If Not m_oSocket.SendArray(baOutput) Then
                    GoTo QH
                End If
                If m_oSocket.LastError <> WSAEWOULDBLOCK Then
                    m_oSocket.PostEvent ucsSfdWrite
                    m_oSocket.frLastError = WSAEWOULDBLOCK
                End If
            End If
            If pvTlsIsClosed(m_uCtx) Then
                If pvOnBeforeNotify(ucsSfdClose) Then
                    RaiseEvent OnClose
                End If
                GoTo QH
            End If
            If Not bResult Then
                RaiseEvent OnError(-1, ucsSfdRead)
                GoTo QH
            End If
            If Not pvTlsIsReady(m_uCtx) Then
                GoTo QH
            End If
            If Not m_uCtx.IsServer And (m_uCtx.LocalFeatures And ucsTlsIgnoreServerCertificateErrors) = 0 Then
                If m_uCtx.RemoteCertificates Is Nothing Then
                    RaiseEvent OnError(-1, ucsSfdRead)
                    GoTo QH
                End If
                If Not pvPkiCertChainValidate(m_uCtx.RemoteHostName, m_uCtx.RemoteCertificates, frRootStore, m_uCtx.LastError) Then
                    RaiseEvent OnError(-1, ucsSfdRead)
                    GoTo QH
                End If
            End If
            If pvOnBeforeNotify(ucsSfdConnect) Then
                RaiseEvent OnConnect
            End If
        End If
        bResult = pvTlsReceive(m_uCtx, baRecv, -1, m_baRecv, m_lRecvSize)
    Loop While Flush
    If m_lRecvSize > lPrevSize Then
        If pvOnBeforeNotify(ucsSfdRead) Then
            RaiseEvent OnReceive
        End If
    End If
    If pvTlsIsClosed(m_uCtx) Then
        If pvOnBeforeNotify(ucsSfdClose) Then
            RaiseEvent OnClose
        End If
    ElseIf Not bResult Then
        RaiseEvent OnError(-1, ucsSfdRead)
        GoTo QH
    End If
    '--- success
    pvOnReceive = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOnSend() As Boolean
    Const FUNC_NAME     As String = "pvOnSend"
    Dim lBytes          As Long
    
    On Error GoTo EH
    If Not pvTlsIsReady(m_uCtx) Then
        GoTo QH
    End If
    Do While m_lSendPos < m_lSendSize
        lBytes = m_oSocket.Send(VarPtr(m_baSend(m_lSendPos)), IIf(m_lSendSize - m_lSendPos > 32768, 32768, m_lSendSize - m_lSendPos))
        If m_oSocket.LastError = WSAEWOULDBLOCK Then
            Exit Do
        ElseIf lBytes < 0 Then
            GoTo QH
        Else
            m_lSendPos = m_lSendPos + lBytes
        End If
    Loop
    If m_oSocket.LastError <> WSAEWOULDBLOCK Then
        m_oSocket.PostEvent ucsSfdWrite
        m_oSocket.frLastError = WSAEWOULDBLOCK
    End If
    '--- success
    pvOnSend = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOnShutdown() As Boolean
    Const FUNC_NAME     As String = "pvOnShutdown"
    Dim baOutput()      As Byte
    
    On Error GoTo EH
    If Not pvTlsShutdown(m_uCtx, baOutput, 0) Then
        GoTo QH
    End If
    If Not m_oSocket.SendArray(baOutput) Then
        GoTo QH
    End If
    If m_oSocket.LastError <> WSAEWOULDBLOCK Then
        m_oSocket.PostEvent ucsSfdWrite
        m_oSocket.frLastError = WSAEWOULDBLOCK
    End If
    '--- success
    pvOnShutdown = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOnBeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum) As Boolean
    Dim bCancel         As Boolean
    
    RaiseEvent BeforeNotify(EventMask, bCancel)
    pvOnBeforeNotify = Not bCancel
End Function

'=========================================================================
' TLS support
'=========================================================================

Private Function pvTlsInitClient( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum = ucsTlsSupportAll) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsSetLastError uEmpty, vbNullString
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        pvTlsArrayRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
        If (LocalFeatures And ucsTlsSupportTls13) <> 0 Then
            '--- note: uCtx.ClientPublic has to be ready for pvTlsBuildClientHello
            pvTlsSetupExchEccGroup uEmpty, TLS_GROUP_X25519
        End If
    End With
    uCtx = uEmpty
    '--- success
    pvTlsInitClient = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Variant) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsSetLastError uEmpty, vbNullString
        .IsServer = True
        .State = ucsTlsStateExpectClientHello
        .RemoteHostName = RemoteHostName
        .LocalFeatures = ucsTlsSupportTls13
        Set .LocalCertificates = Certificates
        If Not IsMissing(PrivateKey) Then
            .LocalPrivateKey = PrivateKey
        End If
        pvTlsArrayRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
    End With
    uCtx = uEmpty
    '--- success
    pvTlsInitServer = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If .State = ucsTlsStateHandshakeStart Then
            .SendPos = pvTlsBuildClientHello(uCtx, .SendBuffer, .SendPos)
            .State = ucsTlsStateExpectServerHello
        Else
            If lSize < 0 Then
                lSize = pvArraySize(baInput)
            End If
            If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
                pvTlsSetLastError uCtx, .LastError, .LastAlertCode
                GoTo QH
            End If
        End If
        '--- success
        pvTlsHandshake = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baPlainText)
        End If
        If lSize = 0 Then
            '--- flush
            pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
            Erase .SendBuffer
            .SendPos = 0
            '--- success
            pvTlsSend = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        .SendPos = pvTlsBuildApplicationData(uCtx, .SendBuffer, .SendPos, baPlainText, lSize, TLS_CONTENT_TYPE_APPDATA)
        '--- success
        pvTlsSend = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If lSize = 0 And lPos = 0 Then
            '--- flush
            pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
            Erase .DecrBuffer
            .DecrPos = 0
            '--- success
            pvTlsReceive = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
        If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
            pvTlsSetLastError uCtx, .LastError, .LastAlertCode
            GoTo QH
        End If
        '--- success
        pvTlsReceive = True
QH:
        '--- swap-out
        pvArraySwap baPlainText, lPos, .DecrBuffer, .DecrPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        .SendPos = pvTlsBuildAlert(uCtx, .SendBuffer, .SendPos, uscTlsAlertCloseNotify, TLS_ALERT_LEVEL_WARNING)
        .State = ucsTlsStateShutdown
        '--- success
        pvTlsShutdown = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Private Function pvTlsIsClosed(uCtx As UcsTlsContext) As Boolean
    pvTlsIsClosed = (uCtx.State = ucsTlsStateClosed)
End Function

Private Function pvTlsIsReady(uCtx As UcsTlsContext) As Boolean
    pvTlsIsReady = (uCtx.State >= ucsTlsStatePostHandshake)
End Function

Private Function pvTlsGetLastError(uCtx As UcsTlsContext) As String
    pvTlsGetLastError = uCtx.LastError
    If uCtx.LastAlertCode <> -1 Then
        pvTlsGetLastError = IIf(LenB(pvTlsGetLastError) <> 0, pvTlsGetLastError & ". ", vbNullString) & Replace(STR_FORMAT_ALERT, "%1", pvTlsGetLastAlert(uCtx))
    End If
End Function

Private Function pvTlsGetLastAlert(uCtx As UcsTlsContext, Optional AlertCode As UcsTlsAlertDescriptionsEnum) As String
    Static vTexts       As Variant
    
    AlertCode = uCtx.LastAlertCode
    If AlertCode >= 0 Then
        If IsEmpty(vTexts) Then
            vTexts = SplitOrReindex(STR_VL_ALERTS, "|")
        End If
        If AlertCode <= UBound(vTexts) Then
            pvTlsGetLastAlert = vTexts(AlertCode)
        End If
        If LenB(pvTlsGetLastAlert) = 0 Then
            pvTlsGetLastAlert = Replace(STR_UNKNOWN, "%1", AlertCode)
        End If
    End If
End Function

Private Function pvTlsBuildClientHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim vElem           As Variant
    Dim baTemp()        As Byte
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Handshake Header
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_HELLO)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, TLS_LOCAL_LEGACY_VERSION, Size:=2)
                lPos = pvWriteArray(baOutput, lPos, .LocalExchRandom)
                '--- Legacy Session ID
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    If pvArraySize(.LocalSessionID) = 0 And (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- non-empty for TLS 1.2 compatibility
                        pvTlsArrayRandom baTemp, TLS_HELLO_RANDOM_SIZE
                        lPos = pvWriteArray(baOutput, lPos, baTemp)
                    Else
                        lPos = pvWriteArray(baOutput, lPos, .LocalSessionID)
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Cipher Suites
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    For Each vElem In pvTlsPrepareCipherSuitsOrder(.LocalFeatures)
                        lPos = pvWriteLong(baOutput, lPos, vElem, Size:=2)
                    Next
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Legacy Compression Methods
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Extensions
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    If LenB(.RemoteHostName) <> 0 Then
                        '--- Extension - Server Name
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SERVER_NAME, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SERVER_NAME_TYPE_HOSTNAME) '--- FQDN
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteString(baOutput, lPos, .RemoteHostName)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Supported Groups
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_GROUPS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            If pvCryptoIsSupported(ucsTlsAlgoExchX25519) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_X25519 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                                End If
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP256R1 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_SECP256R1, Size:=2)
                                End If
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP384R1 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_SECP384R1, Size:=2)
                                End If
                            End If
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- Extension - EC Point Formats
                        pvArrayByte baTemp, 0, TLS_EXTENSION_TYPE_EC_POINT_FORMAT, 0, 2, 1, 0
                        lPos = pvWriteArray(baOutput, lPos, baTemp)     '--- uncompressed only
                        '--- Extension - Renegotiation Info
                        pvArrayByte baTemp, &HFF, 1, 0, 1, 0
                        lPos = pvWriteArray(baOutput, lPos, baTemp)     '--- empty info
                    End If
                    '--- Extension - Signature Algorithms
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512, Size:=2)
                            If pvCryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA512, Size:=2)
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA512, Size:=2)
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha1) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2)
                            End If
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.LocalFeatures And ucsTlsSupportTls13) <> 0 Then
                        '--- Extension - Key Share
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, .ExchGroup, Size:=2)
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        '--- Extension - Supported Versions
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13, Size:=2)
                                If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS12, Size:=2)
                                End If
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        If .HelloRetryRequest Then
                            '--- Extension - Cookie
                            lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_COOKIE, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                    lPos = pvWriteArray(baOutput, lPos, .HelloRetryCookie)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        End If
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
    End With
    pvTlsBuildClientHello = lPos
QH:
End Function

Private Function pvTlsBuildClientLegacyKeyExchange(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Const FUNC_NAME     As String = "pvTlsBuildClientLegacyKeyExchange"
    Dim baLocalIV()     As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    Dim lRecordPos      As Long
    Dim baTemp()        As Byte
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Handshake Client Key Exchange
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_KEY_EXCHANGE)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                If pvArraySize(.LocalExchRsaEncrPriv) > 0 Then
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteArray(baOutput, lPos, .LocalExchRsaEncrPriv)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                Else
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                End If
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- Legacy Change Cipher Spec
        pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
        lPos = pvWriteArray(baOutput, lPos, baTemp)
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            End If
            '--- Handshake Finish
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                pvTlsKdfLegacyPrf baVerifyData, .DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, 12
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lMessageSize = lPos - lMessagePos
            '--- note: *before* allocating space for the authentication tag
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            '--- encrypt message
            pvArrayAllocate baAad, LNG_LEGACY_AAD_SIZE, FUNC_NAME & ".baAad"
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize
            .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
            lMessagePos = lRecordPos + 5
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
    End With
    pvTlsBuildClientLegacyKeyExchange = lPos
QH:
End Function

Private Function pvTlsBuildClientHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    With uCtx
        '--- Legacy Change Cipher Spec
        pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
        lPos = pvWriteArray(baOutput, lPos, baTemp)
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Client Handshake Finished
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .LocalTrafficSecret, "finished", baEmpty, .DigestSize
                pvTlsHkdfExtract baVerifyData, .DigestAlgo, baTemp, baHandshakeHash
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            '--- Record Type
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
        pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildClientHandshakeFinished = lPos
QH:
End Function

Private Function pvTlsBuildServerHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Handshake Header
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_SERVER_HELLO)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, TLS_LOCAL_LEGACY_VERSION, Size:=2)
                lPos = pvWriteArray(baOutput, lPos, .LocalExchRandom)
                '--- Legacy Session ID
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteArray(baOutput, lPos, .RemoteSessionID)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Cipher Suite
                lPos = pvWriteLong(baOutput, lPos, .CipherSuite, Size:=2)
                '--- Legacy Compression Method
                lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
                '--- Extensions
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    '--- Extension - Key Share
                    If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_TYPE_KEY_SHARE) Then
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, .ExchGroup, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Supported Versions
                    If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS) Then
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13, Size:=2)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
    End With
    pvTlsBuildServerHello = lPos
End Function

Private Function pvTlsBuildServerHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baHandshakeHash() As Byte
    Dim lHandshakePos   As Long
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim baSignature()   As Byte
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    With uCtx
        '--- Legacy Change Cipher Spec
        pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
        lPos = pvWriteArray(baOutput, lPos, baTemp)
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            lMessagePos = lPos
            '--- Server Encrypted Extensions
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    lPos = lPos '--- empty
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            '--- Server Certificate
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                '--- certificate request context
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = lPos '--- empty
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                    For lIdx = 1 To .LocalCertificates.Count
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                            baCert = .LocalCertificates.Item(lIdx)
                            lPos = pvWriteArray(baOutput, lPos, baCert)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        '--- certificate extensions
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = lPos '--- empty
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    Next
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Server Certificate Verify
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, .LocalSignatureType, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                    lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                    lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                    pvTlsSignatureSign .LocalPrivateKey, .LocalSignatureType, baVerifyData, baSignature
                    lPos = pvWriteArray(baOutput, lPos, baSignature)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Server Handshake Finished
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .LocalTrafficSecret, "finished", baEmpty, .DigestSize
                pvTlsHkdfExtract baVerifyData, .DigestAlgo, baTemp, baHandshakeHash
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Record Type
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
        pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildServerHandshakeFinished = lPos
End Function

Private Function pvTlsBuildApplicationData(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, ByVal lSize As Long, ByVal lContentType As Long) As Long
    Const FUNC_NAME     As String = "pvTlsBuildApplicationData"
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    
    With uCtx
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            End If
            lMessagePos = lPos
            If lSize > 0 Then
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baData(0)), lSize)
            End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                lPos = pvWriteLong(baOutput, lPos, lContentType)
            End If
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- encrypt message
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize
        ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            pvArrayAllocate baAad, LNG_LEGACY_AAD_SIZE, FUNC_NAME & ".baAad"
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize
        End If
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildApplicationData = lPos
End Function

Private Function pvTlsBuildAlert(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, ByVal eAlertDesc As UcsTlsAlertDescriptionsEnum, ByVal lAlertLevel As Long) As Long
    Const FUNC_NAME     As String = "pvTlsBuildAlert"
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    
    With uCtx
        '--- for TLS 1.3 -> tunnel alert through application data encryption
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            pvArrayAllocate baLocalIV, 2, FUNC_NAME & ".baLocalIV"
            baLocalIV(0) = lAlertLevel
            baLocalIV(1) = eAlertDesc
            pvTlsBuildAlert = pvTlsBuildApplicationData(uCtx, baOutput, lPos, baLocalIV, UBound(baLocalIV) + 1, TLS_CONTENT_TYPE_ALERT)
            GoTo QH
        End If
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_ALERT)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
                If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                    lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
                End If
            End If
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, lAlertLevel)
            lPos = pvWriteLong(baOutput, lPos, eAlertDesc)
            lMessageSize = lPos - lMessagePos
            If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            End If
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            '--- encrypt message
            pvArrayAllocate baAad, LNG_LEGACY_AAD_SIZE, FUNC_NAME & ".baAad"
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            pvTlsAeadEncrypt .AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize
            .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
        End If
    End With
    pvTlsBuildAlert = lPos
QH:
End Function

Private Function pvTlsParsePayload(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lPrevPos        As Long
    Dim lRecvSize       As Long
    
    On Error GoTo EH
    If lSize > 0 Then
    With uCtx
        .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        lPrevPos = .RecvPos
        .RecvPos = pvTlsParseRecord(uCtx, .RecvBuffer, .RecvPos, sError, eAlertCode)
        If LenB(sError) <> 0 Then
            GoTo QH
        End If
        lRecvSize = lPrevPos - .RecvPos
        If .RecvPos > 0 And lRecvSize > 0 Then
            Call CopyMemory(.RecvBuffer(0), .RecvBuffer(.RecvPos), lRecvSize)
        End If
        .RecvPos = IIf(lRecvSize > 0, lRecvSize, 0)
    End With
    End If
    '--- success
    pvTlsParsePayload = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseRecord(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim lRecordProtocol As Long
    Dim baRemoteIV()    As Byte
    Dim lPos            As Long
    Dim lEnd            As Long
    Dim baAad()         As Byte
    Dim bResult         As Boolean
    
    On Error GoTo EH
    With uCtx
    Do While lPos + 6 <= lSize
        lRecordPos = lPos
        lPos = pvReadLong(baInput, lPos, lRecordType)
        lPos = pvReadLong(baInput, lPos, lRecordProtocol, Size:=2)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lRecordSize)
            If lRecordSize > IIf(lRecordType = TLS_CONTENT_TYPE_APPDATA, TLS_MAX_ENCRYPTED_RECORD_SIZE, TLS_MAX_PLAINTEXT_RECORD_SIZE) Then
                sError = ERR_RECORD_TOO_BIG
                eAlertCode = uscTlsAlertDecodeError
                GoTo QH
            End If
            If lPos + lRecordSize > lSize Then
                '--- back off and bail out early
                lPos = pvReadEndOfBlock(baInput, lPos + lRecordSize, .BlocksStack)
                lPos = lRecordPos
                Exit Do
            End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_ALERT
                lEnd = lPos + lRecordSize
                If lRecordSize > 2 Then
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        '--- note: TLS_CONTENT_TYPE_ALERT encryption is tunneled through TLS_CONTENT_TYPE_APPDATA
                        sError = ERR_RECORD_TOO_BIG
                        eAlertCode = uscTlsAlertDecodeError
                        GoTo QH
                    ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        pvTlsPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad
                        bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                End If
HandleAlertContent:
                If lPos + 1 < lEnd Then
                    Select Case baInput(lPos)
                    Case TLS_ALERT_LEVEL_FATAL
                        sError = ERR_FATAL_ALERT
                        eAlertCode = baInput(lPos + 1)
                        GoTo QH
                    Case TLS_ALERT_LEVEL_WARNING
                        .LastAlertCode = baInput(lPos + 1)
                        Debug.Print pvTlsGetLastAlert(uCtx) & " (TLS_ALERT_LEVEL_WARNING)", Timer
                        If .LastAlertCode = uscTlsAlertCloseNotify Then
                            .State = ucsTlsStateClosed
                        End If
                    End Select
                End If
                '--- note: skip AEAD's authentication tag
                lPos = lRecordPos + lRecordSize + 5
            Case TLS_CONTENT_TYPE_HANDSHAKE
                lEnd = lPos + lRecordSize
                If .State = ucsTlsStateExpectServerFinished Then
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        '--- note: ucsTlsStateExpectServerFinished is TLS 1.2 state only
                        sError = Replace(Replace(ERR_UNEXPECTED_PROTOCOL, "%1", "ucsTlsStateExpectServerFinished"), "%2", .ProtocolVersion)
                        eAlertCode = uscTlsAlertInternalError
                        GoTo QH
                    ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        pvTlsPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad
                        bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                End If
HandleHandshakeContent:
                If .MessSize > 0 Then
                    .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                    If Not pvTlsParseHandshake(uCtx, .MessBuffer, .MessPos, .MessSize, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If .MessPos >= .MessSize Then
                        Erase .MessBuffer
                        .MessSize = 0
                        .MessPos = 0
                    End If
                Else
                    If Not pvTlsParseHandshake(uCtx, baInput, lPos, lEnd, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If lPos < lEnd Then
                        .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                        .MessPos = 0
                    End If
                End If
                '--- note: skip AEAD's authentication tag
                lPos = lRecordPos + lRecordSize + 5
            Case TLS_CONTENT_TYPE_APPDATA
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    pvArrayXor baRemoteIV, .RemoteTrafficIV, .RemoteTrafficSeqNo
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baInput, lRecordPos, LNG_AAD_SIZE, baInput, lPos, lRecordSize)
                ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    pvTlsPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                Else
                    bResult = False
                End If
                If Not bResult Then
                    sError = ERR_DECRYPTION_FAILED
                    eAlertCode = uscTlsAlertBadRecordMac
                    GoTo QH
                End If
                .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    lEnd = lPos + lRecordSize - .TagSize - 1
                    '--- trim zero padding at the end of decrypted record
                    Do While baInput(lEnd) = 0
                        lEnd = lEnd - 1
                    Loop
                    lRecordType = baInput(lEnd)
                    Select Case lRecordType
                    Case TLS_CONTENT_TYPE_ALERT
                        GoTo HandleAlertContent
                    Case TLS_CONTENT_TYPE_HANDSHAKE
                        GoTo HandleHandshakeContent
                    Case TLS_CONTENT_TYPE_APPDATA
                        '--- do nothing
                    Case Else
                        sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
                        eAlertCode = uscTlsAlertHandshakeFailure
                        GoTo QH
                    End Select
                End If
                .DecrPos = pvWriteBuffer(.DecrBuffer, .DecrPos, VarPtr(baInput(lPos)), lEnd - lPos)
                '--- note: skip AEAD's authentication tag or zero padding
                lPos = lRecordPos + lRecordSize + 5
            Case Else
                sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
                eAlertCode = uscTlsAlertHandshakeFailure
                GoTo QH
            End Select
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    Loop
    End With
    pvTlsParseRecord = lPos
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshake(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lMessageType    As Long
    Dim baMessage()     As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim lRequestUpdate  As Long
    Dim lCurveType      As Long
    Dim lNamedCurve     As Long
    Dim lSignatureType  As Long
    Dim lSignatureSize  As Long
    Dim baSignature()   As Byte
    Dim baCert()        As Byte
    Dim lCertSize       As Long
    Dim lCertEnd        As Long
    Dim lSignPos        As Long
    Dim lSignSize       As Long
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    On Error GoTo EH
    With uCtx
        Do While lPos < lEnd
            lMessagePos = lPos
            lPos = pvReadLong(baInput, lPos, lMessageType)
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lMessageSize)
                If lPos + lMessageSize > lEnd Then
                    '--- back off and bail out early
                    lPos = pvReadEndOfBlock(baInput, lPos + lMessageSize, .BlocksStack)
                    lPos = lMessagePos
                    Exit Do
                End If
                Select Case .State
                Case ucsTlsStateExpectServerHello
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                        If Not pvTlsParseHandshakeServerHello(uCtx, baInput, lPos, lRecordProtocol, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If .HelloRetryRequest Then
                            '--- after HelloRetryRequest -> replace HandshakeMessages w/ 'synthetic handshake message'
                            pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                            Erase .HandshakeMessages
                            lVerifyPos = pvWriteLong(.HandshakeMessages, 0, TLS_HANDSHAKE_TYPE_MESSAGE_HASH)
                            lVerifyPos = pvWriteLong(.HandshakeMessages, lVerifyPos, .DigestSize, Size:=3)
                            lVerifyPos = pvWriteArray(.HandshakeMessages, lVerifyPos, baHandshakeHash)
                        Else
                            .State = ucsTlsStateExpectExtensions
                        End If
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectServerHello"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectServerHello
                    If .State = ucsTlsStateExpectServerHello And .HelloRetryRequest Then
                        .SendPos = pvTlsBuildClientHello(uCtx, .SendBuffer, .SendPos)
                    End If
                    If .State = ucsTlsStateExpectExtensions And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        pvTlsDeriveHandshakeSecrets uCtx
                    End If
                Case ucsTlsStateExpectExtensions
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lCertSize)
                                lPos = lPos + lCertSize ' skip RemoteCertReqContext
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        End If
                        Set .RemoteCertificates = New Collection
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                            lCertEnd = lPos + lCertSize
                            Do While lPos < lCertEnd
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                                    lPos = pvReadArray(baInput, lPos, baCert, lCertSize)
                                    .RemoteCertificates.Add baCert
                                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lCertSize)
                                        '--- certificate extensions -> skip
                                        lPos = lPos + lCertSize
                                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                End If
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                        lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lCertSize)
                            lPos = pvReadArray(baInput, lPos, baSignature, lCertSize)
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                            sError = ERR_NO_SERVER_CERTIFICATE
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                        lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                        lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                        If Not pvTlsSignatureVerify(baCert, lSignatureType, baVerifyData, baSignature, sError, eAlertCode) Then
                            GoTo QH
                        End If
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                        pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .RemoteTrafficSecret, "finished", baEmpty, .DigestSize
                        pvTlsHkdfExtract baVerifyData, .DigestAlgo, baTemp, baHandshakeHash
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = ERR_SERVER_HANDSHAKE_FAILED
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case TLS_HANDSHAKE_TYPE_SERVER_KEY_EXCHANGE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            lSignPos = lPos
                            lPos = pvReadLong(baInput, lPos, lCurveType)
                            If lCurveType <> 3 Then '--- 3 = named_curve
                                sError = ERR_SERVER_HANDSHAKE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            lPos = pvReadLong(baInput, lPos, lNamedCurve, Size:=2)
                            pvTlsSetupExchEccGroup uCtx, lNamedCurve
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lSignatureSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            lSignSize = lPos - lSignPos
                            '--- signature
                            lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, baSignature, lSignatureSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                                sError = ERR_NO_SERVER_CERTIFICATE
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            lVerifyPos = pvWriteArray(baVerifyData, 0, .LocalExchRandom)
                            lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, .RemoteExchRandom)
                            lVerifyPos = pvWriteBuffer(baVerifyData, lVerifyPos, VarPtr(baInput(lSignPos)), lSignSize)
                            If Not pvTlsSignatureVerify(baCert, lSignatureType, baVerifyData, baSignature, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            pvTlsDeriveLegacySecrets uCtx
                        End If
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            .State = ucsTlsStateExpectServerFinished
                        End If
                        lPos = lPos + lMessageSize
                    Case Else
                        '--- do nothing
                        lPos = lPos + lMessageSize
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectExtensions
                    If .State = ucsTlsStateExpectServerFinished And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If pvTlsCipherSuiteUseRsaCertificate(.CipherSuite) Then
                            If Not SearchCollection(.RemoteCertificates, 1, baCert) Then
                                sError = ERR_NO_SERVER_CERTIFICATE
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            pvTlsSetupExchRsaCertificate uCtx, baCert
                            pvTlsDeriveLegacySecrets uCtx
                        End If
                        .SendPos = pvTlsBuildClientLegacyKeyExchange(uCtx, .SendBuffer, .SendPos)
                    End If
                    If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        .SendPos = pvTlsBuildClientHandshakeFinished(uCtx, .SendBuffer, .SendPos)
                        pvTlsDeriveApplicationSecrets uCtx
                        '--- not used past handshake
                        Erase .HandshakeMessages
                    End If
                Case ucsTlsStateExpectServerFinished
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                            pvTlsKdfLegacyPrf baVerifyData, .DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, 12
                            If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                                sError = ERR_SERVER_HANDSHAKE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            .State = ucsTlsStatePostHandshake
                            '--- not used past handshake
                            Erase .HandshakeMessages
                        Else
                            GoTo InvalidState
                        End If
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectServerFinished"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                Case ucsTlsStateExpectClientHello
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CLIENT_HELLO
                        If Not pvTlsParseHandshakeClientHello(uCtx, baInput, lPos, lRecordProtocol, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        .State = ucsTlsStateExpectClientFinished
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectClientHello"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectClientHello
                    If .State = ucsTlsStateExpectClientFinished Then
                        .SendPos = pvTlsBuildServerHello(uCtx, .SendBuffer, .SendPos)
                        pvTlsDeriveHandshakeSecrets uCtx
                        .SendPos = pvTlsBuildServerHandshakeFinished(uCtx, .SendBuffer, .SendPos)
                    End If
                Case ucsTlsStateExpectClientFinished
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                        pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .RemoteTrafficSecret, "finished", baEmpty, .DigestSize
                        pvTlsHkdfExtract baVerifyData, .DigestAlgo, baTemp, baHandshakeHash
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = ERR_SERVER_HANDSHAKE_FAILED
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectClientFinished"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    '--- post-process ucsTlsStateExpectClientFinished
                    If .State = ucsTlsStatePostHandshake Then
                        pvTlsDeriveApplicationSecrets uCtx
                        '--- not used past handshake
                        Erase .HandshakeMessages
                        Set .RemoteTickets = New Collection
                    End If
                Case ucsTlsStatePostHandshake
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        If Not .RemoteTickets Is Nothing Then
                            .RemoteTickets.Add baMessage
                        End If
                    Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                        Debug.Print "Received TLS_HANDSHAKE_TYPE_KEY_UPDATE", Timer
                        If lMessageSize = 1 Then
                            lRequestUpdate = baInput(lPos)
                        Else
                            lRequestUpdate = -1
                        End If
                        pvTlsDeriveKeyUpdate uCtx, lRequestUpdate <> 0
                        If lRequestUpdate <> 0 Then
                            '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                            pvArrayByte baTemp, TLS_HANDSHAKE_TYPE_KEY_UPDATE, 0, 0, 1, 0
                            pvTlsBuildApplicationData uCtx, baMessage, 0, baTemp, -1, TLS_CONTENT_TYPE_APPDATA
                            .SendPos = pvWriteArray(.SendBuffer, .SendPos, baMessage)
                        End If
                        lPos = lPos + lMessageSize
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStatePostHandshake"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                Case Else
InvalidState:
                    sError = Replace(ERR_INVALID_STATE_HANDSHAKE, "%1", .State)
                    eAlertCode = uscTlsAlertHandshakeFailure
                    GoTo QH
                End Select
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        Loop
    End With
    '--- success
    pvTlsParseHandshake = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeServerHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Static baHelloRetryRandom() As Byte
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExchGroup      As Long
    Dim lBlockSize      As Long
    
    On Error GoTo EH
    If pvArraySize(baHelloRetryRandom) = 0 Then
        pvArrayByte baHelloRetryRandom, &HCF, &H21, &HAD, &H74, &HE5, &H9A, &H61, &H11, &HBE, &H1D, &H8C, &H2, &H1E, &H65, &HB8, &H91, &HC2, &HA2, &H11, &H16, &H7A, &HBB, &H8C, &H5E, &H7, &H9E, &H9, &HE2, &HC8, &HA8, &H33, &H9C
    End If
    With uCtx
        .ProtocolVersion = lRecordProtocol
        lPos = pvReadLong(baInput, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baInput, lPos, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE)
        If .HelloRetryRequest Then
            '--- clear HelloRetryRequest
            .HelloRetryRequest = False
            .HelloRetryCipherSuite = 0
            .HelloRetryExchGroup = 0
            Erase .HelloRetryCookie
        Else
            .HelloRetryRequest = (StrConv(.RemoteExchRandom, vbUnicode) = StrConv(baHelloRetryRandom, vbUnicode))
        End If
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSize)
            lPos = pvReadArray(baInput, lPos, .RemoteSessionID, lSize)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        lPos = pvReadLong(baInput, lPos, lCipherSuite, Size:=2)
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        Debug.Print "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName, Timer
        If .HelloRetryRequest Then
            .HelloRetryCipherSuite = lCipherSuite
        End If
        lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        Debug.Assert lLegacyCompress = 0
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lExtSize)
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                        If lExtSize < 2 Then
                            sError = ERR_INVALID_SIZE_KEY_SHARE
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadLong(baInput, lPos, lExchGroup, Size:=2)
                        pvTlsSetupExchEccGroup uCtx, lExchGroup
                        If .HelloRetryRequest Then
                            .HelloRetryExchGroup = lExchGroup
                        Else
                            If lExtSize <= 4 Then
                                sError = ERR_INVALID_SIZE_REMOTE_KEY
                                eAlertCode = uscTlsAlertDecodeError
                                GoTo QH
                            End If
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lBlockSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        End If
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        If lExtSize <> 2 Then
                            sError = ERR_INVALID_SIZE_VERSIONS
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadLong(baInput, lPos, .ProtocolVersion, Size:=2)
                    Case TLS_EXTENSION_TYPE_COOKIE
                        If Not .HelloRetryRequest Then
                            sError = ERR_COOKIE_NOT_ALLOWED
                            eAlertCode = uscTlsAlertIllegalParameter
                            GoTo QH
                        End If
                        lPos = pvReadArray(baInput, lPos, .HelloRetryCookie, lExtSize)
                    Case Else
                        lPos = lPos + lExtSize
                    End Select
                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    End With
    '--- success
    pvTlsParseHandshakeServerHello = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeClientHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lCipherPref     As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim lExchGroup      As Long
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lProtocolVersion As Long
    Dim lSignatureType  As Long
    Dim cCipherPrefs    As Collection
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim uCertInfo       As UcsKeyInfo
    Dim lHashSize       As Long
    
    On Error GoTo EH
    Set cCipherPrefs = New Collection
    For Each vElem In pvTlsPrepareCipherSuitsOrder(ucsTlsSupportTls13)
        cCipherPrefs.Add cCipherPrefs.Count, "#" & vElem
    Next
    lCipherPref = 1000
    With uCtx
        If SearchCollection(.LocalCertificates, 1, RetVal:=baCert) Then
            If Not pvCryptoAsn1DecodeCertificate(baCert, uCertInfo) Then
                sError = ERR_UNSUPPORTED_CERTIFICATE
                eAlertCode = uscTlsAlertHandshakeFailure
                GoTo QH
            End If
        End If
        .ProtocolVersion = lRecordProtocol
        lPos = pvReadLong(baInput, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baInput, lPos, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSize)
            lPos = pvReadArray(baInput, lPos, .RemoteSessionID, lSize)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lIdx, Size:=2)
                If SearchCollection(cCipherPrefs, "#" & lIdx, RetVal:=vElem) Then
                    If vElem < lCipherPref Then
                        lCipherSuite = lIdx
                        lCipherPref = vElem
                    End If
                End If
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        If lCipherSuite = 0 Then
            sError = ERR_NO_SUPPORTED_CIPHER_SUITE
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End If
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        Debug.Print "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName, Timer
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        Debug.Assert lLegacyCompress = 0
        '--- extensions
        Set .RemoteExtensions = New Collection
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lExtSize)
                    lExtEnd = lPos + lExtSize
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                        If lExtSize < 4 Then
                            sError = ERR_INVALID_SIZE_KEY_SHARE
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                            lBlockEnd = lPos + lBlockSize
                            Do While lPos < lBlockEnd
                                lPos = pvReadLong(baInput, lPos, lExchGroup, Size:=2)
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                    If lExchGroup = TLS_GROUP_X25519 Then
                                        If lBlockSize <> TLS_X25519_KEY_SIZE Then
                                            sError = ERR_INVALID_REMOTE_KEY
                                            eAlertCode = uscTlsAlertIllegalParameter
                                            GoTo QH
                                        End If
                                        lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lBlockSize)
                                        pvTlsSetupExchEccGroup uCtx, lExchGroup
                                    Else
                                        lPos = lPos + lBlockSize
                                    End If
                                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                            Do While lPos < lExtEnd
                                lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                                lHashSize = pvTlsSignatureHashSize(lSignatureType)
                                Select Case lSignatureType
                                Case TLS_SIGNATURE_RSA_PKCS1_SHA1
                                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                        If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha1) Then
                                            .LocalSignatureType = lSignatureType
                                        End If
                                    End If
                                Case TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
                                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                        If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                            .LocalSignatureType = lSignatureType
                                        End If
                                    End If
                                Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
                                    '--- PSS w/ SHA512 fails on short key lengths (min PSS size is 2 + lHashSize + lSaltSize w/ lSaltSize=lHashSize)
                                    If (uCertInfo.BitLen + 7) \ 8 > 2 + 2 * lHashSize Then
                                        If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                                            .LocalSignatureType = lSignatureType
                                        End If
                                    End If
                                Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
                                    If (uCertInfo.BitLen + 7) \ 8 > 2 + 2 * lHashSize Then
                                        If uCertInfo.AlgoObjId = szOID_RSA_SSA_PSS And pvCryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                                            .LocalSignatureType = lSignatureType
                                        End If
                                    End If
                                Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
                                    If uCertInfo.AlgoObjId = szOID_ECC_PUBLIC_KEY Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                End Select
                                If .LocalSignatureType <> 0 Then
                                    lPos = lExtEnd
                                End If
                            Loop
                            If .LocalSignatureType = 0 Then
                                sError = ERR_NEGOTIATE_SIGNATURE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lBlockSize)
                            Do While lPos < lExtEnd
                                lPos = pvReadLong(baInput, lPos, lProtocolVersion, Size:=2)
                                If lProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                    lPos = lExtEnd
                                End If
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        If lProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 Then
                            sError = ERR_UNSUPPORTED_PROTOCOL
                            eAlertCode = uscTlsAlertProtocolVersion
                            GoTo QH
                        End If
                        .ProtocolVersion = lProtocolVersion
                    Case Else
                        lPos = lPos + lExtSize
                    End Select
                    If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                        .RemoteExtensions.Add lExtType, "#" & lExtType
                    End If
                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    End With
    '--- success
    pvTlsParseHandshakeClientHello = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Sub pvTlsSetupExchEccGroup(uCtx As UcsTlsContext, ByVal lExchGroup As Long)
    Const FUNC_NAME     As String = "pvTlsSetupExchEccGroup"
    
    With uCtx
        If .ExchGroup <> lExchGroup Then
            .ExchGroup = lExchGroup
            Select Case lExchGroup
            Case TLS_GROUP_X25519
                .ExchAlgo = ucsTlsAlgoExchX25519
                If Not CryptoEccCurve25519MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "Curve25519")
                End If
            Case TLS_GROUP_SECP256R1
                .ExchAlgo = ucsTlsAlgoExchSecp256r1
                If Not CryptoEccSecp256r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp256r1")
                End If
            Case TLS_GROUP_SECP384R1
                .ExchAlgo = ucsTlsAlgoExchSecp384r1
                If Not CryptoEccSecp384r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp384r1")
                End If
            Case Else
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_EXCH_GROUP, "%1", "0x" & Hex$(.ExchGroup))
            End Select
        End If
    End With
End Sub

Private Sub pvTlsSetupExchRsaCertificate(uCtx As UcsTlsContext, baCert() As Byte)
    Const FUNC_NAME     As String = "pvTlsSetupExchRsaCertificate"
    Dim uCertInfo       As UcsKeyInfo
    Dim baEnc()         As Byte
    
    With uCtx
        .ExchAlgo = ucsTlsAlgoExchCertificate
        pvTlsArrayRandom .LocalExchPrivate, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2 '--- always 48
        pvWriteLong .LocalExchPrivate, 0, TLS_LOCAL_LEGACY_VERSION, Size:=2
        If Not pvCryptoAsn1DecodeCertificate(baCert, uCertInfo) Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_UNSUPPORTED_CERTIFICATE
        End If
        If Not pvCryptoEmePkcs1Encode(baEnc, .LocalExchPrivate, uCertInfo.BitLen) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "pvCryptoEmePkcs1Encode")
        End If
        If Not CryptoRsaModExp(baEnc, uCertInfo.PubExp, uCertInfo.Modulus, .LocalExchRsaEncrPriv) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoRsaModExp")
        End If
    End With
End Sub

Private Sub pvTlsSetupCipherSuite(uCtx As UcsTlsContext, ByVal lCipherSuite As Long)
    Const FUNC_NAME     As String = "pvTlsSetupCipherSuite"
    
    With uCtx
        If .CipherSuite <> lCipherSuite Then
            .CipherSuite = lCipherSuite
            Select Case lCipherSuite
            Case TLS_CS_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                .AeadAlgo = ucsTlsAlgoAeadChacha20Poly1305
                .KeySize = TLS_CHACHA20_KEY_SIZE
                .IvSize = TLS_CHACHA20POLY1305_IV_SIZE
                .TagSize = TLS_CHACHA20POLY1305_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = TLS_SHA256_DIGEST_SIZE
            Case TLS_CS_AES_128_GCM_SHA256, TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256, TLS_CS_RSA_WITH_AES_128_GCM_SHA256
                .AeadAlgo = ucsTlsAlgoAeadAes128
                .KeySize = TLS_AES128_KEY_SIZE
                .IvSize = TLS_AESGCM_IV_SIZE
                If lCipherSuite <> TLS_CS_AES_128_GCM_SHA256 Then
                    .IvDynamicSize = 8 '--- AES in TLS 1.2
                End If
                .TagSize = TLS_AESGCM_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = TLS_SHA256_DIGEST_SIZE
            Case TLS_CS_AES_256_GCM_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .AeadAlgo = ucsTlsAlgoAeadAes256
                .KeySize = TLS_AES256_KEY_SIZE
                .IvSize = TLS_AESGCM_IV_SIZE
                If lCipherSuite <> TLS_CS_AES_256_GCM_SHA384 Then
                    .IvDynamicSize = 8 '--- AES in TLS 1.2
                End If
                .TagSize = TLS_AESGCM_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha384
                .DigestSize = TLS_SHA384_DIGEST_SIZE
            Case Else
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_CIPHER_SUITE, "%1", "0x" & Hex$(.CipherSuite))
            End Select
        End If
    End With
End Sub

Private Sub pvTlsPrepareLegacyDecryptParams(uCtx As UcsTlsContext, baInput() As Byte, ByVal lRecordPos As Long, ByVal lRecordSize As Long, lPos As Long, lEnd As Long, baRemoteIV() As Byte, baAad() As Byte)
    Const FUNC_NAME     As String = "pvTlsPrepareLegacyDecryptParams"
    Dim lAadPos         As Long
    
    With uCtx
        lEnd = lPos + lRecordSize - .TagSize
        pvArrayXor baRemoteIV, .RemoteTrafficIV, .RemoteTrafficSeqNo
        If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
            pvWriteBuffer baRemoteIV, .IvSize - .IvDynamicSize, VarPtr(baInput(lPos)), .IvDynamicSize
            lPos = lPos + .IvDynamicSize
        End If
        pvArrayAllocate baAad, LNG_LEGACY_AAD_SIZE, FUNC_NAME & ".baAad"
        lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
        lAadPos = pvWriteLong(baAad, lAadPos, .RemoteTrafficSeqNo, Size:=4)
        lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baInput(lRecordPos)), 3)
        lAadPos = pvWriteLong(baAad, lAadPos, lEnd - lPos, Size:=2)
        Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
    End With
End Sub

Private Function pvTlsPrepareCipherSuitsOrder(ByVal eFilter As UcsTlsLocalFeaturesEnum) As Collection
    Const PREF      As Long = &H1000
    Dim oRetVal     As Collection
    
    Set oRetVal = New Collection
    If (eFilter And ucsTlsSupportTls13) <> 0 Then
        If pvCryptoIsSupported(ucsTlsAlgoExchX25519) Then
            '--- first if AES preferred over Chacha20
            If pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CS_AES_128_GCM_SHA256
            End If
            If pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes256) And pvCryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CS_AES_256_GCM_SHA384
            End If
            If pvCryptoIsSupported(ucsTlsAlgoAeadChacha20Poly1305) Then
                oRetVal.Add TLS_CS_CHACHA20_POLY1305_SHA256
            End If
            '--- least preferred AES
            If Not pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CS_AES_128_GCM_SHA256
            End If
            If Not pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes256) And pvCryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CS_AES_256_GCM_SHA384
            End If
        End If
    End If
    If (eFilter And ucsTlsSupportTls12) <> 0 Then
        If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
            '--- first if AES preferred over Chacha20
            If pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
            If pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
            If pvCryptoIsSupported(ucsTlsAlgoAeadChacha20Poly1305) Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
            End If
            '--- least preferred AES
            If Not pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
            If Not pvCryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And pvCryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
        End If
        '--- no perfect forward secrecy -> least preferred
        If pvCryptoIsSupported(ucsTlsAlgoAeadAes128) Then
            oRetVal.Add TLS_CS_RSA_WITH_AES_128_GCM_SHA256
        End If
        If pvCryptoIsSupported(ucsTlsAlgoAeadAes256) Then
            oRetVal.Add TLS_CS_RSA_WITH_AES_256_GCM_SHA384
        End If
    End If
    Set pvTlsPrepareCipherSuitsOrder = oRetVal
End Function

Private Sub pvTlsSetLastError(uCtx As UcsTlsContext, sError As String, Optional ByVal AlertDesc As UcsTlsAlertDescriptionsEnum = -1)
    With uCtx
        .LastError = sError
        .LastAlertCode = AlertDesc
        If LenB(sError) = 0 Then
            Set .BlocksStack = Nothing
        Else
            If AlertDesc >= 0 Then
                .SendPos = pvTlsBuildAlert(uCtx, .SendBuffer, .SendPos, AlertDesc, TLS_ALERT_LEVEL_FATAL)
            End If
            .State = ucsTlsStateClosed
        End If
    End With
End Sub

'= HMAC-based key derivation functions ===================================

Private Sub pvTlsDeriveHandshakeSecrets(uCtx As UcsTlsContext)
    Const FUNC_NAME     As String = "pvTlsDeriveHandshakeSecrets"
    Dim baHandshakeHash() As Byte
    Dim baEarlySecret() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baSharedSecret() As Byte
    Dim baEmpty()       As Byte
    Dim baEmpty2()      As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
        End If
        pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
        pvArrayAllocate baEmpty2, .DigestSize, FUNC_NAME & ".DigestSize"
        pvTlsHkdfExtract baEarlySecret, .DigestAlgo, baEmpty2, baEmpty2
        pvTlsArrayHash baEmptyHash, .DigestAlgo, baEmpty, 0
        pvTlsHkdfExpandLabel baDerivedSecret, .DigestAlgo, baEarlySecret, "derived", baEmptyHash, .DigestSize
        pvTlsSharedSecret baSharedSecret, .ExchAlgo, .LocalExchPrivate, .RemoteExchPublic
        pvTlsHkdfExtract .HandshakeSecret, .DigestAlgo, baDerivedSecret, baSharedSecret
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .HandshakeSecret, IIf(.IsServer, "c", "s") & " hs traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .HandshakeSecret, IIf(.IsServer, "s", "c") & " hs traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
        .LocalTrafficSeqNo = 0
    End With
End Sub

Private Sub pvTlsDeriveApplicationSecrets(uCtx As UcsTlsContext)
    Const FUNC_NAME     As String = "pvTlsDeriveApplicationSecrets"
    Dim baHandshakeHash() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baEmpty()       As Byte
    Dim baEmpty2()       As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
        End If
        pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
        pvTlsArrayHash baEmptyHash, .DigestAlgo, baEmpty, 0
        pvTlsHkdfExpandLabel baDerivedSecret, .DigestAlgo, .HandshakeSecret, "derived", baEmptyHash, .DigestSize
        pvArrayAllocate baEmpty2, .DigestSize, FUNC_NAME & ".baEmpty2"
        pvTlsHkdfExtract .MasterSecret, .DigestAlgo, baDerivedSecret, baEmpty2
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .MasterSecret, IIf(.IsServer, "c", "s") & " ap traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .MasterSecret, IIf(.IsServer, "s", "c") & " ap traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
        .LocalTrafficSeqNo = 0
    End With
End Sub

Private Sub pvTlsDeriveKeyUpdate(uCtx As UcsTlsContext, ByVal bLocalUpdate As Boolean)
    Const FUNC_NAME     As String = "pvTlsDeriveKeyUpdate"
    Dim baEmpty()       As Byte
    
    With uCtx
        If pvArraySize(.RemoteTrafficSecret) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "RemoteTrafficSecret")
        End If
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .RemoteTrafficSecret, "traffic upd", baEmpty, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        If bLocalUpdate Then
            If pvArraySize(.LocalTrafficSecret) = 0 Then
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "LocalTrafficSecret")
            End If
            pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .LocalTrafficSecret, "traffic upd", baEmpty, .DigestSize
            pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
            pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
            .LocalTrafficSeqNo = 0
        End If
    End With
End Sub

Private Sub pvTlsHkdfExpandLabel(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsHkdfExpandLabel"
    Dim lRetValPos      As Long
    Dim baInfo()        As Byte
    Dim lInfoPos        As Long
    Dim baInput()       As Byte
    Dim lInputPos       As Long
    Dim lIdx            As Long
    Dim baLast()        As Byte
    
    If LenB(sLabel) <> 0 Then
        sLabel = "tls13 " & sLabel
        pvWriteReserved baInfo, 0, 3 + Len(sLabel) + 1 + pvArraySize(baContext)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, lSize, Size:=2)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, Len(sLabel))
        lInfoPos = pvWriteString(baInfo, lInfoPos, sLabel)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, pvArraySize(baContext))
        lInfoPos = pvWriteArray(baInfo, lInfoPos, baContext)
    Else
        baInfo = baContext
    End If
    lIdx = 1
    Do While lRetValPos < lSize
        lInputPos = pvWriteArray(baInput, 0, baLast)
        lInputPos = pvWriteArray(baInput, lInputPos, baInfo)
        lInputPos = pvWriteLong(baInput, lInputPos, lIdx)
        pvTlsArrayHmac baLast, eHash, baKey, baInput, 0, Size:=lInputPos
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baLast)
        lIdx = lIdx + 1
    Loop
    If UBound(baRetVal) <> lSize - 1 Then
        pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    End If
End Sub

Private Sub pvTlsHkdfExtract(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte)
    pvTlsArrayHmac baRetVal, eHash, baKey, baInput, 0
End Sub

'= legacy PRF-based key derivation functions =============================

Private Sub pvTlsDeriveLegacySecrets(uCtx As UcsTlsContext)
    Const FUNC_NAME     As String = "pvTlsDeriveLegacySecrets"
    Dim baPreMasterSecret() As Byte
    Dim baRandom()      As Byte
    Dim baExpanded()    As Byte
    Dim lPos            As Long
    Dim baEmpty()       As Byte
    
    With uCtx
        If pvArraySize(.RemoteExchRandom) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_REMOTE_RANDOM
        End If
        Debug.Assert pvArraySize(.LocalExchRandom) = TLS_HELLO_RANDOM_SIZE
        Debug.Assert pvArraySize(.RemoteExchRandom) = TLS_HELLO_RANDOM_SIZE
        pvTlsSharedSecret baPreMasterSecret, .ExchAlgo, .LocalExchPrivate, .RemoteExchPublic
        pvArrayAllocate baRandom, pvArraySize(.LocalExchRandom) + pvArraySize(.RemoteExchRandom), FUNC_NAME & ".baRandom"
        lPos = pvWriteArray(baRandom, 0, .LocalExchRandom)
        lPos = pvWriteArray(baRandom, lPos, .RemoteExchRandom)
        pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "master secret", baRandom, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2   '--- always 48
        lPos = pvWriteArray(baRandom, 0, .RemoteExchRandom)
        lPos = pvWriteArray(baRandom, lPos, .LocalExchRandom)
        pvTlsKdfLegacyPrf baExpanded, .DigestAlgo, .MasterSecret, "key expansion", baRandom, 2 * (.MacSize + .KeySize + .IvSize)
        lPos = pvReadArray(baExpanded, 0, baEmpty, .MacSize) '--- LocalMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, baEmpty, .MacSize) '--- RemoteMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .RemoteTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficIV, .IvSize - .IvDynamicSize)
        pvTlsArrayRandom baRandom, .IvDynamicSize
        pvWriteArray .LocalTrafficIV, .IvSize - .IvDynamicSize, baRandom
        lPos = pvReadArray(baExpanded, lPos, .RemoteTrafficIV, .IvSize - .IvDynamicSize)
        pvTlsArrayRandom baRandom, .IvDynamicSize
        pvWriteArray .RemoteTrafficIV, .IvSize - .IvDynamicSize, baRandom
    End With
End Sub

Private Sub pvTlsKdfLegacyPrf(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSecret() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsKdfLegacyPrf"
    Dim baSeed()        As Byte
    Dim lRetValPos      As Long
    Dim baInput()       As Byte
    Dim lInputPos       As Long
    Dim baLast()        As Byte
    Dim baHmac()        As Byte
    Dim baTemp()        As Byte
    
    lInputPos = pvWriteString(baSeed, 0, sLabel)
    lInputPos = pvWriteArray(baSeed, lInputPos, baContext)
    baLast = baSeed
    Do While lRetValPos < lSize
        baTemp = baLast
        pvTlsArrayHmac baLast, eHash, baSecret, baTemp, 0
        lInputPos = pvWriteArray(baInput, 0, baLast)
        lInputPos = pvWriteArray(baInput, lInputPos, baSeed)
        pvTlsArrayHmac baHmac, eHash, baSecret, baInput, 0, Size:=lInputPos
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baHmac)
    Loop
    If lRetValPos <> lSize Then
        pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    End If
End Sub

'= crypto wrappers =======================================================

Private Sub pvTlsArrayRandom(baRetVal() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsArrayRandom"
    
    If lSize > 0 Then
        pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        CryptoRandomBytes VarPtr(baRetVal(0)), lSize
    Else
        baRetVal = vbNullString
    End If
End Sub

Private Sub pvTlsArrayHash(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1)
    Const FUNC_NAME     As String = "pvTlsArrayHash"
    
    Select Case eHash
    Case 0
        pvReadArray baInput, lPos, baRetVal, Size
    Case ucsTlsAlgoDigestSha256
        If Not CryptoHashSha256(baRetVal, baInput, lPos, Size) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
        End If
    Case ucsTlsAlgoDigestSha384
        If Not CryptoHashSha384(baRetVal, baInput, lPos, Size) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
        End If
    Case ucsTlsAlgoDigestSha512
        If Not CryptoHashSha512(baRetVal, baInput, lPos, Size) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha512")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported hash type " & eHash
    End Select
End Sub

Private Sub pvTlsArrayHmac(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1)
    Const FUNC_NAME     As String = "pvTlsArrayHmac"
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        If Not CryptoHmacSha256(baRetVal, baKey, baInput, lPos, Size) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmacSha256")
        End If
    Case ucsTlsAlgoDigestSha384
        If Not CryptoHmacSha384(baRetVal, baKey, baInput, lPos, Size) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmacSha384")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported hash type " & eHash
    End Select
End Sub

Private Function pvTlsAeadDecrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baRemoteIV() As Byte, baRemoteKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvTlsAeadDecrypt"
    
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Decrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmDecrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported AEAD type " & eAead
    End Select
    '--- success
    pvTlsAeadDecrypt = True
QH:
End Function

Private Sub pvTlsAeadEncrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baLocalIV() As Byte, baLocalKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsAeadEncrypt"
    
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Encrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoAeadChacha20Poly1305Encrypt")
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmEncrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoAeadChacha20Poly1305Encrypt")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported AEAD type " & eAead
    End Select
End Sub

Private Sub pvTlsSharedSecret(baRetVal() As Byte, ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte)
    Const FUNC_NAME     As String = "pvTlsSharedSecret"
    
    Select Case eKeyX
    Case ucsTlsAlgoExchX25519
        If Not CryptoEccCurve25519SharedSecret(baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEccCurve25519SharedSecret")
        End If
    Case ucsTlsAlgoExchSecp256r1
        If Not CryptoEccSecp256r1SharedSecret(baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEccSecp256r1SharedSecret")
        End If
    Case ucsTlsAlgoExchSecp384r1
        If Not CryptoEccSecp384r1SharedSecret(baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEccSecp384r1SharedSecret")
        End If
    Case ucsTlsAlgoExchCertificate
        baRetVal = baPriv
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported exchange curve " & eKeyX
    End Select
End Sub

Private Function pvTlsCipherSuiteName(ByVal lCipherSuite As Long) As String
    Select Case lCipherSuite
    Case TLS_CS_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "TLS_AES_128_GCM_SHA256"
    Case TLS_CS_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "TLS_AES_256_GCM_SHA384"
    Case TLS_CS_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "TLS_CHACHA20_POLY1305_SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "ECDHE-ECDSA-AES128-GCM-SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "ECDHE-ECDSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "ECDHE-RSA-AES128-GCM-SHA256"
    Case TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "ECDHE-RSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "ECDHE-RSA-CHACHA20-POLY1305"
    Case TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "ECDHE-ECDSA-CHACHA20-POLY1305"
    Case TLS_CS_RSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "AES128-GCM-SHA256"
    Case TLS_CS_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "AES256-GCM-SHA384"
    Case Else
        pvTlsCipherSuiteName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lCipherSuite))
    End Select
End Function

Private Function pvTlsCipherSuiteUseRsaCertificate(ByVal lCipherSuite As Long) As Boolean
    Select Case lCipherSuite
    Case TLS_CS_RSA_WITH_AES_128_GCM_SHA256, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteUseRsaCertificate = True
    End Select
End Function

Private Function pvTlsSignatureTypeName(ByVal lSignatureType As Long) As String
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA1"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA256"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA384
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA384"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA512
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA512"
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        pvTlsSignatureTypeName = "ECDSA_SECP256R1_SHA256"
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        pvTlsSignatureTypeName = "ECDSA_SECP384R1_SHA384"
    Case TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        pvTlsSignatureTypeName = "ECDSA_SECP521R1_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA512"
    Case Else
        pvTlsSignatureTypeName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lSignatureType))
    End Select
End Function

Private Sub pvTlsSignatureSign(baPrivKey() As Byte, ByVal lSignatureType As Long, baVerifyData() As Byte, baSignature() As Byte)
    Const FUNC_NAME     As String = "pvTlsSignatureSign"
    Dim uKeyInfo        As UcsKeyInfo
    Dim lHashSize       As Long
    Dim baEnc()         As Byte
    Dim baVerifyHash()  As Byte
    Dim baTemp()        As Byte
    
    Debug.Print "Signing with " & pvTlsSignatureTypeName(lSignatureType) & " signature", Timer
    If Not pvCryptoAsn1DecodePrivateKey(baPrivKey, uKeyInfo) Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_UNSUPPORTED_PRIVATE_KEY
    End If
    lHashSize = pvTlsSignatureHashSize(lSignatureType)
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        Debug.Assert uKeyInfo.AlgoObjId = szOID_RSA_RSA
        If Not pvCryptoEmsaPkcs1Encode(baEnc, baVerifyData, uKeyInfo.BitLen, lHashSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "pvCryptoEmsaPkcs1Encode")
        End If
        If Not CryptoRsaModExp(baEnc, uKeyInfo.PrivExp, uKeyInfo.Modulus, baSignature) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoRsaModExp")
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        Debug.Assert uKeyInfo.AlgoObjId = szOID_RSA_RSA Or uKeyInfo.AlgoObjId = szOID_RSA_SSA_PSS
        If Not pvCryptoEmsaPssEncode(baEnc, baVerifyData, uKeyInfo.BitLen, lHashSize, lHashSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "pvCryptoEmsaPssEncode")
        End If
        If Not CryptoRsaModExp(baEnc, uKeyInfo.PrivExp, uKeyInfo.Modulus, baSignature) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoRsaModExp")
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        Debug.Assert uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256
        If Not CryptoHashSha256(baVerifyHash, baVerifyData, 0) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
        End If
        If Not CryptoEccSecp256r1Sign(baTemp, uKeyInfo.KeyBlob, baVerifyHash) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEccSecp256r1Sign")
        End If
        If Not pvAsn1EncodeEccSignature(baSignature, baTemp, TLS_SECP256R1_KEY_SIZE) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "pvAsn1EncodeEccSignature")
        End If
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        Debug.Assert uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384
        If Not CryptoHashSha384(baVerifyHash, baVerifyData, 0) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
        End If
        If Not CryptoEccSecp384r1Sign(baTemp, uKeyInfo.KeyBlob, baVerifyHash) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEccSecp384r1Sign")
        End If
        If Not pvAsn1EncodeEccSignature(baSignature, baTemp, TLS_SECP384R1_KEY_SIZE) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "pvAsn1EncodeEccSignature")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_SIGNATURE_TYPE, "%1", "0x" & Hex$(lSignatureType))
    End Select
    If pvArraySize(baSignature) = 0 Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_SIGNATURE_FAILED, "%1", pvTlsSignatureTypeName(lSignatureType))
    End If
End Sub

Private Function pvTlsSignatureVerify(baCert() As Byte, ByVal lSignatureType As Long, baVerifyData() As Byte, baSignature() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsSignatureVerify"
    Dim uCertInfo       As UcsKeyInfo
    Dim lHashSize       As Long
    Dim baVerifyHash()  As Byte
    Dim baPlainSig()    As Byte
    Dim lCurveSize      As Long
    Dim bSkip           As Boolean
    Dim baTemp()        As Byte
    Dim bDeprecated     As Boolean
    Dim baDecr()        As Byte
    
    On Error GoTo EH
    If Not pvCryptoAsn1DecodeCertificate(baCert, uCertInfo) Then
        sError = ERR_UNSUPPORTED_CERTIFICATE
        eAlertCode = uscTlsAlertHandshakeFailure
        GoTo QH
    End If
    lHashSize = pvTlsSignatureHashSize(lSignatureType)
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If Not CryptoRsaModExp(baSignature, uCertInfo.PubExp, uCertInfo.Modulus, baDecr) Then
InvalidSignature:
            sError = ERR_INVALID_SIGNATURE
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End If
        If Not pvCryptoEmsaPkcs1Decode(baVerifyData, baDecr, lHashSize) Then
            GoTo InvalidSignature
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If Not CryptoRsaModExp(baSignature, uCertInfo.PubExp, uCertInfo.Modulus, baDecr) Then
            GoTo InvalidSignature
        End If
        If Not pvCryptoEmsaPssDecode(baVerifyData, baDecr, uCertInfo.BitLen, lHashSize, lHashSize) Then
            GoTo InvalidSignature
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        If uCertInfo.AlgoObjId <> szOID_ECC_PUBLIC_KEY Then
            sError = Replace(ERR_UNSUPPORTED_PUBLIC_KEY, "%1", uCertInfo.AlgoObjId)
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End If
        pvTlsArrayHash baVerifyHash, pvTlsSignatureDigestAlgo(lSignatureType), baVerifyData, 0
        lCurveSize = UBound(uCertInfo.KeyBlob) \ 2
        If Not pvAsn1DecodeEccSignature(baPlainSig, baSignature, lCurveSize) Then
            GoTo InvalidSignature
        End If
        If UBound(baVerifyHash) + 1 < lCurveSize Then
            '--- note: when hash size is less than curve size must left-pad w/ zeros (right-align hash) -> deprecated
            '---       incl. ECDSA_SECP384R1_SHA256 only
            baTemp = baVerifyHash
            pvArrayAllocate baVerifyHash, lCurveSize, FUNC_NAME & ".baRetVal"
            Call CopyMemory(baVerifyHash(lCurveSize - UBound(baTemp) - 1), baTemp(0), UBound(baTemp) + 1)
            bDeprecated = True
        ElseIf UBound(baVerifyHash) + 1 > lCurveSize Then
            '--- note: when hash size is above curve size the excess is ignored -> deprecated
            '---       incl. ECDSA_SECP256R1_SHA384, ECDSA_SECP256R1_SHA512 and ECDSA_SECP384R1_SHA512
            bDeprecated = True
        End If
        Select Case lCurveSize
        Case TLS_SECP256R1_KEY_SIZE
            If Not CryptoEccSecp256r1Verify(uCertInfo.KeyBlob, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case TLS_SECP384R1_KEY_SIZE
            If Not CryptoEccSecp384r1Verify(uCertInfo.KeyBlob, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case Else
            sError = Replace(ERR_UNSUPPORTED_CURVE_SIZE, "%1", lCurveSize)
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End Select
    Case Else
        sError = Replace(ERR_UNSUPPORTED_SIGNATURE_TYPE, "%1", "0x" & Hex$(lSignatureType))
        eAlertCode = uscTlsAlertInternalError
        GoTo QH
    End Select
    '--- success
    pvTlsSignatureVerify = True
QH:
    Debug.Print IIf(pvTlsSignatureVerify, IIf(bSkip, "Skipping ", IIf(bDeprecated, "Deprecated ", "Valid ")), "Invalid ") & pvTlsSignatureTypeName(lSignatureType) & " signature" & IIf(bDeprecated, " (lCurveSize=" & lCurveSize & " from server's public key)", vbNullString), Timer
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsSignatureDigestAlgo(ByVal lSignatureType As Long) As UcsTlsCryptoAlgorithmsEnum
    Select Case lSignatureType And &HFF
    Case 1, 2, 3 '--- 1 - RSA, 2 - DSA, 3 - ECDSA
        Select Case lSignatureType \ &H100
        '--- Skipping: 1 - MD-5, 2 - SHA-1, 3 - SHA-224
        Case 4
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case 5
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case 6
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    Case Else
        '--- 8 - Intrinsic for TLS 1.3
        Select Case lSignatureType
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    End Select
End Function

Private Function pvTlsSignatureHashSize(ByVal lSignatureType As Long) As Long
    Select Case pvTlsSignatureDigestAlgo(lSignatureType)
    Case ucsTlsAlgoDigestSha256
        pvTlsSignatureHashSize = 32
    Case ucsTlsAlgoDigestSha384
        pvTlsSignatureHashSize = 48
    Case ucsTlsAlgoDigestSha512
        pvTlsSignatureHashSize = 64
    End Select
End Function

Private Function pvAsn1DecodeEccSignature(baRetVal() As Byte, baDerSig() As Byte, ByVal lCurveSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvAsn1DecodeEccSignature"
    Dim lType           As Long
    Dim lPos            As Long
    Dim lSize           As Long
    Dim cStack          As Collection
    Dim baTemp()        As Byte
    
    pvArrayAllocate baRetVal, 64, FUNC_NAME & ".baRetVal"
    '--- ECDSA-Sig-Value ::= SEQUENCE { r INTEGER, s INTEGER }
    lPos = pvReadLong(baDerSig, 0, lType)
    If lType <> LNG_ANS1_TYPE_SEQUENCE Then
        GoTo QH
    End If
    lPos = pvReadBeginOfBlock(baDerSig, lPos, cStack)
        lPos = pvReadLong(baDerSig, lPos, lType)
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        lPos = pvReadLong(baDerSig, lPos, lSize)
        lPos = pvReadArray(baDerSig, lPos, baTemp, lSize)
        If lSize <= lCurveSize Then
            pvWriteArray baRetVal, lCurveSize - lSize, baTemp
        Else
            pvWriteBuffer baRetVal, 0, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
        lPos = pvReadLong(baDerSig, lPos, lType)
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        lPos = pvReadLong(baDerSig, lPos, lSize)
        lPos = pvReadArray(baDerSig, lPos, baTemp, lSize)
        If lSize <= lCurveSize Then
            pvWriteArray baRetVal, lCurveSize + lCurveSize - lSize, baTemp
        Else
            pvWriteBuffer baRetVal, lCurveSize, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
    lPos = pvReadEndOfBlock(baDerSig, lPos, cStack)
    '--- success
    pvAsn1DecodeEccSignature = True
QH:
End Function

Private Function pvAsn1EncodeEccSignature(baRetVal() As Byte, baPlainSig() As Byte, ByVal lPartSize As Long) As Boolean
    Dim lPos            As Long
    Dim cStack          As Collection
    Dim lStart          As Long
    
    lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_SEQUENCE)
    lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
        lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_INTEGER)
        lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lStart) And &H80) <> 0 Then
                lPos = pvWriteLong(baRetVal, lPos, 0)
            End If
            lPos = pvWriteBuffer(baRetVal, lPos, VarPtr(baPlainSig(lStart)), lPartSize - lStart)
        lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
        lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_INTEGER)
        lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lPartSize + lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lPartSize + lStart) And &H80) <> 0 Then
                lPos = pvWriteLong(baRetVal, lPos, 0)
            End If
            lPos = pvWriteBuffer(baRetVal, lPos, VarPtr(baPlainSig(lPartSize + lStart)), lPartSize - lStart)
        lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
    lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
    pvAsn1EncodeEccSignature = baRetVal
    '--- success
    pvAsn1EncodeEccSignature = True
End Function

'= buffer management =====================================================

Private Function pvWriteBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection, Optional ByVal Size As Long = 1) As Long
    If cStack Is Nothing Then
        Set cStack = New Collection
    End If
    If cStack.Count = 0 Then
        cStack.Add lPos
    Else
        cStack.Add lPos, Before:=1
    End If
    pvWriteBeginOfBlock = pvWriteReserved(baBuffer, lPos, Size)
    '--- note: keep Size in baBuffer
    baBuffer(lPos) = (Size And &HFF)
End Function

Private Function pvWriteEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection) As Long
    Dim lStart          As Long
    
    lStart = cStack.Item(1)
    cStack.Remove 1
    pvWriteLong baBuffer, lStart, lPos - lStart - baBuffer(lStart), Size:=baBuffer(lStart)
    pvWriteEndOfBlock = lPos
End Function

Private Function pvWriteString(baBuffer() As Byte, ByVal lPos As Long, sValue As String) As Long
    pvWriteString = pvWriteArray(baBuffer, lPos, StrConv(sValue, vbFromUnicode))
End Function

Private Function pvWriteArray(baBuffer() As Byte, ByVal lPos As Long, baSrc() As Byte) As Long
    Dim lSize       As Long
    
    lSize = pvArraySize(baSrc)
    If lSize > 0 Then
        lPos = pvWriteBuffer(baBuffer, lPos, VarPtr(baSrc(0)), lSize)
    End If
    pvWriteArray = lPos
End Function

Private Function pvWriteLong(baBuffer() As Byte, ByVal lPos As Long, ByVal lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte

    If Size <= 1 Then
        pvWriteLong = pvWriteBuffer(baBuffer, lPos, VarPtr(lValue), Size)
    Else
        pvWriteLong = pvWriteReserved(baBuffer, lPos, Size)
        Call CopyMemory(baTemp(0), lValue, 4)
        baBuffer(lPos) = baTemp(Size - 1)
        baBuffer(lPos + 1) = baTemp(Size - 2)
        If Size >= 3 Then baBuffer(lPos + 2) = baTemp(Size - 3)
        If Size >= 4 Then baBuffer(lPos + 3) = baTemp(Size - 4)
    End If
End Function

Private Function pvWriteReserved(baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    pvWriteReserved = pvWriteBuffer(baBuffer, lPos, 0, lSize)
End Function

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

Private Function pvReadBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection, Optional ByVal Size As Long = 1, Optional BlockSize As Long) As Long
    If cStack Is Nothing Then
        Set cStack = New Collection
    End If
    pvReadBeginOfBlock = pvReadLong(baBuffer, lPos, BlockSize, Size)
    If cStack.Count = 0 Then
        cStack.Add pvReadBeginOfBlock + BlockSize
    Else
        cStack.Add pvReadBeginOfBlock + BlockSize, Before:=1
    End If
End Function

Private Function pvReadEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection) As Long
    Dim lEnd          As Long
    
    #If baBuffer Then '--- touch args
    #End If
    lEnd = cStack.Item(1)
    cStack.Remove 1
    Debug.Assert lPos = lEnd
    pvReadEndOfBlock = lEnd
End Function

Private Function pvReadLong(baBuffer() As Byte, ByVal lPos As Long, lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte
    
    If lPos + Size <= pvArraySize(baBuffer) Then
        If Size <= 1 Then
            lValue = baBuffer(lPos)
        Else
            baTemp(Size - 1) = baBuffer(lPos + 0)
            baTemp(Size - 2) = baBuffer(lPos + 1)
            If Size >= 3 Then baTemp(Size - 3) = baBuffer(lPos + 2)
            If Size >= 4 Then baTemp(Size - 4) = baBuffer(lPos + 3)
            Call CopyMemory(lValue, baTemp(0), Size)
        End If
    Else
        lValue = 0
    End If
    pvReadLong = lPos + Size
End Function

Private Function pvReadArray(baBuffer() As Byte, ByVal lPos As Long, baDest() As Byte, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvReadArray"
    
    If lSize < 0 Then
        lSize = pvArraySize(baBuffer) - lPos
    End If
    If lSize > 0 Then
        pvArrayAllocate baDest, lSize, FUNC_NAME & ".baDest"
        If lPos + lSize <= pvArraySize(baBuffer) Then
            Call CopyMemory(baDest(0), baBuffer(lPos), lSize)
        ElseIf lPos < pvArraySize(baBuffer) Then
            Call CopyMemory(baDest(0), baBuffer(lPos), pvArraySize(baBuffer) - lPos)
        End If
    Else
        Erase baDest
    End If
    pvReadArray = lPos + lSize
End Function

'= arrays helpers ========================================================

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    Debug.Assert RedimStats(sFuncName, UBound(baRetVal) + 1)
End Sub

Private Sub pvArrayReallocate(baArray() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim Preserve baArray(0 To lSize - 1) As Byte
    Else
        baArray = vbNullString
    End If
    Debug.Assert RedimStats(sFuncName, UBound(baArray) + 1)
End Sub

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

Private Sub pvArrayXor(baRetVal() As Byte, baArray() As Byte, ByVal lSeqNo As Long)
    Dim lIdx            As Long
    
    baRetVal = baArray
    lIdx = pvArraySize(baRetVal)
    Do While lSeqNo <> 0 And lIdx > 0
        lIdx = lIdx - 1
        baRetVal(lIdx) = baRetVal(lIdx) Xor (lSeqNo And &HFF)
        lSeqNo = (lSeqNo And -&H100&) \ &H100&
    Loop
End Sub

Private Sub pvArraySwap(baBuffer() As Byte, lBufferPos As Long, baInput() As Byte, lInputPos As Long)
    Dim lTemp           As Long
    
    Call CopyMemory(lTemp, ByVal ArrPtr(baBuffer), 4)
    Call CopyMemory(ByVal ArrPtr(baBuffer), ByVal ArrPtr(baInput), 4)
    Call CopyMemory(ByVal ArrPtr(baInput), lTemp, 4)
    lTemp = lBufferPos
    lBufferPos = lInputPos
    lInputPos = lTemp
End Sub

Private Sub pvArrayByte(baRetVal() As Byte, ParamArray A() As Variant)
    Const FUNC_NAME     As String = "pvArrayByte"
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    If UBound(A) >= 0 Then
        pvArrayAllocate baRetVal, UBound(A) + 1, FUNC_NAME & ".baRetVal"
        For Each vElem In A
            baRetVal(lIdx) = vElem And &HFF
            lIdx = lIdx + 1
        Next
    End If
End Sub

Private Sub pvArrayHash(ByVal lHashSize As Long, baInput() As Byte, baRetVal() As Byte)
    Const FUNC_NAME     As String = "pvArrayHash"
    
    Select Case lHashSize
    Case 32
        If Not CryptoHashSha256(baRetVal, baInput, 0) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
        End If
    Case 48
        If Not CryptoHashSha384(baRetVal, baInput, 0) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
        End If
    Case 64
        If Not CryptoHashSha512(baRetVal, baInput, 0) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha512")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_INVALID_HASH_SIZE, "%1", lHashSize)
    End Select
End Sub

Private Sub pvArrayIncCounter(baInput() As Byte, ByVal lPos As Long)
    Do While lPos >= 0
        If baInput(lPos) < 255 Then
            baInput(lPos) = baInput(lPos) + 1
            Exit Do
        Else
            baInput(lPos) = 0
            lPos = lPos - 1
        End If
    Loop
End Sub

Private Sub pvArrayReverse(baData() As Byte, Optional ByVal NewSize As Long = -1)
    Const FUNC_NAME     As String = "pvArrayReverse"
    Dim lIdx            As Long
    Dim bTemp           As Byte
    Dim baCopy()        As Byte
    
    If NewSize = 0 Then
        baData = vbNullString
    ElseIf NewSize > 0 Then
        baCopy = baData
        pvArrayAllocate baData, NewSize, FUNC_NAME & ".baData"
        Call CopyMemory(baData(0), baCopy(0), IIf(NewSize < UBound(baCopy) + 1, NewSize, UBound(baCopy) + 1))
    End If
    For lIdx = 0 To UBound(baData) \ 2
        bTemp = baData(lIdx)
        baData(lIdx) = baData(UBound(baData) - lIdx)
        baData(UBound(baData) - lIdx) = bTemp
    Next
End Sub

'=========================================================================
' Crypto
'=========================================================================

Private Function pvCryptoIsSupported(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As Boolean
    Const PREF          As Long = &H1000
    
    Select Case eAlgo
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        #If ImplUseLibSodium Then
            pvCryptoIsSupported = (crypto_aead_aes256gcm_is_available() <> 0 And eAlgo = ucsTlsAlgoAeadAes256)
        #Else
            pvCryptoIsSupported = True
        #End If
    Case PREF + ucsTlsAlgoAeadAes128, PREF + ucsTlsAlgoAeadAes256
        '--- signal if AES preferred over Chacha20
        #If ImplUseLibSodium Then
            pvCryptoIsSupported = (crypto_aead_aes256gcm_is_available() <> 0 And eAlgo = PREF + ucsTlsAlgoAeadAes256)
        #End If
    Case ucsTlsAlgoSignaturePkcsSha1
        pvCryptoIsSupported = False
    Case Else
        pvCryptoIsSupported = True
    End Select
End Function

Private Function pvCryptoEmePkcs1Encode(baRetVal() As Byte, baMessage() As Byte, ByVal lBitLen As Long) As Boolean
    Const FUNC_NAME     As String = "pvCryptoEmePkcs1Encode"
    Dim lIdx            As Long
    
    '--- from RFC 8017, Section  7.2.1
    pvArrayAllocate baRetVal, (lBitLen + 7) \ 8, FUNC_NAME & ".baRetVal"
    If UBound(baMessage) > UBound(baRetVal) - 11 Then
        GoTo QH
    End If
    baRetVal(1) = 2
    CryptoRandomBytes VarPtr(baRetVal(2)), UBound(baRetVal) - UBound(baMessage) - 3
    For lIdx = 2 To UBound(baRetVal) - UBound(baMessage) - 3
        If baRetVal(lIdx) = 0 Then
            baRetVal(lIdx) = 42
        End If
    Next
    Call CopyMemory(baRetVal(UBound(baRetVal) - UBound(baMessage)), baMessage(0), UBound(baMessage) + 1)
    '--- success
    pvCryptoEmePkcs1Encode = True
QH:
End Function

Private Function pvCryptoEmsaPkcs1Encode(baRetVal() As Byte, baMessage() As Byte, ByVal lBitLen As Long, ByVal lHashSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvCryptoEmsaPkcs1Encode"
    Dim baHash()        As Byte
    Dim baDerHash()     As Byte
    Dim lPos            As Long
    
    '--- from RFC 8017, Section 9.2.
    pvArrayHash lHashSize, baMessage, baHash
    If Not pvCryptoAsn1EncodePkcs1SignatureHash(baHash, baDerHash) Then
        GoTo QH
    End If
    pvArrayAllocate baRetVal, (lBitLen + 7) \ 8, FUNC_NAME & ".baRetVal"
    baRetVal(1) = 1
    For lPos = 2 To UBound(baRetVal) - UBound(baDerHash) - 2
        baRetVal(lPos) = &HFF
    Next
    lPos = lPos + 1
    Debug.Assert UBound(baRetVal) - lPos >= UBound(baDerHash)
    Call CopyMemory(baRetVal(lPos), baDerHash(0), UBound(baDerHash) + 1)
    '--- success
    pvCryptoEmsaPkcs1Encode = True
QH:
End Function

Private Function pvCryptoEmsaPkcs1Decode(baMessage() As Byte, baEnc() As Byte, ByVal lHashSize As Long) As Boolean
    Dim baHash()        As Byte
    Dim baDerHash()     As Byte
    Dim lIdx            As Long
    Dim lPos            As Long
    
    If baEnc(0) <> &H0 Or baEnc(1) <> &H1 Then
        GoTo QH
    End If
    pvArrayHash lHashSize, baMessage, baHash
    If Not pvCryptoAsn1EncodePkcs1SignatureHash(baHash, baDerHash) Then
        GoTo QH
    End If
    For lPos = 2 To UBound(baEnc) - UBound(baDerHash) - 3
        If baEnc(lPos) <> &HFF Then
            GoTo QH
        End If
    Next
    If baEnc(lPos) <> &H0 Then
        GoTo QH
    End If
    lPos = lPos + 1
    For lIdx = 0 To UBound(baDerHash)
        If baEnc(lPos) <> baDerHash(lIdx) Then
            GoTo QH
        End If
        lPos = lPos + 1
    Next
    '--- success
    pvCryptoEmsaPkcs1Decode = True
QH:
End Function

Private Function pvCryptoEmsaPssEncode(baRetVal() As Byte, baMessage() As Byte, ByVal lBitLen As Long, ByVal lHashSize As Long, ByVal lSaltSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvCryptoEmsaPssEncode"
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim baHash()        As Byte
    Dim baSalt()        As Byte
    Dim baSeed()        As Byte
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim bMask           As Byte
    
    '--- from RFC 8017, Section 9.1.1.
    lSize = (lBitLen + 7) \ 8
    '--- 2. Let |mHash| = |Hash(M)|, an octet string of length hLen.
    pvArrayHash lHashSize, baMessage, baHash
    '--- 3. If |emLen| < |hLen + sLen + 2|, output "encoding error" and stop.
    If lSize < lHashSize + lSaltSize + 2 Then
        GoTo QH
    End If
    '--- 4. Generate a random octet string salt of length sLen; if |sLen| = 0, then salt is the empty string.
    If lSaltSize > 0 Then
        pvArrayAllocate baSalt, lSaltSize, FUNC_NAME & ".baSalt"
        CryptoRandomBytes VarPtr(baSalt(0)), lSaltSize
    Else
        baSalt = vbNullString
    End If
    '--- 5. Let |M'| = (0x)00 00 00 00 00 00 00 00 || mHash || salt;
    pvArrayAllocate baBuffer, 8 + lHashSize + lSaltSize, FUNC_NAME & ".baBuffer"
    Call CopyMemory(baBuffer(8), baHash(0), lHashSize)
    Call CopyMemory(baBuffer(8 + lHashSize), baSalt(0), lSaltSize)
    '--- 6. Let |H| = Hash(M'), an octet string of length hLen.
    pvArrayHash lHashSize, baBuffer, baHash
    '--- 7. Generate an octet string |PS| consisting of |emLen - sLen - hLen - 2| zero octets. The length of PS may be 0.
    '--- 8. Let |DB| = PS || 0x01 || salt; DB is an octet string of length |emLen - hLen - 1|.
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    baRetVal(lSize - lHashSize - lSaltSize - 2) = &H1
    Call CopyMemory(baRetVal(lSize - lHashSize - lSaltSize - 1), baSalt(0), lSaltSize)
    Call CopyMemory(baRetVal(lSize - lHashSize - 1), baHash(0), lHashSize)
    '--- 9. Let |dbMask| = MGF(H, emLen - hLen - 1).
    '--- 10. Let |maskedDB| = DB \xor dbMask.
    pvArrayAllocate baSeed, lHashSize + 4, FUNC_NAME & ".baSeed" '--- leave 4 more bytes at the end for counter
    Call CopyMemory(baSeed(0), baRetVal(lSize - lHashSize - 1), lHashSize)
    Do
        pvArrayHash lHashSize, baSeed, baHash
        For lIdx = 0 To UBound(baHash)
            baRetVal(lPos) = baRetVal(lPos) Xor baHash(lIdx)
            lPos = lPos + 1
            If lPos >= lSize - lHashSize - 1 Then
                Exit Do
            End If
        Next
        pvArrayIncCounter baSeed, lHashSize + 3
    Loop
    '--- 11. Set the leftmost |8 * emLen - emBits| bits of the leftmost octet in |maskedDB| to zero.
    bMask = &HFF \ (2 ^ (lSize * 8 - lBitLen))
    baRetVal(0) = baRetVal(0) And (bMask \ 2)
    '--- 12. Let |EM| = maskedDB || H || 0xbc.
    baRetVal(lSize - 1) = &HBC
    '--- 13. Output EM.
    pvCryptoEmsaPssEncode = True
QH:
End Function

Private Function pvCryptoEmsaPssDecode(baMessage() As Byte, baEnc() As Byte, ByVal lBitLen As Long, ByVal lHashSize As Long, ByVal lSaltSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvCryptoEmsaPssDecode"
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim baSeed()        As Byte
    Dim baHash()        As Byte
    Dim baSalt()        As Byte
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim bMask           As Byte
    
    '--- from RFC 8017, Section 9.1.2.
    lSize = (lBitLen + 7) \ 8
    '--- 3. If |emLen| < |hLen + sLen + 2|, output "inconsistent" and stop.
    If lSize < lHashSize + lSaltSize + 2 Then
        GoTo QH
    End If
    '--- 4. If the rightmost octet of |EM| does not have hexadecimal value 0xbc, output "inconsistent" and stop.
    If baEnc(lSize - 1) <> &HBC Then
        GoTo QH
    End If
    '--- 5. Let |maskedDB| be the leftmost |emLen - hLen - 1| octets of |EM|, and let |H| be the next |hLen| octets.
    '--- 6. If the leftmost |8 * emLen - emBits| bits of the leftmost octet in |maskedDB| are not all equal to zero,
    '---    output "inconsistent" and stop.
    bMask = &HFF \ (2 ^ (lSize * 8 - lBitLen))
    If (baEnc(0) And Not bMask) <> 0 Then
        GoTo QH
    End If
    '--- 7. Let |dbMask| = MGF(H, emLen - hLen - 1).
    '--- 8. Let |DB| = maskedDB \xor dbMask.
    pvArrayAllocate baSeed, lHashSize + 4, FUNC_NAME & ".baSeed"  '--- leave 4 more bytes at the end for counter
    Call CopyMemory(baSeed(0), baEnc(lSize - lHashSize - 1), lHashSize)
    Do
        pvArrayHash lHashSize, baSeed, baHash
        For lIdx = 0 To UBound(baHash)
            baEnc(lPos) = baEnc(lPos) Xor baHash(lIdx)
            lPos = lPos + 1
            If lPos >= lSize - lHashSize - 1 Then
                Exit Do
            End If
        Next
        pvArrayIncCounter baSeed, lHashSize + 3
    Loop
    '--- 9. Set the leftmost |8 * emLen - emBits| bits of the leftmost octet in |DB| to zero.
    '--- note: troubles w/ sign bit so use (bMask \ 2) to clear MSB
    baEnc(0) = baEnc(0) And (bMask \ 2)
    '--- 10. If the |emLen - hLen - sLen - 2| leftmost octets of |DB| are not zero or if the octet at position
    '---     |emLen - hLen - sLen - 1| (the leftmost position is "position 1") does not have hexadecimal
    '---     value 0x01, output "inconsistent" and stop.
    For lIdx = 0 To lPos - lHashSize - 2
        If baEnc(lIdx) <> 0 Then
            Exit For
        End If
    Next
    If lIdx <> lPos - lHashSize - 1 Then
        GoTo QH
    End If
    If baEnc(lPos - lHashSize - 1) <> &H1 Then
        GoTo QH
    End If
    '--- 11. Let |salt| be the last |sLen| octets of |DB|.
    pvArrayAllocate baSalt, lSaltSize, FUNC_NAME & ".baSalt"
    Call CopyMemory(baSalt(0), baEnc(lPos - lSaltSize), lSaltSize)
    '--- 12. Let |M'| = (0x)00 00 00 00 00 00 00 00 || mHash || salt
    pvArrayAllocate baSeed, 8 + lHashSize + lSaltSize, FUNC_NAME & ".baSeed"
    pvArrayHash lHashSize, baMessage, baHash
    Call CopyMemory(baSeed(8), baHash(0), lHashSize)
    Call CopyMemory(baSeed(8 + lHashSize), baSalt(0), lSaltSize)
    '--- 13. Let |H'| = Hash(M'), an octet string of length |hLen|.
    pvArrayHash lHashSize, baSeed, baHash
    '--- |H| is still not de-masked in decrypted buffer
    pvArrayAllocate baBuffer, lHashSize, FUNC_NAME & ".baBuffer"
    Call CopyMemory(baBuffer(0), baEnc(lPos), lHashSize)
    '--- 14. If |H| = |H'|, output "consistent." Otherwise, output "inconsistent."
    If StrConv(baHash, vbUnicode) <> StrConv(baBuffer, vbUnicode) Then
        GoTo QH
    End If
    '--- success
    pvCryptoEmsaPssDecode = True
QH:
End Function

Private Function pvCryptoAsn1DecodePrivateKey(baPrivKey() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "pvCryptoAsn1DecodePrivateKey"
    Dim lPkiPtr         As Long
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim lSize           As Long
    Dim lHalfSize       As Long
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
        lSize = (uRetVal.BitLen + 7) \ 8
        lHalfSize = (uRetVal.BitLen + 15) \ 16
        pvArrayAllocate uRetVal.Modulus, lSize, FUNC_NAME & ".uRetVal.Modulus"
        Debug.Assert UBound(uRetVal.KeyBlob) - 20 >= UBound(uRetVal.Modulus)
        Call CopyMemory(uRetVal.Modulus(0), uRetVal.KeyBlob(20), UBound(uRetVal.Modulus) + 1)
        pvArrayReverse uRetVal.Modulus
        pvArrayAllocate uRetVal.PrivExp, lSize, FUNC_NAME & ".uRetVal.PrivExp"
        Debug.Assert UBound(uRetVal.KeyBlob) >= 20 + lSize + 5 * lHalfSize + UBound(uRetVal.PrivExp)
        Call CopyMemory(uRetVal.PrivExp(0), uRetVal.KeyBlob(20 + lSize + 5 * lHalfSize), UBound(uRetVal.PrivExp) + 1)
        pvArrayReverse uRetVal.PrivExp
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
    pvCryptoAsn1DecodePrivateKey = True
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

Private Function pvCryptoAsn1DecodeCertificate(baCert() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "pvCryptoAsn1DecodeCertificate"
    Dim pCertContext    As Long
    Dim lPtr            As Long
    Dim uPublicKeyInfo  As CERT_PUBLIC_KEY_INFO
    Dim hProv           As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim hResult         As Long
    Dim sApiSource      As String

    pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pCertContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)                                        '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
    Call CopyMemory(uPublicKeyInfo, ByVal lPtr, Len(uPublicKeyInfo))
    uRetVal.AlgoObjId = pvToString(uPublicKeyInfo.Algorithm.pszObjId)
    pvArrayAllocate uRetVal.KeyBlob, uPublicKeyInfo.PublicKey.cbData, FUNC_NAME & ".uRetVal.KeyBlob"
    Call CopyMemory(uRetVal.KeyBlob(0), ByVal uPublicKeyInfo.PublicKey.pbData, uPublicKeyInfo.PublicKey.cbData)
    If uRetVal.AlgoObjId = szOID_RSA_RSA Then
        If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
            GoTo QH
        End If
        If CryptImportPublicKeyInfo(hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lPtr, hKey) = 0 Then
            GoTo QH
        End If
        If CryptExportKey(hKey, 0, PUBLICKEYBLOB, 0, ByVal 0, lSize) = 0 Then
            GoTo QH
        End If
        pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
        If CryptExportKey(hKey, 0, PUBLICKEYBLOB, 0, baBuffer(0), lSize) = 0 Then
            GoTo QH
        End If
        '--- retrieve RSA key size (in bits)
        Debug.Assert UBound(baBuffer) >= 16
        Call CopyMemory(uRetVal.BitLen, baBuffer(12), 4)                                        '--- 12 = sizeof(PUBLICKEYSTRUC) + offset(RSAPUBKEY, bitlen)
        lSize = (uRetVal.BitLen + 7) \ 8
        '--- retrieve RSA public exponent
        pvArrayAllocate uRetVal.PubExp, 4, FUNC_NAME & ".uRetVal.PubExp"
        Debug.Assert UBound(baBuffer) - 16 >= UBound(uRetVal.PubExp)
        Call CopyMemory(uRetVal.PubExp(0), baBuffer(16), UBound(uRetVal.PubExp) + 1)            '--- 16 = sizeof(PUBLICKEYSTRUC) + offset(RSAPUBKEY, pubexp)
        pvArrayReverse uRetVal.PubExp, lSize
        '--- retrieve RSA key modulus
        pvArrayAllocate uRetVal.Modulus, lSize, FUNC_NAME & ".uRetVal.Modulus"
        Debug.Assert UBound(baBuffer) - 20 >= UBound(uRetVal.Modulus)
        Call CopyMemory(uRetVal.Modulus(0), baBuffer(20), UBound(uRetVal.Modulus) + 1)          '--- 20 = sizeof(PUBLICKEYSTRUC) + sizeof(RSAPUBKEY)
        pvArrayReverse uRetVal.Modulus
    End If
    '--- success
    pvCryptoAsn1DecodeCertificate = True
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvCryptoAsn1EncodePkcs1SignatureHash(baHash() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvCryptoAsn1EncodePkcs1SignatureHash"
    Dim baPrefix()      As Byte
    
    Select Case UBound(baHash) + 1
    Case 32
        pvArrayByte baPrefix, &H30, &H31, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H1, &H5, &H0, &H4, &H20
    Case 48
        pvArrayByte baPrefix, &H30, &H41, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H2, &H5, &H0, &H4, &H30
    Case 64
        pvArrayByte baPrefix, &H30, &H51, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H3, &H5, &H0, &H4, &H40
    End Select
    pvArrayAllocate baRetVal, UBound(baPrefix) + UBound(baHash) + 1, FUNC_NAME & ".baRetVal"
    Call CopyMemory(baRetVal(0), baPrefix(0), UBound(baPrefix) + 1)
    Call CopyMemory(baRetVal(UBound(baPrefix) + 1), baHash(0), UBound(baHash) + 1)
    '--- success
    pvCryptoAsn1EncodePkcs1SignatureHash = True
QH:
End Function

'=========================================================================
' PKI
'=========================================================================

Public Function PkiPemImportCertificates(ByVal vPemFiles As Variant) As Boolean
    Dim cCerts          As Collection
    Dim baPrivKey()     As Byte
    
    If pvPkiPemImportCertificates(vPemFiles, cCerts, baPrivKey) Then
        Set m_uCtx.LocalCertificates = cCerts
        m_uCtx.LocalPrivateKey = baPrivKey
        '--- success
        PkiPemImportCertificates = True
    End If
End Function

Public Function PkiPkcs12ImportCertificates(sPfxFile As String, sPassword As String) As Boolean
    Dim cCerts          As Collection
    Dim baPrivKey()     As Byte
    
    If pvPkiPkcs12ImportCertificates(sPfxFile, sPassword, cCerts, baPrivKey) Then
        Set m_uCtx.LocalCertificates = cCerts
        m_uCtx.LocalPrivateKey = baPrivKey
        '--- success
        PkiPkcs12ImportCertificates = True
    End If
End Function

Public Function PkiPemImportRootCaCertStore(Optional CaBundlePemFile As String, Optional RootFwd As cTlsSocket) As Boolean
    Dim hCertStore      As Long
    
    If LenB(CaBundlePemFile) = 0 Then
        Set m_oRootStoreFwd = RootFwd
        If m_hRootStore <> 0 Then
            Call CertCloseStore(m_hRootStore, 0)
            m_hRootStore = 0
        End If
    Else
        hCertStore = pvPkiPemImportRootCaCertStore(CaBundlePemFile)
        If hCertStore = 0 Then
            GoTo QH
        End If
        If m_hRootStore <> 0 Then
            Call CertCloseStore(m_hRootStore, 0)
        End If
        m_hRootStore = hCertStore
    End If
    '--- success
    PkiPemImportRootCaCertStore = True
QH:
End Function

'= private ===============================================================

Private Function pvPkiPemImportCertificates(ByVal vPemFiles As Variant, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Dim vElem           As Variant
    Dim sPemText        As String
    Dim cKeys           As Collection
    
    If VarType(vPemFiles) = vbString Then
        vPemFiles = Array(vPemFiles)
    End If
    For Each vElem In vPemFiles
        sPemText = StrConv(CStr(ReadBinaryFile(CStr(vElem))), vbUnicode)
        pvPkiPemGetTextPortions sPemText, "PRIVATE KEY", cKeys
        pvPkiPemGetTextPortions sPemText, "EC PRIVATE KEY", cKeys
        pvPkiPemGetTextPortions sPemText, "CERTIFICATE", cCerts
    Next
    If SearchCollection(cKeys, 1, RetVal:=baPrivKey) Then
        '--- success
        pvPkiPemImportCertificates = True
    End If
End Function

Private Function pvPkiPemImportRootCaCertStore(sCaBundlePemFile As String) As Long
    Const FUNC_NAME     As String = "pvPkiPemImportRootCaCertStore"
    Dim hCertStore      As Long
    Dim cCerts          As Collection
    Dim vElem           As Variant
    Dim baCert()        As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    Set cCerts = pvPkiPemGetTextPortions(StrConv(CStr(ReadBinaryFile(CStr(sCaBundlePemFile))), vbUnicode), "CERTIFICATE")
    If cCerts.Count = 0 Then
        GoTo QH
    End If
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For Each vElem In cCerts
        baCert = vElem
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    '--- commit
    pvPkiPemImportRootCaCertStore = hCertStore
    hCertStore = 0
QH:
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiPkcs12ImportCertificates(sPfxFile As String, sPassword As String, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvPkiPkcs12ImportCertificates"
    Dim baPfx()         As Byte
    Dim uBlob           As CRYPT_BLOB_DATA
    Dim hPfxStore       As Long
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    baPfx = ReadBinaryFile(sPfxFile)
    If UBound(baPfx) < 0 Then
        GoTo QH
    End If
    uBlob.cbData = UBound(baPfx) + 1
    uBlob.pbData = VarPtr(baPfx(0))
    hPfxStore = PFXImportCertStore(uBlob, StrPtr(sPassword), CRYPT_EXPORTABLE)
    If hPfxStore = 0 And Err.LastDllError <> NTE_BAD_ALGID Then
        hPfxStore = PFXImportCertStore(baPfx(0), 0, CRYPT_EXPORTABLE)
    End If
    If hPfxStore = 0 Then
        sApiSource = "PFXImportCertStore"
        hResult = Err.LastDllError
        GoTo QH
    End If
    Do
        pCertContext = CertEnumCertificatesInStore(hPfxStore, pCertContext)
        If pCertContext = 0 Then
            Exit Do
        End If
        If pvPkiAppendCertContext(pCertContext, cCerts, baPrivKey) Then
            '--- success
            pvPkiPkcs12ImportCertificates = True
        End If
    Loop
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiGenerSelfSignedCertificate(cCerts As Collection, baPrivKey() As Byte, Optional ByVal Subject As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiGenerSelfSignedCertificate"
    Dim hProv           As Long
    Dim hKey            As Long
    Dim sName           As String
    Dim baName()        As Byte
    Dim lSize           As Long
    Dim uName           As CRYPT_BLOB_DATA
    Dim uExpire         As SYSTEMTIME
    Dim uInfo           As CRYPT_KEY_PROV_INFO
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    Call CryptAcquireContext(0, 0, 0, PROV_RSA_FULL, CRYPT_DELETEKEYSET)
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH
    End If
    If CryptGenKey(hProv, AT_SIGNATURE, RSA1024BIT_KEY Or CRYPT_EXPORTABLE, hKey) = 0 Then
        GoTo QH
    End If
    If Left$(Subject, 3) <> "CN=" Then
        If LenB(Subject) = 0 Then
            Subject = LCase$(Environ$("COMPUTERNAME") & IIf(LenB(Environ$("USERDNSDOMAIN")) <> 0, "." & Environ$("USERDNSDOMAIN"), vbNullString))
        End If
        sName = "CN=""" & Replace(Subject, """", """""") & """" & ",OU=""" & Replace(Environ$("USERDOMAIN") & "\" & Environ$("USERNAME"), """", """""") & """,O=""VbAsyncSocket Self-Signed Certificate"""
    Else
        sName = Subject
    End If
    If CertStrToName(X509_ASN_ENCODING, StrPtr(sName), CERT_OID_NAME_STR, 0, ByVal 0, lSize, 0) = 0 Then
        GoTo QH
    End If
    pvArrayAllocate baName, lSize, FUNC_NAME & ".baName"
    If CertStrToName(X509_ASN_ENCODING, StrPtr(sName), CERT_OID_NAME_STR, 0, baName(0), lSize, 0) = 0 Then
        GoTo QH
    End If
    With uName
        .cbData = lSize
        .pbData = VarPtr(baName(0))
    End With
    Call GetSystemTime(uExpire)
    uExpire.wYear = uExpire.wYear + 1
    With uInfo
        .dwProvType = PROV_RSA_FULL
        .dwKeySpec = AT_SIGNATURE
    End With
    pCertContext = CertCreateSelfSignCertificate(hProv, uName, 0, uInfo, 0, ByVal 0, uExpire, 0)
    If pvPkiAppendCertContext(pCertContext, cCerts, baPrivKey) Then
        '--- success
        pvPkiGenerSelfSignedCertificate = True
    End If
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
        Call CryptAcquireContext(0, 0, 0, PROV_RSA_FULL, CRYPT_DELETEKEYSET)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiCertChainValidate(sRemoteHostName As String, cCerts As Collection, ByVal hRootStore As Long, sError As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiCertChainValidate"
    Dim hCertStore      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim pCertContext    As Long
    Dim pChainContext   As Long
    Dim uChain          As CERT_CHAIN_CONTEXT
    Dim lPtr            As Long
    Dim dwErrorStatus   As Long
    Dim uChainParams    As CERT_CHAIN_PARA
    Dim uInfo           As CERT_INFO
    Dim uExtension      As CERT_EXTENSION
    Dim lAltInfoPtr     As Long
    Dim uAltInfo        As CERT_ALT_NAME_INFO
    Dim uEntry          As CERT_ALT_NAME_ENTRY
    Dim sDnsName        As String
    Dim bValidName      As Boolean
    Dim uEngineConfig   As CERT_CHAIN_ENGINE_CONFIG
    Dim hChainEngine    As Long
    Dim uChainElem      As CERT_CHAIN_ELEMENT
    Dim pExistContext   As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- load server X.509 certificates to an in-memory certificate store
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For lIdx = 1 To cCerts.Count
        baCert = cCerts.Item(lIdx)
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    '--- search remote host FQDN in any X.509 certificate's "Subject Alternative Name" list of DNS names (incl. wildcards)
    Do
        pCertContext = CertEnumCertificatesInStore(hCertStore, pCertContext)
        If pCertContext = 0 Then
            sError = Replace(ERR_NO_MATCHING_ALT_NAME, "%1", sRemoteHostName)
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)               '--- dereference pCertContext->pCertInfo->cExtension
        Call CopyMemory(uInfo, ByVal lPtr, Len(uInfo))
        lPtr = CertFindExtension(szOID_SUBJECT_ALT_NAME2, uInfo.cExtension, uInfo.rgExtension)
        If lPtr <> 0 Then
            Call CopyMemory(uExtension, ByVal lPtr, Len(uExtension))
            If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, szOID_SUBJECT_ALT_NAME2, ByVal uExtension.Value.pbData, _
                        uExtension.Value.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lAltInfoPtr, 0) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptDecodeObjectEx(szOID_SUBJECT_ALT_NAME2)"
                GoTo QH
            End If
            Call CopyMemory(uAltInfo, ByVal lAltInfoPtr, Len(uAltInfo))
            For lIdx = 0 To uAltInfo.cAltEntry - 1
                lPtr = UnsignedAdd(uAltInfo.rgAltEntry, lIdx * Len(uEntry))         '--- dereference lAltInfoPtr->rgAltEntry[lidx].dwAltNameChoice
                Call CopyMemory(uEntry, ByVal lPtr, Len(uEntry))
                If uEntry.dwAltNameChoice = CERT_ALT_NAME_DNS_NAME Then
                    sDnsName = LCase$(pvToStringW(uEntry.pwszDNSName))
                    If Left$(sDnsName, 1) = "*" Then
                        If LCase$(sRemoteHostName) Like sDnsName And Not LCase$(sRemoteHostName) Like "*." & sDnsName Then
                            bValidName = True
                            Exit Do
                        End If
                    Else
                        If LCase$(sRemoteHostName) = sDnsName Then
                            bValidName = True
                            Exit Do
                        End If
                    End If
                End If
            Next
            Call LocalFree(lAltInfoPtr)
            lAltInfoPtr = 0
        End If
    Loop
    '--- build custom chain engine that trusts the additional root CA certificates if provided
    If hRootStore <> 0 Then
        uEngineConfig.cbSize = Len(uEngineConfig) - IIf(OsVersion < ucsOsvWin7, 12, 0)
        uEngineConfig.cAdditionalStore = 1
        uEngineConfig.rghAdditionalStore = VarPtr(hRootStore)
        If CertCreateCertificateChainEngine(uEngineConfig, hChainEngine) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertCreateCertificateChainEngine"
            GoTo QH
        End If
    End If
    '--- for the matched server certificate try to build a chain of certificates from the ones in the in-memory certificate store
    '---    and check this chain for revokation, expiry or missing link to a trust anchor
    uChainParams.cbSize = Len(uChainParams)
    If CertGetCertificateChain(hChainEngine, pCertContext, 0, hCertStore, uChainParams, 0, 0, pChainContext) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertGetCertificateChain"
        GoTo QH
    End If
    Call CopyMemory(uChain, ByVal pChainContext, Len(uChain))       '--- dereference pChainContext->rgpChain[0]->TrustStatus.dwErrorStatus
    Call CopyMemory(lPtr, ByVal uChain.rgElem, 4)
    Call CopyMemory(uChain, ByVal lPtr, Len(uChain))
    dwErrorStatus = uChain.TrustStatus.dwErrorStatus And Not CERT_TRUST_IS_NOT_TIME_NESTED
    If hRootStore <> 0 And uChain.cElems > 0 Then
        '--- check if the last certificate in the chain is from our custom hRootStore and remove untrusted flags from status
        Call CopyMemory(lPtr, ByVal UnsignedAdd(uChain.rgElem, (uChain.cElems - 1) * 4), 4)
        Call CopyMemory(uChainElem, ByVal lPtr, Len(uChainElem))
        pExistContext = CertFindCertificateInStore(hRootStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_EXISTING, ByVal uChainElem.pCertContext, 0)
        If pExistContext <> 0 Then
            Call CertFreeCertificateContext(pExistContext)
            pExistContext = 0
            dwErrorStatus = dwErrorStatus And Not CERT_TRUST_IS_UNTRUSTED_ROOT And Not CERT_TRUST_IS_NOT_SIGNATURE_VALID
        End If
    End If
    If dwErrorStatus <> 0 Then
        If (dwErrorStatus And CERT_TRUST_IS_REVOKED) <> 0 Then
            sError = ERR_TRUST_IS_REVOKED
        ElseIf (dwErrorStatus And CERT_TRUST_IS_PARTIAL_CHAIN) <> 0 Then
            sError = ERR_TRUST_IS_PARTIAL_CHAIN
        ElseIf (dwErrorStatus And CERT_TRUST_IS_UNTRUSTED_ROOT) <> 0 Then
            sError = ERR_TRUST_IS_UNTRUSTED_ROOT
        ElseIf (dwErrorStatus And CERT_TRUST_IS_NOT_TIME_VALID) <> 0 Then
            sError = ERR_TRUST_IS_NOT_TIME_VALID
        ElseIf (dwErrorStatus And CERT_TRUST_REVOCATION_STATUS_UNKNOWN) <> 0 Then
            sError = ERR_TRUST_REVOCATION_STATUS_UNKNOWN
        Else
            sError = "CertGetCertificateChain error mask: 0x" & Hex$(dwErrorStatus)
        End If
        GoTo QH
    End If
    '--- success
    pvPkiCertChainValidate = True
QH:
    If pChainContext <> 0 Then
        Call CertFreeCertificateChain(pChainContext)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If hChainEngine <> 0 Then
        Call CertFreeCertificateChainEngine(hChainEngine)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiPemGetTextPortions(sContents As String, sBoundary As String, Optional RetVal As Collection) As Collection
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim bInside         As Boolean
    Dim lStart          As Long
    Dim lSize           As Long
    Dim sPortion        As String
    
    If RetVal Is Nothing Then
        Set RetVal = New Collection
    End If
    vSplit = Split(Replace(sContents, vbCr, vbNullString), vbLf)
    For lIdx = 0 To UBound(vSplit)
        If Not bInside Then
            If InStr(vSplit(lIdx), "-----BEGIN " & sBoundary & "-----") > 0 Then
                lStart = lIdx + 1
                lSize = 0
                bInside = True
            End If
        Else
            If InStr(vSplit(lIdx), "-----END " & sBoundary & "-----") > 0 Then
                sPortion = String$(lSize, 0)
                lSize = 1
                For lJdx = lStart To lIdx - 1
                    If InStr(vSplit(lJdx), ":") = 0 Then
                        Mid$(sPortion, lSize, Len(vSplit(lJdx))) = vSplit(lJdx)
                        lSize = lSize + Len(vSplit(lJdx))
                    End If
                Next
                If Not SearchCollection(RetVal, sPortion) Then
                    RetVal.Add FromBase64Array(sPortion), sPortion
                End If
                bInside = False
            ElseIf InStr(vSplit(lIdx), ":") = 0 Then
                lSize = lSize + Len(vSplit(lIdx))
            End If
        End If
    Next
    Set pvPkiPemGetTextPortions = RetVal
End Function

Private Function pvPkiAppendCertContext(ByVal pCertContext As Long, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvPkiAppendCertContext"
    Dim uCertContext    As CERT_CONTEXT
    Dim baBuffer()      As Byte
    
    Call CopyMemory(uCertContext, ByVal pCertContext, Len(uCertContext))
    If uCertContext.cbCertEncoded > 0 Then
        pvArrayAllocate baBuffer, uCertContext.cbCertEncoded, FUNC_NAME & ".baBuffer"
        Call CopyMemory(baBuffer(0), ByVal uCertContext.pbCertEncoded, uCertContext.cbCertEncoded)
        If cCerts Is Nothing Then
            Set cCerts = New Collection
        End If
        cCerts.Add baBuffer
    End If
    If pvPkiExportPrivateKey(pCertContext, baPrivKey) Then
        If cCerts.Count > 1 Then
            '--- move certificate w/ private key to the beginning of the collection
            baBuffer = cCerts.Item(cCerts.Count)
            cCerts.Remove cCerts.Count
            cCerts.Add baBuffer, Before:=1
        End If
        '--- success
        pvPkiAppendCertContext = True
    End If
End Function

Private Function pvPkiExportPrivateKey(ByVal pCertContext As Long, baPrivKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportPrivateKey"
    Dim dwFlags         As Long
    Dim hProvOrKey      As Long
    Dim lKeySpec        As Long
    Dim lFree           As Long
    Dim hCngKey         As Long
    Dim hNewKey         As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim uKeyInfo        As CRYPT_KEY_PROV_INFO
    Dim hProv           As Long
    Dim hKey            As Long
    Dim lMagic          As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- note: this function allows using CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG too for key export w/ all CNG API calls
    dwFlags = CRYPT_ACQUIRE_CACHE_FLAG Or CRYPT_ACQUIRE_SILENT_FLAG Or CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG
    If CryptAcquireCertificatePrivateKey(pCertContext, dwFlags, 0, hProvOrKey, lKeySpec, lFree) = 0 Then
        GoTo QH
    End If
    If lKeySpec < 0 Then
        hCngKey = hProvOrKey: hProvOrKey = 0
        hNewKey = pvPkiCloneKeyWithExportPolicy(hCngKey, NCRYPT_ALLOW_EXPORT_FLAG Or NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG)
        hResult = NCryptExportKey(hNewKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
        If hResult < 0 Then
            sApiSource = "NCryptExportKey(PRIVATEBLOB)"
            GoTo QH
        End If
        pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
        hResult = NCryptExportKey(hNewKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
        If hResult < 0 Then
            sApiSource = "NCryptExportKey(PRIVATEBLOB)#2"
            GoTo QH
        End If
        Call CopyMemory(lMagic, baBuffer(0), 4)
        Select Case lMagic
        Case BCRYPT_RSAPRIVATE_MAGIC
            hResult = NCryptExportKey(hNewKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)"
                GoTo QH
            End If
            pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
            hResult = NCryptExportKey(hNewKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)#2"
                GoTo QH
            End If
            If Not pvPkiExportRsaPrivateKey(baPrivKey, baBuffer, CNG_RSA_PRIVATE_KEY_BLOB) Then
                GoTo QH
            End If
        Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_P521_MAGIC
            Call CopyMemory(lSize, baBuffer(4), 4)
            Debug.Assert 8 + 3 * lSize <= UBound(baBuffer) + 1
            Call CopyMemory(baBuffer(0), baBuffer(8 + 2 * lSize), lSize)
            pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
            If Not pvPkiExportEccPrivateKey(baPrivKey, baBuffer, lMagic) Then
                GoTo QH
            End If
        Case Else
            Debug.Print "Unknown CNG private key magic (0x" & Hex$(lMagic) & ")"
        End Select
    Else
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, ByVal 0, lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertGetCertificateContextProperty(CERT_KEY_PROV_INFO_PROP_ID)"
            GoTo QH
        End If
        pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, baBuffer(0), lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertGetCertificateContextProperty(CERT_KEY_PROV_INFO_PROP_ID)#2"
            GoTo QH
        End If
        Call CopyMemory(uKeyInfo, baBuffer(0), Len(uKeyInfo))
        If CryptAcquireContext(hProv, uKeyInfo.pwszContainerName, uKeyInfo.pwszProvName, uKeyInfo.dwProvType, uKeyInfo.dwFlags) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptAcquireContext"
            GoTo QH
        End If
        If CryptGetUserKey(hProv, uKeyInfo.dwKeySpec, hKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptGetUserKey"
            GoTo QH
        End If
        If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, ByVal 0, lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptExportKey(PRIVATEKEYBLOB)"
            GoTo QH
        End If
        pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
        If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, baBuffer(0), lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptExportKey(PRIVATEKEYBLOB)#2"
            GoTo QH
        End If
        Call CopyMemory(lMagic, baBuffer(8), 4)
        Select Case lMagic
        Case BCRYPT_RSAPRIVATE_MAGIC
            If Not pvPkiExportRsaPrivateKey(baPrivKey, baBuffer, PKCS_RSA_PRIVATE_KEY) Then
                GoTo QH
            End If
        Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_P521_MAGIC
            If Not pvPkiExportEccPrivateKey(baPrivKey, baBuffer, lMagic) Then
                GoTo QH
            End If
        Case Else
            Debug.Print "Unknown CAPI private key magic (0x" & Hex$(lMagic) & ")"
        End Select
    End If
    '--- success
    pvPkiExportPrivateKey = True
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If hProvOrKey <> 0 And lFree <> 0 Then
        Call CryptReleaseContext(hProvOrKey, 0)
    End If
    If hCngKey <> 0 And lFree <> 0 Then
        Call NCryptFreeObject(hCngKey)
    End If
    If hNewKey <> 0 Then
        Call NCryptFreeObject(hNewKey)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiExportRsaPrivateKey(baRetVal() As Byte, baPrivBlob() As Byte, ByVal lStructType As Long) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportRsaPrivateKey"
    Dim baRsaPrivKey()  As Byte
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim sObjId          As String
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lStructType, baPrivBlob(0), 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx"
        GoTo QH
    End If
    pvArrayAllocate baRsaPrivKey, lSize, FUNC_NAME & ".baRsaPrivKey"
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lStructType, baPrivBlob(0), 0, 0, baRsaPrivKey(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx#2"
        GoTo QH
    End If
    sObjId = StrConv(szOID_RSA_RSA, vbFromUnicode)
    With uPrivKey
        .Algorithm.pszObjId = StrPtr(sObjId)
        .PrivateKey.pbData = VarPtr(baRsaPrivKey(0))
        .PrivateKey.cbData = UBound(baRsaPrivKey) + 1
    End With
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, uPrivKey, 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
        GoTo QH
    End If
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, uPrivKey, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(PKCS_PRIVATE_KEY_INFO)#2"
        GoTo QH
    End If
    '--- success
    pvPkiExportRsaPrivateKey = True
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiExportEccPrivateKey(baRetVal() As Byte, baPrivBlob() As Byte, ByVal lMagic As Long) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportEccPrivateKey"
    Dim sObjId          As String
    Dim uEccPrivKey     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    sObjId = StrConv(Switch(lMagic = BCRYPT_ECDH_PRIVATE_P521_MAGIC, szOID_ECC_CURVE_P521, _
                            lMagic = BCRYPT_ECDH_PRIVATE_P384_MAGIC, szOID_ECC_CURVE_P384, _
                            True, szOID_ECC_CURVE_P256), vbFromUnicode)
    With uEccPrivKey
        .dwVersion = 1
        .PrivateKey.pbData = VarPtr(baPrivBlob(0))
        .PrivateKey.cbData = UBound(baPrivBlob) + 1
        .szCurveOid = StrPtr(sObjId)
    End With
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, uEccPrivKey, 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(X509_ECC_PRIVATE_KEY)"
        GoTo QH
    End If
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, uEccPrivKey, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(X509_ECC_PRIVATE_KEY)#2"
        GoTo QH
    End If
    '--- success
    pvPkiExportEccPrivateKey = True
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvPkiCloneKeyWithExportPolicy(ByVal hKey As Long, ByVal lPolicy As Long) As Long
    Const FUNC_NAME     As String = "pvPkiCloneKeyWithExportPolicy"
    Const STR_PASSWORD  As String = "0000"
    Dim baPkcs8()       As Byte
    Dim uParams         As NCryptBufferDesc
    Dim sSecret         As String
    Dim sObjId          As String
    Dim uPbeParams      As CRYPT_PKCS12_PBE_PARAMS
    Dim lSize           As Long
    Dim hProv           As Long
    Dim sKeyName        As String
    Dim hRetVal         As Long
    Dim baBuffer()      As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- export PKCS#8 password protected blob
    ReDim uParams.Buffers(0 To 2) As NCryptBuffer
    Debug.Assert RedimStats(FUNC_NAME & ".uParams.Buffers", 0)
    uParams.cBuffers = UBound(uParams.Buffers) + 1
    uParams.pBuffers = VarPtr(uParams.Buffers(0))
    sSecret = STR_PASSWORD
    With uParams.Buffers(0)
        .BufferType = NCRYPTBUFFER_PKCS_SECRET
        .pvBuffer = StrPtr(sSecret)
        .cbBuffer = LenB(sSecret) + 2
    End With
    sObjId = StrConv(szOID_PKCS_12_pbeWithSHA1And3KeyTripleDES, vbFromUnicode)
    With uParams.Buffers(1)
        .BufferType = NCRYPTBUFFER_PKCS_ALG_OID
        .pvBuffer = StrPtr(sObjId)
        .cbBuffer = LenB(sObjId) + 1
    End With
    uPbeParams.cbSalt = 8
    uPbeParams.iIterations = 2048
    With uParams.Buffers(2)
        .BufferType = NCRYPTBUFFER_PKCS_ALG_PARAM
        .pvBuffer = VarPtr(uPbeParams)
        .cbBuffer = 8 + uPbeParams.cbSalt
    End With
    hResult = NCryptExportKey(hKey, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, ByVal 0, 0, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptExportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    pvArrayAllocate baPkcs8, lSize, FUNC_NAME & ".baPkcs8"
    hResult = NCryptExportKey(hKey, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, baPkcs8(0), UBound(baPkcs8) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptExportKey(PKCS8_PRIVATEKEY)#2"
        GoTo QH
    End If
    '--- retrieve more key props
    hResult = NCryptGetProperty(hKey, StrPtr("Provider Handle"), hProv, 4, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Provider Handle)"
        GoTo QH
    End If
    hResult = NCryptGetProperty(hKey, StrPtr("Name"), ByVal 0, 0, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Name)"
        GoTo QH
    End If
    pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
    hResult = NCryptGetProperty(hKey, StrPtr("Name"), baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Name)#2"
        GoTo QH
    End If
    '--- remove trailing terminating zero too
    sKeyName = Replace(CStr(baBuffer), vbNullChar, vbNullString)
    '--- import PKCS#8 blob and set Export Policy before finalizing
    ReDim uParams.Buffers(0 To 1) As NCryptBuffer
    Debug.Assert RedimStats(FUNC_NAME & ".uParams.Buffers", 0)
    uParams.cBuffers = UBound(uParams.Buffers) + 1
    uParams.pBuffers = VarPtr(uParams.Buffers(0))
    sSecret = STR_PASSWORD
    With uParams.Buffers(0)
        .BufferType = NCRYPTBUFFER_PKCS_SECRET
        .pvBuffer = StrPtr(sSecret)
        .cbBuffer = LenB(sSecret) + 2
    End With
    With uParams.Buffers(1)
        .BufferType = NCRYPTBUFFER_PKCS_KEY_NAME
        .pvBuffer = StrPtr(sKeyName)
        .cbBuffer = LenB(sKeyName) + 2
    End With
    hResult = NCryptImportKey(hProv, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, hRetVal, baPkcs8(0), UBound(baPkcs8) + 1, NCRYPT_OVERWRITE_KEY_FLAG Or NCRYPT_DO_NOT_FINALIZE_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptImportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    hResult = NCryptSetProperty(hRetVal, StrPtr("Export Policy"), lPolicy, 4, NCRYPT_PERSIST_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptSetProperty(Export Policy)"
        GoTo QH
    End If
    hResult = NCryptFinalizeKey(hRetVal, 0)
    If hResult < 0 Then
        sApiSource = "NCryptFinalizeKey"
        GoTo QH
    End If
    pvPkiCloneKeyWithExportPolicy = hRetVal
QH:
    If hProv <> 0 Then
        Call NCryptFreeObject(hProv)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

Private Function pvToStringW(ByVal lPtr As Long) As String
    If lPtr Then
        pvToStringW = String$(lstrlenW(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvToStringW), ByVal lPtr, LenB(pvToStringW))
    End If
End Function

'= shared ================================================================

#If Not ImplUseShared Then
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
        Debug.Assert RedimStats("SplitOrReindex.vResult", 0)
        For lIdx = 0 To UBound(vTemp) Step 2
            vResult(vTemp(lIdx)) = vTemp(lIdx + 1)
        Next
        SplitOrReindex = vResult
    End If
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function ReadBinaryFile(sFile As String) As Byte()
    Const FUNC_NAME     As String = "ReadBinaryFile"
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    baBuffer = vbNullString
    If GetFileAttributes(sFile) <> -1 Then
        nFile = FreeFile
        Open sFile For Binary Access Read Shared As nFile
        If LOF(nFile) > 0 Then
            pvArrayAllocate baBuffer, LOF(nFile), FUNC_NAME & ".baBuffer"
            Get nFile, , baBuffer
        End If
        Close nFile
    End If
    ReadBinaryFile = baBuffer
End Function

Private Property Get OsVersion() As UcsOsVersionEnum
    Static lVersion     As Long
    Dim aVer(0 To 37)   As Long
    
    If lVersion = 0 Then
        aVer(0) = 4 * UBound(aVer)              '--- [0] = dwOSVersionInfoSize
        If GetVersionEx(aVer(0)) <> 0 Then
            lVersion = aVer(1) * 100 + aVer(2)  '--- [1] = dwMajorVersion, [2] = dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property
#End If

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnConnect()
    If m_bUseTls Then
        If Not pvOnConnect() Then
            GoTo QH
        End If
    Else
        RaiseEvent OnConnect
    End If
QH:
End Sub

Private Sub m_oSocket_OnReceive()
    If m_bUseTls Then
        If m_lRecvSize = 0 Or m_uCtx.State = ucsTlsStateShutdown Then
            If Not pvOnReceive() Then
                GoTo QH
            End If
        End If
    Else
        RaiseEvent OnReceive
    End If
QH:
End Sub

Private Sub m_oSocket_OnSend()
    If m_bUseTls Then
        If Not pvTlsIsReady(m_uCtx) Then
            GoTo QH
        End If
        If m_lSendPos >= m_lSendSize Then
            m_lSendPos = 0
            m_lSendSize = 0
            m_lLastSendBytes = m_lSendBytes
            m_lSendBytes = 0
            If pvOnBeforeNotify(ucsSfdWrite) Then
                RaiseEvent OnSend
            End If
        ElseIf m_lSendPos > 0 Then
            If Not pvOnSend() Then
                GoTo QH
            End If
        End If
    Else
        RaiseEvent OnSend
    End If
QH:
End Sub

Private Sub m_oSocket_BeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Cancel As Boolean)
    If Not m_bUseTls Then
        RaiseEvent BeforeNotify(EventMask, Cancel)
    End If
End Sub

Private Sub m_oSocket_OnAccept()
    If m_bUseTls Then
        If pvOnBeforeNotify(ucsSfdAccept) Then
            RaiseEvent OnAccept
        End If
    Else
        RaiseEvent OnAccept
    End If
End Sub

Private Sub m_oSocket_OnClose()
    If m_bUseTls Then
        If Not pvOnReceive(Flush:=True) Then
            GoTo QH
        End If
        If pvTlsIsClosed(m_uCtx) Then
            GoTo QH
        End If
        If pvOnBeforeNotify(ucsSfdClose) Then
            RaiseEvent OnClose
            m_uCtx.State = ucsTlsStateClosed
            m_bUseTls = False
        End If
    Else
        RaiseEvent OnClose
    End If
QH:
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    RaiseEvent OnError(ErrorCode, EventMask)
End Sub

Private Sub m_oSocket_OnMessagePending(Handled As Boolean)
    RaiseEvent OnMessagePending(Handled)
End Sub

Private Sub m_oSocket_OnResolve(IpAddress As String)
    RaiseEvent OnResolve(IpAddress)
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_oSocket = New cAsyncSocket
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
    If m_hRootStore <> 0 Then
        Call CertCloseStore(m_hRootStore, 0)
    End If
End Sub
