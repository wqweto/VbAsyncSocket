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
' VbAsyncSocket Project (c) 2018-2023 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cTlsSocket"

#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplSync = Not (ASYNCSOCKET_NO_SYNC <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)
#Const ImplTlsServer = (ASYNCSOCKET_NO_TLSSERVER = 0)
#Const ImplEncryptedPrivateKey = True

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
Event OnCertificate(Issuers As Object, Confirmed As Boolean)

Public Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls10 = 2 ^ 0
    ucsTlsSupportTls11 = 2 ^ 1
    ucsTlsSupportTls12 = 2 ^ 2
    ucsTlsSupportTls13 = 2 ^ 3
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 4
    ucsTlsIgnoreServerCertificateRevocation = 2 ^ 5
    ucsTlsSupportAll = ucsTlsSupportTls10 Or ucsTlsSupportTls11 Or ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_DELETEKEYSET                        As Long = &H10
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const X509_OCTET_STRING                         As Long = 25
Private Const X509_INTEGER                              As Long = 27
Private Const X509_SEQUENCE_OF_ANY                      As Long = 34
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const PKCS_ENCRYPTED_PRIVATE_KEY_INFO           As Long = 45
Private Const X509_OBJECT_IDENTIFIER                    As Long = 73
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CNG_RSA_PRIVATE_KEY_BLOB                  As Long = 83
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_SHARE_OID_STRING_FLAG        As Long = &H4
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
Private Const X509_NAME                                 As Long = 7
'--- for CryptSignHash
Private Const AT_KEYEXCHANGE                            As Long = 1
Private Const RSA1024BIT_KEY                            As Long = &H4000000
Private Const NTE_BAD_ALGID                             As Long = &H80090008
'--- for CertGetCertificateContextProperty
Private Const CERT_KEY_PROV_HANDLE_PROP_ID              As Long = 1
Private Const CERT_KEY_PROV_INFO_PROP_ID                As Long = 2
'--- for PFXImportCertStore
Private Const CRYPT_EXPORTABLE                          As Long = &H1
Private Const PKCS12_NO_PERSIST_KEY                     As Long = &H8000&
'--- for CryptExportKey
Private Const PRIVATEKEYBLOB                            As Long = 7
'--- for CryptAcquireCertificatePrivateKey
Private Const CRYPT_ACQUIRE_CACHE_FLAG                  As Long = &H1
Private Const CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG       As Long = &H10000
Private Const CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG      As Long = &H20000
'--- for NCryptImportKey
Private Const NCRYPT_OVERWRITE_KEY_FLAG                 As Long = &H80
Private Const NCRYPT_DO_NOT_FINALIZE_FLAG               As Long = &H400
'--- for NCryptSetProperty
Private Const NCRYPT_PERSIST_FLAG                       As Long = &H80000000
'--- for CertStrToName
Private Const CERT_OID_NAME_STR                         As Long = 2
'--- for CertOpenStore
Private Const CERT_STORE_PROV_MEMORY                    As Long = 2
Private Const CERT_STORE_PROV_SYSTEM                    As Long = 10
Private Const CERT_STORE_CREATE_NEW_FLAG                As Long = &H2000
Private Const CERT_SYSTEM_STORE_CURRENT_USER            As Long = &H10000
Private Const CERT_SYSTEM_STORE_LOCAL_MACHINE           As Long = &H20000
Private Const CERT_SYSTEM_STORE_CURRENT_SERVICE         As Long = &H40000
Private Const CERT_SYSTEM_STORE_SERVICES                As Long = &H50000
Private Const CERT_SYSTEM_STORE_USERS                   As Long = &H60000
'--- for CertAddEncodedCertificateToStore
Private Const CERT_STORE_ADD_USE_EXISTING               As Long = 2
'--- for CertGetCertificateChain
Private Const CERT_CHAIN_REVOCATION_CHECK_CHAIN         As Long = &H20000000
Private Const CERT_TRUST_IS_NOT_TIME_VALID              As Long = &H1
Private Const CERT_TRUST_IS_NOT_TIME_NESTED             As Long = &H2
Private Const CERT_TRUST_IS_REVOKED                     As Long = &H4
Private Const CERT_TRUST_IS_NOT_SIGNATURE_VALID         As Long = &H8
Private Const CERT_TRUST_IS_UNTRUSTED_ROOT              As Long = &H20
Private Const CERT_TRUST_REVOCATION_STATUS_UNKNOWN      As Long = &H40
Private Const CERT_TRUST_IS_PARTIAL_CHAIN               As Long = &H10000
'--- for CertFindCertificateInStore
Private Const CERT_FIND_ANY                             As Long = 0
Private Const CERT_FIND_EXISTING                        As Long = &HD0000
Private Const CERT_FIND_ISSUER_ATTR                     As Long = &H30004
Private Const CERT_FIND_SUBJECT_STR                     As Long = &H80007
Private Const CERT_FIND_HASH_STR                        As Long = &H140000
'--- for CERT_ALT_NAME_ENTRY
Private Const CERT_ALT_NAME_DNS_NAME                    As Long = 3
'--- for CertCreateCertificateChainEngine
Private Const CERT_CHAIN_CACHE_END_CERT                 As Long = 1
'--- for CertSetCertificateContextProperty
Private Const CERT_OCSP_RESPONSE_PROP_ID                As Long = 70
Private Const CERT_SET_PROPERTY_IGNORE_PERSIST_ERROR_FLAG As Long = &H80000000
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
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
'--- errors
Private Const WSAENOTCONN                               As Long = 10057
Private Const ERR_TIMEOUT                               As Long = &H800705B4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpStr As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpStr As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
'--- msvbvm60
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
'--- Crypt32
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvParameters As Long, phCryptProvOrNCryptKey As Long, pdwKeySpec As Long, pfCallerFreeProvOrNCryptKey As Long) As Long
Private Declare Function PFXImportCertStore Lib "crypt32" (pPFX As Any, ByVal szPassword As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertStrToName Lib "crypt32" Alias "CertStrToNameW" (ByVal dwCertEncodingType As Long, ByVal pszX500 As Long, ByVal dwStrType As Long, ByVal pvReserved As Long, pbEncoded As Any, pcbEncoded As Long, ByVal ppszError As Long) As Long
Private Declare Function CertCreateSelfSignCertificate Lib "crypt32" (ByVal hCryptProvOrNCryptKey As Long, pSubjectIssuerBlob As Any, ByVal dwFlags As Long, pKeyProvInfo As Any, ByVal pSignatureAlgorithm As Long, pStartTime As Any, pEndTime As Any, ByVal pExtensions As Long) As Long
Private Declare Function CertOpenStore Lib "crypt32" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
Private Declare Function CertCloseStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertAddEncodedCertificateToStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long, ByVal dwAddDisposition As Long, ppCertContext As Any) As Long
Private Declare Function CertCreateCertificateChainEngine Lib "crypt32" (pConfig As Any, phChainEngine As Long) As Long
Private Declare Function CertFreeCertificateChainEngine Lib "crypt32" (ByVal hChainEngine As Long) As Long
Private Declare Function CertGetCertificateChain Lib "crypt32" (ByVal hChainEngine As Long, ByVal pCertContext As Long, ByVal pTime As Long, ByVal hAdditionalStore As Long, pChainPara As Any, ByVal dwFlags As Long, ByVal pvReserved As Long, ppChainContext As Long) As Long
Private Declare Function CertFreeCertificateChain Lib "crypt32" (ByVal pChainContext As Long) As Long
Private Declare Function CertFindExtension Lib "crypt32" (ByVal pszObjId As String, ByVal cExtensions As Long, ByVal rgExtensions As Long) As Long
Private Declare Function CertFindCertificateInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, pvFindPara As Any, ByVal pPrevCertContext As Long) As Long
Private Declare Function CertSetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, ByVal dwFlags As Long, pvData As Any) As Long
Private Declare Function CertAddCertificateContextToStore Lib "crypt32" (ByVal hCertStore As Long, ByVal pCertContext As Long, ByVal dwAddDisposition As Long, ppStoreContext As Any) As Long
'--- cryptdlg
Private Declare Function CertSelectCertificate Lib "cryptdlg" Alias "CertSelectCertificateW" (pCertSelectInfo As Any) As Long
'--- NCrypt
Private Declare Function NCryptImportKey Lib "ncrypt" (ByVal hProvider As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, phKey As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptExportKey Lib "ncrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Any, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFreeObject Lib "ncrypt" (ByVal hKey As Long) As Long
Private Declare Function NCryptGetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptSetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFinalizeKey Lib "ncrypt" (ByVal hKey As Long, ByVal dwFlags As Long) As Long
'--- BCrypt
#If ImplTlsServer And ImplEncryptedPrivateKey Then
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal hPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Currency, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDecrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, pbIV As Any, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
#End If
#If Not ImplUseShared Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private Type CRYPT_DATA_BLOB
    cbData              As Long
    pbData              As Long
End Type
Private Const sizeof_CRYPT_DATA_BLOB As Long = 8

Private Type CRYPT_BIT_BLOB
    cbData              As Long
    pbData              As Long
    cUnusedBits         As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId            As Long
    Parameters          As CRYPT_DATA_BLOB
End Type

Private Type CERT_PUBLIC_KEY_INFO
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PublicKey           As CRYPT_BIT_BLOB
End Type

Private Type CRYPT_ECC_PRIVATE_KEY_INFO
    dwVersion           As Long
    PrivateKey          As CRYPT_DATA_BLOB
    szCurveOid          As Long
    PublicKey           As CRYPT_DATA_BLOB
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
    dwVersion           As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_DATA_BLOB
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
    SerialNumber        As CRYPT_DATA_BLOB
    SignatureAlgorithm  As CRYPT_ALGORITHM_IDENTIFIER
    Issuer              As CRYPT_DATA_BLOB
    NotBefore           As FILETIME
    NotAfter            As FILETIME
    Subject             As CRYPT_DATA_BLOB
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
    Value               As CRYPT_DATA_BLOB
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

Private Type CERT_RDN_ATTR
    pszObjId            As Long
    dwValueType         As Long
    cbData              As Long
    pbData              As Long
End Type
Private Const sizeof_CERT_RDN_ATTR As Long = 16

Private Type CERT_RDN
    cRDNAttr            As Long
    rgRDNAttr           As Long
    Buffer()            As CERT_RDN_ATTR
End Type

Private Type CERT_SELECT_STRUCT_W
    dwSize              As Long
    hWndParent          As Long
    hInstance           As Long
    pTemplateName       As Long
    dwFlags             As Long
    szTitle             As Long
    cCertStore          As Long
    arrayCertStore      As Long
    szPurposeOid        As Long
    cCertContext        As Long
    arrayCertContext    As Long
    lCustData           As Long
    pfnHook             As Long
    pfnFilter           As Long
    szHelpFileName      As Long
    dwHelpId            As Long
    hProv               As Long
End Type

Private Type CRYPT_ENCRYPTED_PRIVATE_KEY_INFO
    EncryptionAlgorithm As CRYPT_ALGORITHM_IDENTIFIER
    EncryptedPrivateKey As CRYPT_DATA_BLOB
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CHR1                                  As String = "" '--- CHAR(1)
Private Const DEF_TIMEOUT                               As Long = 5000
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000
Private Const LNG_STATUS_DATA_ERROR                     As Long = &HC000003E
'--- errors
Private Const ERR_NO_MATCHING_ALT_NAME                  As String = "No certificate subject name matches target host name"
Private Const ERR_NO_CERTIFICATE                        As String = "No certificate is available"
Private Const ERR_TRUST_IS_REVOKED                      As String = "Trust for this certificate or one of the certificates in the certificate chain has been revoked"
Private Const ERR_TRUST_IS_NOT_SIGNATURE_VALID          As String = "The certificate or one of the certificates in the certificate chain does not have a valid signature"
Private Const ERR_TRUST_IS_PARTIAL_CHAIN                As String = "The certificate chain is not complete"
Private Const ERR_TRUST_IS_UNTRUSTED_ROOT               As String = "The certificate or certificate chain is based on an untrusted root"
Private Const ERR_TRUST_IS_NOT_TIME_VALID               As String = "The certificate has expired"
Private Const ERR_TRUST_REVOCATION_STATUS_UNKNOWN       As String = "The revocation status of the certificate or one of the certificates in the certificate chain is unknown"
Private Const ERR_UNKNOWN_CNG_MAGIC                     As String = "Unknown CNG private key magic (%1)"
Private Const ERR_UNKNOWN_CAPI_MAGIC                    As String = "Unknown CAPI private key magic (%1)"
Private Const ERR_UNKNOWN_CERTIFICATECHAIN_ERROR_MASK   As String = "Unknown CertGetCertificateChain error mask (%1)"
Private Const ERR_UNKNOWN_OID                           As String = "Unknown %1 OID (%2)"
Private Const ERR_INCORRECT_PASSWORD                    As String = "Incorrect password"
Private Const ERR_NO_SERVER_COMPILED                    As String = "Server-side TLS not compiled (ASYNCSOCKET_NO_TLSSERVER = 1)"

Private WithEvents m_oAsyncSocket As cAsyncSocket
Attribute m_oAsyncSocket.VB_VarHelpID = -1
Private m_oSocket               As Object
Private m_bUseTls               As Boolean
Private m_bIsServer             As Boolean
Private m_sRemoteHostName       As String
Private m_eLocalFeatures        As UcsTlsLocalFeaturesEnum
Private m_uCtx                  As UcsTlsContext
Private m_hRootStore            As Long
Private m_oRootCa               As cTlsSocket
Private m_sAlpnProtocols        As String
Private m_lSendActual           As Long
Private m_lSendBytes            As Long
Private m_lLastSendBytes        As Long
Private m_baRecvBuffer()        As Byte
Private m_lRecvPos              As Long
Private m_baSendBuffer()        As Byte
Private m_lSendPos              As Long
Private m_lCallbackPtr          As Long
Private m_eRaisedEvent          As UcsAsyncSocketEventMaskEnum
Private m_lLastErrNumber        As Long
Private m_sLastErrSource        As String
Private m_sLastError            As String

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

Public Property Get Socket() As Object
    Set Socket = m_oSocket
End Property

Public Property Set Socket(oValue As Object)
    Set m_oSocket = oValue
End Property

Public Property Get SocketHandle() As Long
    SocketHandle = m_oSocket.SocketHandle
End Property

Public Property Get LastError() As VBA.ErrObject
    Const LNG_FACILITY_WIN32 As Long = &H80070000
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    
    Set LastError = Err
    With LastError
        If m_lLastErrNumber <> 0 Then
            .Description = m_sLastError
            .Number = m_lLastErrNumber
            .Source = m_sLastErrSource
        Else
            .Description = TlsGetLastError(m_uCtx, lErrNumber, sErrSource)
            If LenB(.Description) <> 0 And lErrNumber <> 0 Then
                .Number = IIf(lErrNumber < 0, lErrNumber, lErrNumber Or LNG_FACILITY_WIN32)
                .Source = sErrSource
            ElseIf m_oSocket.HasPendingEvent Then
                .Number = 0
            Else
                Set LastError = m_oSocket.LastError
            End If
        End If
    End With
End Property

Public Property Get IsClosed() As Boolean
    If m_bUseTls Then
        IsClosed = TlsIsClosed(m_uCtx) Or TlsIsShutdown(m_uCtx)
    Else
        IsClosed = m_oSocket.IsClosed
    End If
End Property

Public Property Get AvailableBytes() As Long
    If m_bUseTls Then
        AvailableBytes = m_lRecvPos
    Else
        AvailableBytes = m_oSocket.AvailableBytes
    End If
End Property

Public Property Get LastSendBytes() As Long
    If m_bUseTls Then
        LastSendBytes = m_lLastSendBytes
    Else
        LastSendBytes = m_oSocket.LastSendBytes
    End If
End Property

Public Property Get HasPendingResolve() As Boolean
    HasPendingResolve = m_oSocket.HasPendingResolve
End Property

Public Property Get HasPendingEvent() As Boolean
    HasPendingEvent = m_oSocket.HasPendingEvent
End Property

Public Property Get IsServer() As Boolean
    IsServer = m_bIsServer
End Property

Public Property Get RemoteHostName() As String
    RemoteHostName = m_sRemoteHostName
End Property

Public Property Get LocalFeatures() As UcsTlsLocalFeaturesEnum
    LocalFeatures = m_eLocalFeatures
End Property

Public Property Get LocalCertificates() As Collection
    Set LocalCertificates = m_uCtx.LocalCertificates
End Property

Public Property Set LocalCertificates(oValue As Collection)
    Set m_uCtx.LocalCertificates = oValue
End Property

Public Property Get LocalPrivateKey() As Collection
    Set LocalPrivateKey = m_uCtx.LocalPrivateKey
End Property

Public Property Set LocalPrivateKey(oValue As Collection)
    Set m_uCtx.LocalPrivateKey = oValue
End Property

Public Property Get RemoteCertificates() As Collection
    Set RemoteCertificates = m_uCtx.RemoteCertificates
End Property

Public Property Get RemoteCertStatuses() As Collection
    Set RemoteCertStatuses = m_uCtx.RemoteCertStatuses
End Property

Public Property Get AlpnNegotiated() As String
    AlpnNegotiated = m_uCtx.AlpnNegotiated
End Property

Public Property Get SniRequested() As String
    SniRequested = m_uCtx.SniRequested
End Property

Public Property Get SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket) As Long
    SockOpt = m_oSocket.SockOpt(OptionName, Level)
End Property

Public Property Let SockOpt(ByVal OptionName As UcsAsyncSocketOptionNameEnum, Optional ByVal Level As UcsAsyncSocketOptionLevelEnum = ucsSolSocket, ByVal Value As Long)
    m_oSocket.SockOpt(OptionName, Level) = Value
End Property

Public Property Get CallbackWeakRef() As Object
    Call vbaObjSetAddref(CallbackWeakRef, m_lCallbackPtr)
End Property

Public Property Set CallbackWeakRef(oValue As Object)
    m_lCallbackPtr = ObjPtr(oValue)
End Property

Friend Property Get frRootStore() As Long
    If Not m_oRootCa Is Nothing Then
        frRootStore = m_oRootCa.frRootStore
    Else
        frRootStore = m_hRootStore
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function InitServerTls( _
            Optional CertFile As String, _
            Optional Password As String, _
            Optional CertSubject As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection, _
            Optional AlpnProtocols As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum) As Boolean
#If Not ImplTlsServer Then
    ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
#Else
    Const FUNC_NAME     As String = "InitServerTls"
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection
    
    If m_bUseTls Then
        GoTo QH
    End If
    pvSetLastError
    If pvCollectionCount(Certificates) > 0 And pvCollectionCount(PrivateKey) > 0 Then
        Set cCerts = Certificates
        Set cPrivKey = PrivateKey
        GoTo StartTls
    End If
    If LenB(CertFile) <> 0 Then
        If pvPkiPkcs12ImportCertificates(CertFile, Password, cCerts, cPrivKey) Then
            GoTo StartTls
        End If
        If pvPkiPemImportCertificates(Split(CertFile, "|"), Password, cCerts, cPrivKey) Then
            GoTo StartTls
        End If
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
        GoTo QH
    End If
    If LenB(CertSubject) <> 0 Then
        If pvPkiSystemStoreImportCertificate(CertSubject, 0, cCerts, cPrivKey) Then
            GoTo StartTls
        End If
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
        GoTo QH
    End If
StartTls:
    m_bUseTls = True
    m_bIsServer = True
    m_oSocket.GetPeerName m_sRemoteHostName, 0
    m_sAlpnProtocols = AlpnProtocols
    If (LocalFeatures And ucsTlsSupportAll) = 0 Then
        m_eLocalFeatures = LocalFeatures Or ucsTlsSupportAll
    Else
        m_eLocalFeatures = LocalFeatures
    End If
    If Not TlsInitServer(m_uCtx, m_sRemoteHostName, cCerts, cPrivKey, m_sAlpnProtocols, m_eLocalFeatures) Then
        GoTo QH
    End If
    '--- success
    InitServerTls = True
QH:
#End If
End Function

Public Function Accept( _
            ConnectedSocket As cTlsSocket, _
            Optional SocketAddress As String, _
            Optional SocketPort As Long, _
            Optional ByVal UseTls As Variant) As Boolean
    Const FUNC_NAME     As String = "Accept"
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection
    
    On Error GoTo EH
    pvSetLastError
    If ConnectedSocket Is Nothing Then
        Set ConnectedSocket = New cTlsSocket
    End If
    If Not m_oSocket.Accept(ConnectedSocket.Socket, SocketAddress, SocketPort) Then
        GoTo QH
    End If
    If IsMissing(UseTls) Then
        UseTls = m_bUseTls
    End If
    If UseTls Then
        #If Not ImplTlsServer Then
            ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
        #Else
            If pvCollectionCount(LocalCertificates) = 0 Then
                If Not FireOnCertificate(Nothing) Then
                    If Not pvPkiGenerSelfSignedCertificate(cCerts, cPrivKey) Then
                        GoTo QH
                    End If
                    Set LocalCertificates = cCerts
                    Set LocalPrivateKey = cPrivKey
                End If
            End If
            If Not ConnectedSocket.InitServerTls(Certificates:=LocalCertificates, PrivateKey:=LocalPrivateKey, AlpnProtocols:=m_sAlpnProtocols, LocalFeatures:=m_eLocalFeatures) Then
                GoTo QH
            End If
        #End If
    End If
    '--- success
    Accept = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function Connect( _
            HostAddress As String, _
            ByVal HostPort As Long, _
            Optional ByVal UseTls As Boolean = True, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum, _
            Optional RootCa As cTlsSocket, _
            Optional AlpnProtocols As String) As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    m_bUseTls = UseTls
    m_sRemoteHostName = HostAddress
    m_eLocalFeatures = LocalFeatures
    Set m_oRootCa = RootCa
    m_sAlpnProtocols = AlpnProtocols
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
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum, _
            Optional RootCa As cTlsSocket, _
            Optional AlpnProtocols As String) As Boolean
    Const FUNC_NAME     As String = "StartTls"
    Dim bResult         As Boolean
    Dim baEmpty()       As Byte
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection
    
    On Error GoTo EH
    If TlsIsStarted(m_uCtx) Then
        GoTo QH
    End If
    m_bUseTls = True
    m_bIsServer = False
    If LenB(RemoteHostName) <> 0 Then
        m_sRemoteHostName = RemoteHostName
    End If
    If (LocalFeatures And ucsTlsSupportAll) = 0 Then
        m_eLocalFeatures = LocalFeatures Or ucsTlsSupportAll _
            Or IIf(OsVersion < ucsOsvXp, ucsTlsIgnoreServerCertificateErrors, 0) _
            Or IIf(OsVersion < ucsOsvWin7, ucsTlsIgnoreServerCertificateRevocation, 0)
    Else
        m_eLocalFeatures = LocalFeatures
    End If
    If Not RootCa Is Nothing Then
        Set m_oRootCa = RootCa
        If m_hRootStore <> 0 Then
            Call CertCloseStore(m_hRootStore, 0)
            m_hRootStore = 0
        End If
    End If
    If LenB(AlpnProtocols) <> 0 Then
        m_sAlpnProtocols = AlpnProtocols
    End If
    Set cCerts = LocalCertificates
    Set cPrivKey = LocalPrivateKey
    If Not TlsInitClient(m_uCtx, m_sRemoteHostName, m_eLocalFeatures, Me, m_sAlpnProtocols) Then
        GoTo QH
    End If
    Set LocalCertificates = cCerts
    Set LocalPrivateKey = cPrivKey
    bResult = TlsHandshake(m_uCtx, baEmpty, 0, m_baSendBuffer, m_lSendPos)
    If Not pvHandleSend() Then
        GoTo QH
    End If
    If Not bResult Then
        pvFireOnError LastError.Number, ucsSfdConnect
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
        If m_lRecvPos = 0 Then
            If TlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Not pvHandleReceive() Then
                GoTo QH
            End If
        End If
        If m_lRecvPos > 0 Then
            Buffer = m_baRecvBuffer
            Erase m_baRecvBuffer
            m_lRecvPos = 0
        Else
            Buffer = vbNullString
        End If
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

Public Function Receive( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long) As Long
    Const FUNC_NAME     As String = "Receive"
    
    On Error GoTo EH
    If m_bUseTls Then
        If m_lRecvPos = 0 Then
            If TlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Not pvHandleReceive() Then
                GoTo QH
            End If
        End If
        If m_lRecvPos > 0 Then
            If m_lRecvPos <= BufLen Then
                Receive = m_lRecvPos
                Call CopyMemory(ByVal BufPtr, m_baRecvBuffer(0), Receive)
                Erase m_baRecvBuffer
                m_lRecvPos = 0
            Else
                Receive = BufLen
                Call CopyMemory(ByVal BufPtr, m_baRecvBuffer(0), Receive)
                m_lRecvPos = m_lRecvPos - BufLen
                Call CopyMemory(m_baRecvBuffer(0), m_baRecvBuffer(BufLen), m_lRecvPos)
            End If
        End If
    Else
        Receive = m_oSocket.Receive(BufPtr, BufLen, HostAddress, HostPort)
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
    Dim baTemp()        As Byte
    
    On Error GoTo EH
    If m_bUseTls Then
        baTemp = m_oSocket.ToTextArray(Text, CodePage)
        SendText = SendArray(baTemp, HostAddress, HostPort)
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
    
    On Error GoTo EH
    If m_bUseTls Then
        If TlsIsClosed(m_uCtx) Then
            GoTo QH
        End If
        lInputSize = pvArraySize(Buffer)
        m_lSendBytes = m_lSendBytes + lInputSize
        If Not TlsSend(m_uCtx, Buffer, lInputSize, m_baSendBuffer, m_lSendPos) Then
            pvFireOnError LastError.Number, ucsSfdWrite
            GoTo QH
        End If
        If TlsIsReady(m_uCtx) Then
            If Not pvHandleSend() Then
                GoTo QH
            End If
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
        If TlsIsReady(m_uCtx) And Not TlsIsShutdown(m_uCtx) Then
            If Not pvHandleShutdown() Then
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
    If ErrorCode = vbObjectError Then
        GetErrorDescription = TlsGetLastError(m_uCtx)
    Else
        GetErrorDescription = m_oSocket.GetErrorDescription(ErrorCode)
    End If
End Function

Public Sub Close_()
    If m_bUseTls Then
        If TlsIsReady(m_uCtx) And Not TlsIsShutdown(m_uCtx) Then
            If Not pvHandleShutdown() Then
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
            pvFireOnReceive
        Case ucsSfdWrite
            pvFireOnSend
        Case ucsSfdConnect
            pvFireOnConnect
        Case ucsSfdAccept
            pvFireOnAccept
        Case ucsSfdClose
            pvFireOnClose
            TlsTerminate m_uCtx
        End Select
        pvFireAfterNotify EventMask
        If EventMask = ucsSfdRead And AvailableBytes = 0 Then
            m_oSocket.Receive 0, 0 '--- enable FD_READ notification
        End If
    Else
        m_oSocket.PostEvent EventMask, Immediate
    End If
End Sub

'= forwarded =============================================================

Public Function Create( _
            Optional ByVal SocketPort As Long, _
            Optional ByVal SocketType As UcsAsyncSocketTypeEnum = ucsSckStream, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll, _
            Optional SocketAddress As String) As Boolean
    If m_bUseTls Then
        If IsClosed And Not m_oSocket.IsClosed Then
            m_oSocket.Close_
        End If
    End If
    Create = m_oSocket.Create(SocketPort, SocketType, EventMask, SocketAddress)
End Function

Public Function Listen(Optional ByVal ConnectionBacklog As Long = 5) As Boolean
    Listen = m_oSocket.Listen(ConnectionBacklog)
End Function

Public Function GetSockName(SocketAddress As String, SocketPort As Long) As Boolean
    GetSockName = m_oSocket.GetSockName(SocketAddress, SocketPort)
End Function

Public Function GetPeerName(PeerAddress As String, PeerPort As Long) As Boolean
    GetPeerName = m_oSocket.GetPeerName(PeerAddress, PeerPort)
End Function

Public Function GetLocalHost(HostName As String, HostAddress As String, Optional AdaptersInfo As Variant, Optional ForwardTable As Variant) As Boolean
    GetLocalHost = m_oSocket.GetLocalHost(HostName, HostAddress, AdaptersInfo, ForwardTable)
End Function

Public Function Attach( _
            ByVal SocketHandle As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Attach = m_oSocket.Attach(SocketHandle, EventMask)
End Function

Public Function Detach() As Long
    Detach = m_oSocket.Detach()
End Function

Public Function Bind(Optional SocketAddress As String, Optional ByVal SocketPort As Long) As Boolean
    Bind = m_oSocket.Bind(SocketAddress, SocketPort)
End Function

Public Function FromTextArray(baText() As Byte, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    FromTextArray = m_oSocket.FromTextArray(baText, CodePage)
End Function

Public Function ToTextArray(sText As String, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As Byte()
    ToTextArray = m_oSocket.ToTextArray(sText, CodePage)
End Function

Public Sub SyncProcessMsgQueue()
    m_oSocket.SyncProcessMsgQueue
End Sub

Public Sub SetLastError(ByVal lValue As Long)
    m_oSocket.SetLastError lValue
End Sub

'= sync ==================================================================

#If ImplSync Then
Public Function SyncConnect( _
            HostAddress As String, _
            ByVal HostPort As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal UseTls As Boolean = True, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum, _
            Optional RootCa As cTlsSocket, _
            Optional AlpnProtocols As String) As Boolean
    Const FUNC_NAME     As String = "SyncConnect"
    
    On Error GoTo EH
    If UseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        If Not Connect(HostAddress, HostPort, UseTls:=UseTls, LocalFeatures:=LocalFeatures, RootCa:=RootCa, AlpnProtocols:=AlpnProtocols) Then
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
    Dim baTemp()        As Byte
    
    On Error GoTo EH
    baTemp = m_oSocket.ToTextArray(Text, CodePage)
    SyncSendText = SyncSendArray(baTemp, Timeout:=Timeout)
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
            Optional ByVal NeedLen As Long = 1, _
            Optional ByVal Timeout As Long, _
            Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpUtf8) As String
    Const FUNC_NAME     As String = "SyncReceiveText"
    Dim lElapsed        As Long
    Dim dblTimer        As Double
    Dim baRecv()        As Byte
    Dim lPos            As Long
    Dim baBuffer()      As Byte
    Dim bResult         As Boolean
    
    On Error GoTo EH
    If m_bUseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        baBuffer = vbNullString
        dblTimer = TimerEx
        bResult = True
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
            ElseIf TlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.SetLastError ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not bResult Then
                GoTo QH
            End If
            bResult = SyncWaitForEvent(Timeout - lElapsed, ucsSfdRead)
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
    Dim bResult         As Boolean
    
    On Error GoTo EH
    If m_bUseTls Then
        If Timeout = 0 Then
            Timeout = DEF_TIMEOUT
        End If
        Buffer = vbNullString
        dblTimer = TimerEx
        bResult = True
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
            ElseIf TlsIsClosed(m_uCtx) Then
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.SetLastError ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not bResult Then
                GoTo QH
            End If
            bResult = SyncWaitForEvent(Timeout - lElapsed, ucsSfdRead)
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

Public Function SyncStartTls( _
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum, _
            Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "SyncStartTls"
    
    On Error GoTo EH
    m_eRaisedEvent = 0
    If Not StartTls(RemoteHostName, LocalFeatures) Then
        GoTo QH
    End If
    If Timeout = 0 Then
        Timeout = DEF_TIMEOUT
    End If
    If Not SyncWaitForEvent(Timeout, ucsSfdConnect) Then
        GoTo QH
    End If
    '--- success
    SyncStartTls = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function SyncWaitForEvent( _
            ByVal Timeout As Long, _
            Optional ByVal EventMask As UcsAsyncSocketEventMaskEnum = ucsSfdAll) As Boolean
    Const FUNC_NAME     As String = "SyncWaitForEvent"
    Const LNG_MASK      As Long = ucsSfdConnect Or ucsSfdRead Or ucsSfdWrite
    Dim dblTimer        As Double
    Dim lElapsed        As Long
    Dim lErrNumber      As Long
    
    On Error GoTo EH
    If m_bUseTls And (EventMask And LNG_MASK) <> 0 Then
        dblTimer = TimerEx
        m_oSocket.SyncProcessMsgQueue
        Do
            If (m_eRaisedEvent And EventMask) <> 0 Then
                m_eRaisedEvent = 0
                Exit Do
            End If
            If TlsIsClosed(m_uCtx) Then
                m_oSocket.SetLastError WSAENOTCONN
                GoTo QH
            End If
            If Timeout > 0 Then
                lElapsed = Int((TimerEx - dblTimer) * 1000)
                If lElapsed >= Timeout Then
                    m_oSocket.SetLastError ERR_TIMEOUT
                    GoTo QH
                End If
            End If
            If Not m_oSocket.SyncWaitForEvent(Timeout - lElapsed, ucsSfdAll) Then
                GoTo QH
            End If
            m_oSocket.SyncProcessMsgQueue
            TlsGetLastError m_uCtx, lErrNumber
            If lErrNumber <> 0 Then
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

Private Function pvHandleConnect() As Boolean
    Const FUNC_NAME     As String = "pvHandleConnect"
    
    On Error GoTo EH
    If Not StartTls(m_sRemoteHostName, m_eLocalFeatures) Then
        GoTo QH
    End If
    '--- success
    pvHandleConnect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvHandleReceive(Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvHandleReceive"
    Dim baRecv()        As Byte
    Dim bResult         As Boolean
    Dim lPrevSize       As Long
    Dim sError          As String
    
    On Error GoTo EH
    bResult = True
    lPrevSize = m_lRecvPos
    Do
        If TlsIsClosed(m_uCtx) Then
            Exit Do
        End If
        If Not m_oSocket.ReceiveArray(baRecv) Then
            Exit Do
        End If
        Do While Not TlsIsReady(m_uCtx) And pvArraySize(baRecv) > 0
            bResult = TlsHandshake(m_uCtx, baRecv, -1, m_baSendBuffer, m_lSendPos)
            If Not pvHandleSend() Then
                GoTo QH
            End If
            If Not bResult Then
                pvFireOnError LastError.Number, ucsSfdRead
            End If
            If TlsIsClosed(m_uCtx) Then
                If pvFireBeforeNotify(ucsSfdClose) Then
                    pvFireOnClose
                    pvFireAfterNotify ucsSfdClose
                End If
                GoTo QH
            End If
            If Not TlsIsReady(m_uCtx) Or Not bResult Then
                GoTo QH
            End If
            If Not m_bIsServer Then
                If (LocalFeatures And ucsTlsIgnoreServerCertificateErrors) <> 0 Then
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Will skip server certificate checks for " & RemoteHostName
                    #End If
                ElseIf RemoteCertificates Is Nothing Then
                    pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
                    pvFireOnError vbObjectError, ucsSfdRead
                    GoTo QH
                ElseIf Not pvPkiCertValidate(RemoteHostName, RemoteCertificates, RemoteCertStatuses, frRootStore, sError) Then
                    pvSetLastError vbObjectError, MODULE_NAME & ".pvPkiCertValidate", sError
                    pvFireOnError vbObjectError, ucsSfdRead
                    GoTo QH
                End If
            End If
            If pvFireBeforeNotify(ucsSfdConnect) Then
                pvFireOnConnect
                pvFireAfterNotify ucsSfdConnect
            End If
            If Not m_bIsServer Then
                If pvFireBeforeNotify(ucsSfdWrite) Then
                    pvFireOnSend
                    pvFireAfterNotify ucsSfdWrite
                End If
            End If
            bResult = m_oSocket.ReceiveArray(baRecv)
            If Not bResult Then
                Exit Do
            End If
        Loop
        bResult = TlsReceive(m_uCtx, baRecv, -1, m_baRecvBuffer, m_lRecvPos, m_baSendBuffer, m_lSendPos)
        Call pvHandleSend
        If Not bResult Then
            Exit Do
        End If
    Loop While Flush And pvArraySize(baRecv) > 0
    If m_lRecvPos > lPrevSize Then
        If pvFireBeforeNotify(ucsSfdRead) Then
            pvFireOnReceive
            pvFireAfterNotify ucsSfdRead
        End If
    End If
    If Not bResult Then
        pvFireOnError LastError.Number, ucsSfdRead
        GoTo QH
    End If
    '--- success
    pvHandleReceive = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvHandleSend() As Boolean
    Const FUNC_NAME     As String = "pvHandleSend"
    Dim lBytes          As Long
    Dim lBufferSize     As Long
    
    On Error GoTo EH
    lBufferSize = m_oSocket.SockOpt(ucsSsoSendBuffer)
    Do While m_lSendActual < m_lSendPos
        lBytes = IIf(m_lSendPos - m_lSendActual > lBufferSize, lBufferSize, m_lSendPos - m_lSendActual)
        lBytes = m_oSocket.Send(VarPtr(m_baSendBuffer(m_lSendActual)), lBytes)
        If m_oSocket.HasPendingEvent Then
            Exit Do
        ElseIf lBytes < 0 Then
            GoTo QH
        Else
            m_lSendActual = m_lSendActual + lBytes
            m_eRaisedEvent = m_eRaisedEvent Or ucsSfdWrite
        End If
    Loop
    If m_lSendActual > 0 Then
        If m_lSendActual >= m_lSendPos Then
            m_lSendActual = 0
            m_lSendPos = 0
        End If
    End If
    '--- success
    pvHandleSend = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvHandleShutdown() As Boolean
    Const FUNC_NAME     As String = "pvHandleShutdown"
    
    On Error GoTo EH
    If Not TlsShutdown(m_uCtx, m_baSendBuffer, m_lSendPos) Then
        pvFireOnError LastError.Number, ucsSfdWrite
        GoTo QH
    End If
    If Not pvHandleSend() Then
        GoTo QH
    End If
    '--- success
    pvHandleShutdown = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvSetLastError( _
            Optional ByVal ErrNumber As Long, _
            Optional ErrSource As String, _
            Optional ErrDescription As String)
    m_lLastErrNumber = ErrNumber
    m_sLastErrSource = ErrSource
    If LenB(ErrDescription) <> 0 Then
        m_sLastError = ErrDescription
    Else
        m_sLastError = GetErrorDescription(ErrNumber)
    End If
    '--- ToDo: use TlsSetLastError
    m_uCtx.LastErrNumber = m_lLastErrNumber
    m_uCtx.LastErrSource = m_sLastErrSource
    m_uCtx.LastError = m_sLastError
End Function

Private Sub pvFireOnResolve(IpAddress As String)
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    m_eRaisedEvent = m_eRaisedEvent Or [_ucsSfdResolve]
    RaiseEvent OnResolve(IpAddress)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnResolve(IpAddress)
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnAccept()
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    RaiseEvent OnAccept
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnAccept
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnClose()
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    m_eRaisedEvent = m_eRaisedEvent Or ucsSfdClose
    RaiseEvent OnClose
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnClose
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnConnect()
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    m_eRaisedEvent = m_eRaisedEvent Or ucsSfdConnect
    RaiseEvent OnConnect
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnConnect
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnReceive()
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    m_eRaisedEvent = m_eRaisedEvent Or ucsSfdRead
    RaiseEvent OnReceive
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnReceive
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnSend()
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    m_eRaisedEvent = m_eRaisedEvent Or ucsSfdWrite
    RaiseEvent OnSend
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnSend
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Dim oCallback       As Object
    
    With m_oSocket.LastError
        If .Number <> 0 Then
            '--- ToDo: use TlsSetLastError
            m_uCtx.LastErrNumber = .Number
            m_uCtx.LastErrSource = .Source
            m_uCtx.LastError = .Description
        End If
    End With
    Set oCallback = CallbackWeakRef
    RaiseEvent OnError(ErrorCode, EventMask)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnError(ErrorCode, EventMask)
        On Error GoTo 0
    End If
End Sub

Private Sub pvFireOnMessagePending(Handled As Boolean)
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    RaiseEvent OnMessagePending(Handled)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnMessagePending(Handled)
        On Error GoTo 0
    End If
End Sub

Private Function pvFireBeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum) As Boolean
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    RaiseEvent BeforeNotify(EventMask, pvFireBeforeNotify)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyBeforeNotify(EventMask, pvFireBeforeNotify)
        On Error GoTo 0
    End If
    pvFireBeforeNotify = Not pvFireBeforeNotify
End Function

Private Sub pvFireAfterNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    RaiseEvent AfterNotify(EventMask)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyAfterNotify(EventMask)
        On Error GoTo 0
    End If
End Sub

Public Function FireOnCertificate(Issuers As Collection) As Boolean
Attribute FireOnCertificate.VB_MemberFlags = "40"
    Dim oCallback       As Object
    
    Set oCallback = CallbackWeakRef
    RaiseEvent OnCertificate(Issuers, FireOnCertificate)
    If Not oCallback Is Nothing Then
        On Error Resume Next
        Call oCallback.NotifyOnCertificate(Issuers, FireOnCertificate)
        On Error GoTo 0
    End If
End Function

'= callbacks =============================================================

Public Sub NotifyOnConnect()
Attribute NotifyOnConnect.VB_MemberFlags = "40"
    m_oAsyncSocket_OnConnect
End Sub

Public Sub NotifyOnReceive()
Attribute NotifyOnReceive.VB_MemberFlags = "40"
    m_oAsyncSocket_OnReceive
End Sub

Public Sub NotifyOnSend()
Attribute NotifyOnSend.VB_MemberFlags = "40"
    m_oAsyncSocket_OnSend
End Sub

Public Sub NotifyOnAccept()
Attribute NotifyOnAccept.VB_MemberFlags = "40"
    m_oAsyncSocket_OnAccept
End Sub

Public Sub NotifyOnClose()
Attribute NotifyOnClose.VB_MemberFlags = "40"
    m_oAsyncSocket_OnClose
End Sub

Public Sub NotifyOnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
Attribute NotifyOnError.VB_MemberFlags = "40"
    m_oAsyncSocket_OnError ErrorCode, EventMask
End Sub

Public Sub NotifyOnMessagePending(Handled As Boolean)
Attribute NotifyOnMessagePending.VB_MemberFlags = "40"
    m_oAsyncSocket_OnMessagePending Handled
End Sub

Public Sub NotifyBeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Cancel As Boolean)
Attribute NotifyBeforeNotify.VB_MemberFlags = "40"
    m_oAsyncSocket_BeforeNotify EventMask, Cancel
End Sub

Public Sub NotifyAfterNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
Attribute NotifyAfterNotify.VB_MemberFlags = "40"
    m_oAsyncSocket_AfterNotify EventMask
End Sub

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oAsyncSocket_OnConnect()
    Const FUNC_NAME     As String = "m_oAsyncSocket_OnConnect"
    
    On Error GoTo EH
    If m_bUseTls Then
        If Not pvHandleConnect() Then
            GoTo QH
        End If
    Else
        pvFireOnConnect
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oAsyncSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oAsyncSocket_OnReceive"
    
    On Error GoTo EH
    If m_bUseTls Then
        If m_lRecvPos = 0 Or TlsIsShutdown(m_uCtx) Then
            If Not pvHandleReceive() Then
                GoTo QH
            End If
        End If
    Else
        pvFireOnReceive
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oAsyncSocket_OnSend()
    Const FUNC_NAME     As String = "m_oAsyncSocket_OnSend"
    
    On Error GoTo EH
    If m_bUseTls Then
        If Not TlsIsReady(m_uCtx) Then
            GoTo QH
        End If
        If Not pvHandleSend() Then
            GoTo QH
        End If
        If m_lSendPos = 0 Then
            m_lLastSendBytes = m_lSendBytes
            m_lSendBytes = 0
            If pvFireBeforeNotify(ucsSfdWrite) Then
                pvFireOnSend
                pvFireAfterNotify ucsSfdWrite
            End If
        End If
    Else
        pvFireOnSend
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oAsyncSocket_OnAccept()
    Const FUNC_NAME     As String = "m_oAsyncSocket_OnAccept"
    
    On Error GoTo EH
    If m_bUseTls Then
        If pvFireBeforeNotify(ucsSfdAccept) Then
            pvFireOnAccept
            pvFireAfterNotify ucsSfdAccept
        End If
    Else
        pvFireOnAccept
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oAsyncSocket_OnClose()
    Const FUNC_NAME     As String = "m_oAsyncSocket_OnClose"
    
    On Error GoTo EH
    If m_bUseTls Then
        If Not pvHandleReceive(Flush:=True) Then
            '--- do nothing
        End If
        If pvFireBeforeNotify(ucsSfdClose) Then
            pvFireOnClose
            TlsTerminate m_uCtx
            pvFireAfterNotify ucsSfdClose
        End If
    Else
        pvFireOnClose
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oAsyncSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    pvFireOnError ErrorCode, EventMask
End Sub

Private Sub m_oAsyncSocket_OnMessagePending(Handled As Boolean)
    pvFireOnMessagePending Handled
End Sub

Private Sub m_oAsyncSocket_OnResolve(IpAddress As String)
    pvFireOnResolve IpAddress
End Sub

Private Sub m_oAsyncSocket_BeforeNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum, Cancel As Boolean)
    If Not m_bUseTls Then
        Cancel = Not pvFireBeforeNotify(EventMask)
    End If
End Sub

Private Sub m_oAsyncSocket_AfterNotify(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oAsyncSocket_AfterNotify"
    
    On Error GoTo EH
    If m_bUseTls Then
        If EventMask = ucsSfdRead And AvailableBytes = 0 Then
            m_oSocket.Receive 0, 0 '--- enable FD_READ notification
        End If
    Else
        pvFireAfterNotify EventMask
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_oSocket = New cAsyncSocket
    Set m_oAsyncSocket = m_oSocket
End Sub

Private Sub Class_Terminate()
    Set m_oSocket = Nothing
    Set m_oAsyncSocket = Nothing
    If m_hRootStore <> 0 Then
        Call CertCloseStore(m_hRootStore, 0)
    End If
End Sub

'=========================================================================
' PKI
'=========================================================================

Public Function ImportSystemStoreCertificates(CertSubject As Variant, Optional ByVal hWndOwner As Long) As Boolean
    Const FUNC_NAME     As String = "ImportSystemStoreCertificates"
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection
    
    pvSetLastError
    If pvPkiSystemStoreImportCertificate(CertSubject, hWndOwner, cCerts, cPrivKey) Then
        Set LocalCertificates = cCerts
        Set LocalPrivateKey = cPrivKey
        '--- success
        ImportSystemStoreCertificates = True
    Else
        If m_lLastErrNumber = 0 Then
            pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
        End If
    End If
End Function

Public Function ImportPemCertificates(CertFile As String, Optional Password As String) As Boolean
#If Not ImplTlsServer Then
    ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
#Else
    Const FUNC_NAME     As String = "ImportPemCertificates"
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection

    pvSetLastError
    If pvPkiPemImportCertificates(Split(CertFile, "|"), Password, cCerts, cPrivKey) Then
        Set LocalCertificates = cCerts
        Set LocalPrivateKey = cPrivKey
        '--- success
        ImportPemCertificates = True
    Else
        If m_lLastErrNumber = 0 Then
            pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
        End If
    End If
#End If
End Function

Public Function ImportPkcs12Certificates(CertFile As String, Optional Password As String) As Boolean
#If Not ImplTlsServer Then
    ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
#Else
    Const FUNC_NAME     As String = "ImportPkcs12Certificates"
    Dim cCerts          As Collection
    Dim cPrivKey        As Collection

    pvSetLastError
    If pvPkiPkcs12ImportCertificates(CertFile, Password & "", cCerts, cPrivKey) Then
        Set LocalCertificates = cCerts
        Set LocalPrivateKey = cPrivKey
        '--- success
        ImportPkcs12Certificates = True
    Else
        If m_lLastErrNumber = 0 Then
            pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_NO_CERTIFICATE
        End If
    End If
#End If
End Function

Public Function ImportPemRootCaCertStore(Optional CaBundleFile As String, Optional RootCa As cTlsSocket) As Boolean
#If Not ImplTlsServer Then
    ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
#Else
    Dim hCertStore      As Long
    
    pvSetLastError
    If LenB(CaBundleFile) <> 0 Then
        hCertStore = pvPkiPemImportRootCaCertStore(CaBundleFile)
        If hCertStore = 0 Then
            GoTo QH
        End If
        If m_hRootStore <> 0 Then
            Call CertCloseStore(m_hRootStore, 0)
        End If
        m_hRootStore = hCertStore
        Set m_oRootCa = Nothing
    Else
        Set m_oRootCa = RootCa
        If m_hRootStore <> 0 Then
            Call CertCloseStore(m_hRootStore, 0)
            m_hRootStore = 0
        End If
    End If
    '--- success
    ImportPemRootCaCertStore = True
QH:
#End If
End Function

'= private ===============================================================

#If ImplTlsServer Then
Private Function pvPkiPemImportCertificates(ByVal vPemFiles As Variant, sPassword As String, cCerts As Collection, cPrivKey As Collection) As Boolean
    Dim vElem           As Variant
    Dim sPemText        As String
    Dim vKey            As Variant
    Dim baPrivKey()     As Byte
    Dim baEncrKey()     As Byte
    
    If VarType(vPemFiles) = vbString Then
        vPemFiles = Array(vPemFiles)
    End If
    For Each vElem In vPemFiles
        sPemText = StrConv(CStr(ReadBinaryFile(CStr(vElem))), vbUnicode)
        pvPkiPemGetTextPortions sPemText, "PRIVATE KEY", cPrivKey
        #If ImplEncryptedPrivateKey Then
            For Each vKey In pvPkiPemGetTextPortions(sPemText, "ENCRYPTED PRIVATE KEY")
                baEncrKey = vKey
                If Not pvPkiPemDecryptPrivateKey(baPrivKey, baEncrKey, sPassword) Then
                    GoTo QH
                End If
                cPrivKey.Add baPrivKey
            Next
        #End If
        pvPkiPemGetTextPortions sPemText, "RSA PRIVATE KEY", cPrivKey
        pvPkiPemGetTextPortions sPemText, "EC PRIVATE KEY", cPrivKey
        pvPkiPemGetTextPortions sPemText, "CERTIFICATE", cCerts
    Next
    If pvCollectionCount(cPrivKey) > 0 Then
        '--- success
        pvPkiPemImportCertificates = True
    End If
QH:
End Function
#End If

#If ImplTlsServer And ImplEncryptedPrivateKey Then
Private Function pvPkiPemDecryptPrivateKey(baRetVal() As Byte, baEncrKey() As Byte, sPassword As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiPemDecryptPrivateKey"
    Const szOID_PKCS_5_PBES2        As String = "1.2.840.113549.1.5.13"
    Const szOID_PKCS_5_PBKDF2       As String = "1.2.840.113549.1.5.12"
    Const szOID_RSA_HMAC_SHA1       As String = "1.2.840.113549.2"
    Const szOID_RSA_HMAC_SHA256     As String = "1.2.840.113549.2.9"
    Const szOID_RSA_HMAC_SHA512     As String = "1.2.840.113549.2.11"
    Const szOID_RSA_DES_EDE3_CBC    As String = "1.2.840.113549.3.7"
    Const szOID_NIST_AES128_CBC     As String = "2.16.840.1.101.3.4.1.2"
    Const szOID_NIST_AES256_CBC     As String = "2.16.840.1.101.3.4.1.42"
    Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
    Const BCRYPT_BLOCK_PADDING      As Long = 1
    Const IDX_KDF_OID               As Long = 1
    Const IDX_KDF_SALT              As Long = 2
    Const IDX_KDF_ITER              As Long = 3
    Const IDX_KDF_HASH_OID          As Long = 4
    Const IDX_KDF_ENCR_OID          As Long = 5
    Const IDX_KDF_ENCR_IV           As Long = 6
    Dim lEncrPtr        As Long
    Dim uEncrInfo       As CRYPT_ENCRYPTED_PRIVATE_KEY_INFO
    Dim sOID            As String
    Dim cParams         As Collection
    Dim sHashAlg        As String
    Dim baPass()        As Byte
    Dim lPassPtr        As Long
    Dim baSalt()        As Byte
    Dim lSaltPtr        As Long
    Dim lIterations     As Long
    Dim sEncrAlg        As String
    Dim lKeySize        As Long
    Dim baIV()          As Byte
    Dim baKey()         As Byte
    Dim hAlg            As Long
    Dim lSize           As Long
    Dim baObj()         As Byte
    Dim hKey            As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_ENCRYPTED_PRIVATE_KEY_INFO, baEncrKey(0), UBound(baEncrKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lEncrPtr, 0) = 0 Then
        pvSetLastError Err.LastDllError, MODULE_NAME & "." & FUNC_NAME
        GoTo QH
    End If
    Call CopyMemory(uEncrInfo, ByVal lEncrPtr, LenB(uEncrInfo))
    sOID = pvToStringA(uEncrInfo.EncryptionAlgorithm.pszObjId)
    If sOID <> szOID_PKCS_5_PBES2 Then
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(Replace(ERR_UNKNOWN_OID, "%1", "encryption scheme"), "%2", sOID)
        GoTo QH
    End If
    Set cParams = pvPkiAsn1Scanf(uEncrInfo.EncryptionAlgorithm.Parameters.pbData, uEncrInfo.EncryptionAlgorithm.Parameters.cbData, "((o(xi(o)))(ox))")
    sOID = cParams.Item(IDX_KDF_OID)
    If sOID <> szOID_PKCS_5_PBKDF2 Then
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(Replace(ERR_UNKNOWN_OID, "%1", "key derivation function"), "%2", sOID)
        GoTo QH
    End If
    sOID = cParams.Item(IDX_KDF_HASH_OID)
    Select Case sOID
    Case szOID_RSA_HMAC_SHA1
        sHashAlg = "SHA1"
    Case szOID_RSA_HMAC_SHA256
        sHashAlg = "SHA256"
    Case szOID_RSA_HMAC_SHA512
        sHashAlg = "SHA512"
    Case Else
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(Replace(ERR_UNKNOWN_OID, "%1", "hash"), "%2", sOID)
        GoTo QH
    End Select
    sOID = cParams.Item(IDX_KDF_ENCR_OID)
    Select Case sOID
    Case szOID_RSA_DES_EDE3_CBC
        sEncrAlg = "3DES"
        lKeySize = 3 * 8
    Case szOID_NIST_AES128_CBC
        sEncrAlg = "AES"
        lKeySize = 16
    Case szOID_NIST_AES256_CBC
        sEncrAlg = "AES"
        lKeySize = 32
    Case Else
        pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, Replace(Replace(ERR_UNKNOWN_OID, "%1", "encryption"), "%2", sOID)
        GoTo QH
    End Select
    baPass = StrConv(sPassword, vbFromUnicode)
    baSalt = cParams.Item(IDX_KDF_SALT)
    lIterations = cParams.Item(IDX_KDF_ITER)
    baIV = cParams.Item(IDX_KDF_ENCR_IV)
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr(sHashAlg), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider(" & sHashAlg & ")"
        GoTo QH
    End If
    If UBound(baPass) >= 0 Then
        lPassPtr = VarPtr(baPass(0))
    End If
    If UBound(baSalt) >= 0 Then
        lSaltPtr = VarPtr(baSalt(0))
    End If
    pvArrayAllocate baKey, lKeySize, FUNC_NAME & ".baKey"
    hResult = BCryptDeriveKeyPBKDF2(hAlg, ByVal lPassPtr, UBound(baPass) + 1, ByVal lSaltPtr, UBound(baSalt) + 1, lIterations / 10000@, baKey(0), lKeySize, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    Call BCryptCloseAlgorithmProvider(hAlg, 0)
    hAlg = 0
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr(sEncrAlg), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider(" & sEncrAlg & ")"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ChainingMode"), StrPtr("ChainingModeCBC"), LenB("ChainingModeCBC"), 0)
    If hResult < 0 Then
        sApiSource = "BCryptSetProperty(ChainingMode)"
        GoTo QH
    End If
    hResult = BCryptGetProperty(hAlg, StrPtr("ObjectLength"), lSize, LenB(lSize), 0, 0)
    If hResult < 0 Or lSize = 0 Then
        sApiSource = "BCryptGetProperty(ObjectLength)"
        GoTo QH
    End If
    pvArrayAllocate baObj, lSize, FUNC_NAME & ".baObj"
    hResult = BCryptGenerateSymmetricKey(hAlg, hKey, baObj(0), UBound(baObj) + 1, baKey(0), UBound(baKey) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptGenerateSymmetricKey"
        GoTo QH
    End If
    pvArrayAllocate baRetVal, uEncrInfo.EncryptedPrivateKey.cbData, FUNC_NAME & ".baRetVal"
    hResult = BCryptDecrypt(hKey, ByVal uEncrInfo.EncryptedPrivateKey.pbData, uEncrInfo.EncryptedPrivateKey.cbData, 0, baIV(0), UBound(baIV) + 1, baRetVal(0), UBound(baRetVal) + 1, lSize, BCRYPT_BLOCK_PADDING)
    If hResult < 0 Then
        If hResult = LNG_STATUS_DATA_ERROR Then
            pvSetLastError vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_INCORRECT_PASSWORD
            GoTo QH
        End If
        sApiSource = "BCryptDecrypt"
        GoTo QH
    End If
    pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    '--- success
    pvPkiPemDecryptPrivateKey = True
QH:
    On Error GoTo 0
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    If lEncrPtr <> 0 Then
        Call LocalFree(lEncrPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function
#End If

#If ImplTlsServer Then
Private Function pvPkiPemImportRootCaCertStore(sPemFile As String) As Long
    Const FUNC_NAME     As String = "pvPkiPemImportRootCaCertStore"
    Dim hCertStore      As Long
    Dim cCerts          As Collection
    Dim vElem           As Variant
    Dim baCert()        As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    Set cCerts = pvPkiPemGetTextPortions(StrConv(CStr(ReadBinaryFile(sPemFile)), vbUnicode), "CERTIFICATE")
    If pvCollectionCount(cCerts) = 0 Then
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
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, ByVal 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    '--- commit
    pvPkiPemImportRootCaCertStore = hCertStore
    hCertStore = 0
QH:
    On Error GoTo 0
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function
#End If

#If ImplTlsServer Then
Private Function pvPkiPkcs12ImportCertificates(sPfxFile As String, sPassword As String, cCerts As Collection, cPrivKey As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiPkcs12ImportCertificates"
    Dim baPfx()         As Byte
    Dim uBlob           As CRYPT_DATA_BLOB
    Dim dwFlags         As Long
    Dim hPfxStore       As Long
    Dim pCertContext    As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    baPfx = ReadBinaryFile(sPfxFile)
    If UBound(baPfx) < 0 Then
        GoTo QH
    End If
    uBlob.cbData = UBound(baPfx) + 1
    uBlob.pbData = VarPtr(baPfx(0))
    dwFlags = CRYPT_EXPORTABLE Or IIf(OsVersion <= ucsOsvXp, 0, PKCS12_NO_PERSIST_KEY)
    hPfxStore = PFXImportCertStore(uBlob, StrPtr(sPassword), dwFlags)
    If hPfxStore = 0 And Err.LastDllError <> NTE_BAD_ALGID Then
        hPfxStore = PFXImportCertStore(baPfx(0), 0, dwFlags)
    End If
    If hPfxStore = 0 Then
        pvSetLastError Err.LastDllError, MODULE_NAME & "." & FUNC_NAME
        GoTo QH
    End If
    Do
        pCertContext = CertEnumCertificatesInStore(hPfxStore, pCertContext)
        If pCertContext = 0 Then
            Exit Do
        End If
        If pvPkiAppendCertContext(pCertContext, cCerts, cPrivKey, DeleteKeyset:=(dwFlags And PKCS12_NO_PERSIST_KEY) = 0) Then
            '--- success
            pvPkiPkcs12ImportCertificates = True
        End If
    Loop
QH:
    On Error GoTo 0
    If hPfxStore <> 0 Then
        Call CertCloseStore(hPfxStore, 0)
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function
#End If

Private Function pvPkiSystemStoreImportCertificate(ByVal SubjectOrIssuers As Variant, ByVal hWndParent As Long, cCerts As Collection, cPrivKey As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiSystemStoreImportCertificate"
    Dim hMemStore       As Long
    Dim hCertStore      As Long
    Dim vElem           As Variant
    Dim uCertRdn        As CERT_RDN
    Dim pCertContext    As Long
    Dim lStoreFlags     As Long
    Dim sStoreName      As String
    Dim sCertName       As String
    Dim dwFindType      As Long
    Dim pvFindPara      As Long
    Dim uSelect         As CERT_SELECT_STRUCT_W
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    If IsObject(SubjectOrIssuers) Then
        hCertStore = CertOpenStore(CERT_STORE_PROV_SYSTEM, 0, 0, CERT_SYSTEM_STORE_CURRENT_USER, StrPtr("MY"))
        If hCertStore = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertOpenStore(CERT_SYSTEM_STORE_CURRENT_USER, MY)"
            GoTo QH
        End If
        For Each vElem In SubjectOrIssuers
            If pvPkiIssuerToCertRdn(vElem, uCertRdn) Then
                dwFindType = CERT_FIND_ISSUER_ATTR
                pvFindPara = VarPtr(uCertRdn)
                pCertContext = CertFindCertificateInStore(hCertStore, X509_ASN_ENCODING, 0, dwFindType, ByVal pvFindPara, 0)
            End If
        Next
    Else
        If Not pvPkiParseSelectorString(SubjectOrIssuers, lStoreFlags, sStoreName, sCertName) Then
            GoTo QH
        End If
        hCertStore = CertOpenStore(CERT_STORE_PROV_SYSTEM, 0, 0, lStoreFlags, StrPtr(sStoreName))
        If hCertStore = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertOpenStore(&H" & Hex$(lStoreFlags) & ", " & sStoreName & ")"
            GoTo QH
        End If
        If LenB(sCertName) = 0 Then
            dwFindType = CERT_FIND_ANY
            pCertContext = CertFindCertificateInStore(hCertStore, X509_ASN_ENCODING, 0, dwFindType, ByVal pvFindPara, 0)
        Else
            pvFindPara = StrPtr(sCertName)
            dwFindType = CERT_FIND_SUBJECT_STR
            pCertContext = CertFindCertificateInStore(hCertStore, X509_ASN_ENCODING, 0, dwFindType, ByVal pvFindPara, 0)
            If pCertContext = 0 Then
                dwFindType = CERT_FIND_HASH_STR
                pCertContext = CertFindCertificateInStore(hCertStore, X509_ASN_ENCODING, 0, dwFindType, ByVal pvFindPara, 0)
            End If
        End If
    End If
    If pCertContext <> 0 Then
        If hWndParent <> 0 Then
            hMemStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
            If hMemStore = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CertOpenStore(CERT_STORE_PROV_MEMORY)"
                GoTo QH
            End If
        End If
        Do While pCertContext <> 0
            '--- check if (persistent) private key is availabe for the certificate
            If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, ByVal 0, lSize) <> 0 And lSize <> 0 Then
                If hMemStore = 0 Then
                    Exit Do
                End If
                Call CertAddCertificateContextToStore(hMemStore, pCertContext, CERT_STORE_ADD_USE_EXISTING, ByVal 0)
            End If
            pCertContext = CertFindCertificateInStore(hCertStore, X509_ASN_ENCODING, 0, dwFindType, ByVal pvFindPara, pCertContext)
        Loop
        If hWndParent <> 0 Then
            With uSelect
                .dwSize = LenB(uSelect)
                .hWndParent = hWndParent
                .cCertStore = 1
                .arrayCertStore = VarPtr(hMemStore)
                .cCertContext = 1
                pCertContext = CertEnumCertificatesInStore(hMemStore, 0)
                .arrayCertContext = VarPtr(pCertContext)
            End With
            If CertSelectCertificate(uSelect) = 0 Or pCertContext = 0 Then
                hResult = Err.LastDllError
                If hResult < 0 Then
                    sApiSource = "CertSelectCertificate"
                End If
                GoTo QH
            End If
        End If
        If pvPkiAppendCertChain(pCertContext, hCertStore, cCerts, cPrivKey) Then
            '--- success
            pvPkiSystemStoreImportCertificate = True
        End If
    End If
QH:
    On Error GoTo 0
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If hMemStore <> 0 Then
        Call CertCloseStore(hMemStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

#If ImplTlsServer Then
Private Function pvPkiGenerSelfSignedCertificate(cCerts As Collection, cPrivKey As Collection, Optional ByVal Subject As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiGenerSelfSignedCertificate"
    Dim hProv           As Long
    Dim hKey            As Long
    Dim sName           As String
    Dim baName()        As Byte
    Dim lSize           As Long
    Dim uName           As CRYPT_DATA_BLOB
    Dim uProvInfo       As CRYPT_KEY_PROV_INFO
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    With uProvInfo
        .dwProvType = PROV_RSA_FULL
        .dwFlags = CRYPT_VERIFYCONTEXT
        .dwKeySpec = AT_KEYEXCHANGE
    End With
    If CryptAcquireContext(hProv, uProvInfo.pwszContainerName, uProvInfo.pwszProvName, uProvInfo.dwProvType, uProvInfo.dwFlags) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH
    End If
    If CryptGenKey(hProv, uProvInfo.dwKeySpec, RSA1024BIT_KEY Or CRYPT_EXPORTABLE, hKey) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptGenKey"
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
        hResult = Err.LastDllError
        sApiSource = "CertStrToName"
        GoTo QH
    End If
    pvArrayAllocate baName, lSize, FUNC_NAME & ".baName"
    If CertStrToName(X509_ASN_ENCODING, StrPtr(sName), CERT_OID_NAME_STR, 0, baName(0), lSize, 0) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertStrToName#2"
        GoTo QH
    End If
    With uName
        .cbData = lSize
        .pbData = VarPtr(baName(0))
    End With
    pCertContext = CertCreateSelfSignCertificate(hProv, uName, 0, uProvInfo, 0, ByVal 0, ByVal 0, 0)
    If pCertContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateSelfSignCertificate"
        GoTo QH
    End If
    '--- note: store temp key container to be able to call CryptGetUserKey in pvPkiExportPrivateKey
    If CertSetCertificateContextProperty(pCertContext, CERT_KEY_PROV_HANDLE_PROP_ID, 0, ByVal hProv) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertSetCertificateContextProperty(CERT_KEY_PROV_HANDLE_PROP_ID)"
        GoTo QH
    End If
    If pvPkiAppendCertContext(pCertContext, cCerts, cPrivKey) Then
        '--- success
        pvPkiGenerSelfSignedCertificate = True
    End If
QH:
    On Error GoTo 0
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function
#End If

Private Function pvPkiCertValidate(ByVal sHostName As String, cCerts As Collection, cStatuses As Collection, ByVal hRootStore As Long, sError As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiCertValidate"
    Dim hCertStore      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim pCertContext    As Long
    Dim vElem           As Variant
    Dim sDnsName        As String
    Dim bMatched        As Boolean
    Dim uProp           As CRYPT_DATA_BLOB
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    '--- load server X.509 certificates to an in-memory certificate store
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For lIdx = 1 To pvCollectionCount(cCerts)
        baCert = cCerts.Item(lIdx)
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, pCertContext) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
        If SearchCollection(cStatuses, lIdx, RetVal:=baCert) Then
            If pvArraySize(baCert) > 0 Then
                uProp.cbData = pvArraySize(baCert)
                uProp.pbData = VarPtr(baCert(0))
                If CertSetCertificateContextProperty(pCertContext, CERT_OCSP_RESPONSE_PROP_ID, CERT_SET_PROPERTY_IGNORE_PERSIST_ERROR_FLAG, uProp) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CertSetCertificateContextProperty(CERT_OCSP_RESPONSE_PROP_ID)"
                    GoTo QH
                End If
            End If
        End If
        Call CertFreeCertificateContext(pCertContext)
        pCertContext = 0
    Next
    '--- search remote host FQDN in any X.509 certificate's "Subject Alternative Name" list of DNS names (incl. wildcards)
    sHostName = LCase$(sHostName)
    sError = vbNullString
    Do
        pCertContext = CertEnumCertificatesInStore(hCertStore, pCertContext)
        If pCertContext = 0 Then
            If LenB(sError) = 0 Then
                sError = Replace(ERR_NO_MATCHING_ALT_NAME, "%1", sHostName)
            End If
            Exit Do
        End If
        bMatched = False
        For Each vElem In pvPkiCertGetSubjectAltName2(pCertContext, CERT_ALT_NAME_DNS_NAME)
            sDnsName = LCase$(vElem)
            If Left$(sDnsName, 1) = "*" Then
                If sHostName Like sDnsName And Not sHostName Like "*." & sDnsName Then
                    bMatched = True
                    Exit For
                End If
            Else
                If sHostName = sDnsName Or "www." & sHostName = sDnsName Then
                    bMatched = True
                    Exit For
                End If
            End If
        Next
        If bMatched Then
            If pvPkiCertBuildChain(pCertContext, hRootStore, sError) Then
                '--- success
                pvPkiCertValidate = True
                Exit Do
            End If
        End If
    Loop
QH:
    On Error GoTo 0
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiCertGetSubjectAltName2(ByVal pCertContext As Long, ByVal lNameForm As Long) As Variant
    Dim lIdx            As Long
    Dim lPtr            As Long
    Dim uInfo           As CERT_INFO
    Dim uExtension      As CERT_EXTENSION
    Dim lAltInfoPtr     As Long
    Dim uAltInfo        As CERT_ALT_NAME_INFO
    Dim uEntry          As CERT_ALT_NAME_ENTRY
    Dim vRetVal         As Variant

    vRetVal = Array()
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), LenB(lPtr))          '--- dereference pCertContext->pCertInfo->cExtension
    Call CopyMemory(uInfo, ByVal lPtr, LenB(uInfo))
    lPtr = CertFindExtension(szOID_SUBJECT_ALT_NAME2, uInfo.cExtension, uInfo.rgExtension)
    If lPtr = 0 Then
        GoTo QH
    End If
    Call CopyMemory(uExtension, ByVal lPtr, LenB(uExtension))
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, szOID_SUBJECT_ALT_NAME2, ByVal uExtension.Value.pbData, uExtension.Value.cbData, _
            CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG Or CRYPT_DECODE_SHARE_OID_STRING_FLAG, 0, lAltInfoPtr, 0) = 0 Then
        GoTo QH
    End If
    Debug.Assert lAltInfoPtr <> 0
    Call CopyMemory(uAltInfo, ByVal lAltInfoPtr, LenB(uAltInfo))
    For lIdx = 0 To uAltInfo.cAltEntry - 1
        lPtr = UnsignedAdd(uAltInfo.rgAltEntry, lIdx * LenB(uEntry))                '--- dereference lAltInfoPtr->rgAltEntry[lidx].dwAltNameChoice
        Call CopyMemory(uEntry, ByVal lPtr, LenB(uEntry))
        If uEntry.dwAltNameChoice = lNameForm Then
            ReDim Preserve vRetVal(0 To UBound(vRetVal) + 1) As Variant
            vRetVal(UBound(vRetVal)) = pvToStringW(uEntry.pwszDNSName)
        End If
    Next
    Call LocalFree(lAltInfoPtr)
    lAltInfoPtr = 0
QH:
    pvPkiCertGetSubjectAltName2 = vRetVal
End Function

Private Function pvPkiCertBuildChain(ByVal pCertContext As Long, ByVal hRootStore As Long, sError As String) As Boolean
    Const FUNC_NAME     As String = "pvPkiCertBuildChain"
    Dim hChainEngine    As Long
    Dim uEngineConfig   As CERT_CHAIN_ENGINE_CONFIG
    Dim pChainContext   As Long
    Dim lPtr            As Long
    Dim dwErrorStatus   As Long
    Dim dwFlags         As Long
    Dim uChainParams    As CERT_CHAIN_PARA
    Dim uChain          As CERT_CHAIN_CONTEXT
    Dim uChainSmall     As CERT_CHAIN_CONTEXT
    Dim uChainElem      As CERT_CHAIN_ELEMENT
    Dim pExistContext   As Long
    Dim lIdx            As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    '--- build custom chain engine that trusts the additional root CA certificates if provided
    If hRootStore <> 0 Then
        If OsVersion <= ucsOsvWin7 Then
            uEngineConfig.cbSize = LenB(uEngineConfig) - 12
            uEngineConfig.cAdditionalStore = 1
            uEngineConfig.rghAdditionalStore = VarPtr(hRootStore)
        Else
            uEngineConfig.cbSize = LenB(uEngineConfig)
            uEngineConfig.hExclusiveRoot = hRootStore
        End If
        uEngineConfig.dwFlags = CERT_CHAIN_CACHE_END_CERT
        If CertCreateCertificateChainEngine(uEngineConfig, hChainEngine) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertCreateCertificateChainEngine"
            GoTo QH
        End If
    End If
    '--- for the matched server certificate try to build a chain of certificates from the ones in the in-memory certificate store
    '---    and check this chain for revokation, expiry or missing link to a trust anchor
    uChainParams.cbSize = LenB(uChainParams)
    dwFlags = IIf((m_eLocalFeatures And ucsTlsIgnoreServerCertificateRevocation) <> 0, 0, CERT_CHAIN_REVOCATION_CHECK_CHAIN)
    If CertGetCertificateChain(hChainEngine, pCertContext, 0, 0, uChainParams, dwFlags, 0, pChainContext) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertGetCertificateChain"
        GoTo QH
    End If
    Debug.Assert pChainContext <> 0
    Call CopyMemory(uChain, ByVal pChainContext, LenB(uChain))
    For lIdx = 0 To uChain.cElems - 1
        '--- dereference pChainContext->rgpChain[lidx]->TrustStatus.dwErrorStatus
        Call CopyMemory(lPtr, ByVal UnsignedAdd(uChain.rgElem, lIdx * LenB(lPtr)), LenB(lPtr))
        Debug.Assert lPtr <> 0
        Call CopyMemory(uChainSmall, ByVal lPtr, LenB(uChainSmall))
        dwErrorStatus = uChainSmall.TrustStatus.dwErrorStatus And Not CERT_TRUST_IS_NOT_TIME_NESTED
        If hRootStore <> 0 And uChainSmall.cElems > 0 Then
            '--- check if the last certificate in the chain is from our custom hRootStore and remove untrusted flags from status
            Call CopyMemory(lPtr, ByVal UnsignedAdd(uChainSmall.rgElem, (uChainSmall.cElems - 1) * LenB(lPtr)), LenB(lPtr))
            Call CopyMemory(uChainElem, ByVal lPtr, LenB(uChainElem))
            pExistContext = CertFindCertificateInStore(hRootStore, X509_ASN_ENCODING, 0, CERT_FIND_EXISTING, ByVal uChainElem.pCertContext, 0)
            If pExistContext <> 0 Then
                Call CertFreeCertificateContext(pExistContext)
                pExistContext = 0
                dwErrorStatus = dwErrorStatus And Not CERT_TRUST_IS_UNTRUSTED_ROOT And Not CERT_TRUST_IS_NOT_SIGNATURE_VALID
            End If
        End If
        If dwErrorStatus = 0 Then
            '--- success
            pvPkiCertBuildChain = True
            Exit For
        ElseIf (dwErrorStatus And CERT_TRUST_IS_NOT_TIME_VALID) <> 0 Then
            sError = ERR_TRUST_IS_NOT_TIME_VALID
        ElseIf (dwErrorStatus And CERT_TRUST_IS_REVOKED) <> 0 Then
            sError = ERR_TRUST_IS_REVOKED
        ElseIf (dwErrorStatus And CERT_TRUST_IS_NOT_SIGNATURE_VALID) <> 0 Then
            sError = ERR_TRUST_IS_NOT_SIGNATURE_VALID
        ElseIf (dwErrorStatus And CERT_TRUST_IS_UNTRUSTED_ROOT) <> 0 Then
            sError = ERR_TRUST_IS_UNTRUSTED_ROOT
        ElseIf (dwErrorStatus And CERT_TRUST_REVOCATION_STATUS_UNKNOWN) <> 0 Then
            sError = ERR_TRUST_REVOCATION_STATUS_UNKNOWN
        ElseIf (dwErrorStatus And CERT_TRUST_IS_PARTIAL_CHAIN) <> 0 Then
            sError = ERR_TRUST_IS_PARTIAL_CHAIN
        Else
            sError = Replace(ERR_UNKNOWN_CERTIFICATECHAIN_ERROR_MASK, "%1", "&H" & Hex$(dwErrorStatus))
        End If
    Next
QH:
    On Error GoTo 0
    If pChainContext <> 0 Then
        Call CertFreeCertificateChain(pChainContext)
    End If
    If hChainEngine <> 0 Then
        Call CertFreeCertificateChainEngine(hChainEngine)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

#If ImplTlsServer Then
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
#End If

Private Function pvPkiAppendCertChain(ByVal pCertContext As Long, ByVal hCertStore As Long, cCerts As Collection, cPrivKey As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiAppendCertChain"
    Dim pChainContext   As Long
    Dim uChain          As CERT_CHAIN_CONTEXT
    Dim uChainParams    As CERT_CHAIN_PARA
    Dim uChainElem      As CERT_CHAIN_ELEMENT
    Dim lPtr            As Long
    Dim lIdx            As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant

    On Error GoTo EH
    uChainParams.cbSize = LenB(uChainParams)
    If CertGetCertificateChain(0, pCertContext, 0, hCertStore, uChainParams, 0, 0, pChainContext) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertGetCertificateChain"
        GoTo QH
    End If
    Debug.Assert pChainContext <> 0
    Call CopyMemory(uChain, ByVal pChainContext, LenB(uChain))                   '--- dereference pChainContext->rgpChain[0]
    Call CopyMemory(lPtr, ByVal uChain.rgElem, LenB(lPtr))
    Debug.Assert lPtr <> 0
    Call CopyMemory(uChain, ByVal lPtr, LenB(uChain))
    For lIdx = 0 To IIf(uChain.cElems = 1, 0, uChain.cElems - 2)
        Call CopyMemory(lPtr, ByVal UnsignedAdd(uChain.rgElem, lIdx * LenB(lPtr)), LenB(lPtr))
        Call CopyMemory(uChainElem, ByVal lPtr, LenB(uChainElem))
        If pvPkiAppendCertContext(uChainElem.pCertContext, cCerts, cPrivKey) Then
            '--- success
            pvPkiAppendCertChain = True
        End If
    Next
QH:
    On Error GoTo 0
    If pChainContext <> 0 Then
        Call CertFreeCertificateChain(pChainContext)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiAppendCertContext(ByVal pCertContext As Long, cCerts As Collection, cPrivKey As Collection, Optional ByVal DeleteKeyset As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvPkiAppendCertContext"
    Dim uCertContext    As CERT_CONTEXT
    Dim baBuffer()      As Byte
    
    Debug.Assert pCertContext <> 0
    Call CopyMemory(uCertContext, ByVal pCertContext, LenB(uCertContext))
    If uCertContext.cbCertEncoded > 0 Then
        pvArrayAllocate baBuffer, uCertContext.cbCertEncoded, FUNC_NAME & ".baBuffer"
        Debug.Assert uCertContext.pbCertEncoded <> 0
        Call CopyMemory(baBuffer(0), ByVal uCertContext.pbCertEncoded, uCertContext.cbCertEncoded)
        If cCerts Is Nothing Then
            Set cCerts = New Collection
        End If
        cCerts.Add baBuffer
    End If
    If pvPkiExportPrivateKey(pCertContext, DeleteKeyset, cPrivKey) Then
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

Private Function pvPkiExportPrivateKey(ByVal pCertContext As Long, ByVal bDeleteKeyset As Boolean, cPrivKey As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportPrivateKey"
    Dim baPrivKey()     As Byte
    Dim dwFlags         As Long
    Dim hProv           As Long
    Dim lKeySpec        As Long
    Dim lFree           As Long
    Dim hNProv          As Long
    Dim hNKey           As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim uProvInfo       As CRYPT_KEY_PROV_INFO
    Dim hProvGetUserKey As Long
    Dim hKey            As Long
    Dim lMagic          As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant

    On Error GoTo EH
    dwFlags = CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG Or CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG
    If CryptAcquireCertificatePrivateKey(pCertContext, dwFlags, 0, hProv, lKeySpec, lFree) = 0 Then
        dwFlags = dwFlags Or CRYPT_ACQUIRE_CACHE_FLAG
        If CryptAcquireCertificatePrivateKey(pCertContext, dwFlags, 0, hProv, lKeySpec, lFree) = 0 Then
            pvSetLastError Err.LastDllError, MODULE_NAME & "." & FUNC_NAME
            GoTo QH
        End If
    End If
    Set cPrivKey = New Collection
    If lKeySpec < 0 Then
        hNKey = pvPkiCloneKeyWithExportPolicy(hProv, NCRYPT_ALLOW_EXPORT_FLAG Or NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG)
        If hNKey = 0 Then
            Debug.Assert (dwFlags And CRYPT_ACQUIRE_CACHE_FLAG) = 0
            cPrivKey.Add pvNCryptGetStringProperty(hProv, "Name")
            hResult = NCryptGetProperty(hProv, StrPtr("Provider Handle"), hNProv, LenB(hNProv), 0, 0)
            Debug.Assert Not hResult < 0
            cPrivKey.Add pvNCryptGetStringProperty(hNProv, "Name")
        Else
            hResult = NCryptExportKey(hNKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(PRIVATEBLOB)"
                GoTo QH
            End If
            pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
            hResult = NCryptExportKey(hNKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(PRIVATEBLOB)#2"
                GoTo QH
            End If
            Debug.Assert UBound(baBuffer) + 1 >= LenB(lMagic)
            Call CopyMemory(lMagic, baBuffer(0), LenB(lMagic))
            Select Case lMagic
            Case BCRYPT_RSAPRIVATE_MAGIC
                hResult = NCryptExportKey(hNKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
                If hResult < 0 Then
                    sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)"
                    GoTo QH
                End If
                pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
                hResult = NCryptExportKey(hNKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
                If hResult < 0 Then
                    sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)#2"
                    GoTo QH
                End If
                If Not pvPkiExportRsaPrivateKey(baPrivKey, baBuffer, CNG_RSA_PRIVATE_KEY_BLOB) Then
                    GoTo QH
                End If
            Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_P521_MAGIC
                Debug.Assert UBound(baBuffer) + 1 >= 8
                Call CopyMemory(lSize, baBuffer(4), LenB(lSize))
                Debug.Assert UBound(baBuffer) + 1 >= 8 + 3 * lSize
                Call CopyMemory(baBuffer(0), baBuffer(8 + 2 * lSize), lSize)
                pvArrayReallocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
                If Not pvPkiExportEccPrivateKey(baPrivKey, baBuffer, lMagic) Then
                    GoTo QH
                End If
            Case Else
                #If ImplUseDebugLog Then
                    DebugLog MODULE_NAME, FUNC_NAME, Replace(ERR_UNKNOWN_CNG_MAGIC, "%1", "&H" & Hex$(lMagic)), vbLogEventTypeWarning
                #End If
            End Select
            cPrivKey.Add baPrivKey
        End If
   Else
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, ByVal 0, lSize) = 0 Then
            '--- assume there is an ephemeral key in CERT_KEY_PROV_HANDLE_PROP_ID context property
            With uProvInfo
                .dwProvType = PROV_RSA_FULL
                .dwFlags = CRYPT_VERIFYCONTEXT
                .dwKeySpec = lKeySpec
            End With
        Else
            pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
            If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, baBuffer(0), lSize) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CertGetCertificateContextProperty(CERT_KEY_PROV_INFO_PROP_ID)"
                GoTo QH
            End If
            Debug.Assert UBound(baBuffer) + 1 >= LenB(uProvInfo)
            Call CopyMemory(uProvInfo, baBuffer(0), LenB(uProvInfo))
        End If
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_HANDLE_PROP_ID, hProvGetUserKey, LenB(hProvGetUserKey)) = 0 Then
            hProvGetUserKey = hProv
            lKeySpec = uProvInfo.dwKeySpec
        End If
        If CryptGetUserKey(hProvGetUserKey, lKeySpec, hKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptGetUserKey"
            GoTo QH
        End If
        If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, ByVal 0, lSize) = 0 Then
ExportFailed:
            cPrivKey.Add pvToStringW(uProvInfo.pwszContainerName)
            cPrivKey.Add pvToStringW(uProvInfo.pwszProvName)
            cPrivKey.Add uProvInfo.dwProvType
            cPrivKey.Add lKeySpec
        Else
            pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
            If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, baBuffer(0), lSize) = 0 Then
                '--- NT4 fails late on non-exportable keys
                GoTo ExportFailed
            End If
            Debug.Assert UBound(baBuffer) + 1 >= 12
            Call CopyMemory(lMagic, baBuffer(8), LenB(lMagic))
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
                #If ImplUseDebugLog Then
                    DebugLog MODULE_NAME, FUNC_NAME, Replace(ERR_UNKNOWN_CAPI_MAGIC, "%1", "&H" & Hex$(lMagic)), vbLogEventTypeWarning
                #End If
            End Select
            cPrivKey.Add baPrivKey
        End If
    End If
    '--- success
    pvPkiExportPrivateKey = True
QH:
    On Error GoTo 0
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 And lFree <> 0 Then
        If lKeySpec < 0 Then
            Call NCryptFreeObject(hProv)
        Else
            Call CryptReleaseContext(hProv, 0)
        End If
    End If
    If hNProv <> 0 Then
        Call NCryptFreeObject(hNProv)
    End If
    If hNKey <> 0 Then
        Call NCryptFreeObject(hNKey)
    End If
    If bDeleteKeyset And uProvInfo.pwszContainerName <> 0 Then
        Call CryptAcquireContext(0, uProvInfo.pwszContainerName, uProvInfo.pwszProvName, uProvInfo.dwProvType, uProvInfo.dwFlags Or CRYPT_DELETEKEYSET)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiExportRsaPrivateKey(baRetVal() As Byte, baPrivBlob() As Byte, ByVal lStructType As Long) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportRsaPrivateKey"
    Dim baRsaPrivKey()  As Byte
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim sObjId          As String
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
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
    On Error GoTo 0
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiExportEccPrivateKey(baRetVal() As Byte, baPrivBlob() As Byte, ByVal lMagic As Long) As Boolean
    Const FUNC_NAME     As String = "pvPkiExportEccPrivateKey"
    Dim sObjId          As String
    Dim uEccPrivKey     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
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
    On Error GoTo 0
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
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
    Dim hNewKey         As Long
    Dim baBuffer()      As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    '--- export PKCS#8 password protected blob
    ReDim uParams.Buffers(0 To 2) As NCryptBuffer
    Debug.Assert RedimStats(MODULE_NAME & "." & FUNC_NAME & ".uParams.Buffers", 0)
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
        GoTo QH
    End If
    pvArrayAllocate baPkcs8, lSize, FUNC_NAME & ".baPkcs8"
    hResult = NCryptExportKey(hKey, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, baPkcs8(0), UBound(baPkcs8) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptExportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    '--- retrieve more key props
    hResult = NCryptGetProperty(hKey, StrPtr("Provider Handle"), hProv, LenB(hProv), 0, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Provider Handle)"
        GoTo QH
    End If
    hResult = NCryptGetProperty(hKey, StrPtr("Name"), ByVal 0, 0, lSize, 0)
    If Not hResult < 0 Then
        pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
        hResult = NCryptGetProperty(hKey, StrPtr("Name"), baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
        If hResult < 0 Then
            sApiSource = "NCryptGetProperty(Name)"
            GoTo QH
        End If
        '--- remove trailing terminating zero too
        sKeyName = Replace(CStr(baBuffer), vbNullChar, vbNullString)
    End If
    '--- import PKCS#8 blob and set Export Policy before finalizing
    ReDim uParams.Buffers(0 To 1) As NCryptBuffer
    Debug.Assert RedimStats(MODULE_NAME & "." & FUNC_NAME & ".uParams.Buffers", 0)
    uParams.cBuffers = IIf(LenB(sKeyName) <> 0, 2, 1)
    uParams.pBuffers = VarPtr(uParams.Buffers(0))
    sSecret = STR_PASSWORD
    With uParams.Buffers(0)
        .BufferType = NCRYPTBUFFER_PKCS_SECRET
        .pvBuffer = StrPtr(sSecret)
        .cbBuffer = LenB(sSecret) + 2
    End With
    If LenB(sKeyName) <> 0 Then
        With uParams.Buffers(1)
            .BufferType = NCRYPTBUFFER_PKCS_KEY_NAME
            .pvBuffer = StrPtr(sKeyName)
            .cbBuffer = LenB(sKeyName) + 2
        End With
    End If
    hResult = NCryptImportKey(hProv, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, hNewKey, baPkcs8(0), UBound(baPkcs8) + 1, NCRYPT_OVERWRITE_KEY_FLAG Or NCRYPT_DO_NOT_FINALIZE_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptImportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    hResult = NCryptSetProperty(hNewKey, StrPtr("Export Policy"), lPolicy, LenB(lPolicy), NCRYPT_PERSIST_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptSetProperty(Export Policy)"
        GoTo QH
    End If
    hResult = NCryptFinalizeKey(hNewKey, 0)
    If hResult < 0 Then
        sApiSource = "NCryptFinalizeKey"
        GoTo QH
    End If
    pvPkiCloneKeyWithExportPolicy = hNewKey
    hNewKey = 0
QH:
    On Error GoTo 0
    If hNewKey <> 0 Then
        Call NCryptFreeObject(hNewKey)
    End If
    If hProv <> 0 Then
        Call NCryptFreeObject(hProv)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiIssuerToCertRdn(vIssuer As Variant, uOutput As CERT_RDN) As Boolean
    Const FUNC_NAME     As String = "pvPkiIssuerToCertRdn"
    Dim baRDN()         As Byte
    Dim lPtr            As Long
    Dim uNameInfo       As CRYPT_DATA_BLOB
    Dim uCertRdn()      As CRYPT_DATA_BLOB
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    If Not IsArray(vIssuer) Then
        GoTo QH
    End If
    baRDN = vIssuer
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_NAME, baRDN(0), UBound(baRDN) + 1, _
            CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG Or CRYPT_DECODE_SHARE_OID_STRING_FLAG, 0, lPtr, 0) = 0 Then
        GoTo QH
    End If
    Debug.Assert lPtr <> 0
    Call CopyMemory(uNameInfo, ByVal lPtr, LenB(uNameInfo))
    If uNameInfo.cbData > 0 Then
        ReDim uCertRdn(0 To uNameInfo.cbData - 1) As CRYPT_DATA_BLOB
        Debug.Assert uNameInfo.pbData <> 0
        Call CopyMemory(uCertRdn(0), ByVal uNameInfo.pbData, LenB(uCertRdn(0)) * uNameInfo.cbData)
        For lIdx = 0 To UBound(uCertRdn)
            lPos = lPos + uCertRdn(lIdx).cbData
        Next
    End If
    If lPos > 0 Then
        ReDim uOutput.Buffer(0 To lPos - 1) As CERT_RDN_ATTR
        lPos = 0
        For lIdx = 0 To UBound(uCertRdn)
            Debug.Assert uCertRdn(lIdx).pbData <> 0
            Call CopyMemory(uOutput.Buffer(lPos), ByVal uCertRdn(lIdx).pbData, uCertRdn(lIdx).cbData * sizeof_CERT_RDN_ATTR)
            lPos = lPos + uCertRdn(lIdx).cbData
        Next
        uOutput.cRDNAttr = lPos
        uOutput.rgRDNAttr = VarPtr(uOutput.Buffer(0))
        '--- success
        pvPkiIssuerToCertRdn = True
    End If
QH:
    On Error GoTo 0
    If lPtr <> 0 Then
        Call LocalFree(lPtr)
    End If
    If IsArray(vErr) Then
        ErrRaise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPkiParseSelectorString(ByVal Subject As String, lStoreFlags As Long, sStoreName As String, sCertName As String) As Boolean
    Dim lPos            As Long
    
    lStoreFlags = CERT_SYSTEM_STORE_CURRENT_USER
    sStoreName = "MY"
    lPos = InStr(Subject, "\")
    If lPos = 0 Then
        lPos = Len(Subject) + 1
    End If
    Select Case Left$(Subject, lPos - 1)
    Case "CURRENT_USER"
        lStoreFlags = CERT_SYSTEM_STORE_CURRENT_USER
    Case "LOCAL_MACHINE"
        lStoreFlags = CERT_SYSTEM_STORE_LOCAL_MACHINE
    Case "CURRENT_SERVICE"
        lStoreFlags = CERT_SYSTEM_STORE_CURRENT_SERVICE
    Case "SERVICES"
        lStoreFlags = CERT_SYSTEM_STORE_SERVICES
    Case "USERS"
        lStoreFlags = CERT_SYSTEM_STORE_USERS
    Case Else
        lPos = 0
    End Select
    If lPos > 0 Then
        Subject = Mid$(Subject, lPos + 1)
    End If
    lPos = InStrRev(Subject, "\")
    If lPos > 0 Then
        sStoreName = Left$(Subject, lPos - 1)
        sCertName = Mid$(Subject, lPos + 1)
    Else
        sCertName = Subject
    End If
    '--- success
    pvPkiParseSelectorString = True
End Function

#If ImplTlsServer And ImplEncryptedPrivateKey Then
Private Function pvPkiAsn1Scanf(ByVal lPtr As Long, ByVal lSize As Long, sFormat As String) As Collection
    Dim uInput          As CRYPT_DATA_BLOB
    Dim cOutput         As Collection
    
    uInput.pbData = lPtr
    uInput.cbData = lSize
    Set cOutput = New Collection
    pvPkiAsn1Parse uInput, sFormat, cOutput
    Set pvPkiAsn1Scanf = cOutput
End Function

Private Function pvPkiAsn1Parse(uInput As CRYPT_DATA_BLOB, sFormat As String, cOutput As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiAsn1DecodeRaw"
    Const ASN_SEQUENCE  As Long = &H30
    Dim sChar           As String
    Dim lPtr            As Long
    Dim uBlob           As CRYPT_DATA_BLOB
    Dim uDers()         As CRYPT_DATA_BLOB
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    Dim lType           As Long
    
    Select Case Left$(sFormat, 1)
    Case "("
        sFormat = Mid$(sFormat, 2)
        If CryptDecodeObjectEx(X509_ASN_ENCODING, X509_SEQUENCE_OF_ANY, ByVal uInput.pbData, uInput.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPtr, LenB(lPtr)) <> 0 Then
            Call CopyMemory(uBlob, ByVal lPtr, LenB(uBlob))
            If uBlob.cbData > 0 Then
                ReDim uDers(0 To uBlob.cbData - 1) As CRYPT_DATA_BLOB
                Call CopyMemory(uDers(0), ByVal uBlob.pbData, uBlob.cbData * sizeof_CRYPT_DATA_BLOB)
            End If
            For lIdx = 0 To uBlob.cbData - 1
                Select Case Left$(sFormat, 1)
                Case ")"
                    Exit For
                Case "("
                    If uDers(lIdx).cbData > 0 Then
                        Call CopyMemory(lType, ByVal uDers(lIdx).pbData, 1)
                    End If
                    If lType = ASN_SEQUENCE Then
                        pvPkiAsn1Parse uDers(lIdx), sFormat, cOutput
                    End If
                Case Else
                    If pvPkiAsn1Parse(uDers(lIdx), sFormat, cOutput) Then
                        sFormat = Mid$(sFormat, 2)
                    End If
                End Select
            Next
        End If
        Do
            sChar = Left$(sFormat, 1)
            If sChar <> "(" Then
                sFormat = Mid$(sFormat, 2)
            End If
            Select Case sChar
            Case ")"
                Exit Do
            Case "("
                ReDim uDers(0 To 0) As CRYPT_DATA_BLOB
                pvPkiAsn1Parse uDers(0), sFormat, cOutput
            Case "o"
                cOutput.Add vbNullString
            Case "x"
                baBuffer = vbNullString
                cOutput.Add baBuffer
            Case "i"
                cOutput.Add 0
            End Select
        Loop
    Case "o"
        If CryptDecodeObjectEx(X509_ASN_ENCODING, X509_OBJECT_IDENTIFIER, ByVal uInput.pbData, uInput.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPtr, LenB(lPtr)) = 0 Then
            GoTo QH
        End If
        Call CopyMemory(lIdx, ByVal lPtr, LenB(lIdx))
        cOutput.Add pvToStringA(lIdx)
    Case "x"
        If CryptDecodeObjectEx(X509_ASN_ENCODING, X509_OCTET_STRING, ByVal uInput.pbData, uInput.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPtr, LenB(lPtr)) = 0 Then
            GoTo QH
        End If
        Call CopyMemory(uBlob, ByVal lPtr, LenB(uBlob))
        If uBlob.cbData > 0 Then
            pvArrayAllocate baBuffer, uBlob.cbData, FUNC_NAME & ".baBuffer"
            Call CopyMemory(baBuffer(0), ByVal uBlob.pbData, uBlob.cbData)
        Else
            baBuffer = vbNullString
        End If
        cOutput.Add baBuffer
    Case "i"
        If CryptDecodeObjectEx(X509_ASN_ENCODING, X509_INTEGER, ByVal uInput.pbData, uInput.cbData, 0, 0, lIdx, LenB(lIdx)) = 0 Then
            GoTo QH
        End If
        cOutput.Add lIdx
    End Select
    '--- success
    pvPkiAsn1Parse = True
QH:
End Function
#End If

Private Function pvNCryptGetStringProperty(ByVal hNKey As String, sPropName As String) As String
    Const FUNC_NAME     As String = "pvNCryptGetStringProperty"
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim hResult         As Long
    
    hResult = NCryptGetProperty(hNKey, StrPtr(sPropName), ByVal 0, 0, lSize, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    pvArrayAllocate baBuffer, lSize, FUNC_NAME & ".baBuffer"
    hResult = NCryptGetProperty(hNKey, StrPtr(sPropName), baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    '--- remove trailing terminating zero too
    pvNCryptGetStringProperty = Replace(CStr(baBuffer), vbNullChar, vbNullString)
QH:
End Function

Private Function pvToStringA(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToStringA = String$(lstrlenA(lPtr), 0)
        Call CopyMemory(ByVal pvToStringA, ByVal lPtr, Len(pvToStringA))
    End If
End Function

Private Function pvToStringW(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToStringW = String$(lstrlenW(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvToStringW), ByVal lPtr, LenB(pvToStringW))
    End If
End Function

Private Function pvCollectionCount(oCol As Collection) As Long
    If Not oCol Is Nothing Then
        pvCollectionCount = oCol.Count
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
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), LenB(lPtr))
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

Private Function pvWriteArray(baBuffer() As Byte, ByVal lPos As Long, baSrc() As Byte) As Long
    Dim lSize       As Long
    
    lSize = pvArraySize(baSrc)
    If lSize > 0 Then
        lPos = pvWriteBuffer(baBuffer, lPos, VarPtr(baSrc(0)), lSize)
    End If
    pvWriteArray = lPos
End Function

Private Function pvWriteBuffer(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvWriteBuffer"
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baBuffer)
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), LenB(lBufPtr))
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

'= shared ================================================================

#If Not ImplUseShared Then
Private Function RedimStats(sFuncName As String, ByVal lSize As Long) As Boolean
    #If sFuncName And lSize Then
    #End If
    RedimStats = True
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
    If GetFileAttributes(StrPtr(sFile)) <> -1 Then
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
    Dim aVer(0 To 69)   As Long
    
    If lVersion = 0 Then
        aVer(0) = LenB(aVer(0)) * UBound(aVer)  '--- [0] = dwOSVersionInfoSize
        If GetVersionEx(aVer(0)) <> 0 Then
            lVersion = aVer(1) * 100 + aVer(2)  '--- [1] = dwMajorVersion, [2] = dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property
#End If
