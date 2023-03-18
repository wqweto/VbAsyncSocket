Attribute VB_Name = "mdTlsThunks"
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
'   https://www.davidwong.fr/tls13
'
' Some of the cryptographic thunks are based on the following sources
'
'  1. https://github.com/esxgx/easy-ecc by Kenneth MacKay
'     which is distributed under the BSD 2-clause license
'
'  2. https://github.com/ctz/cifra by Joseph Birr-Pixton
'     CC0 1.0 Universal license (Public Domain Dedication)
'
'  3. https://github.com/github/putty by Simon Tatham
'     which is distributed under the MIT licence
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdTlsThunks"

#Const ImplTlsServer = (ASYNCSOCKET_NO_TLSSERVER = 0)
#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)
#Const ImplCaptureTraffic = CLng(ASYNCSOCKET_CAPTURE_TRAFFIC) '--- bitmask: 1 - traffic, 2 - derived secrets
#Const ImplExoticCiphers = False
#Const ImplTlsServerAllowInsecureRenegotiation = False
#Const ImplTestCrypto = (ASYNCSOCKET_TEST_CRYPTO <> 0)

'=========================================================================
' API
'=========================================================================

'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40
Private Const PTR_SIZE                                  As Long = 4
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const PROV_RSA_AES                              As Long = 24
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_PUBLIC_KEY_INFO                      As Long = 8
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
'--- for CryptExportKey
Private Const PUBLICKEYBLOB                             As Long = 6
'--- for CryptCreateHash
Private Const CALG_SSL3_SHAMD5                          As Long = &H8008&
Private Const CALG_SHA1                                 As Long = &H8004&
Private Const CALG_SHA_256                              As Long = &H800C&
Private Const CALG_SHA_384                              As Long = &H800D&
Private Const CALG_SHA_512                              As Long = &H800E&
Private Const HP_HASHVAL                                As Long = 2
Private Const HP_HASHSIZE                               As Long = 4
'--- for NCryptSignHash
Private Const BCRYPT_PAD_PKCS1                          As Long = &H2
Private Const BCRYPT_PAD_PSS                            As Long = &H8
'--- for NCryptDecrypt
Private Const NCRYPT_PAD_PKCS1_FLAG                     As Long = &H2
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_RSA_SSA_PSS                         As String = "1.2.840.113549.1.1.10"
Private Const szOID_ECC_PUBLIC_KEY                      As String = "1.2.840.10045.2.1"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpStr As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
'--- msvbvm60
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32" Alias "CryptSignHashW" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal szDescription As Long, ByVal dwFlags As Long, pbSignature As Any, pdwSigLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
'--- Crypt32
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
'--- NCrypt
Private Declare Function NCryptOpenStorageProvider Lib "ncrypt" (phProvider As Long, ByVal pszProviderName As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptOpenKey Lib "ncrypt" (ByVal hProvider As Long, phKey As Long, ByVal pszKeyName As Long, ByVal dwLegacyKeySpec As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptSignHash Lib "ncrypt" (ByVal hKey As Long, ByVal pPaddingInfo As Long, pbHashValue As Any, ByVal cbHashValue As Long, pbSignature As Any, ByVal cbSignature As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptDecrypt Lib "ncrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFreeObject Lib "ncrypt" (ByVal hObject As Long) As Long

Private Type CRYPT_DATA_BLOB
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

Private Type CRYPT_PRIVATE_KEY_INFO
    Version             As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_DATA_BLOB
    pAttributes         As Long
End Type

Private Type BCRYPT_PSS_PADDING_INFO
    pszAlgId            As Long
    cbSalt              As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS                             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|21|Decryption failed|22|Record overflow|30|Decompression failure|40|Handshake failure|41|No certificate|42|Bad certificate|43|Unsupported certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown certificate authority|50|Decode error|51|Decrypt error|70|Protocol version|71|Insufficient security|80|Internal error|90|User canceled|100|No renegotiation|109|Missing extension|110|Unsupported extension|111|Certificate unobtainable|112|Unrecognized name|113|Bad certificate status response|114|Bad certificate hash value|116|Certificate required|120|No application protocol"
Private Const STR_VL_STATES                             As String = "0|New|1|Closed|2|HandshakeStart|3|ExpectServerHello|4|ExpectExtensions|5|ExpectServerFinished|6|ExpectClientHello|7|ExpectClientKeyExchange|8|ExpectClientFinished|9|PostHandshake|10|Shutdown"
Private Const STR_VL_MESSAGE_NAMES                      As String = "1|client_hello|2|server_hello|4|new_session_ticket|5|end_of_early_data|8|encrypted_extensions|11|certificate|12|server_key_exchange|13|certificate_request|14|server_hello_done|15|certificate_verify|16|client_key_exchange|20|finished|21|certificate_url|22|certificate_status|24|key_update|25|compressed_certificate|254|message_hash"
Private Const STR_VL_EXTENSION_NAMES                    As String = "0|server_name|1|max_fragment_length|2|client_certificate_url|3|trusted_ca_keys|4|truncated_hmac|5|status_request|6|user_mapping|7|client_authz|8|server_authz|9|cert_type|10|supported_groups|11|ec_point_formats|12|srp|13|signature_algorithms|14|use_srtp|15|heartbeat|16|application_layer_protocol_negotiation|17|status_request_v2|18|signed_certificate_timestamp|19|client_certificate_type|20|server_certificate_type|21|padding|22|encrypt_then_mac|23|extended_master_secret|24|token_binding|25|cached_info|26|tls_lts|27|compress_certificate|28|record_size_limit|29|pwd_protect|30|pwd_clear|31|password_salt|32|ticket_pinning|33|tls_cert_with_extern_psk|34|delegated_credentials|35|session_ticket|41|pre_shared_key|42|early_data|43|supported_versions|44|cookie|45|psk_key_exchange_modes|47|certificate_authorities|48|oid_filters|49|post_handshake_auth|" & _
                                                                    "50|signature_algorithms_cert|51|key_share|52|transparency_info|53|connection_id|55|external_id_hash|56|external_session_id"
Private Const STR_UNKNOWN                               As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT                          As String = "%1."
'--- TLS
Private Const TLS_PROTOCOL_VERSION_TLS12                As Long = &H303
Private Const TLS_PROTOCOL_VERSION_TLS13                As Long = &H304
Private Const TLS_RECORD_VERSION                        As Long = TLS_PROTOCOL_VERSION_TLS12 '--- always legacy version
Private Const TLS_LOCAL_LEGACY_VERSION                  As Long = TLS_PROTOCOL_VERSION_TLS12
'--- TLS ContentType from https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-5
Private Const TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC       As Long = 20
Private Const TLS_CONTENT_TYPE_ALERT                    As Long = 21
Private Const TLS_CONTENT_TYPE_HANDSHAKE                As Long = 22
Private Const TLS_CONTENT_TYPE_APPDATA                  As Long = 23
'--- TLS HandshakeType from https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-7
Private Const TLS_HANDSHAKE_HELLO_REQUEST               As Long = 0
Private Const TLS_HANDSHAKE_CLIENT_HELLO                As Long = 1
Private Const TLS_HANDSHAKE_SERVER_HELLO                As Long = 2
Private Const TLS_HANDSHAKE_NEW_SESSION_TICKET          As Long = 4
'Private Const TLS_HANDSHAKE_END_OF_EARLY_DATA           As Long = 5
Private Const TLS_HANDSHAKE_ENCRYPTED_EXTENSIONS        As Long = 8
Private Const TLS_HANDSHAKE_CERTIFICATE                 As Long = 11
Private Const TLS_HANDSHAKE_SERVER_KEY_EXCHANGE         As Long = 12
Private Const TLS_HANDSHAKE_CERTIFICATE_REQUEST         As Long = 13
Private Const TLS_HANDSHAKE_SERVER_HELLO_DONE           As Long = 14
Private Const TLS_HANDSHAKE_CERTIFICATE_VERIFY          As Long = 15
Private Const TLS_HANDSHAKE_CLIENT_KEY_EXCHANGE         As Long = 16
Private Const TLS_HANDSHAKE_FINISHED                    As Long = 20
Private Const TLS_HANDSHAKE_CERTIFICATE_STATUS          As Long = 22
Private Const TLS_HANDSHAKE_KEY_UPDATE                  As Long = 24
'Private Const TLS_HANDSHAKE_COMPRESSED_CERTIFICATE      As Long = 25
Private Const TLS_HANDSHAKE_MESSAGE_HASH                As Long = 254
'--- TLS Extensions from https://www.iana.org/assignments/tls-extensiontype-values/tls-extensiontype-values.xhtml
Private Const TLS_EXTENSION_SERVER_NAME                 As Long = 0
Private Const TLS_EXTENSION_STATUS_REQUEST              As Long = 5
Private Const TLS_EXTENSION_SUPPORTED_GROUPS            As Long = 10
Private Const TLS_EXTENSION_EC_POINT_FORMAT             As Long = 11
Private Const TLS_EXTENSION_SIGNATURE_ALGORITHMS        As Long = 13
Private Const TLS_EXTENSION_ALPN                        As Long = 16
Private Const TLS_EXTENSION_ENCRYPT_THEN_MAC            As Long = 22
Private Const TLS_EXTENSION_EXTENDED_MASTER_SECRET      As Long = 23
'Private Const TLS_EXTENSION_COMPRESS_CERTIFICATE        As Long = 27
Private Const TLS_EXTENSION_SESSION_TICKET              As Long = 35
'Private Const TLS_EXTENSION_PRE_SHARED_KEY              As Long = 41
'Private Const TLS_EXTENSION_EARLY_DATA                  As Long = 42
Private Const TLS_EXTENSION_SUPPORTED_VERSIONS          As Long = 43
Private Const TLS_EXTENSION_COOKIE                      As Long = 44
'Private Const TLS_EXTENSION_PSK_KEY_EXCHANGE_MODES      As Long = 45
Private Const TLS_EXTENSION_CERTIFICATE_AUTHORITIES     As Long = 47
Private Const TLS_EXTENSION_POST_HANDSHAKE_AUTH         As Long = 49
Private Const TLS_EXTENSION_SIGNATURE_ALGORITHMS_CERT   As Long = 50
Private Const TLS_EXTENSION_KEY_SHARE                   As Long = 51
Private Const TLS_EXTENSION_RENEGOTIATION_INFO          As Long = &HFF01&
'--- TLS Cipher Suites from http://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-4
Private Const TLS_CS_AES_128_GCM_SHA256                 As Long = &H1301
Private Const TLS_CS_AES_256_GCM_SHA384                 As Long = &H1302
Private Const TLS_CS_CHACHA20_POLY1305_SHA256           As Long = &H1303
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256 As Long = &HC02B&
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384 As Long = &HC02C&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256  As Long = &HC02F&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384  As Long = &HC030&
Private Const TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA8&
Private Const TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA9&
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA   As Long = &HC009&
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA   As Long = &HC00A&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA     As Long = &HC013&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA     As Long = &HC014&
Private Const TLS_CS_RSA_WITH_AES_128_CBC_SHA           As Long = &H2F
Private Const TLS_CS_RSA_WITH_AES_256_CBC_SHA           As Long = &H35
Private Const TLS_CS_RSA_WITH_AES_128_GCM_SHA256        As Long = &H9C
Private Const TLS_CS_RSA_WITH_AES_256_GCM_SHA384        As Long = &H9D
Private Const TLS_CS_EMPTY_RENEGOTIATION_INFO_SCSV      As Long = &HFF
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256 As Long = &HC023&
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384 As Long = &HC024&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA256  As Long = &HC027&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA384  As Long = &HC028&
Private Const TLS_CS_RSA_WITH_AES_128_CBC_SHA256        As Long = &H3C
Private Const TLS_CS_RSA_WITH_AES_256_CBC_SHA256        As Long = &H3D
'--- TLS Supported Groups from https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-8
Private Const TLS_GROUP_SECP256R1                       As Long = 23
Private Const TLS_GROUP_SECP384R1                       As Long = 24
Private Const TLS_GROUP_SECP521R1                       As Long = 25
Private Const TLS_GROUP_X25519                          As Long = 29
Private Const TLS_GROUP_X448                            As Long = 30
Private Const TLS_GROUP_FFDHE_FIRST                     As Long = &H100
Private Const TLS_GROUP_FFDHE_LAST                      As Long = &H104
Private Const TLS_GROUP_FFDHE_PRIVATE_USE_FIRST         As Long = &H1FC
Private Const TLS_GROUP_FFDHE_PRIVATE_USE_LAST          As Long = &H1FF
Private Const TLS_GROUP_ECDHE_PRIVATE_USE_FIRST         As Long = &HFE00&
Private Const TLS_GROUP_ECDHE_PRIVATE_USE_LAST          As Long = &HFEFF&
'--- TLS Signature Scheme from https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-signaturescheme
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA1              As Long = &H201 '--- TLS 1.2
Private Const TLS_SIGNATURE_ECDSA_SHA1                  As Long = &H203
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA224            As Long = &H301
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA256            As Long = &H401
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA384            As Long = &H501
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA512            As Long = &H601
Private Const TLS_SIGNATURE_ECDSA_SECP256R1_SHA256      As Long = &H403 '--- TLS 1.3
Private Const TLS_SIGNATURE_ECDSA_SECP384R1_SHA384      As Long = &H503
Private Const TLS_SIGNATURE_ECDSA_SECP521R1_SHA512      As Long = &H603
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA512         As Long = &H806
'Private Const TLS_SIGNATURE_ED25519                     As Long = &H807
'Private Const TLS_SIGNATURE_ED448                       As Long = &H808
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA256          As Long = &H809
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA384          As Long = &H80A
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA512          As Long = &H80B
Private Const TLS_ALERT_LEVEL_WARNING                   As Long = 1
Private Const TLS_ALERT_LEVEL_FATAL                     As Long = 2
Private Const TLS_COMPRESS_NULL                         As Long = 0
Private Const TLS_SERVER_NAME_TYPE_HOSTNAME             As Long = 0
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = TLS_MAX_PLAINTEXT_RECORD_SIZE + 1 + 255 '-- 1 byte content type + 255 bytes AEAD padding
Private Const TLS_HELLO_RANDOM_SIZE                     As Long = 32
Private Const TLS_LEGACY_SECRET_SIZE                    As Long = 48
Private Const TLS_LEGACY_SESSIONID_SIZE                 As Long = 32
Private Const TLS_AAD_SIZE                              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const TLS_LEGACY_AAD_SIZE                       As Long = 13    '--- for TLS 1.2
Private Const TLS_VERIFY_DATA_SIZE                      As Long = 12
'Private Const TLS_PSK_KE_MODE_PSK_DHE                   As Long = 1
'--- crypto constants
Private Const LNG_X25519_KEYSZ                          As Long = 32
Private Const LNG_SECP256R1_KEYSZ                       As Long = 32
Private Const LNG_SECP384R1_KEYSZ                       As Long = 48
Private Const LNG_MD5_HASHSZ                            As Long = 16
'Private Const LNG_MD5_BLOCKSZ                           As Long = 64
Private Const LNG_SHA1_HASHSZ                           As Long = 20
Private Const LNG_SHA1_BLOCKSZ                          As Long = 64
Private Const LNG_SHA224_HASHSZ                         As Long = 28
'Private Const LNG_SHA224_BLOCKSZ                        As Long = 64
Private Const LNG_SHA256_HASHSZ                         As Long = 32
Private Const LNG_SHA256_BLOCKSZ                        As Long = 64
Private Const LNG_SHA384_HASHSZ                         As Long = 48
Private Const LNG_SHA384_BLOCKSZ                        As Long = 128
Private Const LNG_SHA384_CONTEXTSZ                      As Long = 200
Private Const LNG_SHA512_HASHSZ                         As Long = 64
Private Const LNG_SHA512_BLOCKSZ                        As Long = 128
Private Const LNG_CHACHA20_KEYSZ                        As Long = 32
Private Const LNG_CHACHA20POLY1305_IVSZ                 As Long = 12
Private Const LNG_CHACHA20POLY1305_TAGSZ                As Long = 16
Private Const LNG_AES128_KEYSZ                          As Long = 16
Private Const LNG_AES256_KEYSZ                          As Long = 32
Private Const LNG_AESGCM_IVSZ                           As Long = 12
Private Const LNG_AESGCM_TAGSZ                          As Long = 16
Private Const LNG_AESCBC_IVSZ                           As Long = 16
Private Const LNG_ANS1_TYPE_SEQUENCE                    As Long = &H30
Private Const LNG_ANS1_TYPE_INTEGER                     As Long = &H2
Private Const LNG_HMAC_INNER_PAD                        As Long = &H36
Private Const LNG_HMAC_OUTER_PAD                        As Long = &H5C
'--- errors
Private Const ERR_CONNECTION_CLOSED                     As String = "Connection closed"
Private Const ERR_GENER_KEYPAIR_FAILED                  As String = "Failed generating key pair (%1)"
Private Const ERR_UNSUPPORTED_EXCH_GROUP                As String = "Unsupported exchange group (%1)"
Private Const ERR_UNSUPPORTED_CIPHER_SUITE              As String = "Unsupported cipher suite (%1)"
Private Const ERR_UNSUPPORTED_SIGNATURE_SCHEME          As String = "Unsupported signature scheme (%1)"
Private Const ERR_UNSUPPORTED_CERTIFICATE               As String = "Unsupported certificate"
Private Const ERR_UNSUPPORTED_PUBLIC_KEY                As String = "Unsupported public key OID (%1)"
Private Const ERR_UNSUPPORTED_PRIVATE_KEY               As String = "Unsupported private key"
Private Const ERR_UNSUPPORTED_CURVE_SIZE                As String = "Unsupported curve size (%1)"
Private Const ERR_UNSUPPORTED_CURVE_TYPE                As String = "Unsupported curve type (%1)"
Private Const ERR_UNSUPPORTED_PROTOCOL                  As String = "Invalid protocol version (%1)"
Private Const ERR_ENCRYPTION_FAILED                     As String = "Encryption failed"
Private Const ERR_SIGNATURE_FAILED                      As String = "Certificate signature failed (%1)"
Private Const ERR_DECRYPTION_FAILED                     As String = "Decryption failed"
Private Const ERR_SERVER_HANDSHAKE_FAILED               As String = "Handshake verification failed"
Private Const ERR_RECORD_MAC_FAILED                     As String = "MAC verification failed"
Private Const ERR_HELLO_RETRY_FAILED                    As String = "HelloRetryRequest failed"
Private Const ERR_NEGOTIATE_SIGNATURE_FAILED            As String = "Negotiate signature type failed"
Private Const ERR_SECURE_RENEGOTIATION_FAILED           As String = "Secure renegotiation failed"
Private Const ERR_CALL_FAILED                           As String = "Call failed (%1)"
Private Const ERR_RECORD_OVERFLOW                       As String = "Record size too big"
Private Const ERR_FATAL_ALERT                           As String = "Received fatal alert"
Private Const ERR_UNEXPECTED_RECORD_TYPE                As String = "Unexpected record type (%1)"
Private Const ERR_UNEXPECTED_MSG_TYPE                   As String = "Unexpected message type for %1 state (%2)"
Private Const ERR_UNEXPECTED_EXTENSION                  As String = "Unexpected extension (%1)"
Private Const ERR_INVALID_STATE_HANDSHAKE               As String = "Invalid state for handshake content (%1)"
Private Const ERR_INVALID_REMOTE_KEY                    As String = "Invalid remote key"
Private Const ERR_INVALID_SIZE                          As String = "Invalid data size"
Private Const ERR_INVALID_SIZE_EXTENSION                As String = "Invalid data size for %1"
Private Const ERR_INVALID_SIGNATURE                     As String = "Invalid certificate signature"
Private Const ERR_INVALID_HASH_SIZE                     As String = "Invalid hash size (%1)"
Private Const ERR_NO_HANDSHAKE_MESSAGES                 As String = "Missing handshake messages"
Private Const ERR_NO_PREVIOUS_SECRET                    As String = "Missing previous %1"
Private Const ERR_NO_REMOTE_RANDOM                      As String = "Missing remote random"
Private Const ERR_NO_SERVER_CERTIFICATE                 As String = "Missing server certificate"
Private Const ERR_NO_SUPPORTED_CIPHER_SUITE             As String = "Missing supported ciphersuite (%1)"
Private Const ERR_NO_CERTIFICATE                        As String = "Missing certificate"
Private Const ERR_NO_SERVER_COMPILED                    As String = "Server-side TLS not compiled (ASYNCSOCKET_NO_TLSSERVER = 1)"
Private Const ERR_NO_SUPPORTED_GROUPS                   As String = "Missing supported remote group"
Private Const ERR_NO_ALPN_NEGOTIATED                    As String = "No application protocol negotiated"
Private Const ERR_NO_EXTENSION                          As String = "Missing extension (%1)"
Private Const ERR_UNSUPPORTED_REQUEST_UPDATE            As String = "Unsupported request update (%1)"
Private Const ERR_INVALID_COMPRESSION                   As String = "Invalid compression (%1)"
Private Const ERR_INVALID_SERVER_NAME                   As String = "Invalid server name (%1)"
Private Const ERR_MISSING_TRAFFIC_KEYS                  As String = "Missing remote traffic keys"
'--- numeric
Private Const LNG_FACILITY_WIN32                        As Long = &H80070000

Private m_baBuffer()                As Byte
Private m_lBuffIdx                  As Long
Private m_uData                     As UcsCryptoData
Private m_baHelloRetryRandom()      As Byte
Public g_oRequestSocket             As Object

Private Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls10 = 2 ^ 0
    ucsTlsSupportTls11 = 2 ^ 1
    ucsTlsSupportTls12 = 2 ^ 2
    ucsTlsSupportTls13 = 2 ^ 3
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 4
    ucsTlsSupportAll = ucsTlsSupportTls10 Or ucsTlsSupportTls11 Or ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum

Private Enum UcsTlsStatesEnum '--- sync w/ STR_VL_STATES
    ucsTlsStateNew = 0
    ucsTlsStateClosed = 1
    ucsTlsStateHandshakeStart = 2
    ucsTlsStateExpectServerHello = 3
    ucsTlsStateExpectEncryptedExtensions = 4
    ucsTlsStateExpectServerFinished = 5     '--- not used in TLS 1.3
#If ImplTlsServer Then
    ucsTlsStateExpectClientHello = 6
    ucsTlsStateExpectClientKeyExchange = 7  '--- not used in TLS 1.3
    ucsTlsStateExpectClientFinished = 8
#End If
    ucsTlsStatePostHandshake = 9
    ucsTlsStateShutdown = 10
End Enum

Private Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoExchX25519 = 1
    ucsTlsAlgoExchSecp256r1
    ucsTlsAlgoExchSecp384r1
    ucsTlsAlgoExchSecp521r1
    ucsTlsAlgoExchCertificate
    '--- ciphers
    ucsTlsAlgoBulkChacha20Poly1305 = 11 '--- next 3 are authenticated encryption w/ additional data
    ucsTlsAlgoBulkAesGcm128
    ucsTlsAlgoBulkAesGcm256
    ucsTlsAlgoBulkAesCbc128             '--- next 2 are legacy non-AEAD
    ucsTlsAlgoBulkAesCbc256
    '--- hash
    ucsTlsAlgoDigestMd5 = 21
    ucsTlsAlgoDigestSha1
    ucsTlsAlgoDigestSha224
    ucsTlsAlgoDigestSha256
    ucsTlsAlgoDigestSha384
    ucsTlsAlgoDigestSha512
    '--- padding
    ucsTlsAlgoPaddingPkcs = 31
    ucsTlsAlgoPaddingPss
End Enum

'--- TLS Alerts https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-6
Private Enum UcsTlsAlertDescriptionsEnum
    uscTlsAlertCloseNotify = 0
    uscTlsAlertUnexpectedMessage = 10
    uscTlsAlertBadRecordMac = 20
    uscTlsAlertRecordOverflow = 22
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
    uscTlsAlertNoRenegotiation = 100
    uscTlsAlertMissingExtension = 109
    uscTlsAlertUnrecognizedName = 112
    uscTlsAlertCertificateRequired = 116
    uscTlsAlertNoApplicationProtocol = 120
End Enum

Private Type UcsBuffer
    Data()              As Byte
    Pos                 As Long
    Size                As Long
    Stack               As Collection
End Type

Public Type UcsTlsContext
    '--- config
    IsServer            As Boolean
    RemoteHostName      As String
    LocalFeatures       As UcsTlsLocalFeaturesEnum
    ClientCertCallback  As Long
    AlpnProtocols       As String
    '--- state
    State               As UcsTlsStatesEnum
    LastErrNumber       As Long
    LastError           As String
    LastErrSource       As String
    LastAlertCode       As UcsTlsAlertDescriptionsEnum
    BlocksStack         As Collection
    AlpnNegotiated      As String
    SniRequested        As String
    PrevRecordType      As Long
    '--- handshake
    ProtocolVersion     As Long
    LocalSessionID()    As Byte
    LocalSessionTicket() As Byte
    LocalExchRandom()   As Byte
    LocalExchPrivate()  As Byte
    LocalExchPublic()   As Byte
    LocalExchRsaEncrPriv() As Byte
    LocalCertificates   As Collection
    LocalPrivateKey     As Collection
    LocalLegacyVerifyData() As Byte
    RemoteProtocolVersion As Long
    RemoteSessionID()   As Byte
    RemoteExchRandom()  As Byte
    RemoteExchPublic()  As Byte
    RemoteCertificates  As Collection
    RemoteExtensions    As Collection
    RemoteTickets       As Collection
    RemoteSupportedGroups As Collection
    RemoteCertStatuses  As Collection
    RemoteLegacyVerifyData() As Byte
    RemoteLegacyRenegInfo() As Byte
    '--- crypto settings
    ExchGroup           As Long
    ExchAlgo            As UcsTlsCryptoAlgorithmsEnum
    CipherSuite         As Long
    SignatureScheme     As Long
    MacAlgo             As UcsTlsCryptoAlgorithmsEnum   '--- not used w/ AEAD ciphers
    MacSize             As Long                         '--- not used w/ AEAD ciphers
    BulkAlgo            As UcsTlsCryptoAlgorithmsEnum
    KeySize             As Long
    IvSize              As Long
    IvExplicitSize      As Long                         '--- only for AES in TLS 1.2
    TagSize             As Long
    DigestAlgo          As UcsTlsCryptoAlgorithmsEnum
    DigestSize          As Long
    UseRsaKeyTransport  As Boolean
    '--- bulk secrets
    HandshakeMessages   As UcsBuffer                    '--- ToDo: reduce to HandshakeHash only
    HandshakeSecret()   As Byte
    MasterSecret()      As Byte
    LocalMacKey()       As Byte                         '--- not used w/ AEAD ciphers
    RemoteMacKey()      As Byte                         '--- not used w/ AEAD ciphers
    RemoteTrafficSecret() As Byte
    RemoteTrafficKey()  As Byte
    RemoteTrafficIV()   As Byte
    RemoteTrafficSeqNo  As Long
    RemoteEncryptThenMac As Boolean
    RemoteLegacyNextMacKey() As Byte
    RemoteLegacyNextTrafficKey() As Byte
    RemoteLegacyNextTrafficIV() As Byte
    LocalTrafficSecret() As Byte
    LocalTrafficKey()   As Byte
    LocalTrafficIV()    As Byte
    LocalTrafficSeqNo   As Long
    LocalEncryptThenMac As Boolean
    LocalLegacyNextMacKey() As Byte
    LocalLegacyNextTrafficKey() As Byte
    LocalLegacyNextTrafficIV() As Byte
    '--- hello retry request
    HelloRetryRequest   As Boolean
    HelloRetryCipherSuite As Long
    HelloRetryExchGroup As Long
    HelloRetryCookie()  As Byte
    '--- client certificate request
    CertRequestContext() As Byte
    CertRequestSignatureScheme As Long
    CertRequestCaDn     As Collection
    '--- I/O buffers
    RecvBuffer          As UcsBuffer
    DecrBuffer          As UcsBuffer
    SendBuffer          As UcsBuffer
    MessBuffer          As UcsBuffer
#If ImplCaptureTraffic <> 0 Then
    TrafficDump         As Collection
#End If
End Type

Private Type UcsKeyInfo
    AlgoObjId           As String
    KeyBlob()           As Byte
    BitLen              As Long
    Modulus()           As Byte
    PubExp()            As Byte
    Prime1()            As Byte
    Prime2()            As Byte
    Coefficient()       As Byte
    PrivExp()           As Byte
    hNKey               As Long
    hProv               As Long
    hKey                As Long
    dwKeySpec           As Long
End Type

Private Enum UcsThunkPfnIndexEnum
    [_ucsPfnNotUsed]
    ucsPfnCurve25519ScalarMultiply
    ucsPfnCurve25519ScalarMultBase
    ucsPfnSecp256r1MakeKey
    ucsPfnSecp256r1SharedSecret
    ucsPfnSecp256r1UncompressKey
    ucsPfnSecp256r1Sign
    ucsPfnSecp256r1Verify
    ucsPfnSecp384r1MakeKey
    ucsPfnSecp384r1SharedSecret
    ucsPfnSecp384r1UncompressKey
    ucsPfnSecp384r1Sign
    ucsPfnSecp384r1Verify
    ucsPfnSha256Init
    ucsPfnSha256Update
    ucsPfnSha256Final
    ucsPfnSha384Init
    ucsPfnSha384Update
    ucsPfnSha384Final
    ucsPfnSha512Init
    ucsPfnSha512Update
    ucsPfnSha512Final
    ucsPfnChacha20Poly1305Encrypt
    ucsPfnChacha20Poly1305Decrypt
    ucsPfnAesGcmEncrypt
    ucsPfnAesGcmDecrypt
    ucsPfnAesCbcEncrypt
    ucsPfnAesCbcDecrypt
    ucsPfnRsaModExp
    ucsPfnRsaCrtModExp
    [_ucsPfnMax]
End Enum

Private Type UcsCryptoData
    Thunk               As Long
    Glob()              As Byte
    Pfn(1 To [_ucsPfnMax] - 1) As Long
    HashCtx(0 To LNG_SHA384_CONTEXTSZ - 1) As Byte
    HashPad(0 To LNG_SHA512_BLOCKSZ - 1) As Byte
    HashFinal(0 To LNG_SHA512_HASHSZ - 1) As Byte
    hRandomProv         As Long
End Type

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

Private Sub ErrRaise(ByVal Number As Long, Optional Source As Variant, Optional Description As Variant)
    Err.Raise Number, Source, Description
End Sub

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
            Optional ByVal LocalFeatures As Long = ucsTlsSupportAll, _
            Optional ClientCertCallback As Object, _
            Optional AlpnProtocols As String) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not pvCryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsClearLastError uEmpty
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        .ClientCertCallback = ObjPtr(ClientCertCallback)
        .AlpnProtocols = AlpnProtocols
        #If ImplCaptureTraffic <> 0 Then
            Set .TrafficDump = New Collection
        #End If
    End With
    uCtx = uEmpty
    '--- success
    TlsInitClient = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection, _
            Optional AlpnProtocols As String, _
            Optional ByVal LocalFeatures As Long = ucsTlsSupportAll) As Boolean
#If Not ImplTlsServer Then
    ErrRaise vbObjectError, , ERR_NO_SERVER_COMPILED
#Else
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not pvCryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsClearLastError uEmpty
        .IsServer = True
        .State = ucsTlsStateExpectClientHello
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        Set .LocalCertificates = Certificates
        Set .LocalPrivateKey = PrivateKey
        .AlpnProtocols = AlpnProtocols
        #If ImplCaptureTraffic <> 0 Then
            Set .TrafficDump = New Collection
        #End If
    End With
    uCtx = uEmpty
    '--- success
    TlsInitServer = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
#End If
End Function

Public Function TlsTerminate(uCtx As UcsTlsContext)
    uCtx.State = ucsTlsStateClosed
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsHandshake"
    
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsClearLastError uCtx
        '--- swap-in
        pvArraySwap .SendBuffer.Data, .SendBuffer.Size, baOutput, lOutputPos
        If .State = ucsTlsStateHandshakeStart Then
            pvTlsBuildClientHello uCtx, .SendBuffer
            .State = ucsTlsStateExpectServerHello
        Else
            If lSize < 0 Then
                lSize = pvArraySize(baInput)
            End If
            #If (ImplCaptureTraffic And 1) <> 0 Then
                If lSize <> 0 Then
                    .TrafficDump.Add FUNC_NAME & ".Input" & vbCrLf & TlsDesignDumpArray(baInput, Size:=lSize)
                End If
            #End If
            If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
                pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, .LastError, .LastAlertCode
                '--- treat as warnings
                If .LastAlertCode = uscTlsAlertCertificateUnknown Or .LastAlertCode = uscTlsAlertBadCertificate Then
                    TlsHandshake = True
                End If
                GoTo QH
            End If
        End If
        '--- success
        TlsHandshake = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lOutputPos, .SendBuffer.Data, .SendBuffer.Size
        pvArrayWriteEOF baOutput, lOutputPos
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lOutputPos <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Output" & vbCrLf & TlsDesignDumpArray(baOutput, Size:=lOutputPos)
            End If
        #End If
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsReceive"
    
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lSize <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Input (undecrypted)" & vbCrLf & TlsDesignDumpArray(baInput, Size:=lSize)
            End If
        #End If
        If lSize = 0 Then
            '--- flush
            If .DecrBuffer.Size > 0 Then
                If lPos = 0 Then
                    pvArraySwap .DecrBuffer.Data, .DecrBuffer.Size, baPlainText, lPos
                Else
                    lPos = pvArrayWriteBlob(baPlainText, lPos, VarPtr(.DecrBuffer.Data(0)), .DecrBuffer.Size)
                    Call pvArrayWriteEOF(baPlainText, lPos)
                    .DecrBuffer.Size = 0
                End If
            End If
            '--- success
            TlsReceive = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsClearLastError uCtx
        '--- swap-in
        pvArraySwap .DecrBuffer.Data, .DecrBuffer.Size, baPlainText, lPos
        pvArraySwap .SendBuffer.Data, .SendBuffer.Size, baOutput, lOutputPos
        If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, .LastError, .LastAlertCode
            GoTo QH
        End If
        '--- success
        TlsReceive = True
QH:
        '--- swap-out
        pvArraySwap baPlainText, lPos, .DecrBuffer.Data, .DecrBuffer.Size
        pvArrayWriteEOF baPlainText, lPos
        pvArraySwap baOutput, lOutputPos, .SendBuffer.Data, .SendBuffer.Size
        pvArrayWriteEOF baOutput, lOutputPos
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lOutputPos <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Output (encrypted)" & vbCrLf & TlsDesignDumpArray(baOutput, Size:=lOutputPos)
            End If
        #End If
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Const FUNC_NAME     As String = "TlsSend"
    Dim lPos            As Long
    
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baPlainText)
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsClearLastError uCtx
        '--- swap-in
        pvArraySwap .SendBuffer.Data, .SendBuffer.Size, baOutput, lOutputPos
        Do While lPos < lSize
            pvTlsBuildApplicationData uCtx, .SendBuffer, baPlainText, lPos, Clamp(lSize - lPos, 0, TLS_MAX_PLAINTEXT_RECORD_SIZE), TLS_CONTENT_TYPE_APPDATA
            lPos = lPos + TLS_MAX_PLAINTEXT_RECORD_SIZE
        Loop
        '--- success
        TlsSend = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lOutputPos, .SendBuffer.Data, .SendBuffer.Size
        pvArrayWriteEOF baOutput, lOutputPos
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lOutputPos <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Output (encrypted)" & vbCrLf & TlsDesignDumpArray(baOutput, Size:=lOutputPos)
            End If
        #End If
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lOutputPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            Exit Function
        End If
        pvTlsClearLastError uCtx
        '--- swap-in
        pvArraySwap .SendBuffer.Data, .SendBuffer.Size, baOutput, lOutputPos
        pvTlsBuildAlert uCtx, .SendBuffer, uscTlsAlertCloseNotify, TLS_ALERT_LEVEL_WARNING
        .State = ucsTlsStateShutdown
        '--- success
        TlsShutdown = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lOutputPos, .SendBuffer.Data, .SendBuffer.Size
        pvArrayWriteEOF baOutput, lOutputPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Number, Err.Source, Err.Description
    Resume QH
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext, Optional LastErrNumber As Long, Optional LastErrSource As String) As String
    LastErrNumber = uCtx.LastErrNumber
    LastErrSource = uCtx.LastErrSource
    TlsGetLastError = uCtx.LastError
    If uCtx.LastAlertCode <> -1 And uCtx.LastAlertCode <> 0 Then
        TlsGetLastError = IIf(LenB(TlsGetLastError) <> 0, TlsGetLastError & ". ", vbNullString) & Replace(STR_FORMAT_ALERT, "%1", pvTlsGetLastAlert(uCtx))
        '--- warnings
        Select Case uCtx.LastAlertCode
        Case uscTlsAlertCloseNotify, uscTlsAlertUserCanceled, uscTlsAlertNoRenegotiation
            LastErrNumber = 0
        End Select
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

Private Function pvTlsGetStateAsText(ByVal lState As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_STATES, "|")
    End If
    If 0 <= lState And lState <= UBound(vTexts) Then
        pvTlsGetStateAsText = vTexts(lState)
    End If
    If LenB(pvTlsGetStateAsText) = 0 Then
        pvTlsGetStateAsText = Replace(STR_UNKNOWN, "%1", lState)
    End If
End Function

Private Function pvTlsGetMessageName(ByVal lMessageType As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_MESSAGE_NAMES, "|")
    End If
    If 0 <= lMessageType And lMessageType <= UBound(vTexts) Then
        pvTlsGetMessageName = vTexts(lMessageType)
    End If
    If LenB(pvTlsGetMessageName) = 0 Then
        pvTlsGetMessageName = Replace(STR_UNKNOWN, "%1", lMessageType)
    Else
        pvTlsGetMessageName = pvTlsGetMessageName & " (" & lMessageType & ")"
    End If
End Function

Private Function pvTlsGetExtensionName(ByVal lExtType As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_EXTENSION_NAMES, "|")
    End If
    If 0 <= lExtType And lExtType <= UBound(vTexts) Then
        pvTlsGetExtensionName = vTexts(lExtType)
    ElseIf lExtType = &HFF01& Then
        pvTlsGetExtensionName = "renegotiation_info"
    End If
    If LenB(pvTlsGetExtensionName) = 0 Then
        pvTlsGetExtensionName = Replace(STR_UNKNOWN, "%1", lExtType)
    Else
        pvTlsGetExtensionName = pvTlsGetExtensionName & " (" & lExtType & ")"
    End If
End Function

Private Sub pvTlsBuildClientHello(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim lMessagePos     As Long
    Dim vElem           As Variant
    Dim baTemp()        As Byte
    
    With uCtx
        If (.LocalFeatures And ucsTlsSupportTls13) <> 0 And .ExchGroup = 0 Then
            '--- populate preferred .ExchGroup and .LocalExchPublic
            If pvCryptoIsSupported(ucsTlsAlgoExchX25519) Then
                pvTlsSetupExchGroup uCtx, TLS_GROUP_X25519
            ElseIf pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                pvTlsSetupExchGroup uCtx, TLS_GROUP_SECP256R1
            ElseIf pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                pvTlsSetupExchGroup uCtx, TLS_GROUP_SECP384R1
            End If
        End If
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Handshake Header
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CLIENT_HELLO
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteLong uOutput, TLS_LOCAL_LEGACY_VERSION, Size:=2
                pvTlsGetRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
                pvBufferWriteArray uOutput, .LocalExchRandom
                '--- Legacy Session ID
                pvBufferWriteBlockStart uOutput
                    If pvArraySize(.LocalSessionID) = 0 And (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- non-empty for TLS 1.2 compatibility
                        pvTlsGetRandom baTemp, TLS_HELLO_RANDOM_SIZE
                        pvBufferWriteArray uOutput, baTemp
                    Else
                        pvBufferWriteArray uOutput, .LocalSessionID
                    End If
                pvBufferWriteBlockEnd uOutput
                '--- Cipher Suites
                pvBufferWriteBlockStart uOutput, Size:=2
                    For Each vElem In pvTlsGetSortedCipherSuites(.LocalFeatures)
                        pvBufferWriteLong uOutput, vElem, Size:=2
                    Next
                pvBufferWriteBlockEnd uOutput
                '--- Legacy Compression Methods
                pvBufferWriteBlockStart uOutput
                    pvBufferWriteLong uOutput, TLS_COMPRESS_NULL
                pvBufferWriteBlockEnd uOutput
                '--- Extensions
                pvBufferWriteBlockStart uOutput, Size:=2
                    If LenB(.RemoteHostName) <> 0 Then
                        '--- Extension - Server Name
                        pvBufferWriteLong uOutput, TLS_EXTENSION_SERVER_NAME, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteLong uOutput, TLS_SERVER_NAME_TYPE_HOSTNAME '--- FQDN
                                pvBufferWriteBlockStart uOutput, Size:=2
                                    pvBufferWriteString uOutput, .RemoteHostName
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                    End If
                    If LenB(.AlpnProtocols) <> 0 Then
                        '--- Extension - ALPN
                        pvBufferWriteLong uOutput, TLS_EXTENSION_ALPN, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                For Each vElem In Split(.AlpnProtocols, "|")
                                    pvBufferWriteBlockStart uOutput
                                        pvBufferWriteString uOutput, Left$(vElem, 255)
                                    pvBufferWriteBlockEnd uOutput
                                Next
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                    End If
                    '--- Extension - Supported Groups
                    pvBufferWriteLong uOutput, TLS_EXTENSION_SUPPORTED_GROUPS, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            If pvCryptoIsSupported(ucsTlsAlgoExchX25519) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_X25519 Then
                                    pvBufferWriteLong uOutput, TLS_GROUP_X25519, Size:=2
                                End If
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP256R1 Then
                                    pvBufferWriteLong uOutput, TLS_GROUP_SECP256R1, Size:=2
                                End If
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP384R1 Then
                                    pvBufferWriteLong uOutput, TLS_GROUP_SECP384R1, Size:=2
                                End If
                            End If
                        pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockEnd uOutput
                    '--- Extension - OCSP Status Request
                    pvArrayByte baTemp, 0, TLS_EXTENSION_STATUS_REQUEST, 0, 5, 1, 0, 0, 0, 0
                    pvBufferWriteArray uOutput, baTemp
                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- Extension - EC Point Formats
                        pvArrayByte baTemp, 0, TLS_EXTENSION_EC_POINT_FORMAT, 0, 2, 1, 0
                        pvBufferWriteArray uOutput, baTemp      '--- uncompressed only
                        '--- Extension - Extended Master Secret
                        pvArrayByte baTemp, 0, TLS_EXTENSION_EXTENDED_MASTER_SECRET, 0, 0
                        pvBufferWriteArray uOutput, baTemp      '--- supported
                        '--- Extension - Encrypt-then-MAC
                        pvArrayByte baTemp, 0, TLS_EXTENSION_ENCRYPT_THEN_MAC, 0, 0
                        pvBufferWriteArray uOutput, baTemp      '--- supported
                        '--- Extension - Renegotiation Info
                        pvBufferWriteLong uOutput, TLS_EXTENSION_RENEGOTIATION_INFO, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput
                                pvBufferWriteArray uOutput, .LocalLegacyVerifyData
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                        '--- Extension - Session Ticket
                        pvBufferWriteLong uOutput, TLS_EXTENSION_SESSION_TICKET, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteLong uOutput, 0, Size:=4
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteArray uOutput, .LocalSessionTicket
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                    End If
                    '--- Extension - Signature Algorithms
                    pvBufferWriteLong uOutput, TLS_EXTENSION_SIGNATURE_ALGORITHMS, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, Size:=2
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1) Then
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512, Size:=2
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoPaddingPss) Then
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_PSS_SHA256, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PSS_PSS_SHA512, Size:=2
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoPaddingPkcs) Then
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PKCS1_SHA224, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PKCS1_SHA384, Size:=2
                                pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PKCS1_SHA512, Size:=2
                            End If
                            '--- legacy SHA-1 based signatures
                            If pvCryptoIsSupported(ucsTlsAlgoDigestSha1) Then
                                If pvCryptoIsSupported(ucsTlsAlgoPaddingPkcs) Then
                                    pvBufferWriteLong uOutput, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2
                                End If
                                If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Or pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                                    pvBufferWriteLong uOutput, TLS_SIGNATURE_ECDSA_SHA1, Size:=2
                                End If
                            End If
                        pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockEnd uOutput
                    If (.LocalFeatures And ucsTlsSupportTls13) <> 0 Then
                        '--- Extension - Post Handshake Auth
                        pvArrayByte baTemp, 0, TLS_EXTENSION_POST_HANDSHAKE_AUTH, 0, 0
                        pvBufferWriteArray uOutput, baTemp     '--- supported
                        '--- Extension - Key Share
                        pvBufferWriteLong uOutput, TLS_EXTENSION_KEY_SHARE, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteLong uOutput, .ExchGroup, Size:=2
                                pvBufferWriteBlockStart uOutput, Size:=2
                                    pvBufferWriteArray uOutput, .LocalExchPublic
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                        '--- Extension - Supported Versions
                        pvBufferWriteLong uOutput, TLS_EXTENSION_SUPPORTED_VERSIONS, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput
                                pvBufferWriteLong uOutput, TLS_PROTOCOL_VERSION_TLS13, Size:=2
                                If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                    pvBufferWriteLong uOutput, TLS_PROTOCOL_VERSION_TLS12, Size:=2
                                End If
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                        If .HelloRetryRequest And SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_COOKIE) Then
                            '--- Extension - Cookie
                            pvBufferWriteLong uOutput, TLS_EXTENSION_COOKIE, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteBlockStart uOutput
                                    pvBufferWriteArray uOutput, .HelloRetryCookie
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        End If
                    End If
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
QH:
End Sub

Private Sub pvTlsBuildClientLegacyKeyExchange(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim lMessagePos     As Long
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim baSignature()   As Byte
    Dim lIdx            As Long
    Dim baCert()        As Byte
    
    With uCtx
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            If .CertRequestSignatureScheme <> 0 Then
                '--- Client Certificate
                lMessagePos = uOutput.Size
                pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE
                pvBufferWriteBlockStart uOutput, Size:=3
                    pvBufferWriteBlockStart uOutput, Size:=3
                        For lIdx = 1 To pvCollectionCount(.LocalCertificates)
                            pvBufferWriteBlockStart uOutput, Size:=3
                                baCert = .LocalCertificates.Item(lIdx)
                                pvBufferWriteArray uOutput, baCert
                            pvBufferWriteBlockEnd uOutput
                        Next
                    pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockEnd uOutput
                pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            End If
            '--- Handshake Client Key Exchange
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CLIENT_KEY_EXCHANGE
            pvBufferWriteBlockStart uOutput, Size:=3
                If pvArraySize(.LocalExchRsaEncrPriv) > 0 Then
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvBufferWriteArray uOutput, .LocalExchRsaEncrPriv
                    pvBufferWriteBlockEnd uOutput
                Else
                    pvBufferWriteBlockStart uOutput
                        pvBufferWriteArray uOutput, .LocalExchPublic
                    pvBufferWriteBlockEnd uOutput
                End If
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            '--- note: get handshake hash early (before certificate verify)
            pvTlsDeriveLegacySecrets uCtx
            If .CertRequestSignatureScheme > 0 Then
                '--- Client Certificate Verify
                lMessagePos = uOutput.Size
                pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE_VERIFY
                pvBufferWriteBlockStart uOutput, Size:=3
                    pvBufferWriteLong uOutput, .CertRequestSignatureScheme, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvArrayWriteEOF .HandshakeMessages.Data, .HandshakeMessages.Size
                        pvTlsSignatureSign baSignature, .LocalCertificates, .LocalPrivateKey, .CertRequestSignatureScheme, .HandshakeMessages.Data
                        pvBufferWriteArray uOutput, baSignature
                    pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockEnd uOutput
                pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            End If
        pvBufferWriteRecordEnd uOutput, uCtx
        '--- Legacy Change Cipher Spec
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, uCtx
            pvBufferWriteLong uOutput, 1
        pvBufferWriteRecordEnd uOutput, uCtx
        '--- commit next epoch local secrets
        .LocalMacKey = .LocalLegacyNextMacKey
        .LocalTrafficKey = .LocalLegacyNextTrafficKey
        .LocalTrafficIV = .LocalLegacyNextTrafficIV
        .LocalTrafficSeqNo = 0
        .LocalEncryptThenMac = (.MacSize > 0) And SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_ENCRYPT_THEN_MAC)
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Client Handshake Finished
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_FINISHED
            pvBufferWriteBlockStart uOutput, Size:=3
                pvTlsGetHandshakeHash uCtx, baHandshakeHash
                pvTlsKdfLegacyPrf baVerifyData, .DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                pvBufferWriteArray uOutput, baVerifyData
                '--- save for secure renegotiation check
                If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_RENEGOTIATION_INFO) Then
                    .LocalLegacyVerifyData = baVerifyData
                End If
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
QH:
End Sub

Private Sub pvTlsBuildClientHandshakeFinished(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim lMessagePos     As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim baSignature()   As Byte
    Dim baHandshakeHash() As Byte
    Dim uVerify         As UcsBuffer
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    With uCtx
        If .CertRequestSignatureScheme <> 0 Then
            '--- Record Header
            pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_APPDATA, uCtx
                '--- Client Certificate
                lMessagePos = uOutput.Size
                pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE
                pvBufferWriteBlockStart uOutput, Size:=3
                    '--- certificate request context
                    pvBufferWriteBlockStart uOutput
                        pvBufferWriteArray uOutput, .CertRequestContext
                    pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockStart uOutput, Size:=3
                        For lIdx = 1 To pvCollectionCount(.LocalCertificates)
                            pvBufferWriteBlockStart uOutput, Size:=3
                                baCert = .LocalCertificates.Item(lIdx)
                                pvBufferWriteArray uOutput, baCert
                            pvBufferWriteBlockEnd uOutput
                            '--- certificate extensions
                            pvBufferWriteBlockStart uOutput, Size:=2
                                '--- empty
                            pvBufferWriteBlockEnd uOutput
                        Next
                    pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockEnd uOutput
                pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
                If .CertRequestSignatureScheme > 0 Then
                    '--- Client Certificate Verify
                    lMessagePos = uOutput.Size
                    pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE_VERIFY
                    pvBufferWriteBlockStart uOutput, Size:=3
                        pvBufferWriteLong uOutput, .CertRequestSignatureScheme, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvTlsGetHandshakeHash uCtx, baHandshakeHash
                            pvBufferWriteString uVerify, Space$(64) & "TLS 1.3, client CertificateVerify" & Chr$(0)
                            pvBufferWriteArray uVerify, baHandshakeHash
                            pvBufferWriteEOF uVerify
                            pvTlsSignatureSign baSignature, .LocalCertificates, .LocalPrivateKey, .CertRequestSignatureScheme, uVerify.Data
                            pvBufferWriteArray uOutput, baSignature
                        pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockEnd uOutput
                    pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
                End If
                '--- Record Type
                pvBufferWriteLong uOutput, TLS_CONTENT_TYPE_HANDSHAKE
            pvBufferWriteRecordEnd uOutput, uCtx
        End If
        '--- Legacy Change Cipher Spec
        pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
        pvBufferWriteArray uOutput, baTemp
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_APPDATA, uCtx
            '--- Client Handshake Finished
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_FINISHED
            pvBufferWriteBlockStart uOutput, Size:=3
                pvTlsGetHandshakeHash uCtx, baHandshakeHash
                pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .LocalTrafficSecret, "finished", baEmpty, .DigestSize
                pvTlsHkdfExtract uVerify.Data, .DigestAlgo, baTemp, baHandshakeHash
                pvBufferWriteArray uOutput, uVerify.Data
            pvBufferWriteBlockEnd uOutput
            '--- Record Type
            pvBufferWriteLong uOutput, TLS_CONTENT_TYPE_HANDSHAKE
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
QH:
End Sub

Private Sub pvTlsBuildServerHello(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim lMessagePos     As Long
    Dim baTemp()        As Byte
    
    With uCtx
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Handshake Header
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_SERVER_HELLO
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteLong uOutput, TLS_LOCAL_LEGACY_VERSION, Size:=2
                pvTlsGetRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
                If .HelloRetryRequest Then
                    If pvArraySize(m_baHelloRetryRandom) = 0 Then
                        pvTlsGetHelloRetryRandom m_baHelloRetryRandom
                    End If
                    pvBufferWriteArray uOutput, m_baHelloRetryRandom
                Else
                    pvBufferWriteArray uOutput, .LocalExchRandom
                End If
                '--- Legacy Session ID
                pvBufferWriteBlockStart uOutput
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        pvBufferWriteArray uOutput, .RemoteSessionID
                    Else
                        pvBufferWriteArray uOutput, .LocalSessionID
                    End If
                pvBufferWriteBlockEnd uOutput
                '--- Cipher Suite
                pvBufferWriteLong uOutput, IIf(.HelloRetryRequest, .HelloRetryCipherSuite, .CipherSuite), Size:=2
                '--- Legacy Compression Method
                pvBufferWriteLong uOutput, TLS_COMPRESS_NULL
                '--- Extensions
                pvBufferWriteBlockStart uOutput, Size:=2
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        '--- Extension - Key Share
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_KEY_SHARE) Or .HelloRetryRequest Then
                            pvBufferWriteLong uOutput, TLS_EXTENSION_KEY_SHARE, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                If .HelloRetryRequest Then
                                    pvBufferWriteLong uOutput, .HelloRetryExchGroup, Size:=2
                                Else
                                    pvBufferWriteLong uOutput, .ExchGroup, Size:=2
                                    pvBufferWriteBlockStart uOutput, Size:=2
                                        pvBufferWriteArray uOutput, .LocalExchPublic
                                    pvBufferWriteBlockEnd uOutput
                                End If
                            pvBufferWriteBlockEnd uOutput
                        End If
                        '--- Extension - Supported Versions
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_SUPPORTED_VERSIONS) Or .HelloRetryRequest Then
                            pvBufferWriteLong uOutput, TLS_EXTENSION_SUPPORTED_VERSIONS, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteLong uOutput, TLS_PROTOCOL_VERSION_TLS13, Size:=2
                            pvBufferWriteBlockEnd uOutput
                        End If
                        If .HelloRetryRequest And pvArraySize(.HelloRetryCookie) > 0 Then
                            '--- Extension - HRR Cookie
                            pvBufferWriteLong uOutput, TLS_EXTENSION_COOKIE, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteBlockStart uOutput, Size:=2
                                    pvBufferWriteArray uOutput, .HelloRetryCookie
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        End If
                    End If
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_EXTENDED_MASTER_SECRET) Then
                            pvArrayByte baTemp, 0, TLS_EXTENSION_EXTENDED_MASTER_SECRET, 0, 0
                            pvBufferWriteArray uOutput, baTemp      '--- supported
                        End If
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_ENCRYPT_THEN_MAC) Then
                            pvArrayByte baTemp, 0, TLS_EXTENSION_ENCRYPT_THEN_MAC, 0, 0
                            pvBufferWriteArray uOutput, baTemp      '--- supported
                        End If
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_EC_POINT_FORMAT) Then
                            pvArrayByte baTemp, 0, TLS_EXTENSION_EC_POINT_FORMAT, 0, 2, 1, 0
                            pvBufferWriteArray uOutput, baTemp      '--- uncompressed only
                        End If
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_RENEGOTIATION_INFO) Then
                            pvBufferWriteLong uOutput, TLS_EXTENSION_RENEGOTIATION_INFO, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteBlockStart uOutput
                                    pvBufferWriteArray uOutput, .RemoteLegacyVerifyData
                                    pvBufferWriteArray uOutput, .LocalLegacyVerifyData
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        End If
                        If LenB(.AlpnNegotiated) <> 0 Then
                            pvBufferWriteLong uOutput, TLS_EXTENSION_ALPN, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteBlockStart uOutput, Size:=2
                                    pvBufferWriteBlockStart uOutput
                                        pvBufferWriteString uOutput, .AlpnNegotiated
                                    pvBufferWriteBlockEnd uOutput
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        End If
                    End If
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
        If (.HelloRetryRequest Or .HelloRetryCipherSuite = 0) And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            '--- Legacy Change Cipher Spec
            pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
            pvBufferWriteArray uOutput, baTemp
        End If
    End With
End Sub

Private Sub pvTlsBuildServerLegacyKeyExchange(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim lMessagePos     As Long
    Dim baSignature()   As Byte
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim uVerify         As UcsBuffer
    Dim lSignPos        As Long
    Dim lSignSize       As Long
    
    With uCtx
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Server Certificate
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteBlockStart uOutput, Size:=3
                    For lIdx = 1 To pvCollectionCount(.LocalCertificates)
                        pvBufferWriteBlockStart uOutput, Size:=3
                            baCert = .LocalCertificates.Item(lIdx)
                            pvBufferWriteArray uOutput, baCert
                        pvBufferWriteBlockEnd uOutput
                    Next
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
        If Not .UseRsaKeyTransport Then
            '--- Record Header
            pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
                '--- Server Key Exchange
                lMessagePos = uOutput.Size
                pvBufferWriteLong uOutput, TLS_HANDSHAKE_SERVER_KEY_EXCHANGE
                pvBufferWriteBlockStart uOutput, Size:=3
                    '--- Curve Info
                    lSignPos = uOutput.Size
                    pvBufferWriteLong uOutput, 3, Size:=1 '--- 3 = named_curve
                    pvBufferWriteLong uOutput, .ExchGroup, Size:=2
                    '--- Public Key
                    pvBufferWriteBlockStart uOutput
                        pvBufferWriteArray uOutput, .LocalExchPublic
                    pvBufferWriteBlockEnd uOutput
                    lSignSize = uOutput.Size - lSignPos
                    '--- Signature
                    pvBufferWriteArray uVerify, .RemoteExchRandom
                    pvBufferWriteArray uVerify, .LocalExchRandom
                    pvBufferWriteBlob uVerify, VarPtr(uOutput.Data(lSignPos)), lSignSize
                    pvBufferWriteEOF uVerify
                    pvTlsSignatureSign baSignature, .LocalCertificates, .LocalPrivateKey, .SignatureScheme, uVerify.Data
                    pvBufferWriteLong uOutput, .SignatureScheme, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvBufferWriteArray uOutput, baSignature
                    pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockEnd uOutput
                pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            pvBufferWriteRecordEnd uOutput, uCtx
        End If
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Server Hello Done
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_SERVER_HELLO_DONE
            pvBufferWriteBlockStart uOutput, Size:=3
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
End Sub

Private Sub pvTlsBuildServerHandshakeFinished(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim baHandshakeHash() As Byte
    Dim lMessagePos     As Long
    Dim uVerify         As UcsBuffer
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim baSignature()   As Byte
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    With uCtx
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_APPDATA, uCtx
            '--- Server Encrypted Extensions
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_ENCRYPTED_EXTENSIONS
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteBlockStart uOutput, Size:=2
                    If LenB(.AlpnNegotiated) <> 0 Then
                        pvBufferWriteLong uOutput, TLS_EXTENSION_ALPN, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput, Size:=2
                                pvBufferWriteBlockStart uOutput
                                    pvBufferWriteString uOutput, .AlpnNegotiated
                                pvBufferWriteBlockEnd uOutput
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                    End If
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            '--- Server Certificate
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE
            pvBufferWriteBlockStart uOutput, Size:=3
                '--- certificate request context
                pvBufferWriteBlockStart uOutput
                    '--- empty
                pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockStart uOutput, Size:=3
                    For lIdx = 1 To pvCollectionCount(.LocalCertificates)
                        pvBufferWriteBlockStart uOutput, Size:=3
                            baCert = .LocalCertificates.Item(lIdx)
                            pvBufferWriteArray uOutput, baCert
                        pvBufferWriteBlockEnd uOutput
                        '--- certificate extensions
                        pvBufferWriteBlockStart uOutput, Size:=2
                            '--- empty
                        pvBufferWriteBlockEnd uOutput
                    Next
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            '--- Server Certificate Verify
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE_VERIFY
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteLong uOutput, .SignatureScheme, Size:=2
                pvBufferWriteBlockStart uOutput, Size:=2
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvBufferWriteString uVerify, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0)
                    pvBufferWriteArray uVerify, baHandshakeHash
                    pvBufferWriteEOF uVerify
                    pvTlsSignatureSign baSignature, .LocalCertificates, .LocalPrivateKey, .SignatureScheme, uVerify.Data
                    pvBufferWriteArray uOutput, baSignature
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            '--- Server Handshake Finished
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_FINISHED
            pvBufferWriteBlockStart uOutput, Size:=3
                pvTlsGetHandshakeHash uCtx, baHandshakeHash
                pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .LocalTrafficSecret, "finished", baEmpty, .DigestSize
                pvTlsHkdfExtract uVerify.Data, .DigestAlgo, baTemp, baHandshakeHash
                pvBufferWriteArray uOutput, uVerify.Data
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            '--- Record Type
            pvBufferWriteLong uOutput, TLS_CONTENT_TYPE_HANDSHAKE
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
End Sub

Private Sub pvTlsBuildServerLegacyFinished(uCtx As UcsTlsContext, uOutput As UcsBuffer)
    Dim baHandshakeHash() As Byte
    Dim uVerify         As UcsBuffer
    
    With uCtx
        '--- Change Cipher Spec
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, uCtx
            pvBufferWriteLong uOutput, 1
        pvBufferWriteRecordEnd uOutput, uCtx
        '--- commit next epoch local secrets
        .LocalMacKey = .LocalLegacyNextMacKey
        .LocalTrafficKey = .LocalLegacyNextTrafficKey
        .LocalTrafficIV = .LocalLegacyNextTrafficIV
        .LocalTrafficSeqNo = 0
        .LocalEncryptThenMac = (.MacSize > 0) And SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_ENCRYPT_THEN_MAC)
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Server Handshake Finished
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_FINISHED
            pvBufferWriteBlockStart uOutput, Size:=3
                pvTlsGetHandshakeHash uCtx, baHandshakeHash
                pvTlsKdfLegacyPrf uVerify.Data, .DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_RENEGOTIATION_INFO) Then
                    .LocalLegacyVerifyData = uVerify.Data
                End If
                pvBufferWriteArray uOutput, uVerify.Data
            pvBufferWriteBlockEnd uOutput
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
End Sub

Private Sub pvTlsBuildApplicationData(uCtx As UcsTlsContext, uOutput As UcsBuffer, baData() As Byte, ByVal lPos As Long, ByVal lSize As Long, ByVal lContentType As Long)
    With uCtx
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_APPDATA, uCtx
            If lSize > 0 Then
                pvBufferWriteBlob uOutput, VarPtr(baData(lPos)), lSize
            End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                '--- Record Type
                pvBufferWriteLong uOutput, lContentType
            End If
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
End Sub

Private Sub pvTlsBuildAlert(uCtx As UcsTlsContext, uOutput As UcsBuffer, ByVal eAlertDesc As UcsTlsAlertDescriptionsEnum, ByVal lAlertLevel As Long)
    Dim baHandshakeHash() As Byte
    Dim baTemp()        As Byte
    
    With uCtx
#If ImplTlsServer Then
        If .State = ucsTlsStateExpectClientFinished And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            '--- alerts must be protected with application data traffic secrets (not handshake)
            .State = ucsTlsStatePostHandshake
            pvTlsGetHandshakeHash uCtx, baHandshakeHash
            pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
        End If
#End If
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            '--- for TLS 1.3 -> tunnel alert through application data traffic protection
            pvArrayByte baTemp, lAlertLevel, eAlertDesc
            pvTlsBuildApplicationData uCtx, uOutput, baTemp, 0, UBound(baTemp) + 1, TLS_CONTENT_TYPE_ALERT
            GoTo QH
        End If
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_ALERT, uCtx
            pvBufferWriteLong uOutput, lAlertLevel
            pvBufferWriteLong uOutput, eAlertDesc
        pvBufferWriteRecordEnd uOutput, uCtx
    End With
QH:
End Sub

Private Function pvTlsParsePayload(uCtx As UcsTlsContext, baInput() As Byte, lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lNewSize        As Long
    
    On Error GoTo EH
    If lSize > 0 Then
    With uCtx
        pvBufferWriteBlob .RecvBuffer, VarPtr(baInput(0)), lSize
        If Not pvTlsParseRecord(uCtx, .RecvBuffer, sError, eAlertCode) Then
            GoTo QH
        End If
        lNewSize = .RecvBuffer.Size - .RecvBuffer.Pos
        If lNewSize > 0 Then
            Debug.Assert pvArraySize(.RecvBuffer.Data) >= .RecvBuffer.Pos + lNewSize
            Call CopyMemory(.RecvBuffer.Data(0), .RecvBuffer.Data(.RecvBuffer.Pos), lNewSize)
        End If
        .RecvBuffer.Pos = 0
        .RecvBuffer.Size = IIf(lNewSize > 0, lNewSize, 0)
    End With
    End If
    '--- success
    pvTlsParsePayload = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsParseRecord(uCtx As UcsTlsContext, uInput As UcsBuffer, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseRecord"
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim lRecordProtocol As Long
    Dim baRemoteIV()    As Byte
    Dim lEnd            As Long
    Dim uAad            As UcsBuffer
    Dim bResult         As Boolean
    Dim baHmac()        As Byte
    Dim lPadding        As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    With uCtx
    Do While uInput.Pos + 4 < uInput.Size
        lRecordPos = uInput.Pos
        pvBufferReadLong uInput, lRecordType
        If lRecordType < TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC Or lRecordType > TLS_CONTENT_TYPE_APPDATA Then
            GoTo UnexpectedRecordType
        End If
        pvBufferReadLong uInput, lRecordProtocol, Size:=2
        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lRecordSize
            If uInput.Pos + lRecordSize > uInput.Size Then
                '--- back off and bail out early
                uInput.Stack.Remove 1
                uInput.Pos = lRecordPos
                Exit Do
            End If
            '--- try to decrypt record
            If pvArraySize(.RemoteTrafficKey) > 0 Then
                '--- check ciphertext size
                If lRecordSize > TLS_MAX_ENCRYPTED_RECORD_SIZE + .MacSize + .IvExplicitSize Then
                    GoTo RecordOverflow
                End If
                If lRecordSize < .TagSize + .MacSize Then
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        GoTo Unencrypted
                    End If
                    GoTo RecordMacFailed
                End If
                lEnd = uInput.Pos + lRecordSize - .TagSize
                bResult = False
                pvArrayXor baRemoteIV, .RemoteTrafficIV, .RemoteTrafficSeqNo
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    If lRecordType <> TLS_CONTENT_TYPE_APPDATA Then
                        GoTo UnexpectedRecordType
                    End If
                    bResult = pvTlsBulkDecrypt(.BulkAlgo, baRemoteIV, .RemoteTrafficKey, uInput.Data, lRecordPos, TLS_AAD_SIZE, uInput.Data, uInput.Pos, lRecordSize)
                ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    If .IvExplicitSize > 0 Then '--- AES in TLS 1.2
                        pvArrayWriteBlob baRemoteIV, .IvSize - .IvExplicitSize, VarPtr(uInput.Data(uInput.Pos)), .IvExplicitSize
                        uInput.Pos = uInput.Pos + .IvExplicitSize
                    End If
                    uAad.Size = 0
                    pvBufferWriteLong uAad, 0, Size:=4
                    pvBufferWriteLong uAad, .RemoteTrafficSeqNo, Size:=4
                    pvBufferWriteBlob uAad, VarPtr(uInput.Data(lRecordPos)), 3
                    If Not .RemoteEncryptThenMac Then
                        pvBufferWriteLong uAad, lEnd - uInput.Pos, Size:=2
                        Debug.Assert uAad.Size = TLS_LEGACY_AAD_SIZE
                    Else
                        lEnd = lEnd - .MacSize
                        '--- prepare encrypted data for MAC
                        pvBufferWriteLong uAad, lEnd - uInput.Pos + .IvExplicitSize, Size:=2
                        pvBufferWriteBlob uAad, VarPtr(uInput.Data(uInput.Pos - .IvExplicitSize)), lEnd - uInput.Pos + .IvExplicitSize
                    End If
                    bResult = pvTlsBulkDecrypt(.BulkAlgo, baRemoteIV, .RemoteTrafficKey, uAad.Data, 0, uAad.Size, uInput.Data, uInput.Pos, lEnd - uInput.Pos + .TagSize)
                End If
                If Not bResult Then
                    GoTo DecryptionFailed
                End If
                #If (ImplCaptureTraffic And 1) <> 0 Then
                    If lEnd - uInput.Pos <> 0 Then
                        .TrafficDump.Add FUNC_NAME & ".Input (decrypted)" & vbCrLf & TlsDesignDumpArray(uInput.Data, lRecordPos, lEnd - lRecordPos)
                    End If
                #End If
                .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    If lEnd - uInput.Pos > TLS_MAX_PLAINTEXT_RECORD_SIZE + 1 Then
                        GoTo RecordOverflow
                    End If
                    '--- trim zero padding at the end of decrypted record
                    Do While lEnd > uInput.Pos
                        lEnd = lEnd - 1
                        If uInput.Data(lEnd) <> 0 Then
                            Exit Do
                        End If
                    Loop
                    lRecordType = uInput.Data(lEnd)
                ElseIf .MacSize > 0 Then
                    If Not .RemoteEncryptThenMac Then
                        '--- remove padding and prepare decrypted data for MAC
                        lPadding = uInput.Data(lEnd - 1)
                        If lEnd - (lPadding + 1) - .MacSize < uInput.Pos Then
                            GoTo RecordMacFailed
                        End If
                        For lIdx = 2 To lPadding + 1
                            If uInput.Data(lEnd - lIdx) <> lPadding Then
                                GoTo RecordMacFailed
                            End If
                        Next
                        lEnd = lEnd - (lPadding + 1) - .MacSize
                        uAad.Size = uAad.Size - 2
                        pvBufferWriteLong uAad, lEnd - uInput.Pos, Size:=2
                        pvBufferWriteBlob uAad, VarPtr(uInput.Data(uInput.Pos)), lEnd - uInput.Pos
                    End If
                    '--- calc MAC and compare
                    pvTlsGetHmac baHmac, .MacAlgo, .RemoteMacKey, uAad.Data, 0, uAad.Size
                    pvArrayAllocate baRemoteIV, .MacSize, FUNC_NAME & ".baRemoteIV"
                    Call CopyMemory(baRemoteIV(0), ByVal VarPtr(uInput.Data(lEnd)), .MacSize)
                    If Not pvArrayEqual(baHmac, baRemoteIV) Then
                        GoTo RecordMacFailed
                    End If
                    If .RemoteEncryptThenMac Then
                        '--- remove padding from decrypted data
                        lPadding = uInput.Data(lEnd - 1)
                        If lEnd - (lPadding + 1) < uInput.Pos Then
                            GoTo RecordMacFailed
                        End If
                        For lIdx = 2 To lPadding + 1
                            If uInput.Data(lEnd - lIdx) <> lPadding Then
                                GoTo RecordMacFailed
                            End If
                        Next
                        lEnd = lEnd - (lPadding + 1)
                    End If
                End If
            Else
Unencrypted:
                lEnd = uInput.Pos + lRecordSize
            End If
            '--- check plaintext size
            If lEnd - uInput.Pos > TLS_MAX_PLAINTEXT_RECORD_SIZE Then
                GoTo RecordOverflow
            End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                If uInput.Pos + 1 <> lEnd Then
                    GoTo UnexpectedRecordSize
                End If
                If .State = ucsTlsStatePostHandshake Then
                    GoTo UnexpectedRecordType
                End If
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    If pvArraySize(.RemoteLegacyNextTrafficKey) = 0 Then
                        GoTo UnexpectedRecordType
                    End If
                    '--- commit next epoch remote secrets
                    .RemoteMacKey = .RemoteLegacyNextMacKey
                    .RemoteTrafficKey = .RemoteLegacyNextTrafficKey
                    .RemoteTrafficIV = .RemoteLegacyNextTrafficIV
                    .RemoteTrafficSeqNo = 0
                    .RemoteEncryptThenMac = (.MacSize > 0) And SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_ENCRYPT_THEN_MAC)
                    Erase .RemoteLegacyNextTrafficKey
                End If
            Case TLS_CONTENT_TYPE_ALERT
                If uInput.Pos + 2 <> lEnd Then
                    GoTo UnexpectedRecordSize
                End If
                Select Case uInput.Data(uInput.Pos)
                Case TLS_ALERT_LEVEL_FATAL
                    sError = ERR_FATAL_ALERT
                    eAlertCode = uInput.Data(uInput.Pos + 1)
                    GoTo QH
                Case TLS_ALERT_LEVEL_WARNING
                    .LastAlertCode = uInput.Data(uInput.Pos + 1)
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, pvTlsGetLastAlert(uCtx) & " (TLS_ALERT_LEVEL_WARNING)"
                    #End If
                    If .LastAlertCode = uscTlsAlertCloseNotify And .State <> ucsTlsStateShutdown Then
                        pvTlsSetLastError uCtx, AlertCode:=uscTlsAlertCloseNotify
                    End If
                End Select
            Case TLS_CONTENT_TYPE_HANDSHAKE
                If uInput.Pos = lEnd Then
                    GoTo UnexpectedRecordSize
                End If
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    '--- RFC 8446 section 5.1: Handshake messages MUST NOT be interleaved with other record types
                    If .PrevRecordType <> 0 And .PrevRecordType <> lRecordType Then
                        GoTo UnexpectedRecordType
                    End If
                End If
                If .MessBuffer.Size > 0 Then
                    pvBufferWriteBlob .MessBuffer, VarPtr(uInput.Data(uInput.Pos)), lEnd - uInput.Pos
                    If Not pvTlsParseHandshake(uCtx, .MessBuffer, .MessBuffer.Size, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If .MessBuffer.Pos >= .MessBuffer.Size Then
                        Erase .MessBuffer.Data
                        .MessBuffer.Size = 0
                        .MessBuffer.Pos = 0
                    End If
                Else
                    If Not pvTlsParseHandshake(uCtx, uInput, lEnd, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If uInput.Pos < lEnd Then
                        pvBufferWriteBlob .MessBuffer, VarPtr(uInput.Data(uInput.Pos)), lEnd - uInput.Pos
                        .MessBuffer.Pos = 0
                    End If
                End If
            Case TLS_CONTENT_TYPE_APPDATA
                If .IsServer And .State < ucsTlsStatePostHandshake Then
                    GoTo UnexpectedRecordType
                End If
                pvBufferWriteBlob .DecrBuffer, VarPtr(uInput.Data(uInput.Pos)), lEnd - uInput.Pos
                .PrevRecordType = lRecordType
            Case Else
                GoTo UnexpectedRecordType
            End Select
            '--- note: skip AEAD's authentication tag or zero padding
            uInput.Pos = lRecordPos + lRecordSize + 5
        pvBufferReadBlockEnd uInput
    Loop
    End With
    '--- success
    pvTlsParseRecord = True
QH:
    Exit Function
RecordOverflow:
    sError = ERR_RECORD_OVERFLOW
    eAlertCode = uscTlsAlertRecordOverflow
    GoTo QH
DecryptionFailed:
    sError = ERR_DECRYPTION_FAILED
    eAlertCode = uscTlsAlertBadRecordMac
    GoTo QH
UnexpectedRecordType:
    sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
    eAlertCode = uscTlsAlertUnexpectedMessage
    GoTo QH
UnexpectedRecordSize:
    sError = ERR_RECORD_OVERFLOW
    eAlertCode = uscTlsAlertUnexpectedMessage
    GoTo QH
RecordMacFailed:
    sError = ERR_RECORD_MAC_FAILED
    eAlertCode = uscTlsAlertBadRecordMac
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsParseHandshake(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshake"
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lMessageEnd     As Long
    Dim lMessageType    As Long
    Dim baMessage()     As Byte
    Dim baHandshakeHash() As Byte
    Dim uVerify         As UcsBuffer
    Dim lRequestUpdate  As Long
    Dim lCurveType      As Long
    Dim lNamedCurve     As Long
    Dim lSignatureScheme As Long
    Dim lSignatureSize  As Long
    Dim baSignature()   As Byte
    Dim baCert()        As Byte
    Dim lCertSize       As Long
    Dim lCertEnd        As Long
    Dim lSignPos        As Long
    Dim lSignSize       As Long
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim lStringSize     As Long
    Dim lExchGroup      As Long
    Dim lStatusType     As Long
    
    On Error GoTo EH
    lExtType = -1
    With uCtx
    Do While uInput.Pos + 3 < lInputEnd
        lMessagePos = uInput.Pos
        pvBufferReadLong uInput, lMessageType
        pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lMessageSize
            lMessageEnd = uInput.Pos + lMessageSize
            If lMessageEnd > lInputEnd Then
                '--- back off and bail out early
                uInput.Stack.Remove 1
                uInput.Pos = lMessagePos
                Exit Do
            End If
            #If ImplUseDebugLog Then
'                DebugLog MODULE_NAME, FUNC_NAME, ".State=" & pvTlsGetStateAsText(.State) & ", lMessageType=" & pvTlsGetMessageName(lMessageType)
            #End If
            Select Case .State
            Case ucsTlsStateExpectServerHello
                Select Case lMessageType
                Case TLS_HANDSHAKE_SERVER_HELLO
                    If Not pvTlsParseHandshakeServerHello(uCtx, uInput, uInput.Pos + lMessageSize, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If .HelloRetryRequest Then
                        '--- on HelloRetryRequest replace HandshakeMessages w/ 'synthetic handshake message'
                        pvTlsGetHandshakeHash uCtx, baHandshakeHash
                        .HandshakeMessages.Size = 0
                        pvBufferWriteLong .HandshakeMessages, TLS_HANDSHAKE_MESSAGE_HASH
                        pvBufferWriteLong .HandshakeMessages, .DigestSize, Size:=3
                        pvBufferWriteArray .HandshakeMessages, baHandshakeHash
                    Else
                        .State = ucsTlsStateExpectEncryptedExtensions
                    End If
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                '--- post-process ucsTlsStateExpectServerHello
                If .State = ucsTlsStateExpectServerHello And .HelloRetryRequest Then
                    pvTlsBuildClientHello uCtx, .SendBuffer
                End If
                If .State = ucsTlsStateExpectEncryptedExtensions And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    pvTlsDeriveHandshakeSecrets uCtx
                End If
            Case ucsTlsStateExpectEncryptedExtensions
                Select Case lMessageType
                Case TLS_HANDSHAKE_ENCRYPTED_EXTENSIONS
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                    lBlockEnd = uInput.Pos + lBlockSize
                    If lBlockEnd <> lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    Do While uInput.Pos + 3 < lBlockEnd
                        pvBufferReadLong uInput, lExtType, Size:=2
                        #If ImplUseDebugLog Then
'                            DebugLog MODULE_NAME, FUNC_NAME, "EncryptedExtensions " & pvTlsGetExtensionName(lExtType)
                        #End If
                        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                            lExtEnd = uInput.Pos + lExtSize
                            If lExtEnd > lBlockEnd Then
                                GoTo InvalidSize
                            End If
                            Select Case lExtType
                            Case TLS_EXTENSION_ALPN
                                If lExtSize < 2 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lStringSize
                                    If uInput.Pos + lStringSize <> lExtEnd Or lStringSize = 0 Then
                                        GoTo InvalidSize
                                    End If
                                    pvBufferReadBlockStart uInput, BlockSize:=lStringSize
                                        If uInput.Pos + lStringSize <> lExtEnd Or lStringSize = 0 Then
                                            GoTo InvalidSize
                                        End If
                                        pvBufferReadString uInput, .AlpnNegotiated, lStringSize
                                    pvBufferReadBlockEnd uInput
                                pvBufferReadBlockEnd uInput
                            Case TLS_EXTENSION_SUPPORTED_GROUPS
                                If lExtSize < 2 Then
                                    GoTo InvalidSize
                                End If
                                Set .RemoteSupportedGroups = New Collection
                                Do While uInput.Pos + 1 < lExtEnd
                                    pvBufferReadLong uInput, lExchGroup, Size:=2
                                    .RemoteSupportedGroups.Add lExchGroup, "#" & lExchGroup
                                Loop
                            Case Else
                                uInput.Pos = uInput.Pos + lExtSize
                            End Select
                        pvBufferReadBlockEnd uInput
                    Loop
                    pvBufferReadBlockEnd uInput
                Case TLS_HANDSHAKE_CERTIFICATE
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        pvBufferReadBlockStart uInput, BlockSize:=lCertSize
                            If uInput.Pos + lCertSize > lMessageEnd Then
                                GoTo InvalidSize
                            End If
                            uInput.Pos = uInput.Pos + lCertSize '--- skip RemoteCertReqContext
                        pvBufferReadBlockEnd uInput
                    End If
                    Set .RemoteCertificates = New Collection
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        Set .RemoteCertStatuses = New Collection
                    End If
                    If uInput.Pos + 3 > lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lCertSize
                        lCertEnd = uInput.Pos + lCertSize
                        If lCertEnd <> lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        Do While uInput.Pos + 2 < lCertEnd
                            pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lCertSize
                                If uInput.Pos + lCertSize > lCertEnd Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadArray uInput, baCert, lCertSize
                                .RemoteCertificates.Add baCert
                            pvBufferReadBlockEnd uInput
                            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                baCert = vbNullString
                                If uInput.Pos + 2 > lCertEnd Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                If lBlockEnd > lCertEnd Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos + 1 < lBlockEnd
                                    pvBufferReadLong uInput, lExtType, Size:=2
                                    #If ImplUseDebugLog Then
'                                        DebugLog MODULE_NAME, FUNC_NAME, "CertificateExtensions " & pvTlsGetExtensionName(lExtType)
                                    #End If
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                                        lExtEnd = uInput.Pos + lExtSize
                                        If lExtEnd > lBlockEnd Then
                                            GoTo InvalidSize
                                        End If
                                        Select Case lExtType
                                        Case TLS_EXTENSION_STATUS_REQUEST
                                            pvBufferReadArray uInput, baCert, lExtSize
                                        Case Else
                                            uInput.Pos = uInput.Pos + lExtSize
                                        End Select
                                    pvBufferReadBlockEnd uInput
                                Loop
                                pvBufferReadBlockEnd uInput
                                .RemoteCertStatuses.Add baCert
                            End If
                        Loop
                    pvBufferReadBlockEnd uInput
                Case TLS_HANDSHAKE_CERTIFICATE_VERIFY
                    If uInput.Pos + 2 > lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    pvBufferReadLong uInput, lSignatureScheme, Size:=2
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSignatureSize
                        If uInput.Pos + lSignatureSize <> lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadArray uInput, baSignature, lSignatureSize
                    pvBufferReadBlockEnd uInput
                    If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                        GoTo NoServerCertificate
                    End If
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvBufferWriteString uVerify, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0)
                    pvBufferWriteArray uVerify, baHandshakeHash
                    pvBufferWriteEOF uVerify
                    If Not pvTlsSignatureVerify(baCert, lSignatureScheme, uVerify.Data, baSignature, sError, eAlertCode) Then
                        GoTo QH
                    End If
                Case TLS_HANDSHAKE_FINISHED
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .RemoteTrafficSecret, "finished", baEmpty, .DigestSize
                    pvTlsHkdfExtract uVerify.Data, .DigestAlgo, baTemp, baHandshakeHash
                    If Not pvArrayEqual(uVerify.Data, baMessage) Then
                        GoTo ServerHandshakeFailed
                    End If
                    .State = ucsTlsStatePostHandshake
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12, TLS_HANDSHAKE_SERVER_KEY_EXCHANGE, -1)
                    If uInput.Pos + 4 > lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    lSignPos = uInput.Pos
                    pvBufferReadLong uInput, lCurveType
                    If lCurveType <> 3 Then '--- 3 = named_curve
                        GoTo UnsupportedCurveType
                    End If
                    pvBufferReadLong uInput, lNamedCurve, Size:=2
                    pvTlsSetupExchGroup uCtx, lNamedCurve
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "With exchange group " & pvTlsGetExchGroupName(.ExchGroup)
                    #End If
                    pvBufferReadBlockStart uInput, BlockSize:=lSignatureSize
                        If uInput.Pos + lSignatureSize > lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadArray uInput, .RemoteExchPublic, lSignatureSize
                        If Not pvTlsCheckRemoteKey(.ExchGroup, .RemoteExchPublic) Then
                            GoTo InvalidRemoteKey
                        End If
                    pvBufferReadBlockEnd uInput
                    lSignSize = uInput.Pos - lSignPos
                    '--- signature
                    If uInput.Pos + 4 > lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    pvBufferReadLong uInput, lSignatureScheme, Size:=2
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSignatureSize
                        If uInput.Pos + lSignatureSize <> lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadArray uInput, baSignature, lSignatureSize
                    pvBufferReadBlockEnd uInput
                    If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                        GoTo NoServerCertificate
                    End If
                    pvBufferWriteArray uVerify, .LocalExchRandom
                    pvBufferWriteArray uVerify, .RemoteExchRandom
                    pvBufferWriteBlob uVerify, VarPtr(uInput.Data(lSignPos)), lSignSize
                    pvBufferWriteEOF uVerify
                    If Not pvTlsSignatureVerify(baCert, lSignatureScheme, uVerify.Data, baSignature, sError, eAlertCode) Then
                        GoTo QH
                    End If
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12, TLS_HANDSHAKE_SERVER_HELLO_DONE, -1)
                    .State = ucsTlsStateExpectServerFinished
                    uInput.Pos = uInput.Pos + lMessageSize
                    Set .RemoteTickets = New Collection
                Case TLS_HANDSHAKE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, uInput, sError, eAlertCode) Then
                        GoTo QH
                    End If
                Case TLS_HANDSHAKE_CERTIFICATE_STATUS
                    If uInput.Pos + 1 > lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    Set .RemoteCertStatuses = New Collection
                    pvBufferReadLong uInput, lStatusType, Size:=1
                    If lStatusType = 1 Then '--- ocsp
                        If uInput.Pos + 3 > lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lCertSize
                            If uInput.Pos + lCertSize <> lMessageEnd Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadArray uInput, baCert, lCertSize
                            .RemoteCertStatuses.Add baCert
                        pvBufferReadBlockEnd uInput
                    Else
                        #If ImplUseDebugLog Then
                            DebugLog MODULE_NAME, FUNC_NAME, Replace("Unknown status_type (%1) in certificate_status", "%1", lStatusType)
                        #End If
                    End If
                Case Else
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, Replace("Unknown handshake message (%1) during ExpectEncryptedExtensions", "%1", pvTlsGetMessageName(lMessageType))
                    #End If
                    '--- do nothing
                    uInput.Pos = uInput.Pos + lMessageSize
                End Select
                pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                '--- post-process ucsTlsStateExpectEncryptedExtensions
                If .State = ucsTlsStateExpectServerFinished And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    If .UseRsaKeyTransport Then
                        If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                            GoTo NoServerCertificate
                        End If
                        pvTlsSetupExchRsaCertificate uCtx, baCert
                    End If
                    '--- secure renegotiation check
                    If pvArraySize(.LocalLegacyVerifyData) > 0 Then
                        If pvArraySize(.RemoteLegacyRenegInfo) <> 2 * TLS_VERIFY_DATA_SIZE Or pvArraySize(.RemoteLegacyVerifyData) <> TLS_VERIFY_DATA_SIZE Then
                            GoTo SecureRenegotiationFailed
                        ElseIf InStrB(.RemoteLegacyRenegInfo, .LocalLegacyVerifyData) <> 1 Then
                            GoTo SecureRenegotiationFailed
                        ElseIf InStrB(TLS_VERIFY_DATA_SIZE + 1, .RemoteLegacyRenegInfo, .RemoteLegacyVerifyData) <> TLS_VERIFY_DATA_SIZE + 1 Then
                            GoTo SecureRenegotiationFailed
                        End If
                    End If
                    pvTlsBuildClientLegacyKeyExchange uCtx, .SendBuffer
                End If
                If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvTlsBuildClientHandshakeFinished uCtx, .SendBuffer
                    pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
                    pvTlsResetHandshakeHash uCtx
                End If
            Case ucsTlsStateExpectServerFinished
                Select Case lMessageType
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12, TLS_HANDSHAKE_NEW_SESSION_TICKET, -1)
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    If Not .RemoteTickets Is Nothing Then
                        .RemoteTickets.Add baMessage
                    End If
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12, TLS_HANDSHAKE_FINISHED, -1)
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvTlsKdfLegacyPrf uVerify.Data, .DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                    If Not pvArrayEqual(uVerify.Data, baMessage) Then
                        GoTo ServerHandshakeFailed
                    End If
                    '--- save for secure renegotiation check
                    If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_RENEGOTIATION_INFO) Then
                        .RemoteLegacyVerifyData = baMessage
                    End If
                    .State = ucsTlsStatePostHandshake
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                If .State = ucsTlsStatePostHandshake Then
                    pvTlsResetHandshakeHash uCtx
                Else
                    pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                End If
#If ImplTlsServer Then
            Case ucsTlsStateExpectClientHello
RenegotiateClientHello:
                Select Case lMessageType
                Case TLS_HANDSHAKE_CLIENT_HELLO
                    If Not pvTlsParseHandshakeClientHello(uCtx, uInput, lMessageEnd, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If uInput.Pos <> lMessageEnd Then
                        GoTo InvalidSize
                    End If
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        If (.ExchGroup = 0 Or .CipherSuite = 0) Then
                            If .HelloRetryRequest Then
                                GoTo HelloRetryFailed
                            End If
                            .HelloRetryRequest = True
                            If .ExchGroup <> 0 Then
                                .HelloRetryExchGroup = .ExchGroup
                            Else
                                .HelloRetryExchGroup = pvCollectionFirst(.RemoteSupportedGroups, Array( _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchX25519), "#" & TLS_GROUP_X25519, vbNullString), _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1), "#" & TLS_GROUP_SECP256R1, vbNullString), _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1), "#" & TLS_GROUP_SECP384R1, vbNullString)))
                                If .HelloRetryExchGroup = 0 Then
                                    GoTo HelloRetryFailed
                                End If
                            End If
                            If .CipherSuite <> 0 Then
                                .HelloRetryCipherSuite = .CipherSuite
                            Else
                                Select Case True
                                Case pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128)
                                    .HelloRetryCipherSuite = TLS_CS_AES_128_GCM_SHA256
                                Case pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256)
                                    .HelloRetryCipherSuite = TLS_CS_AES_256_GCM_SHA384
                                Case pvCryptoIsSupported(ucsTlsAlgoBulkChacha20Poly1305)
                                    .HelloRetryCipherSuite = TLS_CS_CHACHA20_POLY1305_SHA256
                                Case Else
                                    GoTo HelloRetryFailed
                                End Select
                            End If
                        Else
                            .HelloRetryRequest = False
                            .State = ucsTlsStateExpectClientFinished
                        End If
                    ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If .ExchGroup = 0 Then
                            If pvCollectionCount(.RemoteSupportedGroups) = 0 Then
                                If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                    lExchGroup = TLS_GROUP_SECP256R1
                                End If
                            Else
                                lExchGroup = pvCollectionFirst(.RemoteSupportedGroups, Array( _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchX25519), "#" & TLS_GROUP_X25519, vbNullString), _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1), "#" & TLS_GROUP_SECP256R1, vbNullString), _
                                        IIf(pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1), "#" & TLS_GROUP_SECP384R1, vbNullString)))
                            End If
                            If lExchGroup = 0 Then
                                GoTo NoSupportedGroups
                            End If
                            pvTlsSetupExchGroup uCtx, lExchGroup
                        End If
                        .State = ucsTlsStateExpectClientKeyExchange
                    End If
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                '--- post-process ucsTlsStateExpectClientHello
                If .HelloRetryRequest Then
                    '--- on HelloRetryRequest replace HandshakeMessages w/ 'synthetic handshake message'
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    .HandshakeMessages.Size = 0
                    pvBufferWriteLong .HandshakeMessages, TLS_HANDSHAKE_MESSAGE_HASH
                    pvBufferWriteLong .HandshakeMessages, .DigestSize, Size:=3
                    pvBufferWriteArray .HandshakeMessages, baHandshakeHash
                End If
                pvTlsBuildServerHello uCtx, .SendBuffer
                If .State = ucsTlsStateExpectClientFinished Then
                    pvTlsDeriveHandshakeSecrets uCtx
                    pvTlsBuildServerHandshakeFinished uCtx, .SendBuffer
                End If
                If .State = ucsTlsStateExpectClientKeyExchange Then
                    pvTlsBuildServerLegacyKeyExchange uCtx, .SendBuffer
                End If
            Case ucsTlsStateExpectClientKeyExchange
                Select Case lMessageType
                Case TLS_HANDSHAKE_CLIENT_KEY_EXCHANGE
                    If .UseRsaKeyTransport Then
                        If uInput.Pos + 2 > lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadBlockStart uInput, BlockSize:=lBlockSize, Size:=2
                            If uInput.Pos + lBlockSize <> lMessageEnd Or lBlockSize = 0 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadArray uInput, baTemp, lBlockSize
                        pvBufferReadBlockEnd uInput
                        pvTlsSetupExchRsaPreMasterSecret uCtx, baTemp
                    Else
                        If uInput.Pos + 1 > lMessageEnd Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadBlockStart uInput, BlockSize:=lBlockSize
                            If uInput.Pos + lBlockSize <> lMessageEnd Or lBlockSize = 0 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadArray uInput, .RemoteExchPublic, lBlockSize
                            If Not pvTlsCheckRemoteKey(.ExchGroup, .RemoteExchPublic) Then
                                GoTo InvalidRemoteKey
                            End If
                        pvBufferReadBlockEnd uInput
                    End If
                    .State = ucsTlsStateExpectClientFinished
                Case TLS_HANDSHAKE_CERTIFICATE
                    '--- ToDo: impl
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                If .State = ucsTlsStateExpectClientFinished Then
                    pvTlsDeriveLegacySecrets uCtx
                End If
            Case ucsTlsStateExpectClientFinished
                Select Case lMessageType
                Case TLS_HANDSHAKE_FINISHED
                    If pvArraySize(.RemoteTrafficKey) = 0 Then
                        GoTo MissingTrafficKeys
                    End If
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .RemoteTrafficSecret, "finished", baEmpty, .DigestSize
                        pvTlsHkdfExtract uVerify.Data, .DigestAlgo, baTemp, baHandshakeHash
                    Else
                        Debug.Assert .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12
                        pvTlsKdfLegacyPrf uVerify.Data, .DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_RENEGOTIATION_INFO) Then
                            .RemoteLegacyVerifyData = uVerify.Data
                        End If
                    End If
                    If Not pvArrayEqual(uVerify.Data, baMessage) Then
                        GoTo ServerHandshakeFailed
                    End If
                    .State = ucsTlsStatePostHandshake
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12, TLS_HANDSHAKE_CERTIFICATE_VERIFY, -1)
                    '--- ToDo: impl
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    pvTlsAppendHandshakeHash uCtx, uInput.Data, lMessagePos, lMessageSize + 4
                End If
                '--- post-process ucsTlsStateExpectClientFinished
                If .State = ucsTlsStatePostHandshake Then
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        pvTlsGetHandshakeHash uCtx, baHandshakeHash
                        pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
                        pvTlsResetHandshakeHash uCtx
                        Set .RemoteTickets = New Collection
                    Else
                        Debug.Assert .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12
                        pvTlsBuildServerLegacyFinished uCtx, .SendBuffer
                        pvTlsResetHandshakeHash uCtx
                    End If
                End If
#End If
            Case ucsTlsStatePostHandshake
                Select Case lMessageType
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 And Not .IsServer, TLS_HANDSHAKE_HELLO_REQUEST, -1)
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Received Hello Request. Will renegotiate"
                    #End If
                    If lMessageSize <> 0 Then
                        GoTo InvalidSize
                    End If
                    .State = ucsTlsStateExpectServerHello
                    .AlpnNegotiated = vbNullString
                    .SniRequested = vbNullString
                    '--- renegotiate ephemeral keys too
                    .ExchGroup = 0
                    .CipherSuite = 0
                    .SignatureScheme = 0
                    .PrevRecordType = 0
                    pvTlsBuildClientHello uCtx, .SendBuffer
#If ImplTlsServer Then
                Case IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 And .IsServer, TLS_HANDSHAKE_CLIENT_HELLO, -1)
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Received Client Hello. Will renegotiate"
                    #End If
                    #If Not ImplTlsServerAllowInsecureRenegotiation Then
                        If pvArraySize(.RemoteLegacyVerifyData) = 0 Then
                            GoTo SecureRenegotiationFailed
                        End If
                    #End If
                    .State = ucsTlsStateExpectClientHello
                    .AlpnNegotiated = vbNullString
                    .SniRequested = vbNullString
                    '--- renegotiate ephemeral keys too
                    .ExchGroup = 0
                    .CipherSuite = 0
                    .SignatureScheme = 0
                    .PrevRecordType = 0
                    GoTo RenegotiateClientHello
#End If
                Case TLS_HANDSHAKE_NEW_SESSION_TICKET
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    If Not .RemoteTickets Is Nothing Then
                        .RemoteTickets.Add baMessage
                    End If
                Case TLS_HANDSHAKE_KEY_UPDATE
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Received TLS_HANDSHAKE_KEY_UPDATE"
                    #End If
                    If lMessageSize <> 1 Then
                        GoTo InvalidSize
                    End If
                    pvBufferReadLong uInput, lRequestUpdate
                    If lRequestUpdate <> 0 And lRequestUpdate <> 1 Then
                        GoTo UnsupportedRequestUpdate
                    End If
                    If lRequestUpdate <> 0 Then
                        '--- ack by TLS_HANDSHAKE_KEY_UPDATE w/ update_not_requested(0)
                        pvArrayByte baTemp, TLS_HANDSHAKE_KEY_UPDATE, 0, 0, 1, 0
                        pvTlsBuildApplicationData uCtx, .SendBuffer, baTemp, 0, UBound(baTemp) + 1, TLS_CONTENT_TYPE_HANDSHAKE
                    End If
                    pvTlsDeriveKeyUpdate uCtx, lRequestUpdate <> 0
                Case TLS_HANDSHAKE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, uInput, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    pvTlsBuildClientHandshakeFinished uCtx, .SendBuffer
                    pvTlsResetHandshakeHash uCtx
                Case Else
                    GoTo UnexpectedMessageType
                End Select
            Case Else
                GoTo InvalidStateHandshake
            End Select
        pvBufferReadBlockEnd uInput
    Loop
    End With
    '--- success
    pvTlsParseHandshake = True
QH:
    Exit Function
UnexpectedMessageType:
    sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", pvTlsGetStateAsText(uCtx.State)), "%2", pvTlsGetMessageName(lMessageType))
    eAlertCode = uscTlsAlertUnexpectedMessage
    GoTo QH
ServerHandshakeFailed:
    sError = ERR_SERVER_HANDSHAKE_FAILED
    eAlertCode = IIf(pvArraySize(uVerify.Data) <> pvArraySize(baMessage), uscTlsAlertDecodeError, uscTlsAlertDecryptError)
    GoTo QH
HelloRetryFailed:
    sError = ERR_HELLO_RETRY_FAILED
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
NoServerCertificate:
    sError = ERR_NO_SERVER_CERTIFICATE
    eAlertCode = uscTlsAlertCertificateUnknown
    GoTo QH
InvalidSize:
    sError = IIf(lExtType < 0, ERR_INVALID_SIZE, Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType)))
    eAlertCode = uscTlsAlertDecodeError
    GoTo QH
UnsupportedCurveType:
    sError = Replace(ERR_UNSUPPORTED_CURVE_TYPE, "%1", lCurveType)
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidStateHandshake:
    sError = Replace(ERR_INVALID_STATE_HANDSHAKE, "%1", pvTlsGetStateAsText(uCtx.State))
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
SecureRenegotiationFailed:
    sError = ERR_SECURE_RENEGOTIATION_FAILED
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidRemoteKey:
    sError = ERR_INVALID_REMOTE_KEY
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
NoSupportedGroups:
    sError = ERR_NO_SUPPORTED_GROUPS
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedRequestUpdate:
    sError = Replace(ERR_UNSUPPORTED_REQUEST_UPDATE, "%1", lRequestUpdate)
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
MissingTrafficKeys:
    sError = ERR_MISSING_TRAFFIC_KEYS
    eAlertCode = uscTlsAlertUnexpectedMessage
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsParseHandshakeServerHello(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshakeServerHello"
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lCipherSuite    As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim lExchGroup      As Long
    Dim lPublicSize     As Long
    Dim lNameSize       As Long
    Dim lCookieSize     As Long
    
    On Error GoTo EH
    lExtType = -1
    If pvArraySize(m_baHelloRetryRandom) = 0 Then
        pvTlsGetHelloRetryRandom m_baHelloRetryRandom
    End If
    With uCtx
        .ProtocolVersion = IIf(lRecordProtocol <= TLS_PROTOCOL_VERSION_TLS12, TLS_PROTOCOL_VERSION_TLS12, TLS_PROTOCOL_VERSION_TLS13)
        pvBufferReadLong uInput, .RemoteProtocolVersion, Size:=2
        pvBufferReadArray uInput, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE
        If .HelloRetryRequest Then
            '--- clear HelloRetryRequest
            .HelloRetryRequest = False
            .HelloRetryCipherSuite = 0
            .HelloRetryExchGroup = 0
            Erase .HelloRetryCookie
        Else
            .HelloRetryRequest = (StrConv(.RemoteExchRandom, vbUnicode) = StrConv(m_baHelloRetryRandom, vbUnicode))
        End If
        pvBufferReadBlockStart uInput, BlockSize:=lBlockSize
            pvBufferReadArray uInput, .RemoteSessionID, lBlockSize
        pvBufferReadBlockEnd uInput
        pvBufferReadLong uInput, lCipherSuite, Size:=2
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        #If ImplUseDebugLog Then
            DebugLog MODULE_NAME, FUNC_NAME, "Using " & pvTlsGetCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName
        #End If
        If .HelloRetryRequest Then
            .HelloRetryCipherSuite = lCipherSuite
        End If
        pvBufferReadLong uInput, lLegacyCompress
        Debug.Assert lLegacyCompress = 0
        Set .RemoteExtensions = New Collection
        If uInput.Pos < lInputEnd Then
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                lBlockEnd = uInput.Pos + lBlockSize
                Do While uInput.Pos + 3 < lBlockEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "ServerHello " & pvTlsGetExtensionName(lExtType)
                    #End If
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        lExtEnd = uInput.Pos + lExtSize
                        If lExtEnd > lBlockEnd Then
                            GoTo InvalidSize
                        End If
                        Select Case lExtType
                        Case IIf((.LocalFeatures And ucsTlsSupportTls13) <> 0, TLS_EXTENSION_KEY_SHARE, -1)
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadLong uInput, lExchGroup, Size:=2
                            pvTlsSetupExchGroup uCtx, lExchGroup
                            If .HelloRetryRequest Then
                                .HelloRetryExchGroup = lExchGroup
                            Else
                                #If ImplUseDebugLog Then
                                    DebugLog MODULE_NAME, FUNC_NAME, "With exchange group " & pvTlsGetExchGroupName(.ExchGroup)
                                #End If
                                If lExtSize < 4 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lPublicSize
                                    If uInput.Pos + lPublicSize <> lExtEnd Or lPublicSize = 0 Then
                                        GoTo InvalidSize
                                    End If
                                    pvBufferReadArray uInput, .RemoteExchPublic, lPublicSize
                                    If Not pvTlsCheckRemoteKey(.ExchGroup, .RemoteExchPublic) Then
                                        GoTo InvalidRemoteKey
                                    End If
                                pvBufferReadBlockEnd uInput
                            End If
                        Case IIf((.LocalFeatures And ucsTlsSupportTls13) <> 0, TLS_EXTENSION_SUPPORTED_VERSIONS, -1)
                            If lExtSize <> 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadLong uInput, .ProtocolVersion, Size:=2
                            If .ProtocolVersion <> TLS_PROTOCOL_VERSION_TLS12 And .ProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 Then
                                GoTo UnsupportedProtocol
                            End If
                        Case IIf((.LocalFeatures And ucsTlsSupportTls13) <> 0, TLS_EXTENSION_COOKIE, -1)
                            If Not .HelloRetryRequest Then
                                GoTo UnexpectedExtension
                            End If
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lCookieSize
                                If uInput.Pos + lCookieSize <> lExtEnd Or lCookieSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadArray uInput, .HelloRetryCookie, lCookieSize
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_ALPN
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lNameSize
                                If uInput.Pos + lNameSize <> lExtEnd Or lNameSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadBlockStart uInput, BlockSize:=lNameSize
                                    If uInput.Pos + lNameSize <> lExtEnd Or lNameSize = 0 Then
                                        GoTo InvalidSize
                                    End If
                                    pvBufferReadString uInput, .AlpnNegotiated, lNameSize
                                pvBufferReadBlockEnd uInput
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_RENEGOTIATION_INFO
                            If lExtSize < 1 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=1, BlockSize:=lNameSize
                                If uInput.Pos + lNameSize <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadArray uInput, .RemoteLegacyRenegInfo, lNameSize
                            pvBufferReadBlockEnd uInput
                        Case Else
                            uInput.Pos = uInput.Pos + lExtSize
                        End Select
                        If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                            .RemoteExtensions.Add lExtType, "#" & lExtType
                        End If
                    pvBufferReadBlockEnd uInput
                Loop
            pvBufferReadBlockEnd uInput
        End If
    End With
    '--- success
    pvTlsParseHandshakeServerHello = True
QH:
    Exit Function
InvalidSize:
    sError = IIf(lExtType < 0, ERR_INVALID_SIZE, Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType)))
    eAlertCode = uscTlsAlertDecodeError
    GoTo QH
UnexpectedExtension:
    sError = Replace(ERR_UNEXPECTED_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
InvalidRemoteKey:
    sError = ERR_INVALID_REMOTE_KEY
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
UnsupportedProtocol:
    sError = Replace(ERR_UNSUPPORTED_PROTOCOL, "%1", "&H" & Hex$(uCtx.ProtocolVersion))
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsParseHandshakeClientHello(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshakeClientHello"
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lCipherSuite    As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim lExchGroup      As Long
    Dim eExchAlgo       As UcsTlsCryptoAlgorithmsEnum
    Dim lBlockSize      As Long
    Dim lProtocolVersion As Long
    Dim lSignatureScheme As Long
    Dim cCipherSuites   As Collection
    Dim lIdx            As Long
    Dim vItem           As Variant
    Dim vElem           As Variant
    Dim uKeyInfo        As UcsKeyInfo
    Dim lNameType       As Long
    Dim lNameSize       As Long
    Dim sName           As String
    Dim cAlpnPrefs      As Collection
    Dim lAlpnPref       As Long
    Dim lCompression    As Long
    Dim cPrevRemoteExt  As Collection
    
    On Error GoTo EH
    lExtType = -1
    With uCtx
        Set cPrevRemoteExt = .RemoteExtensions
        Set .RemoteExtensions = New Collection
        If Not pvAsn1DecodePrivateKey(.LocalCertificates, .LocalPrivateKey, uKeyInfo) Then
            GoTo UnsupportedCertificate
        End If
        .ProtocolVersion = IIf((.LocalFeatures And ucsTlsSupportTls12) <> 0 And Not .HelloRetryRequest, TLS_PROTOCOL_VERSION_TLS12, TLS_PROTOCOL_VERSION_TLS13)
        pvBufferReadLong uInput, .RemoteProtocolVersion, Size:=2
        If .RemoteProtocolVersion < TLS_PROTOCOL_VERSION_TLS12 Then
            GoTo UnsupportedProtocol
        End If
        '--- remote random
        If uInput.Pos + TLS_HELLO_RANDOM_SIZE > lInputEnd Then
            GoTo InvalidSize
        End If
        pvBufferReadArray uInput, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE
        '--- session id
        If uInput.Pos + 1 > lInputEnd Then
            GoTo InvalidSize
        End If
        pvBufferReadBlockStart uInput, BlockSize:=lSize
            If uInput.Pos + lSize > lInputEnd Then
                GoTo InvalidSize
            End If
            If lSize <> 0 And lSize <> TLS_LEGACY_SESSIONID_SIZE Then
                GoTo InvalidSize
            End If
            pvBufferReadArray uInput, .RemoteSessionID, lSize
        pvBufferReadBlockEnd uInput
        '--- ciphersuites
        If uInput.Pos + 2 > lInputEnd Then
            GoTo InvalidSize
        End If
        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
            lEnd = uInput.Pos + lSize
            If lEnd > lInputEnd Then
                GoTo InvalidSize
            End If
            Set cCipherSuites = New Collection
            Do While uInput.Pos + 1 < lEnd
                pvBufferReadLong uInput, lIdx, Size:=2
                If Not SearchCollection(cCipherSuites, "#" & lIdx) Then
                    cCipherSuites.Add lIdx, "#" & lIdx
                End If
                If lIdx = TLS_CS_EMPTY_RENEGOTIATION_INFO_SCSV And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    lIdx = TLS_EXTENSION_RENEGOTIATION_INFO
                    '--- 3.7, Server Behavior: Secure Renegotiation. When a ClientHello is received, the server MUST verify that it does not
                    '---   contain the TLS_EMPTY_RENEGOTIATION_INFO_SCSV SCSV. If the SCSV is present, the server MUST abort the handshake.
                    If SearchCollection(cPrevRemoteExt, "#" & lIdx) Then
                        GoTo SecureRenegotiationFailed
                    End If
                    If Not SearchCollection(.RemoteExtensions, "#" & lIdx) Then
                        .RemoteExtensions.Add lIdx, "#" & lIdx
                    End If
                End If
            Loop
            If uInput.Pos <> lEnd Then
                GoTo InvalidSize
            End If
        pvBufferReadBlockEnd uInput
        '--- compression
        If uInput.Pos + 1 > lInputEnd Then
            GoTo InvalidSize
        End If
        pvBufferReadBlockStart uInput, BlockSize:=lSize
            lEnd = uInput.Pos + lSize
            If lEnd > lInputEnd Then
                GoTo InvalidSize
            End If
            If uInput.Pos < lEnd Then
                lIdx = 1
                Do While uInput.Pos < lEnd
                    pvBufferReadLong uInput, lCompression
                    If lCompression <> 0 And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        GoTo InvalidCompression
                    End If
                    If lCompression = 0 Then
                        lIdx = 0
                    End If
                Loop
                If lIdx > 0 Then
                    GoTo InvalidCompression
                End If
            End If
        pvBufferReadBlockEnd uInput
        '--- extensions
        If uInput.Pos + 1 < lInputEnd Then
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                If lEnd <> lInputEnd Then
                    GoTo InvalidSize
                End If
                Do While uInput.Pos + 3 < lEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "ClientHello " & pvTlsGetExtensionName(lExtType)
                    #End If
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        lExtEnd = uInput.Pos + lExtSize
                        If lExtEnd > lEnd Then
                            GoTo InvalidSize
                        End If
                        Select Case lExtType
                        Case TLS_EXTENSION_SERVER_NAME
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos + 2 < lExtEnd
                                    pvBufferReadLong uInput, lNameType
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lNameSize
                                        If uInput.Pos + lNameSize > lExtEnd Or lNameSize = 0 Then
                                            GoTo InvalidSize
                                        End If
                                        If lNameType = TLS_SERVER_NAME_TYPE_HOSTNAME Then
                                            If LenB(.SniRequested) <> 0 Then
                                                GoTo InvalidServerName
                                            End If
                                            pvBufferReadString uInput, .SniRequested, lNameSize
                                            If Not pvIsValidServerName(.SniRequested) Then
                                                GoTo InvalidServerName
                                            End If
                                        Else
                                            uInput.Pos = uInput.Pos + lNameSize '--- skip
                                        End If
                                    pvBufferReadBlockEnd uInput
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_ALPN
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            Set cAlpnPrefs = New Collection
                            For Each vElem In Split(.AlpnProtocols, "|")
                                cAlpnPrefs.Add cAlpnPrefs.Count, "#" & vElem
                            Next
                            lAlpnPref = 1000
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos < lExtEnd
                                    pvBufferReadBlockStart uInput, BlockSize:=lNameSize
                                        If uInput.Pos + lNameSize > lExtEnd Or lNameSize = 0 Then
                                            GoTo InvalidSize
                                        End If
                                        pvBufferReadString uInput, sName, lNameSize
                                        If SearchCollection(cAlpnPrefs, "#" & sName, RetVal:=vElem) Then
                                            If vElem < lAlpnPref Then
                                                .AlpnNegotiated = sName
                                                lAlpnPref = vElem
                                            End If
                                        End If
                                    pvBufferReadBlockEnd uInput
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                            If LenB(.AlpnNegotiated) = 0 And cAlpnPrefs.Count > 0 Then
                                GoTo NoAlpnNegotiated
                            End If
                        Case IIf((.LocalFeatures And ucsTlsSupportTls13) <> 0, TLS_EXTENSION_KEY_SHARE, -1)
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos + 3 < lExtEnd
                                    pvBufferReadLong uInput, lExchGroup, Size:=2
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                        If lBlockSize = 0 Then
                                            GoTo InvalidRemoteKey
                                        End If
                                        If uInput.Pos + lBlockSize > lExtEnd Or lBlockSize = 0 Then
                                            GoTo InvalidSize
                                        End If
                                        If .HelloRetryRequest And lExchGroup <> .HelloRetryExchGroup Then
                                            lExchGroup = 0
                                        End If
                                        Select Case lExchGroup
                                        Case TLS_GROUP_X25519
                                            eExchAlgo = ucsTlsAlgoExchX25519
                                        Case TLS_GROUP_SECP256R1
                                            eExchAlgo = ucsTlsAlgoExchSecp256r1
                                        Case TLS_GROUP_SECP384R1
                                            eExchAlgo = ucsTlsAlgoExchSecp384r1
                                        Case TLS_GROUP_X448, TLS_GROUP_SECP521R1
                                            eExchAlgo = 0
                                        Case TLS_GROUP_FFDHE_FIRST To TLS_GROUP_FFDHE_LAST
                                            eExchAlgo = 0
                                        Case TLS_GROUP_FFDHE_PRIVATE_USE_FIRST To TLS_GROUP_FFDHE_PRIVATE_USE_LAST
                                            eExchAlgo = 0
                                        Case TLS_GROUP_ECDHE_PRIVATE_USE_FIRST To TLS_GROUP_ECDHE_PRIVATE_USE_LAST
                                            eExchAlgo = 0
                                        Case Else
                                            If (lExchGroup And &HFF) = lExchGroup \ &H100 And (lExchGroup And &HF) = &HA Then
                                                eExchAlgo = 0 '--- grease from RFC8701
                                            Else
                                                GoTo UnsupportedExchGroup
                                            End If
                                        End Select
                                        Select Case True
                                        Case eExchAlgo = 0, Not pvCryptoIsSupported(eExchAlgo)
                                            lExchGroup = 0
                                            uInput.Pos = uInput.Pos + lBlockSize
                                        End Select
                                        If lExchGroup <> 0 Then
                                            pvTlsSetupExchGroup uCtx, lExchGroup
                                            pvBufferReadArray uInput, .RemoteExchPublic, lBlockSize
                                            If Not pvTlsCheckRemoteKey(.ExchGroup, .RemoteExchPublic) Then
                                                GoTo InvalidRemoteKey
                                            End If
                                            #If ImplUseDebugLog Then
                                                DebugLog MODULE_NAME, FUNC_NAME, "With exchange group " & pvTlsGetExchGroupName(.ExchGroup)
                                            #End If
                                        End If
                                    pvBufferReadBlockEnd uInput
                                    If lExchGroup <> 0 Then
                                        uInput.Pos = lExtEnd
                                    End If
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_SIGNATURE_ALGORITHMS
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Or lBlockSize Mod 2 <> 0 Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos + 1 < lExtEnd
                                    pvBufferReadLong uInput, lSignatureScheme, Size:=2
                                    If pvTlsMatchSignatureScheme(uCtx, lSignatureScheme, uKeyInfo) Then
                                        .SignatureScheme = lSignatureScheme
                                        uInput.Pos = lExtEnd
                                    End If
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_SUPPORTED_GROUPS
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Or lBlockSize Mod 2 <> 0 Then
                                    GoTo InvalidSize
                                End If
                                Set .RemoteSupportedGroups = New Collection
                                Do While uInput.Pos + 1 < lExtEnd
                                    pvBufferReadLong uInput, lExchGroup, Size:=2
                                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                        Select Case lExchGroup
                                        Case TLS_GROUP_SECP256R1 To TLS_GROUP_SECP521R1, TLS_GROUP_X25519 To TLS_GROUP_X448
                                            '--- ecc curves
                                        Case TLS_GROUP_FFDHE_FIRST To TLS_GROUP_FFDHE_LAST
                                            '--- ffdhe
                                        Case TLS_GROUP_FFDHE_PRIVATE_USE_FIRST To TLS_GROUP_FFDHE_PRIVATE_USE_LAST, TLS_GROUP_ECDHE_PRIVATE_USE_FIRST To TLS_GROUP_ECDHE_PRIVATE_USE_LAST
                                            '--- private use
                                        Case Else
                                            If (lExchGroup And &HFF) = lExchGroup \ &H100 And (lExchGroup And &HF) = &HA Then
                                                '--- grease from RFC8701
                                            Else
                                                GoTo UnsupportedExchGroup
                                            End If
                                        End Select
                                    End If
                                    .RemoteSupportedGroups.Add lExchGroup, "#" & lExchGroup
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                        Case IIf((.LocalFeatures And ucsTlsSupportTls13) <> 0, TLS_EXTENSION_SUPPORTED_VERSIONS, -1)
                            If lExtSize < 1 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Or lBlockSize Mod 2 <> 0 Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos + 1 < lExtEnd
                                    pvBufferReadLong uInput, lProtocolVersion, Size:=2
                                    If lProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                        uInput.Pos = lExtEnd
                                    ElseIf lProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 And (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                        uInput.Pos = lExtEnd
                                    End If
                                Loop
                                If uInput.Pos <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                            pvBufferReadBlockEnd uInput
                            If lProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 And lProtocolVersion <> TLS_PROTOCOL_VERSION_TLS12 Then
                                GoTo UnsupportedProtocol
                            End If
                            .ProtocolVersion = lProtocolVersion
                        Case IIf((.LocalFeatures And ucsTlsSupportTls12) <> 0, TLS_EXTENSION_RENEGOTIATION_INFO, -1)
                            If lExtSize < 1 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadArray uInput, .RemoteLegacyRenegInfo, lBlockSize
                            pvBufferReadBlockEnd uInput
                            If lBlockSize > 0 Then
                                If Not pvArrayEqual(.RemoteLegacyRenegInfo, .RemoteLegacyVerifyData) Then
                                    GoTo SecureRenegotiationFailed
                                End If
                            End If
                        Case IIf((.LocalFeatures And ucsTlsSupportTls12) <> 0, TLS_EXTENSION_EXTENDED_MASTER_SECRET, -1)
                            If lExtSize <> 0 Then
                                GoTo InvalidSize
                            End If
                        Case Else
                            If .HelloRetryRequest Then
                                If Not SearchCollection(cPrevRemoteExt, "#" & lExtType) Then
                                    GoTo UnexpectedExtension
                                End If
                            End If
                            uInput.Pos = uInput.Pos + lExtSize
                        End Select
                        If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                            .RemoteExtensions.Add lExtType, "#" & lExtType
                        End If
                    pvBufferReadBlockEnd uInput
                Loop
                If uInput.Pos <> lEnd Then
                    lExtType = -1
                    GoTo InvalidSize
                End If
            pvBufferReadBlockEnd uInput
        End If
        '--- match preferred ciphersuites
        For Each vElem In pvTlsGetSortedCipherSuites(.LocalFeatures And IIf(.ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13, ucsTlsSupportTls13, ucsTlsSupportTls12), uKeyInfo.AlgoObjId)
            If SearchCollection(cCipherSuites, "#" & vElem, RetVal:=vItem) Then
                If Not .HelloRetryRequest Or .HelloRetryCipherSuite = vItem Then
                    lCipherSuite = vItem
                    Exit For
                End If
            End If
        Next
        If lCipherSuite = 0 Then
            GoTo NoCipherSuite
        End If
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        If .SignatureScheme = 0 Then
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                GoTo NegotiateSignatureFailed
            ElseIf SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_SIGNATURE_ALGORITHMS_CERT) Then
                GoTo NegotiateSignatureFailed
            End If
            If pvTlsMatchSignatureScheme(uCtx, TLS_SIGNATURE_RSA_PKCS1_SHA1, uKeyInfo) Then
                .SignatureScheme = TLS_SIGNATURE_RSA_PKCS1_SHA1
            ElseIf pvTlsMatchSignatureScheme(uCtx, TLS_SIGNATURE_ECDSA_SHA1, uKeyInfo) Then
                .SignatureScheme = TLS_SIGNATURE_ECDSA_SHA1
            Else
                GoTo NegotiateSignatureFailed
            End If
        End If
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            lExtType = TLS_EXTENSION_KEY_SHARE
            If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                GoTo NoExtension
            End If
        Else
            For Each vElem In Array(TLS_EXTENSION_EXTENDED_MASTER_SECRET, TLS_EXTENSION_ENCRYPT_THEN_MAC, TLS_EXTENSION_RENEGOTIATION_INFO)
                If SearchCollection(cPrevRemoteExt, "#" & vElem) Then
                    If Not SearchCollection(.RemoteExtensions, "#" & vElem) Then
                        GoTo NoExtension
                    End If
                End If
            Next
        End If
        #If ImplUseDebugLog Then
            DebugLog MODULE_NAME, FUNC_NAME, "Using " & pvTlsGetCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName
        #End If
    End With
    '--- success
    pvTlsParseHandshakeClientHello = True
QH:
    If uKeyInfo.hKey <> 0 Then
        Call CryptDestroyKey(uKeyInfo.hKey)
    End If
    If uKeyInfo.hProv <> 0 Then
        Call CryptReleaseContext(uKeyInfo.hProv, 0)
    End If
    If uKeyInfo.hNKey <> 0 Then
        Call NCryptFreeObject(uKeyInfo.hNKey)
    End If
    Exit Function
UnsupportedCertificate:
    sError = ERR_UNSUPPORTED_CERTIFICATE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedProtocol:
    sError = Replace(ERR_UNSUPPORTED_PROTOCOL, "%1", "&H" & Hex$(uCtx.RemoteProtocolVersion))
    eAlertCode = uscTlsAlertProtocolVersion
    GoTo QH
NoCipherSuite:
    For Each vElem In cCipherSuites
        sError = sError & ", &H" & Hex$(vElem)
    Next
    sError = Replace(ERR_NO_SUPPORTED_CIPHER_SUITE, "%1", Mid$(sError, 3))
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidSize:
    sError = IIf(lExtType < 0, ERR_INVALID_SIZE, Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType)))
    eAlertCode = uscTlsAlertDecodeError
    GoTo QH
InvalidRemoteKey:
    sError = ERR_INVALID_REMOTE_KEY
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
NegotiateSignatureFailed:
    sError = ERR_NEGOTIATE_SIGNATURE_FAILED
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnexpectedExtension:
    sError = Replace(ERR_UNEXPECTED_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
NoAlpnNegotiated:
    sError = ERR_NO_ALPN_NEGOTIATED
    eAlertCode = uscTlsAlertNoApplicationProtocol
    GoTo QH
UnsupportedExchGroup:
    sError = Replace(ERR_UNSUPPORTED_EXCH_GROUP, "%1", pvTlsGetExchGroupName(lExchGroup))
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
NoExtension:
    sError = Replace(ERR_NO_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
    eAlertCode = uscTlsAlertMissingExtension
    GoTo QH
InvalidCompression:
    sError = Replace(ERR_INVALID_COMPRESSION, "%1", lCompression)
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
InvalidServerName:
    sError = Replace(ERR_INVALID_SERVER_NAME, "%1", uCtx.SniRequested)
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
SecureRenegotiationFailed:
    sError = ERR_SECURE_RENEGOTIATION_FAILED
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsParseHandshakeCertificateRequest(uCtx As UcsTlsContext, uInput As UcsBuffer, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lSignatureScheme As Long
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim baDName()       As Byte
    Dim lDnSize         As Long
    Dim uKeyInfo        As UcsKeyInfo
    Dim baSignatureSchemes() As Byte
    Dim lSigPos         As Long
    Dim oCallback       As Object
    Dim bConfirmed      As Boolean
    
    On Error GoTo EH
    lExtType = -1
    With uCtx
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            If uInput.Pos + 1 > uInput.Size Then
                GoTo InvalidSize
            End If
            pvBufferReadBlockStart uInput, Size:=1, BlockSize:=lSize
                If uInput.Pos + lSize > uInput.Size Then
                    GoTo InvalidSize
                End If
                pvBufferReadArray uInput, .CertRequestContext, lSize
            pvBufferReadBlockEnd uInput
            If uInput.Pos + 2 > uInput.Size Then
                GoTo InvalidSize
            End If
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                If lEnd > uInput.Size Then
                    GoTo InvalidSize
                End If
                Do While uInput.Pos + 3 < lEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        lExtEnd = uInput.Pos + lExtSize
                        If lExtEnd > lEnd Then
                            GoTo InvalidSize
                        End If
                        Select Case lExtType
                        Case TLS_EXTENSION_SIGNATURE_ALGORITHMS
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadArray uInput, baSignatureSchemes, lBlockSize
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_CERTIFICATE_AUTHORITIES
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                If lBlockEnd <> lExtEnd Or lBlockSize = 0 Or lBlockSize Mod 2 <> 0 Then
                                    GoTo InvalidSize
                                End If
                                Set .CertRequestCaDn = New Collection
                                Do While uInput.Pos + 1 < lBlockEnd
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lDnSize
                                        If uInput.Pos + lDnSize > lBlockEnd Or lDnSize = 0 Then
                                            GoTo InvalidSize
                                        End If
                                        pvBufferReadArray uInput, baDName, lDnSize
                                        .CertRequestCaDn.Add baDName
                                    pvBufferReadBlockEnd uInput
                                Loop
                            pvBufferReadBlockEnd uInput
                        Case Else
                            uInput.Pos = uInput.Pos + lExtSize
                        End Select
                    pvBufferReadBlockEnd uInput
                Loop
            pvBufferReadBlockEnd uInput
        End If
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            If uInput.Pos + 1 > uInput.Size Then
                GoTo InvalidSize
            End If
            pvBufferReadBlockStart uInput, Size:=1, BlockSize:=lSize
                If uInput.Pos + lSize > uInput.Size Then
                    GoTo InvalidSize
                End If
                uInput.Pos = uInput.Pos + lSize '--- skip certificate_types
            pvBufferReadBlockEnd uInput
            If uInput.Pos + 2 > uInput.Size Then
                GoTo InvalidSize
            End If
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                If uInput.Pos + lSize > uInput.Size Or lSize = 0 Then
                    GoTo InvalidSize
                End If
                pvBufferReadArray uInput, baSignatureSchemes, lSize
            pvBufferReadBlockEnd uInput
            If uInput.Pos + 2 > uInput.Size Then
                GoTo InvalidSize
            End If
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                If lEnd > uInput.Size Then
                    GoTo InvalidSize
                End If
                Set .CertRequestCaDn = New Collection
                Do While uInput.Pos + 1 < lEnd
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lDnSize
                        If uInput.Pos + lDnSize > lEnd Or lDnSize = 0 Then
                            GoTo InvalidSize
                        End If
                        pvBufferReadArray uInput, baDName, lDnSize
                        .CertRequestCaDn.Add baDName
                    pvBufferReadBlockEnd uInput
                Loop
            pvBufferReadBlockEnd uInput
        End If
        Do
            If Not pvAsn1DecodePrivateKey(.LocalCertificates, .LocalPrivateKey, uKeyInfo) Then
                GoTo UnsupportedPrivateKey
            End If
            .CertRequestSignatureScheme = -1
            lSigPos = 0
            Do While lSigPos < pvArraySize(baSignatureSchemes)
                lSignatureScheme = baSignatureSchemes(lSigPos) * &H100& + baSignatureSchemes(lSigPos + 1)
                lSigPos = lSigPos + 2
                If pvTlsMatchSignatureScheme(uCtx, lSignatureScheme, uKeyInfo) Then
                    .CertRequestSignatureScheme = lSignatureScheme
                    Exit Do
                End If
            Loop
            bConfirmed = False
            If .CertRequestSignatureScheme = -1 And .ClientCertCallback <> 0 Then
                Call vbaObjSetAddref(oCallback, .ClientCertCallback)
                bConfirmed = oCallback.FireOnCertificate(.CertRequestCaDn)
            End If
        Loop While bConfirmed
    End With
    '--- success
    pvTlsParseHandshakeCertificateRequest = True
QH:
    If uKeyInfo.hKey <> 0 Then
        Call CryptDestroyKey(uKeyInfo.hKey)
    End If
    If uKeyInfo.hProv <> 0 Then
        Call CryptReleaseContext(uKeyInfo.hProv, 0)
    End If
    If uKeyInfo.hNKey <> 0 Then
        Call NCryptFreeObject(uKeyInfo.hNKey)
    End If
    Exit Function
UnsupportedPrivateKey:
    sError = ERR_UNSUPPORTED_PRIVATE_KEY
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidSize:
    sError = IIf(lExtType < 0, ERR_INVALID_SIZE, Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType)))
    eAlertCode = uscTlsAlertDecodeError
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsMatchSignatureScheme(uCtx As UcsTlsContext, ByVal lSignatureScheme As Long, uKeyInfo As UcsKeyInfo) As Boolean
    Dim bHasEnoughBits  As Boolean
     
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        '--- PSS w/ SHA512 fails on short key lengths (min PSS size is 2 + lHashSize + lSaltSize where lSaltSize=lHashSize)
        bHasEnoughBits = (uKeyInfo.BitLen + 7) \ 8 > 2 + 2 * pvTlsSignatureHashSize(lSignatureScheme)
    End Select
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        If (uCtx.LocalFeatures And ucsTlsSupportTls12) <> 0 And uCtx.ProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 Then
            If uKeyInfo.AlgoObjId = szOID_RSA_RSA Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPkcs) And pvCryptoIsSupported(ucsTlsAlgoDigestSha1)
            End If
        End If
    Case TLS_SIGNATURE_RSA_PKCS1_SHA224, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If (uCtx.LocalFeatures And ucsTlsSupportTls12) <> 0 And uCtx.ProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 And (uKeyInfo.hProv = 0 Or OsVersion >= ucsOsvXp) Then
            If uKeyInfo.AlgoObjId = szOID_RSA_RSA Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPkcs)
            End If
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        If bHasEnoughBits And uKeyInfo.hProv = 0 Then
            If uKeyInfo.AlgoObjId = szOID_RSA_RSA Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPss)
            End If
        End If
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If bHasEnoughBits And uKeyInfo.hProv = 0 Then
            If uKeyInfo.AlgoObjId = szOID_RSA_SSA_PSS Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPss)
            End If
        End If
    Case TLS_SIGNATURE_ECDSA_SHA1
        If (uCtx.LocalFeatures And ucsTlsSupportTls12) <> 0 And uCtx.ProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 And uKeyInfo.hProv = 0 Then
            If uKeyInfo.AlgoObjId = szOID_ECC_PUBLIC_KEY Then
                pvTlsMatchSignatureScheme = True
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1)
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1)
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P521 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1)
            End If
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        If uKeyInfo.hProv = 0 Then
            If uKeyInfo.AlgoObjId = szOID_ECC_PUBLIC_KEY Then
                pvTlsMatchSignatureScheme = True
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP256R1_SHA256 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1)
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP384R1_SHA384 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1)
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P521 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP521R1_SHA512 Then
                pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1)
            End If
        End If
    End Select
End Function

Private Function pvTlsCheckRemoteKey(ByVal lExchGroup As Long, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvTlsCheckRemoteKey"
    Dim baCompr()       As Byte
    Dim baUncompr()     As Byte
    
    Select Case lExchGroup
    Case TLS_GROUP_X25519
        If pvArraySize(baPublic) <> LNG_X25519_KEYSZ Then
            GoTo QH
        End If
        '--- check empty key share
        If pvArrayAccumulateOr(baPublic) = 0 Then
            GoTo QH
        End If
        '--- check key share of "1"
        If pvArrayAccumulateOr(baPublic, Pos:=1) = 0 And baPublic(0) = 1 Then
            GoTo QH
        End If
    Case TLS_GROUP_SECP256R1
        If pvArraySize(baPublic) <> 2 * LNG_SECP256R1_KEYSZ + 1 Then
            GoTo QH
        End If
        If baPublic(0) <> 4 Then
            GoTo QH
        End If
        pvArrayAllocate baCompr, 1 + LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baCompr"
        baCompr(0) = 2 '--- compressed positive
        Call CopyMemory(baCompr(1), baPublic(1), LNG_SECP256R1_KEYSZ)
        If Not pvCryptoEcdhSecp256r1UncompressKey(baUncompr, baCompr) Then
            GoTo QH
        End If
        If Not pvArrayEqual(baUncompr, baPublic) Then
            baCompr(0) = 3 '--- compressed negative
            If Not pvCryptoEcdhSecp256r1UncompressKey(baUncompr, baCompr) Then
                GoTo QH
            End If
            If Not pvArrayEqual(baUncompr, baPublic) Then
                GoTo QH
            End If
        End If
    Case TLS_GROUP_SECP384R1
        If pvArraySize(baPublic) <> 2 * LNG_SECP384R1_KEYSZ + 1 Then
            GoTo QH
        End If
        If baPublic(0) <> 4 Then
            GoTo QH
        End If
        pvArrayAllocate baCompr, 1 + LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baCompr"
        baCompr(0) = 2 '--- compressed positive
        Call CopyMemory(baCompr(1), baPublic(1), LNG_SECP384R1_KEYSZ)
        If Not pvCryptoEcdhSecp384r1UncompressKey(baUncompr, baCompr) Then
            GoTo QH
        End If
        If Not pvArrayEqual(baUncompr, baPublic) Then
            baCompr(0) = 3 '--- compressed negative
            If Not pvCryptoEcdhSecp384r1UncompressKey(baUncompr, baCompr) Then
                GoTo QH
            End If
            If Not pvArrayEqual(baUncompr, baPublic) Then
                GoTo QH
            End If
        End If
    End Select
    '--- success
    pvTlsCheckRemoteKey = True
QH:
End Function

Private Sub pvTlsSetupExchGroup(uCtx As UcsTlsContext, ByVal lExchGroup As Long)
    Const FUNC_NAME     As String = "pvTlsSetupExchGroup"
    
    With uCtx
        If .ExchGroup <> lExchGroup Then
            .ExchGroup = lExchGroup
            Select Case lExchGroup
            Case TLS_GROUP_X25519
                .ExchAlgo = ucsTlsAlgoExchX25519
                If Not pvCryptoEcdhCurve25519MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "Curve25519")
                End If
            Case TLS_GROUP_SECP256R1
                .ExchAlgo = ucsTlsAlgoExchSecp256r1
                If Not pvCryptoEcdhSecp256r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp256r1")
                End If
            Case TLS_GROUP_SECP384R1
                .ExchAlgo = ucsTlsAlgoExchSecp384r1
                If Not pvCryptoEcdhSecp384r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp384r1")
                End If
            Case Else
                ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_EXCH_GROUP, "%1", pvTlsGetExchGroupName(.ExchGroup))
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
        pvTlsGetRandom .LocalExchPrivate, TLS_LEGACY_SECRET_SIZE
        Call CopyMemory(.LocalExchPrivate(0), TLS_LOCAL_LEGACY_VERSION, 2)
        If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_UNSUPPORTED_CERTIFICATE
        End If
        If Not pvCryptoEmePkcs1Encode(baEnc, .LocalExchPrivate, uCertInfo.BitLen) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEmePkcs1Encode")
        End If
        If Not pvCryptoRsaModExp(baEnc, uCertInfo.PubExp, uCertInfo.Modulus, .LocalExchRsaEncrPriv) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoRsaModExp")
        End If
    End With
End Sub

Private Sub pvTlsSetupExchRsaPreMasterSecret(uCtx As UcsTlsContext, baEnc() As Byte)
    Const FUNC_NAME     As String = "pvTlsSetupExchRsaPreMasterSecret"
    Dim uKeyInfo        As UcsKeyInfo
    Dim baDec()         As Byte
    Dim lSize           As Long
    Dim lVersion        As Long
    Dim hResult         As Long
    Dim sErrDesc        As String
    
    With uCtx
        .ExchAlgo = ucsTlsAlgoExchCertificate
        If Not pvAsn1DecodePrivateKey(.LocalCertificates, .LocalPrivateKey, uKeyInfo) Then
            sErrDesc = ERR_UNSUPPORTED_PRIVATE_KEY
            GoTo QH
        End If
        If uKeyInfo.hNKey <> 0 Then
            hResult = NCryptDecrypt(uKeyInfo.hNKey, baEnc(0), UBound(baEnc) + 1, 0, ByVal 0, 0, lSize, NCRYPT_PAD_PKCS1_FLAG)
            If hResult < 0 Or lSize = 0 Then
                GoTo UseRandom
            End If
            pvArrayAllocate baDec, lSize, FUNC_NAME & ".baDec"
            hResult = NCryptDecrypt(uKeyInfo.hNKey, baEnc(0), UBound(baEnc) + 1, 0, baDec(0), lSize, lSize, NCRYPT_PAD_PKCS1_FLAG)
            If hResult < 0 Then
                GoTo UseRandom
            End If
            pvArrayReallocate baDec, lSize, FUNC_NAME & ".baDec"
            .LocalExchPrivate = baDec
        ElseIf uKeyInfo.hProv <> 0 And uKeyInfo.hKey <> 0 Then
            baDec = baEnc
            pvArrayReverse baDec
            lSize = UBound(baDec) + 1
            If CryptDecrypt(uKeyInfo.hKey, 0, 1, 0, baDec(0), lSize) = 0 Then
                GoTo UseRandom
            End If
            pvArrayReallocate baDec, lSize, FUNC_NAME & ".baDec"
            .LocalExchPrivate = baDec
        Else
            If Not pvCryptoRsaCrtModExp(baEnc, uKeyInfo.PrivExp, uKeyInfo.Modulus, uKeyInfo.Prime1, uKeyInfo.Prime2, uKeyInfo.Coefficient, baDec) Then
                GoTo UseRandom
            End If
            If Not pvCryptoEmePkcs1Decode(.LocalExchPrivate, baDec) Then
                GoTo UseRandom
            End If
        End If
        If pvArraySize(.LocalExchPrivate) <> TLS_LEGACY_SECRET_SIZE Then
            GoTo UseRandom
        End If
        If pvArraySize(.LocalExchPrivate) >= 2 Then
            Call CopyMemory(lVersion, .LocalExchPrivate(0), 2)
        End If
        If lVersion <> .RemoteProtocolVersion Then
UseRandom:
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, "Will use random LocalExchPrivate"
            #End If
            pvTlsGetRandom .LocalExchPrivate, TLS_LEGACY_SECRET_SIZE
        End If
    End With
QH:
    If uKeyInfo.hKey <> 0 Then
        Call CryptDestroyKey(uKeyInfo.hKey)
    End If
    If uKeyInfo.hProv <> 0 Then
        Call CryptReleaseContext(uKeyInfo.hProv, 0)
    End If
    If uKeyInfo.hNKey <> 0 Then
        Call NCryptFreeObject(uKeyInfo.hNKey)
    End If
    If LenB(sErrDesc) <> 0 Then
        ErrRaise vbObjectError, FUNC_NAME, sErrDesc
    End If
End Sub

Private Sub pvTlsSetupCipherSuite(uCtx As UcsTlsContext, ByVal lCipherSuite As Long)
    Const FUNC_NAME     As String = "pvTlsSetupCipherSuite"
    
    With uCtx
        If .CipherSuite <> lCipherSuite Then
            .CipherSuite = lCipherSuite
            .BulkAlgo = 0
            .MacSize = 0
            .KeySize = 0
            .IvSize = 0
            .IvExplicitSize = 0
            .TagSize = 0
            .DigestAlgo = 0
            .DigestSize = 0
            .UseRsaKeyTransport = False
            Select Case lCipherSuite
            Case TLS_CS_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256, _
                    TLS_CS_AES_128_GCM_SHA256, TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256, TLS_CS_RSA_WITH_AES_128_GCM_SHA256
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = LNG_SHA256_HASHSZ
            Case TLS_CS_AES_256_GCM_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .DigestAlgo = ucsTlsAlgoDigestSha384
                .DigestSize = LNG_SHA384_HASHSZ
            Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA, TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA, TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA, TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA, _
                    TLS_CS_RSA_WITH_AES_128_CBC_SHA, TLS_CS_RSA_WITH_AES_256_CBC_SHA
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = LNG_SHA256_HASHSZ
                .MacAlgo = ucsTlsAlgoDigestSha1
                .MacSize = LNG_SHA1_HASHSZ
#If ImplExoticCiphers Then
            Case TLS_CS_RSA_WITH_AES_128_CBC_SHA256, TLS_CS_RSA_WITH_AES_256_CBC_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256, TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA256
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = LNG_SHA256_HASHSZ
                .MacAlgo = .DigestAlgo
                .MacSize = .DigestSize
            Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA384
                .DigestAlgo = ucsTlsAlgoDigestSha384
                .DigestSize = LNG_SHA384_HASHSZ
                .MacAlgo = .DigestAlgo
                .MacSize = .DigestSize
#End If
            End Select
            Select Case lCipherSuite
            Case TLS_CS_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                .BulkAlgo = ucsTlsAlgoBulkChacha20Poly1305
                .KeySize = LNG_CHACHA20_KEYSZ
                .IvSize = LNG_CHACHA20POLY1305_IVSZ
                .TagSize = LNG_CHACHA20POLY1305_TAGSZ
            Case TLS_CS_AES_128_GCM_SHA256, TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256, TLS_CS_RSA_WITH_AES_128_GCM_SHA256
                .BulkAlgo = ucsTlsAlgoBulkAesGcm128
                .KeySize = LNG_AES128_KEYSZ
                .IvSize = LNG_AESGCM_IVSZ
                If lCipherSuite <> TLS_CS_AES_128_GCM_SHA256 Then
                    .IvExplicitSize = 8
                End If
                .TagSize = LNG_AESGCM_TAGSZ
            Case TLS_CS_AES_256_GCM_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .BulkAlgo = ucsTlsAlgoBulkAesGcm256
                .KeySize = LNG_AES256_KEYSZ
                .IvSize = LNG_AESGCM_IVSZ
                If lCipherSuite <> TLS_CS_AES_256_GCM_SHA384 Then
                    .IvExplicitSize = 8
                End If
                .TagSize = LNG_AESGCM_TAGSZ
            Case TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA, TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA, TLS_CS_RSA_WITH_AES_128_CBC_SHA
                .BulkAlgo = ucsTlsAlgoBulkAesCbc128
                .KeySize = LNG_AES128_KEYSZ
                .IvSize = LNG_AESCBC_IVSZ
                .IvExplicitSize = .IvSize
            Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA, TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA, TLS_CS_RSA_WITH_AES_256_CBC_SHA
                .BulkAlgo = ucsTlsAlgoBulkAesCbc256
                .KeySize = LNG_AES256_KEYSZ
                .IvSize = LNG_AESCBC_IVSZ
                .IvExplicitSize = .IvSize
#If ImplExoticCiphers Then
            Case TLS_CS_RSA_WITH_AES_128_CBC_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256, TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA256
                .BulkAlgo = ucsTlsAlgoBulkAesCbc128
                .KeySize = LNG_AES128_KEYSZ
                .IvSize = LNG_AESCBC_IVSZ
                .IvExplicitSize = .IvSize
            Case TLS_CS_RSA_WITH_AES_256_CBC_SHA256, TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA384
                .BulkAlgo = ucsTlsAlgoBulkAesCbc256
                .KeySize = LNG_AES256_KEYSZ
                .IvSize = LNG_AESCBC_IVSZ
                .IvExplicitSize = .IvSize
#End If
            End Select
            Select Case lCipherSuite
            Case TLS_CS_RSA_WITH_AES_128_CBC_SHA, TLS_CS_RSA_WITH_AES_256_CBC_SHA, _
                    TLS_CS_RSA_WITH_AES_128_GCM_SHA256, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .UseRsaKeyTransport = True
#If ImplExoticCiphers Then
            Case TLS_CS_RSA_WITH_AES_128_CBC_SHA256, TLS_CS_RSA_WITH_AES_256_CBC_SHA256
                .UseRsaKeyTransport = True
#End If
            End Select
            Select Case lCipherSuite
            Case TLS_CS_AES_128_GCM_SHA256, TLS_CS_AES_256_GCM_SHA384, TLS_CS_CHACHA20_POLY1305_SHA256
                .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
            End Select
            If .BulkAlgo = 0 Or .DigestAlgo = 0 Then
                ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_CIPHER_SUITE, "%1", pvTlsGetCipherSuiteName(.CipherSuite))
            End If
        End If
    End With
End Sub

Private Function pvTlsGetSortedCipherSuites(ByVal eFilter As UcsTlsLocalFeaturesEnum, Optional AlgoObjId As String) As Collection
    Const PREF      As Long = &H1000
    Dim oRetVal     As Collection
    Dim bNeedEcdsa  As Boolean
    Dim bNeedRsa    As Boolean
    
    Set oRetVal = New Collection
    If (eFilter And ucsTlsSupportTls13) <> 0 Then
        '--- first if AES preferred over Chacha20
        If pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm128) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128) Then
            oRetVal.Add TLS_CS_AES_128_GCM_SHA256
        End If
        If pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm256) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            oRetVal.Add TLS_CS_AES_256_GCM_SHA384
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkChacha20Poly1305) Then
            oRetVal.Add TLS_CS_CHACHA20_POLY1305_SHA256
        End If
        '--- least preferred AES
        If Not pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm128) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128) Then
            oRetVal.Add TLS_CS_AES_128_GCM_SHA256
        End If
        If Not pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm256) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            oRetVal.Add TLS_CS_AES_256_GCM_SHA384
        End If
    End If
    If (eFilter And ucsTlsSupportTls12) <> 0 Then
        bNeedEcdsa = (AlgoObjId = szOID_ECC_PUBLIC_KEY Or AlgoObjId = szOID_ECC_CURVE_P256 Or AlgoObjId = szOID_ECC_CURVE_P384 Or AlgoObjId = szOID_ECC_CURVE_P521 Or LenB(AlgoObjId) = 0)
        bNeedRsa = (AlgoObjId = szOID_RSA_RSA Or AlgoObjId = szOID_RSA_SSA_PSS Or LenB(AlgoObjId) = 0)
        '--- first if AES preferred over Chacha20
        If pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm128) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
        End If
        If pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm256) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkChacha20Poly1305) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
            End If
        End If
        '--- least preferred AES
        If Not pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm128) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
        End If
        If Not pvCryptoIsSupported(PREF + ucsTlsAlgoBulkAesGcm256) And pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
        End If
        '--- legacy AES in CBC mode
#If ImplExoticCiphers Then
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc128) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA256
            End If
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc256) Then
            If bNeedEcdsa Then
                oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384
            End If
            If bNeedRsa Then
                oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA384
            End If
        End If
#End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc128) And bNeedEcdsa Then
            oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc128) And bNeedRsa Then
            oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc256) And bNeedEcdsa Then
            oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc256) And bNeedRsa Then
            oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA
        End If
        '--- no perfect forward secrecy -> least preferred
        If pvCryptoIsSupported(ucsTlsAlgoExchCertificate) And bNeedRsa Then
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm128) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_128_GCM_SHA256
            End If
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_256_GCM_SHA384
            End If
#If ImplExoticCiphers Then
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc128) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_128_CBC_SHA256
            End If
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc256) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_256_CBC_SHA256
            End If
#End If
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc128) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_128_CBC_SHA
            End If
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesCbc256) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_256_CBC_SHA
            End If
        End If
    End If
    Set pvTlsGetSortedCipherSuites = oRetVal
End Function

Private Sub pvTlsClearLastError(uCtx As UcsTlsContext)
    With uCtx
        .LastErrNumber = 0
        .LastErrSource = vbNullString
        .LastError = vbNullString
        .LastAlertCode = 0
        Set .BlocksStack = Nothing
    End With
End Sub

Private Sub pvTlsSetLastError( _
            uCtx As UcsTlsContext, _
            Optional ByVal ErrNumber As Long, _
            Optional ErrSource As String, _
            Optional ErrDescription As String, _
            Optional ByVal AlertCode As UcsTlsAlertDescriptionsEnum = -1)
    Const FUNC_NAME     As String = "pvTlsSetLastError"
    
    With uCtx
        .LastErrNumber = ErrNumber
        .LastErrSource = ErrSource
        .LastError = ErrDescription
        .LastAlertCode = AlertCode
        If LenB(ErrDescription) = 0 And AlertCode = -1 Then
            Set .BlocksStack = Nothing
        Else
            If AlertCode >= 0 Then
                pvTlsBuildAlert uCtx, .SendBuffer, AlertCode, TLS_ALERT_LEVEL_FATAL
            End If
            .State = ucsTlsStateClosed
        End If
        #If ImplCaptureTraffic <> 0 Then
            Clipboard.Clear
            Clipboard.SetText TlsConcatCollection(.TrafficDump, vbCrLf)
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, "Traffic dump copied to clipboard"
            #End If
        #End If
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
    Dim baZeroes()      As Byte
    
    With uCtx
        If .HandshakeMessages.Size = 0 Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
        End If
        pvTlsGetHandshakeHash uCtx, baHandshakeHash
        pvArrayAllocate baZeroes, .DigestSize, FUNC_NAME & ".baZeroes"
        pvTlsHkdfExtract baEarlySecret, .DigestAlgo, baZeroes, baZeroes
        pvTlsGetHash baEmptyHash, .DigestAlgo, baEmpty
        pvTlsHkdfExpandLabel baDerivedSecret, .DigestAlgo, baEarlySecret, "derived", baEmptyHash, .DigestSize
        pvTlsGetSharedSecret baSharedSecret, .ExchAlgo, .LocalExchPrivate, .RemoteExchPublic
        pvTlsHkdfExtract .HandshakeSecret, .DigestAlgo, baDerivedSecret, baSharedSecret
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .HandshakeSecret, IIf(.IsServer, "c", "s") & " hs traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "CLIENT", "SERVER") & "_HANDSHAKE_TRAFFIC_SECRET", .RemoteTrafficSecret
        pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .HandshakeSecret, IIf(.IsServer, "s", "c") & " hs traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
        .LocalTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "SERVER", "CLIENT") & "_HANDSHAKE_TRAFFIC_SECRET", .LocalTrafficSecret
    End With
End Sub

Private Sub pvTlsDeriveApplicationSecrets(uCtx As UcsTlsContext, baHandshakeHash() As Byte)
    Const FUNC_NAME     As String = "pvTlsDeriveApplicationSecrets"
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baEmpty()       As Byte
    Dim baZeroes()      As Byte
    
    With uCtx
        If .HandshakeMessages.Size = 0 Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
        End If
        pvTlsGetHash baEmptyHash, .DigestAlgo, baEmpty
        pvTlsHkdfExpandLabel baDerivedSecret, .DigestAlgo, .HandshakeSecret, "derived", baEmptyHash, .DigestSize
        pvArrayAllocate baZeroes, .DigestSize, FUNC_NAME & ".baZeroes"
        pvTlsHkdfExtract .MasterSecret, .DigestAlgo, baDerivedSecret, baZeroes
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .MasterSecret, IIf(.IsServer, "c", "s") & " ap traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "CLIENT", "SERVER") & "_TRAFFIC_SECRET_0", .RemoteTrafficSecret
        pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .MasterSecret, IIf(.IsServer, "s", "c") & " ap traffic", baHandshakeHash, .DigestSize
        pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
        .LocalTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "SERVER", "CLIENT") & "_TRAFFIC_SECRET_0", .LocalTrafficSecret
    End With
End Sub

Private Sub pvTlsDeriveKeyUpdate(uCtx As UcsTlsContext, ByVal bLocalUpdate As Boolean)
    Const FUNC_NAME     As String = "pvTlsDeriveKeyUpdate"
    Dim baEmpty()       As Byte
    
    With uCtx
        If pvArraySize(.RemoteTrafficSecret) = 0 Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "RemoteTrafficSecret")
        End If
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .RemoteTrafficSecret, "traffic upd", baEmpty, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "CLIENT", "SERVER") & "_TRAFFIC_SECRET_0", .RemoteTrafficSecret
        If bLocalUpdate Then
            If pvArraySize(.LocalTrafficSecret) = 0 Then
                ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "LocalTrafficSecret")
            End If
            pvTlsHkdfExpandLabel .LocalTrafficSecret, .DigestAlgo, .LocalTrafficSecret, "traffic upd", baEmpty, .DigestSize
            pvTlsHkdfExpandLabel .LocalTrafficKey, .DigestAlgo, .LocalTrafficSecret, "key", baEmpty, .KeySize
            pvTlsHkdfExpandLabel .LocalTrafficIV, .DigestAlgo, .LocalTrafficSecret, "iv", baEmpty, .IvSize
            .LocalTrafficSeqNo = 0
            pvTlsLogSecret uCtx, IIf(.IsServer, "SERVER", "CLIENT") & "_TRAFFIC_SECRET_0", .LocalTrafficSecret
        End If
    End With
End Sub

Private Sub pvTlsHkdfExpandLabel(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsHkdfExpandLabel"
    Dim uOutput         As UcsBuffer
    Dim uInfo           As UcsBuffer
    Dim uInput          As UcsBuffer
    Dim lIdx            As Long
    Dim baLast()        As Byte
    
    sLabel = "tls13 " & sLabel
    pvArrayAllocate uInfo.Data, 3 + Len(sLabel) + 1 + pvArraySize(baContext), FUNC_NAME & ".uInfo"
    pvBufferWriteLong uInfo, lSize, Size:=2
    pvBufferWriteLong uInfo, Len(sLabel)
    pvBufferWriteString uInfo, sLabel
    pvBufferWriteLong uInfo, pvArraySize(baContext)
    pvBufferWriteArray uInfo, baContext
    pvBufferWriteEOF uInfo
    lIdx = 1
    Do While uOutput.Size < lSize
        uInput.Size = 0
        pvBufferWriteArray uInput, baLast
        pvBufferWriteArray uInput, uInfo.Data
        pvBufferWriteLong uInput, lIdx
        pvBufferWriteEOF uInput
        pvTlsGetHmac baLast, eHash, baKey, uInput.Data
        pvBufferWriteArray uOutput, baLast
        lIdx = lIdx + 1
    Loop
    uOutput.Size = lSize
    pvBufferWriteEOF uOutput
    baRetVal = uOutput.Data
End Sub

Private Sub pvTlsHkdfExtract(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte)
    pvTlsGetHmac baRetVal, eHash, baKey, baInput
End Sub

'= legacy PRF-based key derivation functions =============================

Private Sub pvTlsDeriveLegacySecrets(uCtx As UcsTlsContext)
    Const FUNC_NAME     As String = "pvTlsDeriveLegacySecrets"
    Dim baPreMasterSecret() As Byte
    Dim baHandshakeHash() As Byte
    Dim uRandom         As UcsBuffer
    Dim uExpanded       As UcsBuffer
    
    With uCtx
        If pvArraySize(.RemoteExchRandom) = 0 Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_NO_REMOTE_RANDOM
        End If
        pvTlsGetSharedSecret baPreMasterSecret, .ExchAlgo, .LocalExchPrivate, .RemoteExchPublic
        #If (ImplCaptureTraffic And 2) <> 0 Then
            .TrafficDump.Add FUNC_NAME & ".baPreMasterSecret" & vbCrLf & TlsDesignDumpArray(baPreMasterSecret)
        #End If
        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_EXTENDED_MASTER_SECRET) Then
            pvTlsGetHandshakeHash uCtx, baHandshakeHash
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "extended master secret", baHandshakeHash, TLS_LEGACY_SECRET_SIZE
        Else
            If uCtx.IsServer Then
                pvBufferWriteArray uRandom, .RemoteExchRandom
                pvBufferWriteArray uRandom, .LocalExchRandom
            Else
                pvBufferWriteArray uRandom, .LocalExchRandom
                pvBufferWriteArray uRandom, .RemoteExchRandom
            End If
            pvBufferWriteEOF uRandom
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "master secret", uRandom.Data, TLS_LEGACY_SECRET_SIZE
        End If
        pvTlsLogSecret uCtx, "CLIENT_RANDOM", .MasterSecret
        #If (ImplCaptureTraffic And 2) <> 0 Then
            .TrafficDump.Add FUNC_NAME & ".MasterSecret" & vbCrLf & TlsDesignDumpArray(.MasterSecret)
        #End If
        uRandom.Size = 0
        If uCtx.IsServer Then
            pvBufferWriteArray uRandom, .LocalExchRandom
            pvBufferWriteArray uRandom, .RemoteExchRandom
        Else
            pvBufferWriteArray uRandom, .RemoteExchRandom
            pvBufferWriteArray uRandom, .LocalExchRandom
        End If
        pvBufferWriteEOF uRandom
        pvTlsKdfLegacyPrf uExpanded.Data, .DigestAlgo, .MasterSecret, "key expansion", uRandom.Data, 2 * (.MacSize + .KeySize + .IvSize)
        #If (ImplCaptureTraffic And 2) <> 0 Then
            .TrafficDump.Add FUNC_NAME & ".uExpanded.Data" & vbCrLf & TlsDesignDumpArray(uExpanded.Data)
        #End If
        If uCtx.IsServer Then
            pvBufferReadArray uExpanded, .RemoteLegacyNextMacKey, .MacSize  '--- not used w/ AEAD
            pvBufferReadArray uExpanded, .LocalLegacyNextMacKey, .MacSize   '--- not used w/ AEAD
        Else
            pvBufferReadArray uExpanded, .LocalLegacyNextMacKey, .MacSize   '--- not used w/ AEAD
            pvBufferReadArray uExpanded, .RemoteLegacyNextMacKey, .MacSize  '--- not used w/ AEAD
        End If
        If uCtx.IsServer Then
            pvBufferReadArray uExpanded, .RemoteLegacyNextTrafficKey, .KeySize
            pvBufferReadArray uExpanded, .LocalLegacyNextTrafficKey, .KeySize
        Else
            pvBufferReadArray uExpanded, .LocalLegacyNextTrafficKey, .KeySize
            pvBufferReadArray uExpanded, .RemoteLegacyNextTrafficKey, .KeySize
        End If
        If uCtx.IsServer Then
            pvTlsGetRandom .RemoteLegacyNextTrafficIV, .IvSize
            pvBufferReadBlob uExpanded, VarPtr(.RemoteLegacyNextTrafficIV(0)), .IvSize - .IvExplicitSize
            pvTlsGetRandom .LocalLegacyNextTrafficIV, .IvSize
            pvBufferReadBlob uExpanded, VarPtr(.LocalLegacyNextTrafficIV(0)), .IvSize - .IvExplicitSize
        Else
            pvTlsGetRandom .LocalLegacyNextTrafficIV, .IvSize
            pvBufferReadBlob uExpanded, VarPtr(.LocalLegacyNextTrafficIV(0)), .IvSize - .IvExplicitSize
            pvTlsGetRandom .RemoteLegacyNextTrafficIV, .IvSize
            pvBufferReadBlob uExpanded, VarPtr(.RemoteLegacyNextTrafficIV(0)), .IvSize - .IvExplicitSize
        End If
    End With
End Sub

Private Sub pvTlsKdfLegacyPrf(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long)
    Dim uOutput         As UcsBuffer
    Dim uSeed           As UcsBuffer
    Dim uInput          As UcsBuffer
    Dim baLast()        As Byte
    Dim baHmac()        As Byte
    Dim baTemp()        As Byte
    
    '--- for TLS 1.2 -- PRF(secret, label, seed) = P_<hash>(secret, label + seed)
    pvBufferWriteString uSeed, sLabel
    pvBufferWriteArray uSeed, baContext
    pvBufferWriteEOF uSeed
    baLast = uSeed.Data
    Do While uOutput.Size < lSize
        baTemp = baLast
        pvTlsGetHmac baLast, eHash, baKey, baTemp
        pvBufferWriteArray uInput, baLast
        pvBufferWriteArray uInput, uSeed.Data
        pvBufferWriteEOF uInput
        pvTlsGetHmac baHmac, eHash, baKey, uInput.Data
        pvBufferWriteArray uOutput, baHmac
    Loop
    uOutput.Size = lSize
    pvBufferWriteEOF uOutput
    baRetVal = uOutput.Data
End Sub

Private Sub pvTlsGetHandshakeHash(uCtx As UcsTlsContext, baRetVal() As Byte)
    Const FUNC_NAME     As String = "pvTlsGetHandshakeHash"
    
    With uCtx
        pvTlsGetHash baRetVal, .DigestAlgo, .HandshakeMessages.Data, Size:=.HandshakeMessages.Size
        #If (ImplCaptureTraffic And 2) <> 0 Then
            .TrafficDump.Add FUNC_NAME & ".baRetVal" & vbCrLf & TlsDesignDumpArray(baRetVal)
        #End If
    End With
End Sub

Private Sub pvTlsAppendHandshakeHash(uCtx As UcsTlsContext, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsAppendHandshakeHash"
    
    With uCtx
        #If (ImplCaptureTraffic And 2) <> 0 Then
            .TrafficDump.Add FUNC_NAME & ".baInput(" & lSize & ")" & vbCrLf & TlsDesignDumpArray(baInput, lPos, lSize)
        #End If
        pvBufferWriteBlob .HandshakeMessages, VarPtr(baInput(lPos)), lSize
    End With
End Sub

Private Sub pvTlsResetHandshakeHash(uCtx As UcsTlsContext)
    With uCtx
        .HandshakeMessages.Size = 0
        pvBufferWriteEOF .HandshakeMessages
    End With
End Sub

Private Sub pvTlsLogSecret(uCtx As UcsTlsContext, sLabel As String, baSecret() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim sFileName       As String
    Dim nFile           As Integer
    
    On Error GoTo EH
    sFileName = String$(1000, 0)
    Call GetEnvironmentVariable(StrPtr("SSLKEYLOGFILE"), StrPtr(sFileName), Len(sFileName) + 1)
    sFileName = Left$(sFileName, InStr(sFileName, vbNullChar) - 1)
    If LenB(sFileName) <> 0 Then
        If Size < 0 Then
            Size = pvArraySize(baSecret) - Pos
        End If
        nFile = FreeFile
        Open sFileName For Append Shared As nFile
        Print #nFile, sLabel & " " & pvToHex(VarPtr(uCtx.LocalExchRandom(0)), UBound(uCtx.LocalExchRandom) + 1) & " " & pvToHex(VarPtr(baSecret(Pos)), Size)
        Close nFile
    End If
EH:
End Sub

'= crypto wrappers =======================================================

Private Sub pvTlsGetRandom(baRetVal() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsGetRandom"
    
    If lSize > 0 Then
        pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        pvCryptoRandomBytes VarPtr(baRetVal(0)), lSize
    Else
        baRetVal = vbNullString
    End If
End Sub

Private Sub pvTlsGetHash(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const FUNC_NAME     As String = "pvTlsGetHash"

    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        If Not pvCryptoHashSha256(baRetVal, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
        End If
    Case ucsTlsAlgoDigestSha384
        If Not pvCryptoHashSha384(baRetVal, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
        End If
    Case ucsTlsAlgoDigestSha512
        If Not pvCryptoHashSha512(baRetVal, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha512")
        End If
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, "Unsupported hash type " & eHash
    End Select
End Sub

Private Sub pvTlsGetHmac(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const FUNC_NAME     As String = "pvTlsGetHmac"
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha1
        If Not pvCryptoHmacSha1(baRetVal, baKey, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmacSha1")
        End If
    Case ucsTlsAlgoDigestSha256
        If Not pvCryptoHmacSha256(baRetVal, baKey, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmacSha256")
        End If
    Case ucsTlsAlgoDigestSha384
        If Not pvCryptoHmacSha384(baRetVal, baKey, baInput, Pos, Size) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmacSha384")
        End If
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, "Unsupported HMAC type " & eHash
    End Select
End Sub

Private Sub pvTlsGetHelloRetryRandom(baRetVal() As Byte)
    pvArrayByte baRetVal, &HCF, &H21, &HAD, &H74, &HE5, &H9A, &H61, &H11, &HBE, &H1D, &H8C, &H2, &H1E, &H65, &HB8, &H91, &HC2, &HA2, &H11, &H16, &H7A, &HBB, &H8C, &H5E, &H7, &H9E, &H9, &HE2, &HC8, &HA8, &H33, &H9C
End Sub

Private Function pvTlsBulkDecrypt(ByVal eBulk As UcsTlsCryptoAlgorithmsEnum, baRemoteIV() As Byte, baRemoteKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvTlsBulkDecrypt"
    
    Select Case eBulk
    Case ucsTlsAlgoBulkChacha20Poly1305
        If Not pvCryptoBulkChacha20Poly1305Decrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAadSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoBulkAesGcm128, ucsTlsAlgoBulkAesGcm256
        If Not pvCryptoBulkAesGcmDecrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAadSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoBulkAesCbc128, ucsTlsAlgoBulkAesCbc256
        If Not pvCryptoBulkAesCbcDecrypt(baRemoteIV, baRemoteKey, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, "Unsupported bulk type " & eBulk
    End Select
    '--- success
    pvTlsBulkDecrypt = True
QH:
End Function

Private Sub pvTlsBulkEncrypt(ByVal eBulk As UcsTlsCryptoAlgorithmsEnum, baLocalIV() As Byte, baLocalKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsBulkEncrypt"
    
    Select Case eBulk
    Case ucsTlsAlgoBulkChacha20Poly1305
        If Not pvCryptoBulkChacha20Poly1305Encrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAadSize, baBuffer, lPos, lSize) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoBulkChacha20Poly1305Encrypt")
        End If
    Case ucsTlsAlgoBulkAesGcm128, ucsTlsAlgoBulkAesGcm256
        If Not pvCryptoBulkAesGcmEncrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAadSize, baBuffer, lPos, lSize) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoBulkAesGcmEncrypt")
        End If
    Case ucsTlsAlgoBulkAesCbc128, ucsTlsAlgoBulkAesCbc256
        If Not pvCryptoBulkAesCbcEncrypt(baLocalIV, baLocalKey, baBuffer, lPos, lSize) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoBulkAesCbcEncrypt")
        End If
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, "Unsupported bulk type " & eBulk
    End Select
End Sub

Private Sub pvTlsGetSharedSecret(baRetVal() As Byte, ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte)
    Const FUNC_NAME     As String = "pvTlsGetSharedSecret"
    
    Select Case eKeyX
    Case ucsTlsAlgoExchX25519
        If Not pvCryptoEcdhCurve25519SharedSecret(baRetVal, baPriv, baPub) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEcdhCurve25519SharedSecret")
        End If
    Case ucsTlsAlgoExchSecp256r1
        If Not pvCryptoEcdhSecp256r1SharedSecret(baRetVal, baPriv, baPub) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEcdhSecp256r1SharedSecret")
        End If
    Case ucsTlsAlgoExchSecp384r1
        If Not pvCryptoEcdhSecp384r1SharedSecret(baRetVal, baPriv, baPub) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEcdhSecp384r1SharedSecret")
        End If
    Case ucsTlsAlgoExchCertificate
        baRetVal = baPriv
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, "Unsupported exchange " & eKeyX
    End Select
End Sub

Private Function pvTlsGetExchGroupName(ByVal lExchGroup As Long) As String
    Select Case lExchGroup
    Case TLS_GROUP_X25519
        pvTlsGetExchGroupName = "X25519"
    Case TLS_GROUP_X448
        pvTlsGetExchGroupName = "X448"
    Case TLS_GROUP_SECP256R1
        pvTlsGetExchGroupName = "secp256r1"
    Case TLS_GROUP_SECP384R1
        pvTlsGetExchGroupName = "secp384r1"
    Case TLS_GROUP_SECP521R1
        pvTlsGetExchGroupName = "secp521r1"
    Case Else
        pvTlsGetExchGroupName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lExchGroup))
    End Select
End Function

Private Function pvTlsGetCipherSuiteName(ByVal lCipherSuite As Long) As String
    Select Case lCipherSuite
    Case TLS_CS_AES_128_GCM_SHA256
        pvTlsGetCipherSuiteName = "TLS_AES_128_GCM_SHA256"
    Case TLS_CS_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "TLS_AES_256_GCM_SHA384"
    Case TLS_CS_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "TLS_CHACHA20_POLY1305_SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES128-GCM-SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_AES_128_GCM_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES128-GCM-SHA256"
    Case TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-RSA-CHACHA20-POLY1305"
    Case TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-CHACHA20-POLY1305"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES128-SHA"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES256-SHA"
    Case TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES128-SHA"
    Case TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES256-SHA"
    Case TLS_CS_RSA_WITH_AES_128_GCM_SHA256
        pvTlsGetCipherSuiteName = "AES128-GCM-SHA256"
    Case TLS_CS_RSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "AES256-GCM-SHA384"
    Case TLS_CS_RSA_WITH_AES_128_CBC_SHA
        pvTlsGetCipherSuiteName = "AES128-SHA"
    Case TLS_CS_RSA_WITH_AES_256_CBC_SHA
        pvTlsGetCipherSuiteName = "AES256-SHA"
    Case TLS_CS_RSA_WITH_AES_128_CBC_SHA256
        pvTlsGetCipherSuiteName = "AES128-SHA256"
    Case TLS_CS_RSA_WITH_AES_256_CBC_SHA256
        pvTlsGetCipherSuiteName = "AES256-SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES128-SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES256-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_AES_128_CBC_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES128-SHA256"
    Case TLS_CS_ECDHE_RSA_WITH_AES_256_CBC_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES256-SHA384"
    Case Else
        pvTlsGetCipherSuiteName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lCipherSuite))
    End Select
End Function

Private Function pvTlsGetSignatureName(ByVal lSignatureScheme As Long) As String
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        pvTlsGetSignatureName = "RSA_PKCS1_SHA1"
    Case TLS_SIGNATURE_ECDSA_SHA1
        pvTlsGetSignatureName = "ECDSA_SHA1"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA224
        pvTlsGetSignatureName = "RSA_PKCS1_SHA224"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256
        pvTlsGetSignatureName = "RSA_PKCS1_SHA256"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA384
        pvTlsGetSignatureName = "RSA_PKCS1_SHA384"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA512
        pvTlsGetSignatureName = "RSA_PKCS1_SHA512"
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        pvTlsGetSignatureName = "ECDSA_SECP256R1_SHA256"
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        pvTlsGetSignatureName = "ECDSA_SECP384R1_SHA384"
    Case TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        pvTlsGetSignatureName = "ECDSA_SECP521R1_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256
        pvTlsGetSignatureName = "RSA_PSS_RSAE_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384
        pvTlsGetSignatureName = "RSA_PSS_RSAE_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        pvTlsGetSignatureName = "RSA_PSS_RSAE_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        pvTlsGetSignatureName = "RSA_PSS_PSS_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        pvTlsGetSignatureName = "RSA_PSS_PSS_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        pvTlsGetSignatureName = "RSA_PSS_PSS_SHA512"
    Case Else
        pvTlsGetSignatureName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lSignatureScheme))
    End Select
End Function

Private Function pvTlsGetHashName(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As String
    Select Case eAlgo
    Case ucsTlsAlgoDigestMd5
        pvTlsGetHashName = "MD5"
    Case ucsTlsAlgoDigestSha1
        pvTlsGetHashName = "SHA1"
    Case ucsTlsAlgoDigestSha224
        pvTlsGetHashName = "SHA224"
    Case ucsTlsAlgoDigestSha256
        pvTlsGetHashName = "SHA256"
    Case ucsTlsAlgoDigestSha384
        pvTlsGetHashName = "SHA384"
    Case ucsTlsAlgoDigestSha512
        pvTlsGetHashName = "SHA512"
    Case Else
        pvTlsGetHashName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(eAlgo))
    End Select
End Function

Private Function pvTlsGetHashAlgId(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As Long
    Select Case eAlgo
    Case ucsTlsAlgoDigestMd5
        pvTlsGetHashAlgId = CALG_SSL3_SHAMD5
    Case ucsTlsAlgoDigestSha1
        pvTlsGetHashAlgId = CALG_SHA1
    Case ucsTlsAlgoDigestSha256
        pvTlsGetHashAlgId = CALG_SHA_256
    Case ucsTlsAlgoDigestSha384
        pvTlsGetHashAlgId = CALG_SHA_384
    Case ucsTlsAlgoDigestSha512
        pvTlsGetHashAlgId = CALG_SHA_512
    Case Else
        pvTlsGetHashAlgId = -1
    End Select
End Function

Private Sub pvTlsSignatureSign(baRetVal() As Byte, cCerts As Collection, cPrivKey As Collection, ByVal lSignatureScheme As Long, baVerifyData() As Byte)
    Const FUNC_NAME     As String = "pvTlsSignatureSign"
    Dim baSignature()   As Byte
    Dim uKeyInfo        As UcsKeyInfo
    Dim lHashSize       As Long
    Dim baEnc()         As Byte
    Dim baVerifyHash()  As Byte
    Dim sHashAlg        As String
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim lPadPtr         As Long
    Dim lSize           As Long
    Dim dwFlags         As Long
    Dim lAlgId          As Long
    Dim hProvHash       As Long
    Dim hHash           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim sErrDesc        As String
        
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, "Signing with " & pvTlsGetSignatureName(lSignatureScheme) & " signature"
    #End If
    If Not pvAsn1DecodePrivateKey(cCerts, cPrivKey, uKeyInfo) Then
        sErrDesc = ERR_UNSUPPORTED_PRIVATE_KEY
        GoTo QH
    End If
    lHashSize = pvTlsSignatureHashSize(lSignatureScheme)
    If uKeyInfo.hNKey <> 0 Then
        If uKeyInfo.AlgoObjId = szOID_RSA_RSA Then
            sHashAlg = pvTlsGetHashName(pvTlsSignatureDigestAlgo(lSignatureScheme))
            uPadInfo.pszAlgId = StrPtr(sHashAlg)
            If (lSignatureScheme And &HFF) = (TLS_SIGNATURE_RSA_PKCS1_SHA1 And &HFF) Then
                dwFlags = BCRYPT_PAD_PKCS1
            ElseIf (lSignatureScheme \ &H100) = (TLS_SIGNATURE_RSA_PSS_RSAE_SHA256 \ &H100) Then
                dwFlags = BCRYPT_PAD_PSS
                uPadInfo.cbSalt = lHashSize
            End If
            lPadPtr = VarPtr(uPadInfo)
        End If
        pvArrayHash lHashSize, baVerifyData, baVerifyHash
        hResult = NCryptSignHash(uKeyInfo.hNKey, lPadPtr, baVerifyHash(0), UBound(baVerifyHash) + 1, ByVal 0, 0, lSize, dwFlags)
        If hResult < 0 Then
            sApiSource = "NCryptSignHash"
            GoTo QH
        End If
        pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        hResult = NCryptSignHash(uKeyInfo.hNKey, lPadPtr, baVerifyHash(0), UBound(baVerifyHash) + 1, baRetVal(0), lSize, lSize, dwFlags)
        If hResult < 0 Then
            sApiSource = "NCryptSignHash#2"
            GoTo QH
        End If
        pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        If uKeyInfo.AlgoObjId = szOID_ECC_PUBLIC_KEY Then
            baSignature = baRetVal
            If Not pvAsn1EncodeEcdsaSignature(baRetVal, baSignature, (UBound(baSignature) + 1) \ 2) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "Asn1EncodeEcdsaSignature")
                GoTo QH
            End If
        End If
    ElseIf uKeyInfo.hProv <> 0 And uKeyInfo.hKey <> 0 Then
        Debug.Assert (lSignatureScheme And &HFF) = (TLS_SIGNATURE_RSA_PKCS1_SHA1 And &HFF)
        lAlgId = pvTlsGetHashAlgId(pvTlsSignatureDigestAlgo(lSignatureScheme))
        If CryptCreateHash(uKeyInfo.hProv, lAlgId, 0, 0, hHash) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptCreateHash"
            GoTo QH
        End If
        If CryptGetHashParam(hHash, HP_HASHSIZE, lSize, LenB(lSize), 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptGetHashParam(HP_HASHSIZE)"
            GoTo QH
        End If
        If lSize <> lHashSize Then
            sErrDesc = Replace(ERR_INVALID_HASH_SIZE, "%1", lSize & "<>" & lHashSize)
            GoTo QH
        End If
        pvArrayHash lHashSize, baVerifyData, baVerifyHash
        If CryptSetHashParam(hHash, HP_HASHVAL, baVerifyHash(0), 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSetHashParam(HP_HASHVAL)"
            GoTo QH
        End If
        If CryptSignHash(hHash, uKeyInfo.dwKeySpec, 0, 0, ByVal 0, lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSignHash"
            GoTo QH
        End If
        pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        If CryptSignHash(hHash, uKeyInfo.dwKeySpec, 0, 0, baRetVal(0), lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSignHash#2"
            GoTo QH
        End If
        pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
        pvArrayReverse baRetVal
    Else
        Select Case lSignatureScheme
        Case TLS_SIGNATURE_RSA_PKCS1_SHA1, TLS_SIGNATURE_RSA_PKCS1_SHA224, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
            Debug.Assert uKeyInfo.AlgoObjId = szOID_RSA_RSA
            If Not pvCryptoEmsaPkcs1Encode(baEnc, baVerifyData, uKeyInfo.BitLen, lHashSize) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoEmsaPkcs1Encode")
                GoTo QH
            End If
            If Not pvCryptoRsaCrtModExp(baEnc, uKeyInfo.PrivExp, uKeyInfo.Modulus, uKeyInfo.Prime1, uKeyInfo.Prime2, uKeyInfo.Coefficient, baRetVal) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoRsaCrtModExp")
                GoTo QH
            End If
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
                TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
            Debug.Assert uKeyInfo.AlgoObjId = szOID_RSA_RSA Or uKeyInfo.AlgoObjId = szOID_RSA_SSA_PSS
            If Not pvCryptoEmsaPssEncode(baEnc, baVerifyData, uKeyInfo.BitLen, lHashSize, lHashSize) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoEmsaPssEncode")
                GoTo QH
            End If
            If Not pvCryptoRsaCrtModExp(baEnc, uKeyInfo.PrivExp, uKeyInfo.Modulus, uKeyInfo.Prime1, uKeyInfo.Prime2, uKeyInfo.Coefficient, baRetVal) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoRsaCrtModExp")
                GoTo QH
            End If
        Case TLS_SIGNATURE_ECDSA_SHA1
            If Not pvCryptoHashSha1(baVerifyHash, baVerifyData) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha1")
                GoTo QH
            End If
            If uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256 Then
                GoTo Secp256r1Sign
            ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384 Then
                GoTo Secp384r1Sign
            End If
        Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
            Debug.Assert uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256
            If Not pvCryptoHashSha256(baVerifyHash, baVerifyData) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
                GoTo QH
            End If
Secp256r1Sign:
            If Not pvCryptoEcdsaSecp256r1Sign(baSignature, uKeyInfo.KeyBlob, baVerifyHash) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoEcdsaSecp256r1Sign")
                GoTo QH
            End If
            If Not pvAsn1EncodeEcdsaSignature(baRetVal, baSignature, LNG_SECP256R1_KEYSZ) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "Asn1EncodeEcdsaSignature")
                GoTo QH
            End If
        Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
            Debug.Assert uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384
            If Not pvCryptoHashSha384(baVerifyHash, baVerifyData) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
                GoTo QH
            End If
Secp384r1Sign:
            If Not pvCryptoEcdsaSecp384r1Sign(baSignature, uKeyInfo.KeyBlob, baVerifyHash) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "CryptoEcdsaSecp384r1Sign")
                GoTo QH
            End If
            If Not pvAsn1EncodeEcdsaSignature(baRetVal, baSignature, LNG_SECP384R1_KEYSZ) Then
                sErrDesc = Replace(ERR_CALL_FAILED, "%1", "Asn1EncodeEcdsaSignature")
                GoTo QH
            End If
        Case Else
            sErrDesc = Replace(ERR_UNSUPPORTED_SIGNATURE_SCHEME, "%1", pvTlsGetSignatureName(lSignatureScheme))
            GoTo QH
        End Select
    End If
    If pvArraySize(baRetVal) = 0 Then
        sErrDesc = Replace(ERR_SIGNATURE_FAILED, "%1", pvTlsGetSignatureName(lSignatureScheme))
    End If
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If hProvHash <> 0 Then
        Call CryptReleaseContext(hProvHash, 0)
    End If
    If uKeyInfo.hKey <> 0 Then
        Call CryptDestroyKey(uKeyInfo.hKey)
    End If
    If uKeyInfo.hProv <> 0 Then
        Call CryptReleaseContext(uKeyInfo.hProv, 0)
    End If
    If uKeyInfo.hNKey <> 0 Then
        Call NCryptFreeObject(uKeyInfo.hNKey)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
    If LenB(sErrDesc) <> 0 Then
        ErrRaise vbObjectError, FUNC_NAME, sErrDesc
    End If
End Sub

Private Function pvTlsSignatureVerify(baCert() As Byte, ByVal lSignatureScheme As Long, baVerifyData() As Byte, baSignature() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
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
    If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
        GoTo UnsupportedCertificate
    End If
    lHashSize = pvTlsSignatureHashSize(lSignatureScheme)
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1, TLS_SIGNATURE_RSA_PKCS1_SHA224, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If Not pvCryptoRsaModExp(baSignature, uCertInfo.PubExp, uCertInfo.Modulus, baDecr) Then
            GoTo InvalidSignature
        End If
        If Not pvCryptoEmsaPkcs1Decode(baVerifyData, baDecr, lHashSize) Then
            GoTo InvalidSignature
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If Not pvCryptoRsaModExp(baSignature, uCertInfo.PubExp, uCertInfo.Modulus, baDecr) Then
            GoTo InvalidSignature
        End If
        If Not pvCryptoEmsaPssDecode(baVerifyData, baDecr, uCertInfo.BitLen, lHashSize, lHashSize) Then
            GoTo InvalidSignature
        End If
    Case TLS_SIGNATURE_ECDSA_SHA1, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        If uCertInfo.AlgoObjId <> szOID_ECC_PUBLIC_KEY Then
            GoTo UnsupportedPublicKey
        End If
        pvArrayHash lHashSize, baVerifyData, baVerifyHash
        lCurveSize = UBound(uCertInfo.KeyBlob) \ 2
        If Not pvAsn1DecodeEcdsaSignature(baPlainSig, baSignature, lCurveSize) Then
            GoTo InvalidSignature
        End If
        If UBound(baVerifyHash) + 1 < lCurveSize Then
            '--- note: when hash size is less than curve size must left-pad w/ zeros (right-align hash) -> deprecated
            '---       incl. ECDSA_SECP384R1_SHA256 only
            baTemp = baVerifyHash
            pvArrayAllocate baVerifyHash, lCurveSize, FUNC_NAME & ".baVerifyHash"
            Call CopyMemory(baVerifyHash(lCurveSize - UBound(baTemp) - 1), baTemp(0), UBound(baTemp) + 1)
            bDeprecated = True
        ElseIf UBound(baVerifyHash) + 1 > lCurveSize Then
            '--- note: when hash size is above curve size the excess is ignored -> deprecated
            '---       incl. ECDSA_SECP256R1_SHA384, ECDSA_SECP256R1_SHA512 and ECDSA_SECP384R1_SHA512
            bDeprecated = True
        End If
        Select Case lCurveSize
        Case LNG_SECP256R1_KEYSZ
            If Not pvCryptoEcdsaSecp256r1Verify(uCertInfo.KeyBlob, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case LNG_SECP384R1_KEYSZ
            If Not pvCryptoEcdsaSecp384r1Verify(uCertInfo.KeyBlob, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case Else
            GoTo UnsupportedCurveSize
        End Select
    Case Else
        GoTo UnsupportedSignatureScheme
    End Select
    '--- success
    pvTlsSignatureVerify = True
QH:
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, IIf(pvTlsSignatureVerify, IIf(bSkip, "Skipping ", IIf(bDeprecated, "Deprecated ", "Valid ")), "Invalid ") & pvTlsGetSignatureName(lSignatureScheme) & " signature" & IIf(bDeprecated, " (lCurveSize=" & lCurveSize & " from server's public key)", vbNullString)
    #End If
    Exit Function
UnsupportedCertificate:
    sError = ERR_UNSUPPORTED_CERTIFICATE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidSignature:
    sError = ERR_INVALID_SIGNATURE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedPublicKey:
    sError = Replace(ERR_UNSUPPORTED_PUBLIC_KEY, "%1", uCertInfo.AlgoObjId)
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedCurveSize:
    sError = Replace(ERR_UNSUPPORTED_CURVE_SIZE, "%1", lCurveSize)
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedSignatureScheme:
    sError = Replace(ERR_UNSUPPORTED_SIGNATURE_SCHEME, "%1", pvTlsGetSignatureName(lSignatureScheme))
    eAlertCode = uscTlsAlertInternalError
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
    Resume QH
End Function

Private Function pvTlsSignatureHashSize(ByVal lSignatureScheme As Long) As Long
    pvTlsSignatureHashSize = pvTlsDigestHashSize(pvTlsSignatureDigestAlgo(lSignatureScheme))
End Function

Private Function pvTlsSignatureDigestAlgo(ByVal lSignatureScheme As Long) As UcsTlsCryptoAlgorithmsEnum
    Const TLS_SIGNATURE_ALGORITHM_RSA   As Long = 1
    Const TLS_SIGNATURE_ALGORITHM_DSA   As Long = 2
    Const TLS_SIGNATURE_ALGORITHM_ECDSA As Long = 3
    Const TLS_HASH_ALGORITHM_MD5        As Long = 1
    Const TLS_HASH_ALGORITHM_SHA1       As Long = 2
    Const TLS_HASH_ALGORITHM_SHA224     As Long = 3
    Const TLS_HASH_ALGORITHM_SHA256     As Long = 4
    Const TLS_HASH_ALGORITHM_SHA384     As Long = 5
    Const TLS_HASH_ALGORITHM_SHA512     As Long = 6
    
    Select Case (lSignatureScheme And &HFF)
    Case TLS_SIGNATURE_ALGORITHM_RSA, TLS_SIGNATURE_ALGORITHM_DSA, TLS_SIGNATURE_ALGORITHM_ECDSA
        Select Case lSignatureScheme \ &H100
        Case TLS_HASH_ALGORITHM_MD5
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestMd5
        Case TLS_HASH_ALGORITHM_SHA1
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha1
        Case TLS_HASH_ALGORITHM_SHA224
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha224
        Case TLS_HASH_ALGORITHM_SHA256
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case TLS_HASH_ALGORITHM_SHA384
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case TLS_HASH_ALGORITHM_SHA512
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    Case Else
        '--- 8 - Intrinsic for TLS 1.3
        Select Case lSignatureScheme
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    End Select
End Function

Private Function pvTlsDigestHashSize(ByVal eDigestAlgo As UcsTlsCryptoAlgorithmsEnum) As Long
    Select Case eDigestAlgo
    Case ucsTlsAlgoDigestMd5
        pvTlsDigestHashSize = LNG_MD5_HASHSZ
    Case ucsTlsAlgoDigestSha1
        pvTlsDigestHashSize = LNG_SHA1_HASHSZ
    Case ucsTlsAlgoDigestSha224
        pvTlsDigestHashSize = LNG_SHA224_HASHSZ
    Case ucsTlsAlgoDigestSha256
        pvTlsDigestHashSize = LNG_SHA256_HASHSZ
    Case ucsTlsAlgoDigestSha384
        pvTlsDigestHashSize = LNG_SHA384_HASHSZ
    Case ucsTlsAlgoDigestSha512
        pvTlsDigestHashSize = LNG_SHA512_HASHSZ
    Case Else
        '--- do nothing
    End Select
End Function

'= buffer management =====================================================

Private Sub pvBufferWriteRecordStart(uOutput As UcsBuffer, ByVal lContentType As Long, uCtx As UcsTlsContext)
    Dim lRecordPos      As Long
    Dim baLocalIV()     As Byte
    
    With uCtx
        lRecordPos = uOutput.Size
        '--- Record Header
        pvBufferWriteLong uOutput, lContentType
        pvBufferWriteLong uOutput, TLS_RECORD_VERSION, Size:=2
        pvBufferWriteBlockStart uOutput, Size:=2
            If pvArraySize(.LocalTrafficKey) > 0 Then
                pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
                If .IvExplicitSize > 0 Then '--- AES in TLS 1.2
                    pvBufferWriteBlob uOutput, VarPtr(baLocalIV(.IvSize - .IvExplicitSize)), .IvExplicitSize
                End If
                uOutput.Stack.Add Array(lRecordPos, uOutput.Size, baLocalIV), Before:=1
                '--- to be continued in end-of-record. . .
            End If
    End With
End Sub

Private Sub pvBufferWriteRecordEnd(uOutput As UcsBuffer, uCtx As UcsTlsContext)
    Const FUNC_NAME     As String = "pvBufferWriteRecordEnd"
    Dim vRecordData     As Variant
    Dim lRecordPos      As Long
    Dim baLocalIV()     As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim uAad            As UcsBuffer
    Dim baHmac()        As Byte
    Dim lPadding        As Long
    
    With uCtx
        If pvArraySize(.LocalTrafficKey) > 0 Then
                '--- . . . continues from start-of-record
                vRecordData = uOutput.Stack.Item(1)
                uOutput.Stack.Remove 1
                lRecordPos = vRecordData(0)
                lMessagePos = vRecordData(1)
                baLocalIV = vRecordData(2)
                lMessageSize = uOutput.Size - lMessagePos
                pvBufferWriteBlob uOutput, 0, .TagSize
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    pvBufferWriteLong uAad, 0, Size:=4
                    pvBufferWriteLong uAad, .LocalTrafficSeqNo, Size:=4
                    pvBufferWriteBlob uAad, VarPtr(uOutput.Data(lRecordPos)), 3
                    pvBufferWriteLong uAad, lMessageSize, Size:=2
                    Debug.Assert uAad.Size = TLS_LEGACY_AAD_SIZE
                    If .MacSize > 0 Then
                        If Not .LocalEncryptThenMac Then
                            pvBufferWriteBlob uAad, VarPtr(uOutput.Data(lMessagePos)), lMessageSize
                            pvTlsGetHmac baHmac, .MacAlgo, .LocalMacKey, uAad.Data, 0, uAad.Size
                            pvBufferWriteArray uOutput, baHmac
                        Else
                            pvArrayAllocate baHmac, .IvSize, FUNC_NAME & ".baHmac"
                        End If
                        lPadding = .IvSize - (uOutput.Size - lMessagePos) Mod .IvSize
                        Debug.Assert lPadding <= pvArraySize(baHmac)
                        Call FillMemory(baHmac(0), lPadding, lPadding - 1)
                        pvBufferWriteBlob uOutput, VarPtr(baHmac(0)), lPadding
                        lMessageSize = uOutput.Size - lMessagePos
                        If .LocalEncryptThenMac Then
                            pvBufferWriteBlob uOutput, 0, .MacSize
                        End If
                    End If
                End If
            pvBufferWriteBlockEnd uOutput
            #If (ImplCaptureTraffic And 1) <> 0 Then
                If lMessageSize <> 0 Then
                    .TrafficDump.Add FUNC_NAME & ".Output (unencrypted)" & vbCrLf & TlsDesignDumpArray(uOutput.Data, lRecordPos, uOutput.Size - lRecordPos - .TagSize - .MacSize)
                End If
            #End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                pvTlsBulkEncrypt .BulkAlgo, baLocalIV, .LocalTrafficKey, uOutput.Data, lRecordPos, TLS_AAD_SIZE, uOutput.Data, lMessagePos, lMessageSize
            ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                pvTlsBulkEncrypt .BulkAlgo, baLocalIV, .LocalTrafficKey, uAad.Data, 0, uAad.Size, uOutput.Data, lMessagePos, lMessageSize
                If .LocalEncryptThenMac Then
                    uAad.Size = uAad.Size - 2
                    pvBufferWriteLong uAad, lMessageSize + .IvExplicitSize, Size:=2
                    pvBufferWriteBlob uAad, VarPtr(uOutput.Data(lMessagePos - .IvExplicitSize)), lMessageSize + .IvExplicitSize
                    pvTlsGetHmac baHmac, .MacAlgo, .LocalMacKey, uAad.Data, 0, uAad.Size
                    uOutput.Size = uOutput.Size - .MacSize
                    pvBufferWriteArray uOutput, baHmac
                End If
            End If
            .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
        Else
            pvBufferWriteBlockEnd uOutput
        End If
    End With
End Sub

Private Sub pvBufferWriteBlockStart(uOutput As UcsBuffer, Optional ByVal Size As Long = 1)
    Dim lPos            As Long
    
    With uOutput
        If .Stack Is Nothing Then
            Set .Stack = New Collection
        End If
        If .Stack.Count = 0 Then
            .Stack.Add .Size
        Else
            .Stack.Add .Size, Before:=1
        End If
        lPos = .Size
        pvBufferWriteBlob uOutput, 0, Size
        '--- note: keep Size in uOutput.Data
        .Data(lPos) = (Size And &HFF)
    End With
End Sub

Private Sub pvBufferWriteBlockEnd(uOutput As UcsBuffer)
    Dim lPos            As Long
    
    With uOutput
        lPos = .Size
        .Size = .Stack.Item(1)
        .Stack.Remove 1
        pvBufferWriteLong uOutput, lPos - .Size - .Data(.Size), Size:=.Data(.Size)
        .Size = lPos
    End With
End Sub

Private Sub pvBufferWriteString(uOutput As UcsBuffer, sValue As String)
    pvBufferWriteArray uOutput, StrConv(sValue, vbFromUnicode)
End Sub

Private Sub pvBufferWriteArray(uOutput As UcsBuffer, baSrc() As Byte)
    Dim lSize       As Long
    
    With uOutput
        lSize = pvArraySize(baSrc)
        If lSize > 0 Then
            .Size = pvArrayWriteBlob(.Data, .Size, VarPtr(baSrc(0)), lSize)
        End If
    End With
End Sub

Private Sub pvBufferWriteLong(uOutput As UcsBuffer, ByVal lValue As Long, Optional ByVal Size As Long = 1)
    Static baTemp(0 To 3) As Byte
    Dim lPos            As Long

    With uOutput
        If Size <= 1 Then
            pvBufferWriteBlob uOutput, VarPtr(lValue), Size
        Else
            lPos = .Size
            pvBufferWriteBlob uOutput, 0, Size
            Call CopyMemory(baTemp(0), lValue, LenB(lValue))
            .Data(lPos + 0) = baTemp(Size - 1)
            .Data(lPos + 1) = baTemp(Size - 2)
            If Size >= 3 Then .Data(lPos + 2) = baTemp(Size - 3)
            If Size >= 4 Then .Data(lPos + 3) = baTemp(Size - 4)
        End If
    End With
End Sub

Private Sub pvBufferWriteBlob(uOutput As UcsBuffer, ByVal lPtr As Long, ByVal lSize As Long)
    uOutput.Size = pvArrayWriteBlob(uOutput.Data, uOutput.Size, lPtr, lSize)
End Sub

Private Function pvArrayWriteBlob(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const FUNC_NAME     As String = "pvArrayWriteBlob"
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baBuffer)
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), LenB(lBufPtr))
    If lBufPtr = 0 Then
        pvArrayAllocate baBuffer, Clamp(lPos + lSize, 256), FUNC_NAME & ".baBuffer"
    ElseIf UBound(baBuffer) < lPos + lSize - 1 Then
        pvArrayReallocate baBuffer, lPos + lSize, FUNC_NAME & ".baBuffer"
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baBuffer(lPos), ByVal lPtr, lSize)
    End If
    pvArrayWriteBlob = lPos + lSize
End Function

Private Sub pvBufferWriteEOF(uOutput As UcsBuffer)
    uOutput.Size = pvArrayWriteEOF(uOutput.Data, uOutput.Size)
End Sub

Private Function pvArrayWriteEOF(baBuffer() As Byte, ByVal lPos As Long) As Long
    Const FUNC_NAME     As String = "pvArrayWriteEOF"
    
    If pvArraySize(baBuffer) <> lPos Then
        pvArrayReallocate baBuffer, lPos, FUNC_NAME & ".baBuffer"
    End If
End Function

Private Sub pvBufferReadBlockStart(uInput As UcsBuffer, Optional ByVal Size As Long = 1, Optional BlockSize As Long)
    With uInput
        If .Stack Is Nothing Then
            Set .Stack = New Collection
        End If
        pvBufferReadLong uInput, BlockSize, Size
        If .Stack.Count = 0 Then
            .Stack.Add uInput.Pos + BlockSize
        Else
            .Stack.Add uInput.Pos + BlockSize, Before:=1
        End If
    End With
End Sub

Private Sub pvBufferReadBlockEnd(uInput As UcsBuffer)
    Dim lEnd          As Long
    
    With uInput
        lEnd = .Stack.Item(1)
        .Stack.Remove 1
        Debug.Assert .Pos = lEnd
    End With
End Sub

Private Sub pvBufferReadLong(uInput As UcsBuffer, lValue As Long, Optional ByVal Size As Long = 1)
    Static baTemp(0 To 3) As Byte
    
    lValue = 0
    With uInput
        If .Pos + Size <= pvArraySize(.Data) Then
            If Size <= 1 Then
                lValue = .Data(.Pos)
            Else
                baTemp(Size - 1) = .Data(.Pos + 0)
                baTemp(Size - 2) = .Data(.Pos + 1)
                If Size >= 3 Then baTemp(Size - 3) = .Data(.Pos + 2)
                If Size >= 4 Then baTemp(Size - 4) = .Data(.Pos + 3)
                Call CopyMemory(lValue, baTemp(0), Size)
            End If
        End If
        .Pos = .Pos + Size
    End With
End Sub

Private Sub pvBufferReadBlob(uInput As UcsBuffer, ByVal lPtr As Long, ByVal lSize As Long)
    Dim baDest()        As Byte
    
    pvBufferReadArray uInput, baDest, lSize
    lSize = pvArraySize(baDest)
    If lSize > 0 Then
        Call CopyMemory(ByVal lPtr, baDest(0), lSize)
    End If
End Sub

Private Sub pvBufferReadArray(uInput As UcsBuffer, baDest() As Byte, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvBufferReadArray"
    
    With uInput
        If lSize < 0 Then
            lSize = pvArraySize(.Data) - .Pos
        End If
        If lSize > 0 Then
            pvArrayAllocate baDest, lSize, FUNC_NAME & ".baDest"
            If .Pos + lSize <= pvArraySize(.Data) Then
                Call CopyMemory(baDest(0), .Data(.Pos), lSize)
            ElseIf .Pos < pvArraySize(.Data) Then
                Call CopyMemory(baDest(0), .Data(.Pos), pvArraySize(.Data) - .Pos)
            End If
        Else
            Erase baDest
        End If
        .Pos = .Pos + lSize
    End With
End Sub

Private Sub pvBufferReadString(uInput As UcsBuffer, sValue As String, ByVal lSize As Long)
    Dim baTemp()        As Byte
    
    pvBufferReadArray uInput, baTemp(), lSize
    sValue = StrConv(baTemp, vbUnicode)
End Sub

'= arrays helpers ========================================================

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    Debug.Assert RedimStats(MODULE_NAME & "." & sFuncName, UBound(baRetVal) + 1)
End Sub

Private Sub pvArrayReallocate(baArray() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim Preserve baArray(0 To lSize - 1) As Byte
    Else
        baArray = vbNullString
    End If
    Debug.Assert RedimStats(MODULE_NAME & "." & sFuncName, UBound(baArray) + 1)
End Sub

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), LenB(lPtr))
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
    
    Call CopyMemory(lTemp, ByVal ArrPtr(baBuffer), LenB(lTemp))
    Call CopyMemory(ByVal ArrPtr(baBuffer), ByVal ArrPtr(baInput), LenB(lTemp))
    Call CopyMemory(ByVal ArrPtr(baInput), lTemp, LenB(lTemp))
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
    Case LNG_SHA1_HASHSZ
        If Not pvCryptoHashSha1(baRetVal, baInput) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha1")
        End If
    Case LNG_SHA224_HASHSZ
        If Not pvCryptoHashSha224(baRetVal, baInput) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha224")
        End If
    Case LNG_SHA256_HASHSZ
        If Not pvCryptoHashSha256(baRetVal, baInput) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha256")
        End If
    Case LNG_SHA384_HASHSZ
        If Not pvCryptoHashSha384(baRetVal, baInput) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha384")
        End If
    Case LNG_SHA512_HASHSZ
        If Not pvCryptoHashSha512(baRetVal, baInput) Then
            ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHashSha512")
        End If
    Case Else
        ErrRaise vbObjectError, FUNC_NAME, Replace(ERR_INVALID_HASH_SIZE, "%1", lHashSize)
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

Private Function pvArrayAccumulateOr(baData() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte
    Dim lIdx            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baData)
    End If
    For lIdx = Pos To Size - 1
        pvArrayAccumulateOr = pvArrayAccumulateOr Or baData(lIdx)
    Next
End Function

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If pvArraySize(baFirst) = pvArraySize(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

Private Function pvToStringA(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToStringA = String$(lstrlenA(lPtr), 0)
        Call CopyMemory(ByVal pvToStringA, ByVal lPtr, Len(pvToStringA))
    End If
End Function

Private Function pvCollectionCount(oCol As Collection) As Long
    If Not oCol Is Nothing Then
        pvCollectionCount = oCol.Count
    End If
End Function

Private Function pvCollectionFirst(oCol As Collection, vKeys As Variant) As Variant
    Dim vElem       As Variant
    
    For Each vElem In vKeys
        If SearchCollection(oCol, vElem, RetVal:=pvCollectionFirst) Then
            Exit For
        End If
    Next
End Function

Private Function pvToHex(ByVal lPtr As Long, ByVal lSize As Long) As String
    Const FUNC_NAME     As String = "pvToHex"
    Dim aText()         As String
    Dim lByte           As Long
    Dim lIdx            As Long
    
    If lSize <> 0 Then
        ReDim aText(0 To lSize - 1) As String
        Debug.Assert RedimStats(FUNC_NAME & ".aText", 0)
        For lIdx = 0 To lSize - 1
            Call CopyMemory(lByte, ByVal lPtr, 1)
            lPtr = (lPtr Xor &H80000000) + 1 Xor &H80000000
            aText(lIdx) = Right$("0" & Hex$(lByte), 2)
        Next
        pvToHex = Join(aText, vbNullString)
    End If
End Function

Private Function pvIsValidServerName(sName As String) As Boolean
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sName)
        Select Case AscW(Mid$(sName, lIdx, 1))
        Case 65 To 90, 97 To 122, 48 To 57, 45, 46  '--- A-Z, a-z, 0-9, "-", "."
            '--- do nothing
        Case Else
            Exit Function
        End Select
    Next
    '--- success
    pvIsValidServerName = True
End Function

'=========================================================================
' Crypto
'=========================================================================

Private Function pvCryptoIsSupported(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As Boolean
    Const PREF          As Long = &H1000
    
    Select Case eAlgo
    Case ucsTlsAlgoExchSecp521r1
        '--- not supported
    Case ucsTlsAlgoBulkAesGcm128, ucsTlsAlgoBulkAesGcm256
        pvCryptoIsSupported = True
    Case ucsTlsAlgoBulkChacha20Poly1305
        pvCryptoIsSupported = True
    Case PREF + ucsTlsAlgoBulkAesGcm128, PREF + ucsTlsAlgoBulkAesGcm256
        '--- signal if AES preferred over Chacha20
        pvCryptoIsSupported = True
    Case Else
        pvCryptoIsSupported = True
    End Select
End Function

Private Function pvCryptoEmePkcs1Encode(baRetVal() As Byte, baMessage() As Byte, ByVal lBitLen As Long) As Boolean
    Const FUNC_NAME     As String = "CryptoEmePkcs1Encode"
    Dim lIdx            As Long
    
    '--- from RFC 8017, Section  7.2.1
    pvArrayAllocate baRetVal, (lBitLen + 7) \ 8, FUNC_NAME & ".baRetVal"
    If UBound(baMessage) > UBound(baRetVal) - 11 Then
        GoTo QH
    End If
    baRetVal(1) = 2
    pvCryptoRandomBytes VarPtr(baRetVal(2)), UBound(baRetVal) - UBound(baMessage) - 3
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

Private Function pvCryptoEmePkcs1Decode(baRetVal() As Byte, baMessage() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEmePkcs1Decode"
    Dim lIdx            As Long
    
    If baMessage(0) <> 0 Or baMessage(1) <> 2 Then
        GoTo QH
    End If
    For lIdx = 2 To UBound(baMessage)
        If baMessage(lIdx) = 0 Then
            lIdx = lIdx + 1
            Exit For
        End If
    Next
    If lIdx > UBound(baMessage) Then
        GoTo QH
    End If
    pvArrayAllocate baRetVal, UBound(baMessage) + 1 - lIdx, FUNC_NAME & ".baRetVal"
    Call CopyMemory(baRetVal(0), baMessage(lIdx), UBound(baRetVal) + 1)
    '--- success
    pvCryptoEmePkcs1Decode = True
QH:
End Function

Private Function pvCryptoEmsaPkcs1Encode(baRetVal() As Byte, baMessage() As Byte, ByVal lBitLen As Long, ByVal lHashSize As Long) As Boolean
    Const FUNC_NAME     As String = "CryptoEmsaPkcs1Encode"
    Dim baHash()        As Byte
    Dim baDerHash()     As Byte
    Dim lPos            As Long
    
    '--- from RFC 8017, Section 9.2.
    pvArrayHash lHashSize, baMessage, baHash
    If Not pvAsn1EncodePkcs1Signature(baHash, baDerHash) Then
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
    If Not pvAsn1EncodePkcs1Signature(baHash, baDerHash) Then
        GoTo QH
    End If
    For lPos = 2 To UBound(baEnc) - UBound(baDerHash) - 2
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
    Const FUNC_NAME     As String = "CryptoEmsaPssEncode"
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
        pvCryptoRandomBytes VarPtr(baSalt(0)), lSaltSize
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
    Const FUNC_NAME     As String = "CryptoEmsaPssDecode"
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
    If Not pvArrayEqual(baHash, baBuffer) Then
        GoTo QH
    End If
    '--- success
    pvCryptoEmsaPssDecode = True
QH:
End Function

Private Function pvAsn1DecodePrivateKey(cCerts As Collection, cPrivKey As Collection, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "Asn1DecodePrivateKey"
    Const IDX_KEYNAME   As Long = 1
    Const IDX_PROVNAME  As Long = 2
    Const IDX_PROVTYPE  As Long = 3
    Const IDX_KEYSPEC   As Long = 4
    Dim baPrivKey()     As Byte
    Dim baCert()        As Byte
    Dim uCertInfo       As UcsKeyInfo
    Dim hNProv          As Long
    Dim hProv           As Long
    Dim lPkiPtr         As Long
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim lSize           As Long
    Dim lHalfSize       As Long
    Dim uEccKeyInfo     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If pvCollectionCount(cPrivKey) > 1 Then
        If Not SearchCollection(cCerts, 1, RetVal:=baCert) Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_NO_CERTIFICATE
        End If
        If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
            ErrRaise vbObjectError, FUNC_NAME, ERR_UNSUPPORTED_CERTIFICATE
        End If
        uRetVal.AlgoObjId = uCertInfo.AlgoObjId
        uRetVal.BitLen = uCertInfo.BitLen
        With cPrivKey
            If pvCollectionCount(cPrivKey) = IDX_PROVNAME Then
                hResult = NCryptOpenStorageProvider(hNProv, StrPtr(.Item(IDX_PROVNAME)), 0)
                If hResult < 0 Then
                    sApiSource = "NCryptOpenStorageProvider"
                    GoTo QH
                End If
                hResult = NCryptOpenKey(hNProv, uRetVal.hNKey, StrPtr(.Item(IDX_KEYNAME)), 0, 0)
                If hResult < 0 Then
                    sApiSource = "NCryptOpenKey"
                    GoTo QH
                End If
            Else
                If .Item(IDX_PROVTYPE) = PROV_RSA_FULL Then
                    '--- try using PROV_RSA_AES to have SHA-2 available in pvTlsSignatureSign
                    Call CryptAcquireContext(hProv, StrPtr(.Item(IDX_KEYNAME)), 0, PROV_RSA_AES, 0)
                End If
                If hProv = 0 Then
                    If CryptAcquireContext(hProv, StrPtr(.Item(IDX_KEYNAME)), StrPtr(.Item(IDX_PROVNAME)), .Item(IDX_PROVTYPE), 0) = 0 Then
                        hResult = Err.LastDllError
                        sApiSource = "CryptAcquireContext"
                        GoTo QH
                    End If
                End If
                If CryptGetUserKey(hProv, .Item(IDX_KEYSPEC), uRetVal.hKey) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CryptGetUserKey"
                    GoTo QH
                End If
                uRetVal.hProv = hProv: hProv = 0
                uRetVal.dwKeySpec = .Item(IDX_KEYSPEC)
            End If
        End With
    ElseIf SearchCollection(cPrivKey, 1, RetVal:=baPrivKey) Then
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) <> 0 Then
            Call CopyMemory(uPrivKey, ByVal lPkiPtr, LenB(uPrivKey))
            If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uPrivKey.PrivateKey.pbData, uPrivKey.PrivateKey.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)"
                GoTo QH
            End If
            uRetVal.AlgoObjId = pvToStringA(uPrivKey.Algorithm.pszObjId)
            GoTo DecodeRsa
        ElseIf CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) <> 0 Then
            uRetVal.AlgoObjId = szOID_RSA_RSA
DecodeRsa:
            pvArrayAllocate uRetVal.KeyBlob, lKeySize, FUNC_NAME & ".uRetVal.KeyBlob"
            Call CopyMemory(uRetVal.KeyBlob(0), ByVal lKeyPtr, lKeySize)
            Debug.Assert UBound(uRetVal.KeyBlob) >= 16
            Call CopyMemory(uRetVal.BitLen, uRetVal.KeyBlob(12), LenB(uRetVal.BitLen))
            lSize = (uRetVal.BitLen + 7) \ 8
            lHalfSize = (uRetVal.BitLen + 15) \ 16
            '--- modulus (mod = p * q)
            pvArrayAllocate uRetVal.Modulus, lSize, FUNC_NAME & ".uRetVal.Modulus"
            Debug.Assert UBound(uRetVal.KeyBlob) - 20 >= UBound(uRetVal.Modulus)
            Call CopyMemory(uRetVal.Modulus(0), uRetVal.KeyBlob(20), UBound(uRetVal.Modulus) + 1)
            pvArrayReverse uRetVal.Modulus
            '--- prime1 (p)
            pvArrayAllocate uRetVal.Prime1, lHalfSize, FUNC_NAME & ".uRetVal.Prime1"
            Debug.Assert UBound(uRetVal.KeyBlob) >= 20 + lSize + 0 * lHalfSize + UBound(uRetVal.Prime1)
            Call CopyMemory(uRetVal.Prime1(0), uRetVal.KeyBlob(20 + lSize + 0 * lHalfSize), UBound(uRetVal.Prime1) + 1)
            pvArrayReverse uRetVal.Prime1
            '--- prime2 (q)
            pvArrayAllocate uRetVal.Prime2, lHalfSize, FUNC_NAME & ".uRetVal.Prime2"
            Debug.Assert UBound(uRetVal.KeyBlob) >= 20 + lSize + 1 * lHalfSize + UBound(uRetVal.Prime2)
            Call CopyMemory(uRetVal.Prime2(0), uRetVal.KeyBlob(20 + lSize + 1 * lHalfSize), UBound(uRetVal.Prime2) + 1)
            pvArrayReverse uRetVal.Prime2
            '--- coefficient (iqmp)
            pvArrayAllocate uRetVal.Coefficient, lHalfSize, FUNC_NAME & ".uRetVal.Coefficient"
            Debug.Assert UBound(uRetVal.KeyBlob) >= 20 + lSize + 4 * lHalfSize + UBound(uRetVal.Coefficient)
            Call CopyMemory(uRetVal.Coefficient(0), uRetVal.KeyBlob(20 + lSize + 4 * lHalfSize), UBound(uRetVal.Coefficient) + 1)
            pvArrayReverse uRetVal.Coefficient
            '--- privateExponent
            pvArrayAllocate uRetVal.PrivExp, lSize, FUNC_NAME & ".uRetVal.PrivExp"
            Debug.Assert UBound(uRetVal.KeyBlob) >= 20 + lSize + 5 * lHalfSize + UBound(uRetVal.PrivExp)
            Call CopyMemory(uRetVal.PrivExp(0), uRetVal.KeyBlob(20 + lSize + 5 * lHalfSize), UBound(uRetVal.PrivExp) + 1)
            pvArrayReverse uRetVal.PrivExp
        ElseIf CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, 0) <> 0 Then
            Call CopyMemory(uEccKeyInfo, ByVal lKeyPtr, LenB(uEccKeyInfo))
            uRetVal.AlgoObjId = pvToStringA(uEccKeyInfo.szCurveOid)
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
            '--- unsupported private key
            GoTo QH
        End If
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
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If hNProv <> 0 Then
        Call NCryptFreeObject(hNProv)
    End If
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvAsn1DecodeCertificate(baCert() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "Asn1DecodeCertificate"
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
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), LenB(lPtr))  '--- dereference pCertContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)                                            '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
    Call CopyMemory(uPublicKeyInfo, ByVal lPtr, LenB(uPublicKeyInfo))
    uRetVal.AlgoObjId = pvToStringA(uPublicKeyInfo.Algorithm.pszObjId)
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
        Call CopyMemory(uRetVal.BitLen, baBuffer(12), LenB(uRetVal.BitLen))                     '--- 12 = sizeof(PUBLICKEYSTRUC) + offset(RSAPUBKEY, bitlen)
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
    pvAsn1DecodeCertificate = True
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
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvAsn1EncodePkcs1Signature(baHash() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "Asn1EncodePkcs1Signature"
    Dim baPrefix()      As Byte
    
    Select Case UBound(baHash) + 1
    Case LNG_SHA1_HASHSZ
        pvArrayByte baPrefix, &H30, &H21, &H30, &H9, &H6, &H5, &H2B, &HE, &H3, &H2, &H1A, &H5, &H0, &H4, &H14
    Case LNG_SHA224_HASHSZ
        pvArrayByte baPrefix, &H30, &H2D, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H4, &H5, &H0, &H4, &H1C
    Case LNG_SHA256_HASHSZ
        pvArrayByte baPrefix, &H30, &H31, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H1, &H5, &H0, &H4, &H20
    Case LNG_SHA384_HASHSZ
        pvArrayByte baPrefix, &H30, &H41, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H2, &H5, &H0, &H4, &H30
    Case LNG_SHA512_HASHSZ
        pvArrayByte baPrefix, &H30, &H51, &H30, &HD, &H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H3, &H5, &H0, &H4, &H40
    End Select
    pvArrayAllocate baRetVal, UBound(baPrefix) + UBound(baHash) + 2, FUNC_NAME & ".baRetVal"
    Call CopyMemory(baRetVal(0), baPrefix(0), UBound(baPrefix) + 1)
    Call CopyMemory(baRetVal(UBound(baPrefix) + 1), baHash(0), UBound(baHash) + 1)
    '--- success
    pvAsn1EncodePkcs1Signature = True
QH:
End Function

Private Function pvAsn1DecodeEcdsaSignature(baRetVal() As Byte, baDerSig() As Byte, ByVal lCurveSize As Long) As Boolean
    Dim uOutput         As UcsBuffer
    Dim uInput          As UcsBuffer
    Dim lType           As Long
    Dim lSize           As Long
    Dim baTemp()        As Byte
    
    pvArraySwap uInput.Data, 0, baDerSig, 0
    '--- ECDSA-Sig-Value ::= SEQUENCE { r INTEGER, s INTEGER }
    pvBufferReadLong uInput, lType
    If lType <> LNG_ANS1_TYPE_SEQUENCE Then
        GoTo QH
    End If
    If uInput.Pos <= UBound(uInput.Data) Then
        lSize = uInput.Data(uInput.Pos)
    End If
    '--- check for long form encoding of length of sequence
    If lSize > &H83 Then
        GoTo QH
    ElseIf lSize > &H80 Then
        lSize = lSize - &H80 + 1
        uInput.Data(uInput.Pos) = 0
    Else
        lSize = 1
    End If
    pvBufferReadBlockStart uInput, Size:=lSize
        pvBufferReadLong uInput, lType
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        pvBufferReadLong uInput, lSize
        pvBufferReadArray uInput, baTemp, lSize
        If lSize <= lCurveSize Then
            uOutput.Size = lCurveSize - lSize
            pvBufferWriteArray uOutput, baTemp
        Else
            uOutput.Size = 0
            pvBufferWriteBlob uOutput, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
        pvBufferReadLong uInput, lType
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        pvBufferReadLong uInput, lSize
        pvBufferReadArray uInput, baTemp, lSize
        If lSize <= lCurveSize Then
            uOutput.Size = lCurveSize + lCurveSize - lSize
            pvBufferWriteArray uOutput, baTemp
        Else
            uOutput.Size = lCurveSize
            pvBufferWriteBlob uOutput, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
        If uInput.Stack(1) <> uInput.Pos Then
            GoTo QH
        End If
    pvBufferReadBlockEnd uInput
    pvBufferWriteEOF uOutput
    '--- success
    pvAsn1DecodeEcdsaSignature = True
QH:
    pvArraySwap baDerSig, 0, uInput.Data, 0
    baRetVal = uOutput.Data
End Function

Private Function pvAsn1EncodeEcdsaSignature(baRetVal() As Byte, baPlainSig() As Byte, ByVal lPartSize As Long) As Boolean
    Dim uOutput         As UcsBuffer
    Dim lStart          As Long
    
    pvBufferWriteLong uOutput, LNG_ANS1_TYPE_SEQUENCE
    pvBufferWriteBlockStart uOutput
        pvBufferWriteLong uOutput, LNG_ANS1_TYPE_INTEGER
        pvBufferWriteBlockStart uOutput
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lStart) And &H80) <> 0 Then
                pvBufferWriteLong uOutput, 0
            End If
            pvBufferWriteBlob uOutput, VarPtr(baPlainSig(lStart)), lPartSize - lStart
        pvBufferWriteBlockEnd uOutput
        pvBufferWriteLong uOutput, LNG_ANS1_TYPE_INTEGER
        pvBufferWriteBlockStart uOutput
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lPartSize + lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lPartSize + lStart) And &H80) <> 0 Then
                pvBufferWriteLong uOutput, 0
            End If
            pvBufferWriteBlob uOutput, VarPtr(baPlainSig(lPartSize + lStart)), lPartSize - lStart
        pvBufferWriteBlockEnd uOutput
    pvBufferWriteBlockEnd uOutput
    pvBufferWriteEOF uOutput
    '--- success
    pvAsn1EncodeEcdsaSignature = True
    baRetVal = uOutput.Data
End Function

Private Function pvCryptoInit() As Boolean
    Const FUNC_NAME     As String = "CryptoInit"
    Dim lOffset         As Long
    Dim lIdx            As Long
    Dim baThunk()       As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    With m_uData
        If .hRandomProv = 0 Then
            If CryptAcquireContext(.hRandomProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptAcquireContext"
                GoTo QH
            End If
        End If
        If m_uData.Thunk = 0 Then
            '--- prepare thunk/context in executable memory
            pvGetThunkData baThunk
            .Thunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            If .Thunk = 0 Then
                hResult = Err.LastDllError
                sApiSource = "VirtualAlloc"
                GoTo QH
            End If
            Call CopyMemory(ByVal .Thunk, baThunk(0), UBound(baThunk) + 1)
            pvGetGlobData .Glob
            '--- init pfns from thunk addr + offsets stored at beginning of it
            For lIdx = LBound(.Pfn) To UBound(.Pfn)
                Call CopyMemory(lOffset, ByVal UnsignedAdd(.Thunk, LenB(lOffset) * lIdx), LenB(lOffset))
                .Pfn(lIdx) = UnsignedAdd(.Thunk, lOffset)
            Next
            '--- init pfns trampolines
            Call pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
            Call pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
            Call pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
            Call pvPatchTrampoline(AddressOf pvCallSecpSign)
            Call pvPatchTrampoline(AddressOf pvCallSecpVerify)
            Call pvPatchTrampoline(AddressOf pvCallCurve25519Multiply)
            Call pvPatchTrampoline(AddressOf pvCallCurve25519MulBase)
            Call pvPatchTrampoline(AddressOf pvCallSha2Init)
            Call pvPatchTrampoline(AddressOf pvCallSha2Update)
            Call pvPatchTrampoline(AddressOf pvCallSha2Final)
            Call pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Encrypt)
            Call pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Decrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesGcmEncrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesGcmDecrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesCbcEncrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesCbcDecrypt)
            Call pvPatchTrampoline(AddressOf pvCallRsaModExp)
            Call pvPatchTrampoline(AddressOf pvCallRsaCrtModExp)
            '--- init thunk's first 4 bytes -> global data in C/C++
            Call CopyMemory(ByVal .Thunk, VarPtr(.Glob(0)), 4)
            Call CopyMemory(.Glob(0), GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemAlloc"), 4)
            Call CopyMemory(.Glob(4), GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemRealloc"), 4)
            Call CopyMemory(.Glob(8), GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemFree"), 4)
        End If
    End With
    '--- success
    pvCryptoInit = True
QH:
    If LenB(sApiSource) <> 0 Then
        ErrRaise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

#If False Then
Public Sub pvCryptoTerminate()
    With m_uData
        If .hRandomProv <> 0 Then
            Call CryptReleaseContext(.hRandomProv, 0)
            .hRandomProv = 0
        End If
    End With
End Sub
#End If

Private Function pvCryptoEcdhCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhCurve25519MakeKey"
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    pvArrayAllocate baPrivate, LNG_X25519_KEYSZ, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, LNG_X25519_KEYSZ, FUNC_NAME & ".baPublic"
    For lIdx = 1 To MAX_RETRIES
        pvCryptoRandomBytes VarPtr(baPrivate(0)), LNG_X25519_KEYSZ
        If pvArrayAccumulateOr(baPrivate) <> 0 Then
            Exit For
        End If
    Next
    '--- fix privkey randomness
    baPrivate(0) = baPrivate(0) And &HF8
    baPrivate(LNG_X25519_KEYSZ - 1) = (baPrivate(LNG_X25519_KEYSZ - 1) And &H7F) Or &H40
    Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519MulBase)
    pvCallCurve25519MulBase m_uData.Pfn(ucsPfnCurve25519ScalarMultBase), baPublic(0), baPrivate(0)
    '--- success
    pvCryptoEcdhCurve25519MakeKey = True
End Function

Private Function pvCryptoEcdhCurve25519SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhCurve25519SharedSecret"
    
    If UBound(baPrivate) = LNG_X25519_KEYSZ - 1 And UBound(baPublic) = LNG_X25519_KEYSZ - 1 Then
        pvArrayAllocate baRetVal, LNG_X25519_KEYSZ, FUNC_NAME & ".baRetVal"
        Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519Multiply)
        pvCallCurve25519Multiply m_uData.Pfn(ucsPfnCurve25519ScalarMultiply), baRetVal(0), baPrivate(0), baPublic(0)
        '--- success
        pvCryptoEcdhCurve25519SharedSecret = True
    End If
End Function

Private Function pvCryptoEcdhSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp256r1MakeKey"
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    pvArrayAllocate baPrivate, LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, 2 * LNG_SECP256R1_KEYSZ + 1, FUNC_NAME & ".baPublic"
    For lIdx = 1 To MAX_RETRIES
        pvCryptoRandomBytes VarPtr(baPrivate(0)), LNG_SECP256R1_KEYSZ
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdhSecp256r1MakeKey = True
QH:
End Function

Private Function pvCryptoEcdhSecp256r1SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp256r1SharedSecret"
    
    If UBound(baPrivate) = LNG_SECP256R1_KEYSZ - 1 And UBound(baPublic) >= LNG_SECP256R1_KEYSZ Then
        pvArrayAllocate baRetVal, LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baRetVal"
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
        If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
            GoTo QH
        End If
        '--- success
        pvCryptoEcdhSecp256r1SharedSecret = True
    End If
QH:
End Function

Private Function pvCryptoEcdhSecp256r1UncompressKey(baRetVal() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp256r1UncompressKey"

    pvArrayAllocate baRetVal, 1 + 2 * LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdhSecp256r1UncompressKey = True
QH:
End Function

Private Function pvCryptoEcdsaSecp256r1Sign(baRetVal() As Byte, baPrivKey() As Byte, baHash() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdsaSecp256r1Sign"
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim lIdx            As Long
    
    pvArrayAllocate baRandom, LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baRandom"
    pvArrayAllocate baRetVal, 2 * LNG_SECP256R1_KEYSZ, FUNC_NAME & ".baRetVal"
    For lIdx = 1 To MAX_RETRIES
        pvCryptoRandomBytes VarPtr(baRandom(0)), LNG_SECP256R1_KEYSZ
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp256r1Sign), baPrivKey(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdsaSecp256r1Sign = True
QH:
End Function

Private Function pvCryptoEcdsaSecp256r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    pvCryptoEcdsaSecp256r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Private Function pvCryptoEcdhSecp384r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp384r1MakeKey"
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    pvArrayAllocate baPrivate, LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, 2 * LNG_SECP384R1_KEYSZ + 1, FUNC_NAME & ".baPublic"
    For lIdx = 1 To MAX_RETRIES
        pvCryptoRandomBytes VarPtr(baPrivate(0)), LNG_SECP384R1_KEYSZ
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp384r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdhSecp384r1MakeKey = True
QH:
End Function

Private Function pvCryptoEcdhSecp384r1SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp384r1SharedSecret"
    
    If UBound(baPrivate) = LNG_SECP384R1_KEYSZ - 1 And UBound(baPublic) >= LNG_SECP384R1_KEYSZ Then
        pvArrayAllocate baRetVal, LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baRetVal"
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
        If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp384r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
            GoTo QH
        End If
        '--- success
        pvCryptoEcdhSecp384r1SharedSecret = True
    End If
QH:
End Function

Private Function pvCryptoEcdhSecp384r1UncompressKey(baRetVal() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdhSecp384r1UncompressKey"

    pvArrayAllocate baRetVal, 1 + 2 * LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp384r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdhSecp384r1UncompressKey = True
QH:
End Function

Private Function pvCryptoEcdsaSecp384r1Sign(baRetVal() As Byte, baPrivKey() As Byte, baHash() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEcdsaSecp384r1Sign"
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim lIdx            As Long
    
    pvArrayAllocate baRandom, LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baRandom"
    pvArrayAllocate baRetVal, 2 * LNG_SECP384R1_KEYSZ, FUNC_NAME & ".baRetVal"
    For lIdx = 1 To MAX_RETRIES
        pvCryptoRandomBytes VarPtr(baRandom(0)), LNG_SECP384R1_KEYSZ
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp384r1Sign), baPrivKey(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    pvCryptoEcdsaSecp384r1Sign = True
QH:
End Function

Private Function pvCryptoEcdsaSecp384r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    pvCryptoEcdsaSecp384r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp384r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Private Function pvCryptoHashSha1(baRetVal() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha256"
    Dim lPtr            As Long
    Dim hHash           As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA1_HASHSZ, FUNC_NAME & ".baRetVal"
    If CryptCreateHash(m_uData.hRandomProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        GoTo QH
    End If
    If CryptHashData(hHash, ByVal lPtr, Size, 0) = 0 Then
        GoTo QH
    End If
    If CryptGetHashParam(hHash, HP_HASHVAL, baRetVal(0), LNG_SHA1_HASHSZ, 0) = 0 Then
        GoTo QH
    End If
    '--- success
    pvCryptoHashSha1 = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
End Function

Private Function pvCryptoHashSha224(baRetVal() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha224"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Static H224(0 To 7) As Long
    
    If H224(0) = 0 Then
        H224(0) = &HC1059ED8
        H224(1) = &H367CD507
        H224(2) = &H3070DD17
        H224(3) = &HF70E5939
        H224(4) = &HFFC00B31
        H224(5) = &H68581511
        H224(6) = &H64F98FA7
        H224(7) = &HBEFA4FA4
    End If
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA256_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        lCtxPtr = VarPtr(.HashCtx(0))
        pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
        Call CopyMemory(ByVal lCtxPtr, H224(0), 32)
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
    End With
    pvArrayReallocate baRetVal, LNG_SHA224_HASHSZ, FUNC_NAME & ".baRetVal"
    '--- success
    pvCryptoHashSha224 = True
End Function

Private Function pvCryptoHashSha256(baRetVal() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha256"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA256_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        lCtxPtr = VarPtr(.HashCtx(0))
        pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
    End With
    '--- success
    pvCryptoHashSha256 = True
End Function

Private Function pvCryptoHashSha384(baRetVal() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha384"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA384_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
    End With
    '--- success
    pvCryptoHashSha384 = True
End Function

Private Function pvCryptoHashSha512(baRetVal() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha512"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA512_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        pvCallSha2Init .Pfn(ucsPfnSha512Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha512Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha512Final), lCtxPtr, baRetVal(0)
    End With
    '--- success
    pvCryptoHashSha512 = True
End Function

Private Function pvCryptoHmacSha1(baRetVal() As Byte, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmacSha1"
    Dim lPtr            As Long
    Dim lIdx            As Long
    Dim hHash           As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA1_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    With m_uData
        '-- inner hash
        Call FillMemory(.HashPad(0), LNG_SHA1_BLOCKSZ, LNG_HMAC_INNER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
        Next
        If CryptCreateHash(m_uData.hRandomProv, CALG_SHA1, 0, 0, hHash) = 0 Then
            GoTo QH
        End If
        If CryptHashData(hHash, .HashPad(0), LNG_SHA1_BLOCKSZ, 0) = 0 Then
            GoTo QH
        End If
        If CryptHashData(hHash, ByVal lPtr, Size, 0) = 0 Then
            GoTo QH
        End If
        If CryptGetHashParam(hHash, HP_HASHVAL, .HashFinal(0), LNG_SHA1_HASHSZ, 0) = 0 Then
            GoTo QH
        End If
        Call CryptDestroyHash(hHash)
        hHash = 0
        '-- outer hash
        Call FillMemory(.HashPad(0), LNG_SHA1_BLOCKSZ, LNG_HMAC_OUTER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
        Next
        If CryptCreateHash(m_uData.hRandomProv, CALG_SHA1, 0, 0, hHash) = 0 Then
            GoTo QH
        End If
        If CryptHashData(hHash, .HashPad(0), LNG_SHA1_BLOCKSZ, 0) = 0 Then
            GoTo QH
        End If
        If CryptHashData(hHash, .HashFinal(0), LNG_SHA1_HASHSZ, 0) = 0 Then
            GoTo QH
        End If
        pvArrayAllocate baRetVal, LNG_SHA1_HASHSZ, FUNC_NAME & ".baRetVal"
        If CryptGetHashParam(hHash, HP_HASHVAL, baRetVal(0), LNG_SHA1_HASHSZ, 0) = 0 Then
            GoTo QH
        End If
    End With
    '--- success
    pvCryptoHmacSha1 = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
End Function

Private Function pvCryptoHmacSha256(baRetVal() As Byte, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmacSha256"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA256_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        pvArrayAllocate baRetVal, LNG_SHA256_HASHSZ, FUNC_NAME & ".baRetVal"
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        '-- inner hash
        Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
        Next
        pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, .HashFinal(0)
        '-- outer hash
        Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
        Next
        pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
        pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA256_HASHSZ
        pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
    End With
    '--- success
    pvCryptoHmacSha256 = True
End Function

Private Function pvCryptoHmacSha384(baRetVal() As Byte, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmacSha384"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA384_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - Pos
    Else
        Debug.Assert pvArraySize(baInput) >= Pos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(Pos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        pvArrayAllocate baRetVal, LNG_SHA384_HASHSZ, FUNC_NAME & ".baRetVal"
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
        '-- inner hash
        Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
        Next
        pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
        pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
        pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, .HashFinal(0)
        '-- outer hash
        Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
        Next
        pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
        pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
        pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA384_HASHSZ
        pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
    End With
    '--- success
    pvCryptoHmacSha384 = True
End Function

Private Function pvCryptoBulkChacha20Poly1305Encrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAadPtr         As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_CHACHA20POLY1305_TAGSZ
    If lAadSize > 0 Then
        lAadPtr = VarPtr(baAad(lAadPos))
    End If
    Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Encrypt)
    Call pvCallChacha20Poly1305Encrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Encrypt), _
            baKey(0), baNonce(0), _
            lAadPtr, lAadSize, _
            baBuffer(lPos), lSize, _
            baBuffer(lPos), baBuffer(lPos + lSize))
    '--- success
    pvCryptoBulkChacha20Poly1305Encrypt = True
End Function

Private Function pvCryptoBulkChacha20Poly1305Decrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Decrypt)
    If pvCallChacha20Poly1305Decrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Decrypt), _
            baKey(0), baNonce(0), _
            baAad(lAadPos), lAadSize, _
            baBuffer(lPos), lSize - LNG_CHACHA20POLY1305_TAGSZ, _
            baBuffer(lPos + lSize - LNG_CHACHA20POLY1305_TAGSZ), baBuffer(lPos)) = 0 Then
        '--- success
        pvCryptoBulkChacha20Poly1305Decrypt = True
    End If
End Function

Private Function pvCryptoBulkAesGcmEncrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAadPtr         As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_AESGCM_TAGSZ
    If lAadSize > 0 Then
        lAadPtr = VarPtr(baAad(lAadPos))
    End If
    Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmEncrypt)
    Call pvCallAesGcmEncrypt(m_uData.Pfn(ucsPfnAesGcmEncrypt), _
            baBuffer(lPos), baBuffer(lPos + lSize), _
            baBuffer(lPos), lSize, _
            lAadPtr, lAadSize, _
            baNonce(0), baKey(0), UBound(baKey) + 1)
    '--- success
    pvCryptoBulkAesGcmEncrypt = True
End Function

Private Function pvCryptoBulkAesGcmDecrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAadSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmDecrypt)
    If pvCallAesGcmDecrypt(m_uData.Pfn(ucsPfnAesGcmDecrypt), _
            baBuffer(lPos), _
            baBuffer(lPos), lSize - LNG_AESGCM_TAGSZ, _
            baBuffer(lPos + lSize - LNG_AESGCM_TAGSZ), _
            baAad(lAadPos), lAadSize, _
            baNonce(0), baKey(0), UBound(baKey) + 1) = 0 Then
        '--- success
        pvCryptoBulkAesGcmDecrypt = True
    End If
End Function

Private Function pvCryptoBulkAesCbcEncrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESCBC_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    Debug.Assert lSize Mod pvArraySize(baNonce) = 0
    Debug.Assert pvPatchTrampoline(AddressOf pvCallAesCbcEncrypt)
    Call pvCallAesCbcEncrypt(m_uData.Pfn(ucsPfnAesCbcEncrypt), _
            baBuffer(lPos), baBuffer(lPos), lSize, _
            baNonce(0), baKey(0), UBound(baKey) + 1)
    '--- success
    pvCryptoBulkAesCbcEncrypt = True
End Function

Private Function pvCryptoBulkAesCbcDecrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESCBC_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    If lSize Mod pvArraySize(baNonce) = 0 Then
        Debug.Assert pvPatchTrampoline(AddressOf pvCallAesCbcDecrypt)
        Call pvCallAesCbcDecrypt(m_uData.Pfn(ucsPfnAesCbcDecrypt), _
                baBuffer(lPos), baBuffer(lPos), lSize, _
                baNonce(0), baKey(0), UBound(baKey) + 1)
        '--- success
        pvCryptoBulkAesCbcDecrypt = True
    End If
End Function

Private Sub pvCryptoRandomBytes(ByVal lPtr As Long, ByVal lSize As Long)
    Call CryptGenRandom(m_uData.hRandomProv, lSize, lPtr)
End Sub

Private Function pvCryptoRsaModExp(baBase() As Byte, baExp() As Byte, baModulus() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaModExp"
    
    pvArrayAllocate baRetVal, UBound(baBase) + 1, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallRsaModExp)
    Call pvCallRsaModExp(m_uData.Pfn(ucsPfnRsaModExp), UBound(baBase) + 1, baBase(0), baExp(0), baModulus(0), baRetVal(0))
    '--- success
    pvCryptoRsaModExp = True
End Function

Private Function pvCryptoRsaCrtModExp(baBase() As Byte, baExp() As Byte, baModulus() As Byte, baPrime1() As Byte, baPrime2() As Byte, baCoefficient() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaCrtModExp"
    
    pvArrayAllocate baRetVal, UBound(baBase) + 1, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallRsaCrtModExp)
    Call pvCallRsaCrtModExp(m_uData.Pfn(ucsPfnRsaCrtModExp), UBound(baBase) + 1, baBase(0), baExp(0), baModulus(0), baPrime1(0), baPrime2(0), baCoefficient(0), baRetVal(0))
    '--- success
    pvCryptoRsaCrtModExp = True
End Function

'= private ===============================================================

Private Function pvThunkAllocate(sText As String, Optional ByVal ThunkPtr As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    If ThunkPtr <> 0 Then
        pvThunkAllocate = ThunkPtr
    Else
        pvThunkAllocate = VirtualAlloc(0, (Len(sText) \ 4) * 3, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If pvThunkAllocate = 0 Then
            Exit Function
        End If
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
        lPtr = UnsignedAdd(lPtr, 3)
    Next
End Function

Private Function pvPatchTrampoline(ByVal Pfn As Long) As Boolean
    Dim bInIDE          As Boolean
 
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(Pfn, ByVal Pfn + &H16, LenB(Pfn))
    Else
        Call VirtualProtect(Pfn, 8, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0:  58                      pop    eax
    ' 1:  59                      pop    ecx
    ' 2:  50                      push   eax
    ' 3:  ff e1                   jmp    ecx
    ' 5:  90                      nop
    ' 6:  90                      nop
    ' 7:  90                      nop
    Call CopyMemory(ByVal Pfn, -802975883527609.7192@, 8)
    '--- success
    pvPatchTrampoline = True
End Function

Private Function pvSetTrue(bValue As Boolean) As Boolean
    #If TWINBASIC = 0 Then
        bValue = True
    #End If
    pvSetTrue = True
End Function

#If Not ImplUseShared Then
Private Function RedimStats(sFuncName As String, ByVal lSize As Long) As Boolean
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

Public Function FromBase64Array(sText As String) As Byte()
    Const FUNC_NAME     As String = "FromBase64Array"
    Dim baRetVal()      As Byte
    Dim lSize           As Long

    lSize = (Len(sText) \ 4) * 3
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    pvThunkAllocate sText, VarPtr(baRetVal(0))
    If Right$(sText, 2) = "==" Then
        pvArrayReallocate baRetVal, lSize - 2, FUNC_NAME & ".baRetVal"
    ElseIf Right$(sText, 1) = "=" Then
        pvArrayReallocate baRetVal, lSize - 1, FUNC_NAME & ".baRetVal"
    End If
    FromBase64Array = baRetVal
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function pvCallCollectionItem(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM As Long = 7
    
    pvPatchMethodTrampoline AddressOf mdTlsThunks.pvCallCollectionItem, IDX_COLLECTION_ITEM
    pvCallCollectionItem = pvCallCollectionItem(oCol, Index, RetVal)
End Function

Private Function pvCallCollectionRemove(ByVal oCol As Collection, Index As Variant) As Long
    Const IDX_COLLECTION_REMOVE As Long = 10
    
    pvPatchMethodTrampoline AddressOf mdTlsThunks.pvCallCollectionRemove, IDX_COLLECTION_REMOVE
    pvCallCollectionRemove = pvCallCollectionRemove(oCol, Index)
End Function

Private Function pvPatchMethodTrampoline(ByVal Pfn As Long, ByVal lMethodIdx As Long) As Boolean
    Dim bInIDE          As Boolean

    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(Pfn, ByVal Pfn + &H16, LenB(Pfn))
    Else
        Call VirtualProtect(Pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal Pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (Pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * PTR_SIZE, LenB(lMethodIdx))
    '--- success
    pvPatchMethodTrampoline = True
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

Private Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal lMin As Long = -2147483647, _
            Optional ByVal lMax As Long = 2147483647) As Long
    Select Case lValue
    Case lMin To lMax
        Clamp = lValue
    Case Is < lMin
        Clamp = lMin
    Case Is > lMax
        Clamp = lMax
    End Select
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
#End If ' Not ImplUseShared

#If (ImplCaptureTraffic And 1) <> 0 Then
Public Function TlsDesignDumpArray(baData() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As String
    If Size < 0 Then
        Size = UBound(baData) + 1 - Pos
    End If
    If Size > 0 Then
        TlsDesignDumpArray = TlsDesignDumpMemory(VarPtr(baData(Pos)), Size)
    End If
End Function

Public Function TlsDesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim aResult()       As String
    
    ReDim aResult(0 To (lSize + 15) \ 16) As String
    Debug.Assert RedimStats("TlsDesignDumpMemory.aResult", UBound(aResult) + 1)
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(lPtr, 1) = 0 Then
                Call CopyMemory(lValue, ByVal lPtr, 1)
                sHex = sHex & Right$("0" & Hex$(lValue), 2) & " "
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
        If ((lIdx + 1) Mod 4) = 0 Then
            sHex = sHex & " "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            aResult(lIdx \ 16) = Right$("000" & Hex$(lIdx - 15), 4) & " - " & sHex & sChar
            sHex = vbNullString
            sChar = vbNullString
        End If
        lPtr = (lPtr Xor &H80000000) + 1 Xor &H80000000
    Next
    TlsDesignDumpMemory = Join(aResult, vbCrLf)
End Function

Public Function TlsConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        TlsConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(TlsConcatCollection) Then
                Mid$(TlsConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function
#End If ' ImplCaptureTraffic

#If ImplTestCrypto Then

Public Sub TestCryptoAesGcm(oJson As Object)
    Const FUNC_NAME     As String = "TlsTestCryptoAesGcm"
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sComment        As String
    Dim sResult         As String
    Dim lPassed         As Long
    Dim lFailed         As Long
    Dim lSkipped        As Long
    '--- local
    Dim baNonce()       As Byte
    Dim baKey()         As Byte
    Dim baAad()         As Byte
    Dim baMsg()         As Byte
    Dim baCt()          As Byte
    Dim baTag()         As Byte
    Dim baBuffer()      As Byte
    
    pvCryptoInit
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baNonce = FromHex(JsonValue(oTest, "iv"))
            If pvArraySize(baNonce) = LNG_AESGCM_IVSZ Then
                baKey = FromHex(JsonValue(oTest, "key"))
                If pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ Then
                    baAad = FromHex(JsonValue(oTest, "aad"))
                    baMsg = FromHex(JsonValue(oTest, "msg"))
                    baCt = FromHex(JsonValue(oTest, "ct"))
                    baTag = FromHex(JsonValue(oTest, "tag"))
                    baBuffer = baMsg
                    ReDim Preserve baBuffer(0 To UBound(baMsg) + UBound(baTag) + 1) As Byte
                    sResult = "valid"
                    If Not pvCryptoBulkAesGcmEncrypt(baNonce, baKey, baAad, 0, UBound(baAad) + 1, baBuffer, 0, UBound(baMsg) + 1) Then
                        sResult = "invalid"
                    Else
                        pvArrayWriteBlob baCt, UBound(baCt) + 1, VarPtr(baTag(0)), UBound(baTag) + 1
                        If Not pvArrayEqual(baBuffer, baCt) Then
                            sResult = "invalid"
                        End If
                    End If
                    If JsonValue(oTest, "result") = sResult Then
                        lPassed = lPassed + 1
                    Else
                        lFailed = lFailed + 1
                        sComment = JsonValue(oTest, "comment")
                        DebugLog MODULE_NAME, FUNC_NAME, "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
                    End If
                Else
                    lSkipped = lSkipped + 1
                End If
            Else
                lSkipped = lSkipped + 1
            End If
        Next
    Next
    DebugLog MODULE_NAME, FUNC_NAME, "[+] Passed=" & lPassed & ", Failed=" & lFailed & ", Skipped=" & lSkipped
End Sub

Public Sub TestCryptoAesCbc(oJson As Object)
    Const FUNC_NAME     As String = "TlsTestCryptoAesCbc"
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sComment        As String
    Dim sResult         As String
    Dim lPassed         As Long
    Dim lFailed         As Long
    Dim lSkipped        As Long
    '--- local
    Dim baNonce()       As Byte
    Dim baKey()         As Byte
    Dim baMsg()         As Byte
    Dim baCt()          As Byte
    Dim baBuffer()      As Byte
    Dim lPadding        As Long
    Dim baHmac(0 To 15) As Byte
    
    pvCryptoInit
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baNonce = FromHex(JsonValue(oTest, "iv"))
            If pvArraySize(baNonce) = LNG_AESCBC_IVSZ Then
                baKey = FromHex(JsonValue(oTest, "key"))
                If pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ Then
                    baMsg = FromHex(JsonValue(oTest, "msg"))
                    baCt = FromHex(JsonValue(oTest, "ct"))
                    baBuffer = baMsg
                    '--- PKCS5 padding (note: different than padding for TLS in pvBufferWriteRecordEnd)
                    lPadding = LNG_AESCBC_IVSZ - pvArraySize(baMsg) Mod LNG_AESCBC_IVSZ
                    Debug.Assert lPadding <= pvArraySize(baHmac)
                    Call FillMemory(baHmac(0), lPadding, lPadding)
                    pvArrayWriteBlob baBuffer, UBound(baBuffer) + 1, VarPtr(baHmac(0)), lPadding
                    sResult = "valid"
                    If Not pvCryptoBulkAesCbcEncrypt(baNonce, baKey, baBuffer, 0, UBound(baBuffer) + 1) Then
                        sResult = "invalid"
                    ElseIf Not pvArrayEqual(baBuffer, baCt) Then
                        sResult = "invalid"
                    End If
                    If JsonValue(oTest, "result") = sResult Then
                        lPassed = lPassed + 1
                    Else
                        lFailed = lFailed + 1
                        sComment = JsonValue(oTest, "comment")
                        DebugLog MODULE_NAME, FUNC_NAME, "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
                    End If
                Else
                    lSkipped = lSkipped + 1
                End If
            Else
                lSkipped = lSkipped + 1
            End If
        Next
    Next
    DebugLog MODULE_NAME, FUNC_NAME, "[+] Passed=" & lPassed & ", Failed=" & lFailed & ", Skipped=" & lSkipped
End Sub

Public Sub TestCryptoChacha20(oJson As Object)
    Const FUNC_NAME     As String = "TlsTestCryptoChacha20"
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sComment        As String
    Dim sResult         As String
    Dim lPassed         As Long
    Dim lFailed         As Long
    Dim lSkipped        As Long
    '--- local
    Dim baNonce()       As Byte
    Dim baKey()         As Byte
    Dim baAad()         As Byte
    Dim baMsg()         As Byte
    Dim baCt()          As Byte
    Dim baTag()         As Byte
    Dim baBuffer()      As Byte
    
    pvCryptoInit
    For Each oGroup In JsonValue(oJson, "testGroups")
        For Each oTest In JsonValue(oGroup, "tests")
            baNonce = FromHex(JsonValue(oTest, "iv"))
            If pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ Then
                baKey = FromHex(JsonValue(oTest, "key"))
                If pvArraySize(baKey) = LNG_CHACHA20_KEYSZ Then
                    baAad = FromHex(JsonValue(oTest, "aad"))
                    baMsg = FromHex(JsonValue(oTest, "msg"))
                    baCt = FromHex(JsonValue(oTest, "ct"))
                    baTag = FromHex(JsonValue(oTest, "tag"))
                    baBuffer = baMsg
                    ReDim Preserve baBuffer(0 To UBound(baMsg) + UBound(baTag) + 1) As Byte
                    sResult = "valid"
                    If Not pvCryptoBulkChacha20Poly1305Encrypt(baNonce, baKey, baAad, 0, UBound(baAad) + 1, baBuffer, 0, UBound(baMsg) + 1) Then
                        sResult = "invalid"
                    Else
                        pvArrayWriteBlob baCt, UBound(baCt) + 1, VarPtr(baTag(0)), UBound(baTag) + 1
                        If Not pvArrayEqual(baBuffer, baCt) Then
                            sResult = "invalid"
                        End If
                    End If
                    If JsonValue(oTest, "result") = sResult Then
                        lPassed = lPassed + 1
                    Else
                        lFailed = lFailed + 1
                        sComment = JsonValue(oTest, "comment")
                        DebugLog MODULE_NAME, FUNC_NAME, "[-] " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString)
                    End If
                Else
                    lSkipped = lSkipped + 1
                End If
            Else
                lSkipped = lSkipped + 1
            End If
        Next
    Next
    DebugLog MODULE_NAME, FUNC_NAME, "[+] Passed=" & lPassed & ", Failed=" & lFailed & ", Skipped=" & lSkipped
End Sub

Public Sub TestCryptoEcdh(oJson As Object)
    Const FUNC_NAME     As String = "TestCryptoEcdh"
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sComment        As String
    Dim sResult         As String
    Dim lPassed         As Long
    Dim lFailed         As Long
    Dim lSkipped        As Long
    '--- local
    Dim baPublic()      As Byte
    Dim baPrivate()     As Byte
    Dim baShared()      As Byte
    Dim baBuffer()      As Byte
    Dim lPtr            As Long
    Dim lSize           As Long
    Dim uPublicKeyInfo  As CERT_PUBLIC_KEY_INFO
    Dim sFlags          As String
    Dim sCurve          As String

    pvCryptoInit
    For Each oGroup In JsonValue(oJson, "testGroups")
        sCurve = JsonValue(oGroup, "curve")
        If sCurve = "secp256r1" Or sCurve = "secp384r1" Then
            For Each oTest In JsonValue(oGroup, "tests")
                baPublic = FromHex(JsonValue(oTest, "public"))
                baPrivate = FromHex(JsonValue(oTest, "private"))
                baShared = FromHex(JsonValue(oTest, "shared"))
                sFlags = Join(JsonValue(oTest, "flags/*"), ", ")
                If InStr(sFlags, "InvalidAsn") = 0 Then
                    If UBound(baPublic) >= 0 Then
                        If CryptDecodeObjectEx(X509_ASN_ENCODING, X509_PUBLIC_KEY_INFO, baPublic(0), UBound(baPublic) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPtr, lSize) <> 0 Then
                            Call CopyMemory(uPublicKeyInfo, ByVal lPtr, LenB(uPublicKeyInfo))
                            Debug.Assert pvToStringA(uPublicKeyInfo.Algorithm.pszObjId) = szOID_ECC_PUBLIC_KEY
                            If uPublicKeyInfo.PublicKey.cbData = 0 Then
                                baPublic = vbNullString
                            Else
                                ReDim baPublic(0 To uPublicKeyInfo.PublicKey.cbData - 1) As Byte
                                Call CopyMemory(baPublic(0), ByVal uPublicKeyInfo.PublicKey.pbData, UBound(baPublic) + 1)
                            End If
                            Call LocalFree(lPtr)
                        End If
                    End If
                    If UBound(baPrivate) >= 0 Then
                        If baPrivate(0) = 0 Then
                            Call CopyMemory(baPrivate(0), baPrivate(1), UBound(baPrivate))
                            ReDim Preserve baPrivate(0 To UBound(baPrivate) - 1) As Byte
                        End If
                    End If
                    If sCurve = "secp256r1" And UBound(baPrivate) = LNG_SECP256R1_KEYSZ - 1 And UBound(baPublic) >= LNG_SECP256R1_KEYSZ Then
                        sResult = "valid"
                        If Not pvCryptoEcdhSecp256r1SharedSecret(baBuffer, baPrivate, baPublic) Then
                            sResult = "invalid"
                        ElseIf Not pvArrayEqual(baBuffer, baShared) Then
                            sResult = "invalid"
                        End If
                        If JsonValue(oTest, "result") = "acceptable" And sResult = "valid" Then
                            lPassed = lPassed + 1
                        ElseIf JsonValue(oTest, "result") = sResult Then
                            lPassed = lPassed + 1
                        Else
                            lFailed = lFailed + 1
                            sComment = JsonValue(oTest, "comment")
                            DebugLog MODULE_NAME, FUNC_NAME, "[-] " & sCurve & ", " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString) & _
                                IIf(LenB(sFlags) <> 0, ", " & sFlags, vbNullString) & ", sResult=" & sResult & " vs " & JsonValue(oTest, "result")
                        End If
                    ElseIf sCurve = "secp384r1" And UBound(baPrivate) = LNG_SECP384R1_KEYSZ - 1 And UBound(baPublic) >= LNG_SECP384R1_KEYSZ Then
                        sResult = "valid"
                        If Not pvCryptoEcdhSecp384r1SharedSecret(baBuffer, baPrivate, baPublic) Then
                            sResult = "invalid"
                        ElseIf Not pvArrayEqual(baBuffer, baShared) Then
                            sResult = "invalid"
                        End If
                        If JsonValue(oTest, "result") = "acceptable" And sResult = "valid" Then
                            lPassed = lPassed + 1
                        ElseIf JsonValue(oTest, "result") = sResult Then
                            lPassed = lPassed + 1
                        Else
                            lFailed = lFailed + 1
                            sComment = JsonValue(oTest, "comment")
                            DebugLog MODULE_NAME, FUNC_NAME, "[-] " & sCurve & ", " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString) & _
                                IIf(LenB(sFlags) <> 0, ", " & sFlags, vbNullString) & ", sResult=" & sResult & " vs " & JsonValue(oTest, "result")
                        End If
                    Else
                        lSkipped = lSkipped + 1
                    End If
                Else
                    lSkipped = lSkipped + 1
                End If
            Next
        End If
    Next
    DebugLog MODULE_NAME, FUNC_NAME, "[+] Passed=" & lPassed & ", Failed=" & lFailed & ", Skipped=" & lSkipped
End Sub

Public Sub TestCryptoEcdsa(oJson As Object)
    Const FUNC_NAME     As String = "TestCryptoEcdsa"
    Dim oGroup          As Object
    Dim oTest           As Object
    Dim sComment        As String
    Dim sResult         As String
    Dim lPassed         As Long
    Dim lFailed         As Long
    Dim lSkipped        As Long
    '--- local
    Dim baPublic()      As Byte
    Dim baMsg()         As Byte
    Dim baSig()         As Byte
    Dim baBuffer()      As Byte
    Dim sFlags          As String
    Dim sCurve          As String
    Dim lCurveSize      As KeyCodeConstants
    Dim lHashSize       As Long
    Dim baTemp()        As Byte
    
    pvCryptoInit
    For Each oGroup In JsonValue(oJson, "testGroups")
        sCurve = JsonValue(oGroup, "key/curve")
        If sCurve = "secp256r1" Or sCurve = "secp384r1" Then
            baPublic = FromHex(JsonValue(oGroup, "key/uncompressed"))
            lCurveSize = JsonValue(oGroup, "key/keySize") / 8
            lHashSize = Replace(JsonValue(oGroup, "sha"), "SHA-", vbNullString) / 8
            For Each oTest In JsonValue(oGroup, "tests")
                baMsg = FromHex(JsonValue(oTest, "msg"))
                baTemp = FromHex(JsonValue(oTest, "sig"))
                If pvAsn1DecodeEcdsaSignature(baSig, baTemp, lCurveSize) Then
                    sFlags = Join(JsonValue(oTest, "flags/*"), ", ")
                    pvArrayHash lHashSize, baMsg, baBuffer
                    If UBound(baBuffer) + 1 < lCurveSize Then
                        baTemp = baBuffer
                        pvArrayAllocate baBuffer, lCurveSize, FUNC_NAME & ".baBuffer"
                        Call CopyMemory(baBuffer(lCurveSize - UBound(baTemp) - 1), baTemp(0), UBound(baTemp) + 1)
                    End If
                    If sCurve = "secp256r1" Then
                        sResult = "valid"
                        If Not pvCryptoEcdsaSecp256r1Verify(baPublic, baBuffer, baSig) Then
                            sResult = "invalid"
                        End If
                        If JsonValue(oTest, "result") = "acceptable" And sResult = "valid" Then
                            lPassed = lPassed + 1
                        ElseIf JsonValue(oTest, "result") = sResult Then
                            lPassed = lPassed + 1
                        Else
                            lFailed = lFailed + 1
                            sComment = JsonValue(oTest, "comment")
                            DebugLog MODULE_NAME, FUNC_NAME, "[-] " & sCurve & ", " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString) & _
                                IIf(LenB(sFlags) <> 0, ", " & sFlags, vbNullString) & ", sResult=" & sResult & " vs " & JsonValue(oTest, "result")
                        End If
                    Else ' If sCurve = "secp384r1" Then
                        sResult = "valid"
                        If Not pvCryptoEcdsaSecp384r1Verify(baPublic, baBuffer, baSig) Then
                            sResult = "invalid"
                        End If
                        If JsonValue(oTest, "result") = "acceptable" And sResult = "valid" Then
                            lPassed = lPassed + 1
                        ElseIf JsonValue(oTest, "result") = sResult Then
                            lPassed = lPassed + 1
                        Else
                            lFailed = lFailed + 1
                            sComment = JsonValue(oTest, "comment")
                            DebugLog MODULE_NAME, FUNC_NAME, "[-] " & sCurve & ", " & JsonValue(oTest, "tcId") & IIf(LenB(sComment) <> 0, " (" & sComment & ")", vbNullString) & _
                                IIf(LenB(sFlags) <> 0, ", " & sFlags, vbNullString) & ", sResult=" & sResult & " vs " & JsonValue(oTest, "result")
                        End If
                    End If
                Else
                    lSkipped = lSkipped + 1
                End If
            Next
        End If
    Next
    DebugLog MODULE_NAME, FUNC_NAME, "[+] Passed=" & lPassed & ", Failed=" & lFailed & ", Skipped=" & lSkipped
End Sub

#End If ' ImplTestCrypto

'= trampolines ===========================================================

Private Function pvCallCurve25519Multiply(ByVal Pfn As Long, pSecretPtr As Byte, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
End Function

Private Function pvCallCurve25519MulBase(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
End Function

Private Function pvCallSecpMakeKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
    ' int ecc_make_key384(uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384])
End Function

Private Function pvCallSecpSharedSecret(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte, pSecretPtr As Byte) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
    ' int ecdh_shared_secret384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384], uint8_t p_secret[ECC_BYTES_384])
End Function

Private Function pvCallSecpUncompressKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pUncompressedKeyPtr As Byte) As Long
    ' int ecdh_uncompress_key(const uint8_t p_publicKey[ECC_BYTES + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES + 1])
    ' int ecdh_uncompress_key384(const uint8_t p_publicKey[ECC_BYTES_384 + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES_384 + 1])
End Function

Private Function pvCallSecpSign(ByVal Pfn As Long, pPrivKeyPtr As Byte, pHashPtr As Byte, pRandomPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_sign(const uint8_t p_privateKey[ECC_BYTES], const uint8_t p_hash[ECC_BYTES], uint64_t k[NUM_ECC_DIGITS], uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_sign384(const uint8_t p_privateKey[ECC_BYTES_384], const uint8_t p_hash[ECC_BYTES_384], uint64_t k[NUM_ECC_DIGITS_384], uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCallSecpVerify(ByVal Pfn As Long, pPubKeyPtr As Byte, pHashPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_verify(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_hash[ECC_BYTES], const uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_verify384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_hash[ECC_BYTES_384], const uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCallSha2Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha256_init(cf_sha256_context *ctx)
    ' void cf_sha384_init(cf_sha384_context *ctx)
    ' void cf_sha512_init(cf_sha512_context *ctx)
End Function

Private Function pvCallSha2Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha256_update(cf_sha256_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha384_update(cf_sha384_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha512_update(cf_sha512_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvCallSha2Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, pHashPtr As Byte) As Long
    ' void cf_sha256_digest_final(cf_sha256_context *ctx, uint8_t hash[LNG_SHA256_HASHSZ])
    ' void cf_sha384_digest_final(cf_sha384_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
    ' void cf_sha512_digest_final(cf_sha512_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
End Function

Private Function pvCallChacha20Poly1305Encrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, _
            pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            pCipherTextPtr As Byte, pTagPtr As Byte) As Long
    ' void cf_chacha20poly1305_encrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                  const uint8_t *plaintext, size_t nbytes, uint8_t *ciphertext, uint8_t tag[16])
End Function

Private Function pvCallChacha20Poly1305Decrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, _
            pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, _
            pTagPtr As Byte, pPlaintTextPtr As Byte) As Long
    ' int cf_chacha20poly1305_decrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                 const uint8_t *ciphertext, size_t nbytes, const uint8_t tag[16], uint8_t *plaintext)
End Function

Private Function pvCallAesGcmEncrypt( _
            ByVal Pfn As Long, pCipherTextPtr As Byte, pTagPtr As Byte, pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_encrypt(uint8_t *c, uint8_t *mac, const uint8_t *m, const size_t mlen, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, size_t klen)
End Function

Private Function pvCallAesGcmDecrypt( _
            ByVal Pfn As Long, pPlaintTextPtr As Byte, pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, pTagPtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_decrypt(uint8_t *m, const uint8_t *c, const size_t clen, const uint8_t *mac, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, const size_t klen)
End Function

Private Function pvCallAesCbcEncrypt( _
            ByVal Pfn As Long, pCipherTextPtr As Byte, pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' static void cf_aescbc_encrypt(uint8_t *c, const uint8_t *m, const size_t mlen,
    '                               const uint8_t *npub, const uint8_t *k, const size_t klen)
End Function

Private Function pvCallAesCbcDecrypt( _
            ByVal Pfn As Long, pPlaintTextPtr As Byte, pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, _
            pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' static void cf_aescbc_decrypt(uint8_t *m, const uint8_t *c, const size_t clen,
    '                              const uint8_t *npub, const uint8_t *k, const size_t klen)
End Function

Private Function pvCallRsaModExp(ByVal Pfn As Long, ByVal lSize As Long, pBasePtr As Byte, pExpPtr As Byte, pModPtr As Byte, pRetPtr As Byte) As Long
    ' static void rsa_modexp(const uint32_t maxbytes, const uint8_t *base_in, const uint8_t *exp_in, const uint8_t *mod_in, uint8_t *ret_out)
End Function

Private Function pvCallRsaCrtModExp(ByVal Pfn As Long, ByVal lSize As Long, pBasePtr As Byte, pExpPtr As Byte, pModPtr As Byte, pPPtr As Byte, pQPtr As Byte, pIqmpPtr As Byte, pRetPtr As Byte) As Long
    ' static void rsa_crt_modexp(const uint32_t maxbytes, const uint8_t *base_in, const uint8_t *exp_in, const uint8_t *mod_in,
    '                            const uint8_t *p_in, const uint8_t *q_in, const uint8_t *iqmp_in, uint8_t *ret_out)
End Function

Private Sub pvAppendBuffer(ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, ByVal a09 As Long, ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, ByVal a15 As Long, ByVal a16 As Long, ByVal a17 As Long, ByVal a18 As Long, ByVal a19 As Long, ByVal a20 As Long, ByVal a21 As Long, ByVal a22 As Long, ByVal a23 As Long, ByVal a24 As Long, ByVal a25 As Long, ByVal a26 As Long, ByVal a27 As Long, ByVal a28 As Long, ByVal a29 As Long, ByVal a30 As Long, ByVal a31 As Long, ByVal a32 As Long)
    #If a01 And a02 And a03 And a04 And a05 And a06 And a07 And a08 And a09 And a10 And a11 And a12 And a13 And a14 And a15 And a16 And a17 And a18 And a19 And a20 And a21 And a22 And a23 And a24 And a25 And a26 And a27 And a28 And a29 And a30 And a31 And a32 Then '--- touch args
    #End If
    Call CopyMemory(m_baBuffer(m_lBuffIdx), a01, 4 * 32)
    m_lBuffIdx = m_lBuffIdx + 4 * 32
End Sub

Private Sub pvGetGlobData(baBuffer() As Byte)
    ReDim m_baBuffer(0 To 2048 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &H76AC8030, &H76B46E90, &H76AC7300, &H0&, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &H0&, &H0&, &H0&, &H1&, &HFFFFFFFF, &H27D2604B, &H3BCE3C3E, &HCC53B0F6, &H651D06B0, &H769886BC, &HB3EBBD55, &HAA3A93E7, &H5AC635D8, &HD898C296, &HF4A13945, &H2DEB33A0, &H77037D81, &H63A440F2, &HF8BCE6E5, &HE12C4247, &H6B17D1F2, &H37BF51F5, &HCBB64068, &H6B315ECE, &H2BCE3357
    pvAppendBuffer &H7C0F9E16, &H8EE7EB4A, &HFE1A7F9B, &H4FE342E2, &HFC632551, &HF3B9CAC2, &HA7179E84, &HBCE6FAAD, &HFFFFFFFF, &HFFFFFFFF, &H0&, &HFFFFFFFF, &HFFFFFFFF, &H0&, &H0&, &HFFFFFFFF, &HFFFFFFFE, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HD3EC2AEF, &H2A85C8ED, &H8A2ED19D, &HC656398D, &H5013875A, &H314088F, &HFE814112, &H181D9C6E
    pvAppendBuffer &HE3F82D19, &H988E056B, &HE23EE7E4, &HB3312FA7, &H72760AB7, &H3A545E38, &HBF55296C, &H5502F25D, &H82542A38, &H59F741E0, &H8BA79B98, &H6E1D3B62, &HF320AD74, &H8EB1C71E, &HBE8B0537, &HAA87CA22, &H90EA0E5F, &H7A431D7C, &H1D7E819D, &HA60B1CE, &HB5F0B8C0, &HE9DA3113, &H289A147C, &HF8F41DBD, &H9292DC29, &H5D9E98BF, &H96262C6F, &H3617DE4A, &HCCC52973, &HECEC196A, &H48B0A77A, &H581A0DB2
    pvAppendBuffer &HF4372DDF, &HC7634D81, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA
    pvAppendBuffer &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3
    pvAppendBuffer &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2, &HD728AE22, &H428A2F98, &H23EF65CD, &H71374491, &HEC4D3B2F, &HB5C0FBCF, &H8189DBBC, &HE9B5DBA5, &HF348B538, &H3956C25B, &HB605D019, &H59F111F1, &HAF194F9B, &H923F82A4, &HDA6D8118, &HAB1C5ED5, &HA3030242, &HD807AA98, &H45706FBE, &H12835B01, &H4EE4B28C, &H243185BE, &HD5FFB4E2, &H550C7DC3
    pvAppendBuffer &HF27B896F, &H72BE5D74, &H3B1696B1, &H80DEB1FE, &H25C71235, &H9BDC06A7, &HCF692694, &HC19BF174, &H9EF14AD2, &HE49B69C1, &H384F25E3, &HEFBE4786, &H8B8CD5B5, &HFC19DC6, &H77AC9C65, &H240CA1CC, &H592B0275, &H2DE92C6F, &H6EA6E483, &H4A7484AA, &HBD41FBD4, &H5CB0A9DC, &H831153B5, &H76F988DA, &HEE66DFAB, &H983E5152, &H2DB43210, &HA831C66D, &H98FB213F, &HB00327C8, &HBEEF0EE4, &HBF597FC7
    pvAppendBuffer &H3DA88FC2, &HC6E00BF3, &H930AA725, &HD5A79147, &HE003826F, &H6CA6351, &HA0E6E70, &H14292967, &H46D22FFC, &H27B70A85, &H5C26C926, &H2E1B2138, &H5AC42AED, &H4D2C6DFC, &H9D95B3DF, &H53380D13, &H8BAF63DE, &H650A7354, &H3C77B2A8, &H766A0ABB, &H47EDAEE6, &H81C2C92E, &H1482353B, &H92722C85, &H4CF10364, &HA2BFE8A1, &HBC423001, &HA81A664B, &HD0F89791, &HC24B8B70, &H654BE30, &HC76C51A3
    pvAppendBuffer &HD6EF5218, &HD192E819, &H5565A910, &HD6990624, &H5771202A, &HF40E3585, &H32BBD1B8, &H106AA070, &HB8D2D0C8, &H19A4C116, &H5141AB53, &H1E376C08, &HDF8EEB99, &H2748774C, &HE19B48A8, &H34B0BCB5, &HC5C95A63, &H391C0CB3, &HE3418ACB, &H4ED8AA4A, &H7763E373, &H5B9CCA4F, &HD6B2B8A3, &H682E6FF3, &H5DEFB2FC, &H748F82EE, &H43172F60, &H78A5636F, &HA1F0AB72, &H84C87814, &H1A6439EC, &H8CC70208
    pvAppendBuffer &H23631E28, &H90BEFFFA, &HDE82BDE9, &HA4506CEB, &HB2C67915, &HBEF9A3F7, &HE372532B, &HC67178F2, &HEA26619C, &HCA273ECE, &H21C0C207, &HD186B8C7, &HCDE0EB1E, &HEADA7DD6, &HEE6ED178, &HF57D4F7F, &H72176FBA, &H6F067AA, &HA2C898A6, &HA637DC5, &HBEF90DAE, &H113F9804, &H131C471B, &H1B710B35, &H23047D84, &H28DB77F5, &H40C72493, &H32CAAB7B, &H15C9BEBC, &H3C9EBE0A, &H9C100D4C, &H431D67C4
    pvAppendBuffer &HCB3E42B6, &H4CC5D4BE, &HFC657E2A, &H597F299C, &H3AD6FAEC, &H5FCB6FAB, &H4A475817, &H6C44198C, &H61707865, &H3120646E, &H79622D36, &H6B206574, &H70786500, &H20646E61, &H622D3233, &H20657479, &H6B&, &H5&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    pvAppendBuffer &H0&, &HFC&, &H7B777C63, &HC56F6BF2, &H2B670130, &H76ABD7FE, &H7DC982CA, &HF04759FA, &HAFA2D4AD, &HC072A49C, &H2693FDB7, &HCCF73F36, &HF1E5A534, &H1531D871, &HC323C704, &H9A059618, &HE2801207, &H75B227EB, &H1A2C8309, &HA05A6E1B, &HB3D63B52, &H842FE329, &HED00D153, &H5BB1FC20, &H39BECB6A, &HCF584C4A, &HFBAAEFD0, &H85334D43, &H7F02F945, &HA89F3C50, &H8F40A351, &HF5389D92
    pvAppendBuffer &H21DAB6BC, &HD2F3FF10, &HEC130CCD, &H1744975F, &H3D7EA7C4, &H73195D64, &HDC4F8160, &H88902A22, &H14B8EE46, &HDB0B5EDE, &HA3A32E0, &H5C240649, &H62ACD3C2, &H79E49591, &H6D37C8E7, &HA94ED58D, &HEAF4566C, &H8AE7A65, &H2E2578BA, &HC6B4A61C, &H1F74DDE8, &H8A8BBD4B, &H66B53E70, &HEF60348, &HB9573561, &H9E1DC186, &H1198F8E1, &H948ED969, &HE9871E9B, &HDF2855CE, &HD89A18C, &H6842E6BF
    pvAppendBuffer &HF2D9941, &H16BB54B0, &H402018D, &H40201008, &H52361B80, &H30D56A09, &HBF38A536, &H819EA340, &H7CFBD7F3, &H9B8239E3, &H3487FF2F, &HC444438E, &H54CBE9DE, &HA632947B, &HEE3D23C2, &H420B954C, &H84EC3FA, &H2866A12E, &H76B224D9, &H6D49A25B, &H7225D18B, &H8664F6F8, &HD4169868, &H5DCC5CA4, &H6C92B665, &HFD504870, &H5EDAB9ED, &HA7574615, &H90849D8D, &H8C00ABD8, &HF70AD3BC, &HB80558E4
    pvAppendBuffer &HD00645B3, &HCA8F1E2C, &HC1020F3F, &H103BDAF, &H3A6B8A13, &H4F411191, &H97EADC67, &HF0CECFF2, &H9673E6B4, &HE72274AC, &HE28535AD, &H1CE837F9, &H476EDF75, &H1D711AF1, &H6F89C529, &HAA0E62B7, &HFC1BBE18, &HC64B3E56, &H9A2079D2, &H78FEC0DB, &H1FF45ACD, &H8833A8DD, &HB131C707, &H27591012, &H605FEC80, &H19A97F51, &H2D0D4AB5, &H939F7AE5, &HA0EF9CC9, &HAE4D3BE0, &HC8B0F52A, &H833CBBEB
    pvAppendBuffer &H17619953, &HBA7E042B, &HE126D677, &H55631469, &H7D0C21, &H0&, &H1&, &H1&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 1952 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub

Private Sub pvGetThunkData(baBuffer() As Byte)
    ReDim m_baBuffer(0 To 34176 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &H480238, &H29DA&, &H2CD8&, &H3A3A&, &H3E54&, &H3F11&, &H3F8B&, &H421F&, &H3AE4&, &H3EAE&, &H3F4E&, &H40CB&, &H45E3&, &H3522&, &H3572&, &H3437&, &H35CE&, &H3659&, &H35A5&, &H378B&, &H3816&, &H365E&, &H2927&, &H28EA&, &H1FCC&, &H1F7D&, &H1F2A&, &H1ED7&, &H6CBA&, &H6B09&, &HA9C45C7, &H89000000
    pvAppendBuffer &H4D89AC55, &HB845C7B0, &H11&, &H12BC45C7, &H8B000000, &HFF788584, &H895FFFFF, &H5D89F075, &H45C75EE8, &H1C0&, &HC445C700, &H2&, &H3C845C7, &HC7000000, &H4CC45, &H45C70000, &H5D0&, &HD445C700, &H6&, &H7D845C7, &HC7000000, &H8DC45, &H45C70000, &H9E0&, &HE445C700, &HA&, &H89F45589, &H448BF84D, &H8B5BC085, &H4C25DE5, &HE800&, &H2D580000
    pvAppendBuffer &H4740FE, &H47400005, &HC3008B00, &HE8&, &H112D5800, &H5004741, &H474000, &HEC8B55C3, &H5348EC83, &H56105D8B, &H565E046A, &H74D0E853, &HC0850000, &H163850F, &H56570000, &H8D0C75FF, &HE850D845, &H7E77&, &H8D087D8B, &H5056D845, &HB8458D57, &H7E38E850, &H8D560000, &H5050D845, &H7E5AE8, &HFF535600, &H75FF0C75, &H7E20E80C, &H53560000, &H7E45E853, &HE8560000, &HFFFFFF79
    pvAppendBuffer &H5010C083, &HE8575753, &H78B7&, &HFF67E856, &HC083FFFF, &H53535010, &H78A5E853, &HE8560000, &HFFFFFF55, &H5010C083, &HE8535753, &H7E39&, &H57575356, &H7DD9E8, &H3AE85600, &H83FFFFFF, &H575010C0, &H78E85357, &H56000078, &HFFFF28E8, &H10C083FF, &H57575350, &H7866E8, &H57006A00, &H823AE8, &H56C20B00, &HAE82574, &H83FFFFFF, &H575010C0, &H730FE857, &H46A0000, &HE8F08B57
    pvAppendBuffer &H7FA6&, &H91FE6C1, &H46A1C77, &H5706EB5E, &H7F95E8, &H53575600, &H7DA2E8, &HD6E85600, &H83FFFFFE, &H8D5010C0, &H5350B845, &H7DB7E853, &HE8560000, &HFFFFFEC1, &H5010C083, &H50B8458D, &HA2E85353, &H5600007D, &HFFFEACE8, &H10C083FF, &H458D5350, &HE85050B8, &H7D8D&, &HB8458D56, &HE8575750, &H7D2A&, &HFE8BE856, &HC083FFFF, &H458D5010, &H505750D8, &H7D6CE8, &H57535600
    pvAppendBuffer &H7F6AE8, &H75FF5600, &H60E8530C, &H5600007F, &H50D8458D, &HE80C75FF, &H7F53&, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H5368EC83, &H56105D8B, &H565E066A, &H7348E853, &HC0850000, &H177850F, &H56570000, &H8D0C75FF, &HE850C845, &H7CEF&, &H8D087D8B, &H5056C845, &H98458D57, &H7CB0E850, &H8D560000, &H5050C845, &H7CD2E8, &HFF535600, &H75FF0C75, &H7C98E80C, &H53560000, &H7CBDE853
    pvAppendBuffer &HE8560000, &HFFFFFDF1, &HB005&, &H57535000, &H772DE857, &HE8560000, &HFFFFFDDD, &HB005&, &H53535000, &H7719E853, &HE8560000, &HFFFFFDC9, &HB005&, &H57535000, &H7CABE853, &H53560000, &H4BE85757, &H5600007C, &HFFFDACE8, &HB005FF, &H57500000, &HE8E85357, &H56000076, &HFFFD98E8, &HB005FF, &H53500000, &HD4E85757, &H6A000076, &HA8E85700, &HB000080, &H277456C2, &HFFFD78E8
    pvAppendBuffer &HB005FF, &H57500000, &H717BE857, &H66A0000, &HE8F08B57, &H7E12&, &H91FE6C1, &H66A2C77, &H5706EB5E, &H7E01E8, &H53575600, &H7C0EE8, &H42E85600, &H5FFFFFD, &HB0&, &H98458D50, &HE8535350, &H7C21&, &HFD2BE856, &HB005FFFF, &H50000000, &H5098458D, &HAE85353, &H5600007C, &HFFFD14E8, &HB005FF, &H53500000, &H5098458D, &H7BF3E850, &H8D560000, &H57509845, &H7B90E857
    pvAppendBuffer &HE8560000, &HFFFFFCF1, &HB005&, &H458D5000, &H505750C8, &H7BD0E8, &H57535600, &H7DCEE8, &H75FF5600, &HC4E8530C, &H5600007D, &H50C8458D, &HE80C75FF, &H7DB7&, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H8758B56, &HE856046A, &H71B2&, &H1474C085, &H468D046A, &HA3E85020, &H85000071, &H330574C0, &H2EB40C0, &H5D5EC033, &H550004C2, &H8B56EC8B, &H66A0875, &H7184E856, &HC0850000
    pvAppendBuffer &H66A1474, &H5030468D, &H7175E8, &H74C08500, &H40C03305, &HC03302EB, &H4C25D5E, &HEC8B5500, &HA8EC81, &H8B530000, &H458D0C5D, &H6A5756B8, &HE8505304, &H7D37&, &H6A20438D, &H45895004, &H78858DF8, &H50FFFFFF, &H7D22E8, &H1475FF00, &HFF58858D, &H8D50FFFF, &H8D509845, &HFFFF7885, &H458D50FF, &H82E850B8, &H6A000008, &H1075FF04, &H7C54E8, &H83F63300, &H894602E8, &HC0851445
    pvAppendBuffer &HFF50577E, &H14E81075, &HB00007F, &H8B0475C2, &H3302EBCE, &H5E1C1C9, &HFF589D8D, &HD903FFFF, &H398458D, &H78B58DC1, &H53FFFFFF, &H4589D9F7, &HF10350FC, &H3B87D8D, &HE85756F9, &H4F3&, &HFF535756, &HBBE8FC75, &H8B000002, &H89481445, &H16A1445, &H7FC0855E, &HFF006AA9, &HBCE81075, &HB00007E, &H330274C2, &H5E6C1F6, &HFF589D8D, &HDE03FFFF, &H53107589, &H398458D, &H78BD8DC6
    pvAppendBuffer &H2BFFFFFF, &HB8758DFE, &H5010752B, &HA0E85657, &H6A000004, &HFB5FE804, &HC083FFFF, &H458D5010, &H458D5098, &H458D50B8, &H3AE850D8, &H6A00007A, &H458D5704, &HE85050D8, &H79D6&, &H75FF046A, &HD8458D0C, &HC7E85050, &H6A000079, &HFB27E804, &HC083FFFF, &H458D5010, &HE85050D8, &H74A2&, &H75FF046A, &HD8458DF8, &HA3E85050, &H6A000079, &H458D5604, &HE85050D8, &H7996&, &H458D5657
    pvAppendBuffer &H10450398, &HFBE85053, &H8D000001, &H8D50D845, &HFFFF5885, &H458D50FF, &H52E85098, &H8B00000A, &H458D0875, &H50046A98, &H7BC1E856, &H46A0000, &HFF58858D, &H8D50FFFF, &HE8502046, &H7BAF&, &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &HF8EC81, &H8B530000, &H458D0C5D, &H6A575698, &HE8505306, &H7B8B&, &H6A30438D, &H45895006, &H38858DF8, &H50FFFFFF, &H7B76E8, &H1475FF00, &HFF08858D
    pvAppendBuffer &H8D50FFFF, &HFFFF6885, &H858D50FF, &HFFFFFF38, &H98458D50, &H74DE850, &H66A0000, &HE81075FF, &H7AA5&, &HE883DB33, &H45894302, &H7EC08514, &H75FF505A, &H7D65E810, &HC20B0000, &HC38B0475, &HC03302EB, &H8D30C06B, &HFFFF089D, &H688D8DFF, &H8DFFFFFF, &HFFFF38B5, &H987D8DFF, &HC803D803, &H89D8F753, &H351FC4D, &H56F803F0, &H4CCE857, &H57560000, &HFC75FF53, &H219E8, &H14458B00
    pvAppendBuffer &H14458948, &H855B016A, &H6AA67FC0, &H1075FF00, &H7D0AE8, &H74C20B00, &H6BDB3302, &H8D8D30C3, &HFFFFFF68, &HFF089D8D, &HBD8DFFFF, &HFFFFFF38, &H398758D, &H53C803D8, &H89F82B51, &HF02B104D, &H77E85657, &H6A000004, &HF9ABE806, &HB005FFFF, &H50000000, &HFF68858D, &H8D50FFFF, &H8D509845, &HE850C845, &H7881&, &H8D57066A, &H5050C845, &H781DE8, &HFF066A00, &H458D0C75, &HE85050C8
    pvAppendBuffer &H780E&, &H6EE8066A, &H5FFFFF9, &HB0&, &HC8458D50, &HE7E85050, &H6A000072, &HF875FF06, &H50C8458D, &H77E8E850, &H66A0000, &HC8458D56, &HDBE85050, &H57000077, &H75FF5356, &H154E810, &H458D0000, &H858D50C8, &HFFFFFF08, &H68858D50, &H50FFFFFF, &H8E3E8, &H8758B00, &HFF68858D, &H66AFFFF, &H4E85650, &H6A00007A, &H8858D06, &H50FFFFFF, &H5030468D, &H79F2E8, &H5B5E5F00
    pvAppendBuffer &HC25DE58B, &H8B550010, &H20EC83EC, &H6A575653, &HE8565E04, &HFFFFF8E1, &H83085D8B, &H535010C0, &H8D1075FF, &HE850E045, &H77BD&, &HE0458D56, &H88E85050, &H56000077, &H50E0458D, &H4FE85353, &H56000077, &H50E0458D, &HFF1075FF, &H3FE81075, &H56000077, &HFFF8A0E8, &HC758BFF, &H8B10C083, &H5650147D, &H7EE85757, &H6A000077, &H458D5704, &H48E850E0, &H6A000077, &HF87BE804, &HC083FFFF
    pvAppendBuffer &H8D535010, &H5050E045, &H775CE8, &HE8046A00, &HFFFFF865, &H5010C083, &H8D1075FF, &H5050E045, &H7744E8, &HE8046A00, &HFFFFF84D, &H5010C083, &H1075FF53, &HE81075FF, &H772D&, &H75FF046A, &HE8565610, &H76CA&, &H2AE8046A, &H83FFFFF8, &H8D5010C0, &H5350E045, &HE81075FF, &H7709&, &H75FF046A, &HE8575710, &H76A6&, &H6E8046A, &H83FFFFF8, &H565010C0, &HEAE85757, &H6A000076
    pvAppendBuffer &HE0458D04, &H1075FF50, &H78E2E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H30EC83EC, &H6A575653, &HE8575F06, &HFFFFF7D1, &HBE085D8B, &HB0&, &H5350C603, &H8D1075FF, &HE850D045, &H76A9&, &HD0458D57, &H74E85050, &H57000076, &H50D0458D, &H3BE85353, &H57000076, &H50D0458D, &HFF1075FF, &H2BE81075, &H57000076, &HFFF78CE8, &H147D8BFF, &H758BC603, &H5756500C, &H766BE857, &H66A0000
    pvAppendBuffer &HD0458D57, &H7635E850, &H66A0000, &HFFF768E8, &HB005FF, &H53500000, &H50D0458D, &H7647E850, &H66A0000, &HFFF750E8, &HB005FF, &HFF500000, &H458D1075, &HE85050D0, &H762D&, &H36E8066A, &H5FFFFF7, &HB0&, &H75FF5350, &H1075FF10, &H7614E8, &HFF066A00, &H56561075, &H75B1E8, &HE8066A00, &HFFFFF711, &HB005&, &H458D5000, &HFF5350D0, &HEEE81075, &H6A000075, &H1075FF06
    pvAppendBuffer &H8BE85757, &H6A000075, &HF6EBE806, &HB005FFFF, &H50000000, &HE8575756, &H75CD&, &H458D066A, &H75FF50D0, &H77C5E810, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC83EC8B, &H57565360, &H535B046A, &HFFF6B4E8, &H107D8BFF, &H5010C083, &H8D0875FF, &H5057C045, &H7590E8, &H458D5300, &HE85050C0, &H755B&, &HC0458D53, &H875FF50, &HE80875FF, &H751E&, &HC0458D53, &HE8575750, &H7512&
    pvAppendBuffer &HF673E853, &H5D8BFFFF, &H10C0830C, &H5014758B, &H458D5653, &HA8E850C0, &H6A00006F, &HF657E804, &HC083FFFF, &H56535010, &H753BE856, &H46A0000, &HFFF644E8, &H10C083FF, &H875FF50, &H57E0458D, &H7523E850, &H46A0000, &H50E0458D, &HBFE85353, &H6A000074, &HF61FE804, &HC083FFFF, &HFF575010, &H458D0875, &H58E850E0, &H6A00006F, &HE8575604, &H74CB&, &HFEE8046A, &H83FFFFF5, &H8D5010C0
    pvAppendBuffer &H5750E045, &H74DFE857, &H46A0000, &HFFF5E8E8, &H10C083FF, &H75FF5750, &HA0458D08, &H74C7E850, &H46A0000, &H50A0458D, &H63E85656, &H6A000074, &HF5C3E804, &H106AFFFF, &H50C7035F, &HE8565653, &H74A5&, &H565E046A, &H50C0458D, &H50A0458D, &H746AE8, &H9EE85600, &H3FFFFF5, &H458D50C7, &H458D50E0, &HE85050A0, &H747D&, &HF587E856, &HC703FFFF, &H875FF50, &H50A0458D, &H50E0458D
    pvAppendBuffer &H7464E8, &H458D5600, &H458D50C0, &HE85050E0, &H73FE&, &HF55FE856, &HC703FFFF, &H458D5350, &HE85350E0, &H7441&, &HA0458D56, &H875FF50, &H763AE8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H90EC81EC, &H53000000, &H66A5756, &H26E8565E, &H8BFFFFF5, &HB0BB107D, &H3000000, &H75FF50C3, &HA0458D08, &HFEE85057, &H56000073, &H50A0458D, &H73C9E850, &H8D560000, &HFF50A045, &H75FF0875
    pvAppendBuffer &H738CE808, &H8D560000, &H5750A045, &H7380E857, &HE8560000, &HFFFFF4E1, &H314758B, &HC5D8BC3, &H8D565350, &HE850A045, &H6E17&, &HC6E8066A, &H5FFFFF4, &HB0&, &H56565350, &H73A8E8, &HE8066A00, &HFFFFF4B1, &HB005&, &H75FF5000, &HD0458D08, &H8EE85057, &H6A000073, &HD0458D06, &HE8535350, &H732A&, &H8AE8066A, &H5FFFFF4, &HB0&, &H75FF5750, &HD0458D08, &H6DC1E850
    pvAppendBuffer &H66A0000, &H34E85756, &H6A000073, &HF467E806, &HB005FFFF, &H50000000, &H50D0458D, &H46E85757, &H6A000073, &HF44FE806, &HB005FFFF, &H50000000, &H875FF57, &HFF70858D, &HE850FFFF, &H7329&, &H858D066A, &HFFFFFF70, &HE8565650, &H72C2&, &H22E8066A, &HBFFFFFF4, &HB0&, &H5350C703, &H2E85656, &H6A000073, &H8D565E06, &H8D50A045, &HFFFF7085, &HC4E850FF, &H56000072, &HFFF3F8E8
    pvAppendBuffer &H50C703FF, &H50D0458D, &HFF70858D, &H5050FFFF, &H72D4E8, &HDEE85600, &H3FFFFF3, &H75FF50C7, &H70858D08, &H50FFFFFF, &H50D0458D, &H72B8E8, &H458D5600, &H458D50A0, &HE85050D0, &H7252&, &HF3B3E856, &HC703FFFF, &H458D5350, &HE85350D0, &H7295&, &H70858D56, &H50FFFFFF, &HE80875FF, &H748B&, &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &H6A20EC83, &H875FF04, &HE81075FF, &H746F&
    pvAppendBuffer &H75FF046A, &H1475FF0C, &H7462E8, &H8D046A00, &HE850E045, &H67EE&, &HE46583, &H187D83, &H1E045C7, &H74000000, &HFF046A0B, &HE8501875, &H743B&, &H50E0458D, &HFF0C75FF, &HAEE80875, &H8D000002, &HFF50E045, &H75FF0C75, &HF34BE808, &H458DFFFF, &H75FF50E0, &H1075FF14, &H290E8, &H5DE58B00, &H550014C2, &HEC83EC8B, &HFF066A30, &H75FF0875, &H73F5E810, &H66A0000, &HFF0C75FF
    pvAppendBuffer &HE8E81475, &H6A000073, &HD0458D06, &H6774E850, &H65830000, &H7D8300D4, &H45C70018, &H1D0&, &H6A0B7400, &H1875FF06, &H73C1E850, &H458D0000, &H75FF50D0, &H875FF0C, &H27FE8, &HD0458D00, &HC75FF50, &HE80875FF, &HFFFFF459, &H50D0458D, &HFF1475FF, &H61E81075, &H8B000002, &H14C25DE5, &H448B5300, &H4C8B0C24, &HE1F71024, &H448BD88B, &H64F70824, &HD8031424, &H824448B, &HD303E1F7
    pvAppendBuffer &H10C25B, &H7340F980, &H20F98015, &HA50F0673, &HC3E0D3C2, &HC033D08B, &HD31FE180, &HC033C3E2, &H80C3D233, &H157340F9, &H7320F980, &HD0AD0F06, &H8BC3EAD3, &H80D233C2, &HE8D31FE1, &H33C033C3, &H8B55C3D2, &H10558BEC, &H758B5653, &H7D8B570C, &H6AF22B08, &H5BFA2B10, &H3160C8B, &H16448B0A, &H4421304, &H8D170C89, &H44890852, &HEB83FC17, &H5FE57501, &HC25D5B5E, &H8B55000C, &H1C558BEC
    pvAppendBuffer &H5620458B, &H5708758B, &H30C7D8B, &H104513D7, &H46891689, &H10453B04, &H4720D77, &H773D73B, &H3340C033, &H3304EBC9, &H3C88BC0, &H135F2445, &H4503284D, &H8468914, &H4D13C68B, &HC4E8918, &H24C25D5E, &HEC8B5500, &H8B0C558B, &H28B084D, &H428B0131, &H4413104, &H3108428B, &H428B0841, &HC41310C, &H8C25D, &H83EC8B55, &H4D8B0CEC, &H8B565308, &H18B1075, &HC104C183, &H335702EE
    pvAppendBuffer &H107589FF, &H8DFC4D89, &H48504, &H45890000, &H74F685F8, &HC4D292C, &H7589D98B, &HC758B08, &H8D087D8B, &HE8501E04, &H5A48&, &H5B8D0389, &H1EF8304, &H758BED75, &H8BFE8B10, &H458BFC4D, &HC7DB33F8, &H10845, &HF83B0000, &H96830F, &HC78B0000, &H48DC62B, &H10458981, &HFCB9548B, &H3B0C5589, &HFF0575DE, &HDB330845, &H3475DB85, &HFFF104E8, &H58805FF, &H8B500000, &HC0C10C45
    pvAppendBuffer &H70E85008, &H89000064, &HEAE80C45, &H8BFFFFF0, &H558B084D, &H84B60F0C, &H68808, &H18E0C100, &H20EBD033, &H7606FE83, &H4FB831E, &HC6E81975, &H5FFFFF0, &H588&, &HC458B50, &H6435E850, &HD08B0000, &H8B10458B, &HFC458B08, &HC89CA33, &H458B47B8, &HFC4D8B10, &H4304C083, &H3B104589, &H820FF87D, &HFFFFFF74, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H8D20EC83, &H46AE045, &H501075FF
    pvAppendBuffer &H6F3AE8, &H8D046A00, &HFF50E045, &H75FF0875, &H6EFCE808, &H46A0000, &H8D1075FF, &H5050E045, &H6EEDE8, &H8D046A00, &HFF50E045, &H75FF0C75, &H6EDCE80C, &HE58B0000, &HCC25D, &H83EC8B55, &H458D30EC, &HFF066AD0, &HE8501075, &H6EEF&, &H458D066A, &H75FF50D0, &H875FF08, &H6EB1E8, &HFF066A00, &H458D1075, &HE85050D0, &H6EA2&, &H458D066A, &H75FF50D0, &HC75FF0C, &H6E91E8
    pvAppendBuffer &H5DE58B00, &H55000CC2, &HEC83EC8B, &H8B565318, &H8B570C75, &H78B087D, &H27F063B, &H588D068B, &H5D895301, &H524DE8EC, &HD08B0000, &HC033C933, &H21E85589, &H2141F845, &H4589F445, &H7CD93BFC, &H8DC78B67, &HC62B045E, &H5589D62B, &H3BD08BF0, &H8B057F0F, &H2EB1A3C, &HE3BFF33, &H38B077F, &HEB0C4589, &HC658304, &H1C03300, &H7D8B0C7D, &H1C013F8, &H7D8B0C7D, &HF44513F0, &HF46583
    pvAppendBuffer &H8BF84589, &HC0850C45, &H8B1F0489, &H458B087D, &H3B0974FC, &H8B057EC8, &HFC4589C1, &H4C38341, &H7EEC4D3B, &HE8558BAA, &H2895E5F, &H8B5BC28B, &H8C25DE5, &HEC8B5500, &H5310EC83, &H7D8B5756, &H8B378B0C, &H2E3C1DE, &H89F87589, &H26E8F05D, &H53FFFFEF, &HD08B10FF, &H5589C933, &H7EF6850C, &H2B078B0F, &H87048BC1, &H418A0489, &HF17CCE3B, &H8B085D8B, &H7FDE3B1B, &H15E8D03, &HE7C1FB8B
    pvAppendBuffer &HFC7D8902, &HFFEEF0E8, &H10FF57FF, &H5589D08B, &H7EDB85F4, &H33CB8B08, &HF3FA8BC0, &H84D8BAB, &H3947FF33, &H8B1C7C39, &HFCC283C3, &H8B02E0C1, &H8BD003F7, &H8946B104, &HFC528D02, &HF37E313B, &H8BF8758B, &H8B0C45, &H8458950, &H4F7E8, &H8558B00, &H85F84589, &H8B1A74C0, &H3BE2D3C8, &H6A127EF7, &HC82B5920, &H8B0C458B, &HE8D30440, &H458BD00B, &HE85250F8, &H57C8&, &H1475FF50
    pvAppendBuffer &HC75FF56, &H53F4758B, &H3A22E856, &H458B0000, &H74C08510, &H7C383923, &HFC4D8B1F, &HFCC1834B, &HDB85CE03, &H118B0478, &HD23302EB, &H83B81489, &H4B4704E9, &HEA7E383B, &H8BF075FF, &HE8530C5D, &H60EB&, &HFFEE30E8, &H50FF53FF, &HFC75FF08, &H60D9E856, &H1EE80000, &H56FFFFEE, &H5F0850FF, &HE58B5B5E, &H10C25D, &H56EC8B55, &H570C758B, &H80E836FF, &H6A000050, &H57F88B00, &H875FF56
    pvAppendBuffer &HFFFEB0E8, &H13F83FF, &H78B1076, &H873C83, &H89480875, &H1F88307, &HC78BF077, &HC25D5E5F, &H8B550008, &HFF006AEC, &H75FF0C75, &H4E808, &HC25D0000, &H8B550008, &H20EC83EC, &H85D8B53, &HC758B56, &H8B1B8B57, &HE45D89FB, &H4589068B, &H7FD83BE0, &H57F88B02, &H4FF6E8, &HC1CF8B00, &HC10302E1, &H8902E0C1, &H8AE8EC45, &HFFFFFFED, &H10FFEC75, &HF06583, &H4D89C88B, &H7EFF85F8
    pvAppendBuffer &H8BC78B6E, &H2E0C1D7, &H3081C8D, &HFC5D89C6, &H8BF44589, &HD88B0845, &H5D89DE2B, &HF45D8BE8, &HB7F103B, &H8BE8458B, &H45891804, &H8304EBF4, &H8B00F465, &H758BF045, &H813489F4, &H3B0C758B, &H8B077F16, &HF44D890B, &H658304EB, &H4D8B00F4, &H458340FC, &HEB8304FC, &HF4758B04, &HF045894A, &H458BC73B, &H8B318908, &HB47CF84D, &H8BE45D8B, &H4E0C1C7, &H8950C103, &H48D0845, &H8B5057F9
    pvAppendBuffer &H2E0C1C7, &H5150C103, &H3B3FE8, &HE0458B00, &H8BD80340, &H5D891045, &H74C0850C, &H3B008B0C, &H8D067FD8, &H5D890158, &H48E8530C, &H8B00004F, &H41C933F0, &HC18BD233, &H2A7C0E39, &H8D085D8B, &HC3833F0C, &H7FC13BFC, &HEB3B8B04, &H89FF3302, &HFF85863C, &HD08B0274, &H4EB8340, &HE57E063B, &H330C5D8B, &H458B41C9, &H85168910, &H836274C0, &H8300F065, &H3B00F465, &H8B567CD9, &H45E8DF8
    pvAppendBuffer &H7D89FE2B, &H7F0E3BFC, &H893B8B0A, &H7D8B087D, &H8304EBFC, &H3B000865, &H8B057F08, &H2EB1F3C, &HC033FF33, &H13087D03, &HF07D03C0, &H45133B89, &HF46583F4, &HF0458900, &H674FF85, &H27ECA3B, &H458BD18B, &H7D8B4110, &H4C383FC, &H7E0C4D3B, &HEC75FFB4, &H57F87D8B, &HD0E81689, &HE800005E, &HFFFFEC15, &H850FF57, &H5EC68B5F, &H5DE58B5B, &H55000CC2, &H4D8BEC8B, &HC985560C, &H758B3278
    pvAppendBuffer &HC1068B08, &HC83B02E0, &HC18B267D, &H3E28399, &HF8C1C203, &H3E18102, &H79800000, &HC9834905, &H448B41FC, &HE1C10486, &HFE8D303, &H2EBC0B6, &H5D5EC033, &H550008C2, &H5653EC8B, &H570C758B, &H8B087D8B, &H830F8B16, &H97501F9, &HF704478B, &H23C01BD8, &H1FA83C8, &H468B0975, &H1BD8F704, &H8BD023C0, &H7FCA3BC1, &H85C28B02, &H8D3174C0, &HFE2B861C, &H67EC13B, &H86583, &H348B06EB
    pvAppendBuffer &H875891F, &H47EC23B, &H7EBF633, &H7539338B, &H391B7208, &H11770875, &H8304EB83, &HD47501E8, &H5E5FC033, &H8C25D5B, &H40C03300, &HC883F4EB, &H55EFEBFF, &H8B53EC8B, &H57560C5D, &H9903438D, &H8D03E283, &HFFC1023C, &H7D895702, &H4DA5E80C, &HF08B0000, &H7C01FF83, &HC4D8B0A, &H33047E8D, &H85ABF3C0, &H8B3A74DB, &H3E7C1FB, &H4B084D8B, &H8A08EF83, &H4D894101, &H85D38B08, &H830379D2
    pvAppendBuffer &HFAC103C2, &H81CF8B02, &H1FE1&, &H49057980, &H41E0C983, &HD3C0B60F, &H964409E0, &H75DB8504, &H13E83CB, &H68B1076, &H863C83, &H89480875, &H1F88306, &H8B5FF077, &H5D5B5EC6, &H550008C2, &H4D8BEC8B, &H78C9850C, &H758B563D, &HC1068B08, &HC83B05E0, &HC18B2F7D, &H1FE28399, &HF8C1C203, &H1FE18105, &H79800000, &HC9834905, &HD23341E0, &H83E2D342, &H7400107D, &H86540906, &HF706EB04
    pvAppendBuffer &H865421D2, &HC25D5E04, &H8B55000C, &H18EC83EC, &H85D8B53, &H7D8B5756, &H3B338B0C, &H8B027F37, &HD0E85637, &H8300004C, &H8B00F065, &H89C033D0, &H8B40E855, &H7CF03BC8, &HF8558971, &H2904478D, &HD233F87D, &H29FC5D89, &H8942FC7D, &HB3BF445, &H5D8B0B7F, &H31C8BFC, &HEBEC5D89, &HEC658304, &H7F0F3B00, &HEB188B04, &H8BDB3302, &HC033F87D, &H5D03D3F7, &H3C013EC, &HD083DA, &H458BD08B
    pvAppendBuffer &H71C89F4, &H850C7D8B, &H3B0874DB, &H37EF04D, &H8BF04D89, &H8341085D, &H458904C0, &H7ECE3BF4, &HC5589AD, &H8BE8558B, &H4D8B0C45, &H5F0A89F0, &HC0855B5E, &HE8520A75, &H3294&, &H2EBC033, &HE58BC28B, &H8C25D, &H56EC8B55, &H87D8B57, &H106AF633, &H59206A5A, &HC78BCA2B, &HC085E8D3, &HCA8B0675, &HF203E7D3, &HE975FAD1, &H5EC68B5F, &H4C25D, &H8BEC8B55, &H83990845, &HC2031FE2
    pvAppendBuffer &H5605F8C1, &HECE85040, &H6A00004B, &H875FF01, &HE856F08B, &HFFFFFEA7, &H5D5EC68B, &H550004C2, &H5151EC8B, &H53084D8B, &HDB335756, &HD90481, &H8B000100, &H5483D914, &H8B0004D9, &H3304D944, &HC2AC0FC9, &H10F8C110, &H89F85589, &HFB83FC45, &H4103750F, &HD23304EB, &HFB83C033, &H17B8D0F, &HF61B006A, &HAF0FDEF7, &H6AD12BF7, &HD88325, &H6AE85250, &H3FFFFF6, &H4D8BF845, &HFC551308
    pvAppendBuffer &H8301E883, &H40100DA, &HFC458BF1, &H4F15411, &HFF8558B, &HC110D0A4, &H142910E2, &HD94419D9, &H83DF8B04, &H847210FB, &H8B5B5E5F, &H4C25DE5, &HEC8B5500, &H5610EC83, &H570C758B, &H51B2E856, &H45890000, &H4468DF0, &H51A6E850, &H45890000, &H8468DF4, &H519AE850, &H45890000, &HC468DF8, &H518EE850, &H4D8B0000, &HFC458908, &H798D318B, &HC1C68B04, &HF80304E0, &H57F0458D, &HF6C3E850
    pvAppendBuffer &H21EBFFFF, &H3B6AE8, &HF0458D00, &H3BF7E850, &H8D570000, &HE850F045, &HFFFFF6A9, &H50F0458D, &H3B18E8, &H10EF8300, &H50F0458D, &H7501EE83, &H3B3DE8D3, &H458D0000, &HCAE850F0, &H5700003B, &H50F0458D, &HFFF67CE8, &H10758BFF, &HF075FF56, &H6B72E8, &H4468D00, &HF475FF50, &H6B66E8, &H8468D00, &HF875FF50, &H6B5AE8, &HC468D00, &HFC75FF50, &H6B4EE8, &H8B5E5F00, &HCC25DE5
    pvAppendBuffer &HEC8B5500, &H5310EC83, &HC758B56, &HE1E85657, &H89000050, &H468DF045, &HD5E85004, &H89000050, &H468DF445, &HC9E85008, &H89000050, &H468DF845, &HBDE8500C, &H8B000050, &H4589085D, &HF0458DFC, &H5604738D, &HF5FBE850, &HFF33FFFF, &H4710C683, &H2D763B39, &H50F0458D, &H5ABDE8, &HF0458D00, &H59B3E850, &H458D0000, &H13E850F0, &H5600003C, &H50F0458D, &HFFF5CCE8, &H10C683FF, &H723B3B47
    pvAppendBuffer &HF0458DD3, &H5A90E850, &H458D0000, &H86E850F0, &H56000059, &H50F0458D, &HFFF5A8E8, &H10758BFF, &HF075FF56, &H6A9EE8, &H4468D00, &HF475FF50, &H6A92E8, &H8468D00, &HF875FF50, &H6A86E8, &HC468D00, &HFC75FF50, &H6A7AE8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H575653EC, &H8B1C75FF, &H75FF107D, &H57E85718, &H85000004, &HE81F74C0, &HFFFFE718, &H475BA0BE, &H4000BB00, &HF32B0047
    pvAppendBuffer &H5E8F003, &HB9FFFFE7, &H475AAC, &H75FF2CEB, &H147D8B1C, &H571875FF, &H3CE8&, &HE6EAE800, &H1BEFFFF, &HBB004759, &H474000, &HF003F32B, &HFFE6D7E8, &H5831B9FF, &HCB2B0047, &H4D8BC103, &H8418908, &H890C458B, &H1C70471, &H10&, &H5E5F3889, &H18C25D5B, &HEC8B5500, &H8758B56, &HF468&, &H56006A00, &H3AC0E8, &HCC48300, &H10107D83, &H7D832074, &H10741810, &H20107D83
    pvAppendBuffer &H6C72575, &HE&, &H12EB206A, &HC06C7, &H186A0000, &H6C708EB, &HA&, &H75FF106A, &HC1E8560C, &H5EFFFFF4, &HCC25D, &H51DC8B53, &HF0E48351, &H5504C483, &H89046B8B, &H8B04246C, &H2CEC83EC, &H8B084B8B, &H6A560C43, &H918B10, &HF000002, &H18B3810, &H2280F5E, &HF66D603, &H290FF8EF, &HE883F07D, &H483C740A, &H7401E883, &HE883481E, &H95850F01, &HF000000, &HD6030A28
    pvAppendBuffer &HDE380F66, &H2280FF9, &HF66D603, &HFF8DE38, &HD6030A28, &HDE380F66, &H2280FF9, &HF66D603, &HFF8DE38, &HFF07D29, &HD6033A28, &H332280F, &H75290FD6, &H75280FE0, &H380F66F0, &H280FF7DE, &HFD6032A, &HFF07529, &H66F07D28, &H7DDE380F, &H22280FE0, &HF66D603, &H66FDDE38, &HFCDE380F, &H31A280F, &H380F66D6, &H280FFBDE, &H66D60312, &HFADE380F, &H30A280F, &H380F66D6, &HF66F9DE
    pvAppendBuffer &H663ADE38, &H3CDE380F, &H380F6632, &H10327CDF, &H5E10438B, &H8B38110F, &HE38B5DE5, &HCC25B, &H51DC8B53, &HF0E48351, &H5504C483, &H89046B8B, &H8B04246C, &H2CEC83EC, &H8B084B8B, &H6A560C43, &HFC918B10, &HF000001, &H18B3810, &H2280F5E, &HF66D603, &H290FF8EF, &HE883F07D, &H483C740A, &H7401E883, &HE883481E, &H95850F01, &HF000000, &HD6030A28, &HDC380F66, &H2280FF9, &HF66D603
    pvAppendBuffer &HFF8DC38, &HD6030A28, &HDC380F66, &H2280FF9, &HF66D603, &HFF8DC38, &HFF07D29, &HD6033A28, &H332280F, &H75290FD6, &H75280FE0, &H380F66F0, &H280FF7DC, &HFD6032A, &HFF07529, &H66F07D28, &H7DDC380F, &H22280FE0, &HF66D603, &H66FDDC38, &HFCDC380F, &H31A280F, &H380F66D6, &H280FFBDC, &H66D60312, &HFADC380F, &H30A280F, &H380F66D6, &HF66F9DC, &H663ADC38, &H3CDC380F, &H380F6632
    pvAppendBuffer &H10327CDD, &H5E10438B, &H8B38110F, &HE38B5DE5, &HCC25B, &H81EC8B55, &H10CEC, &HC458B00, &H5756C933, &H1F445C7, &H8B080402, &H6788DF1, &H10F845C7, &H8D804020, &H4BD14, &HC7660000, &H361BFC45, &H89E47D89, &HD285E855, &HE7840F, &H7D8B0000, &HF4558D0C, &H8DF05589, &HFFFEF495, &H2E0C1FF, &H458BD02B, &HEC558908, &HFEF4958D, &HC22BFFFF, &H3B084589, &H8D1B73F7, &H8D8DB004
    pvAppendBuffer &HFFFFFEF4, &HE850C103, &H4D10&, &H33EC558B, &H86E9C9, &H848B0000, &HFFFEF0B5, &H89D233FF, &HC68B0C45, &HFF83F7F7, &H8B0A7508, &H3C0724C6, &H7401B004, &H85C18A02, &H663275D2, &HC4D6E0F, &H570FC033, &H6E0F66D2, &HF0458BC0, &HC8620F66, &HD1620F66, &HDF3A0F66, &HB60F00C2, &H3A0F6600, &H3301C216, &HF045FFD0, &HEB0C5589, &H74C08427, &H6E0F6623, &HC0330C4D, &H66D2570F, &H66C06E0F
    pvAppendBuffer &H66C8620F, &H66D1620F, &HC2DF3A0F, &H3A0F6600, &HC4516, &H8BEC558B, &HC453302, &HF4B58489, &H83FFFFFE, &H458B04C2, &H55894608, &HE8753BEC, &HFF41820F, &H7D8BFFFF, &H10558BE4, &HFEF8858D, &HF08BFFFF, &H778DD62B, &H8758901, &H8310758B, &H6601086D, &H8586E0F, &H406E0F66, &H6E0F6604, &H6E0F6608, &HF66FC50, &HF66D062, &HF66CB62, &H290FD162, &H408D0214, &H68D47510, &HF0&
    pvAppendBuffer &HF4858D51, &H50FFFFFE, &H3720E8, &H83CF8B00, &HE1C10CC4, &H144D0304, &H280FD233, &H74D28506, &H2BC78B0B, &H660574C2, &HC0DB380F, &H4201290F, &H8310C683, &HD73B10E9, &H5E5FE076, &HC25DE58B, &H8B550010, &H10EC83EC, &H33575653, &HF07D8DC0, &H53C93340, &HF38BA20F, &H8907895B, &H4F890477, &HC578908, &HF845F7, &H74020000, &HF845F75F, &H80000, &H7D8B5674, &H8BC78B10, &HCE8B0875
    pvAppendBuffer &H83D1F799, &HEFC103E2, &HC1C20302, &HC08302F8, &H8D068906, &HD0F70446, &H2E8C140, &H4003E083, &H2E9C141, &H8303E183, &H148D40C1, &HFC968986, &H8D000001, &H52518E0C, &HC75FF57, &H2008E89, &HCDE80000, &H33FFFFFD, &H2EB40C0, &H5E5FC033, &H5DE58B5B, &H55000CC2, &HEC81EC8B, &H228&, &H8D1C75FF, &HFFFDD885, &H1875FFFF, &H458D5050, &H458D501C, &HDBE850F4, &HFFFFFFFA, &H458D1475
    pvAppendBuffer &H1C75FFF4, &HDC458D50, &H42FE850, &H458B0000, &H4E8C110, &H875FF50, &HFFDC458D, &HE8500C75, &H339&, &HC25DE58B, &H8B550018, &H28EC81EC, &HFF000002, &H858D1C75, &HFFFFFDD8, &H501875FF, &H1C458D50, &HF4458D50, &HFA88E850, &H75FFFFFF, &HF4458D14, &H501C75FF, &H50DC458D, &H3DCE8, &H10458B00, &H5004E8C1, &H8D0875FF, &H75FFDC45, &H57E8500C, &H8B000003, &H18C25DE5, &HEC8B5500
    pvAppendBuffer &H210EC81, &H75FF0000, &HF0858D28, &HFFFFFFFD, &H50502475, &H5028458D, &H50F4458D, &HFFFA35E8, &H875FFFF, &H6AF4458D, &H1475FF10, &H75FF0C6A, &H1C75FF20, &HFF1875FF, &H75FF1075, &H2875FF0C, &HD40E850, &HE58B0000, &H24C25D, &H81EC8B55, &H210EC, &H2875FF00, &HFDF0858D, &H75FFFFFF, &H8D505024, &H8D502845, &HE850F445, &HFFFFF9E6, &H75FF106A, &HF4458D0C, &H6A0875FF, &H2075FF0C
    pvAppendBuffer &HFF1C75FF, &H75FF1875, &H1075FF14, &H502875FF, &HE9FE8, &H5DE58B00, &H550024C2, &H5751EC8B, &H33187D8B, &HFC5589D2, &H6374FF85, &HC5D8B53, &H8B0B8B56, &HF12B1075, &H3B184D89, &H8B0273FE, &H75D285F7, &H45B60F1A, &H8B505614, &HC1030845, &H34D7E850, &H4D8B0000, &HCC48318, &H85FC558B, &H3B0B75C9, &H6751075, &H8942D233, &H48DFC55, &H10453B0E, &H75FF1175, &H2075FF08, &H831C55FF
    pvAppendBuffer &H558B0023, &H102EBFC, &H75FE2B33, &H5F5B5EA4, &HC25DE58B, &H8B55001C, &H758B56EC, &H83C68B20, &H840F00E8, &H86&, &HFF2875FF, &H16A2475, &H7401E883, &H1E88361, &H5014458D, &HFF1075FF, &H75FF0C75, &HE8487408, &H67&, &H8B2875FF, &H75FF184D, &H1C4D3824, &H468D2074, &HFF5150FE, &H75FF1075, &H875FF0C, &HFFFF2EE8, &H2875FFFF, &HFF1C458D, &H16A2475, &H468D25EB, &HFF5150FF
    pvAppendBuffer &H75FF1075, &H875FF0C, &HFFFF0EE8, &HE81FEBFF, &H1F&, &H458AD7EB, &H1445301C, &H5014458D, &HFF1075FF, &H75FF0C75, &H5E808, &H5D5E0000, &H550024C2, &H75FFEC8B, &H1C75FF20, &HFF1C75FF, &H75FF1875, &H1075FF14, &HFF0C75FF, &H4E80875, &H5D000000, &H55001CC2, &H458BEC8B, &H5D8B530C, &H38835614, &H18758B00, &H107D8B57, &HF6856274, &HCF8B5E74, &H4D89082B, &H72CE3B18, &H89CE8B05
    pvAppendBuffer &H8B1875, &H51084503, &H79E85053, &H8B000033, &HC4830C45, &H184D8B0C, &HF12BD903, &H38390801, &H75FF2E75, &H2475FF08, &H575F685, &HEB2055FF, &H1C55FF03, &H830C458B, &H14EB0020, &H2475FF53, &H575F73B, &HEB2055FF, &H1C55FF03, &HF72BDF03, &HE873F73B, &H2974F685, &H8B0C458B, &H3BF82B00, &H8B0272FE, &H84503FE, &HE8505357, &H3318&, &H30C458B, &HCC483DF, &HF72B3801, &H75107D8B
    pvAppendBuffer &H5B5E5FDA, &H20C25D, &H56EC8B55, &H851C758B, &H534574F6, &H570C5D8B, &H75003B83, &H875FF10, &HFF2475FF, &H458B2055, &HEB038910, &H10458B03, &HC72B3B8B, &H272FE3B, &H4503FE8B, &HFF505708, &H75FF1475, &H62D4E818, &H3B290000, &H1187D01, &HF72B147D, &H5B5FC275, &H20C25D5E, &HEC8B5500, &H5310EC83, &H56145D8B, &H8B08758B, &H89008B06, &HDB850845, &H8B575074, &H7D290C7D, &H8D068B10
    pvAppendBuffer &H5751F04D, &HFF0476FF, &H75FF0850, &H8468D08, &HF0458D50, &H83E85050, &HFF000062, &H468D0875, &HE8505708, &H3260&, &H8D0875FF, &H8B50F045, &HC7031045, &H324EE850, &H7D030000, &H18C48308, &H7501EB83, &H5B5E5FB8, &HC25DE58B, &H8B550010, &H10EC83EC, &H8758B56, &H147D8B57, &H88B068B, &H85084D89, &H8B4E74FF, &H568D0C45, &H5D8B5308, &H89C32B10, &H52510C45, &H8D50C303, &HE850F045
    pvAppendBuffer &H621A&, &H4E8D068B, &H4D8D5108, &H76FF51F0, &H450FF04, &H8D0875FF, &H53500846, &H31E7E8, &H84D8B00, &H8B08568D, &HC4830C45, &H83D9030C, &HC27501EF, &H8B5E5F5B, &H10C25DE5, &HEC8B5500, &H8B08558B, &H458B0C4D, &H4428910, &H8908428D, &HFF31FF0A, &HE8501475, &H31AC&, &H5D0CC483, &H550010C2, &HA1E8EC8B, &HB9FFFFDD, &H4768B3, &H4000E981, &HC1030047, &H51084D8B, &H1475FF50
    pvAppendBuffer &HFF74418D, &H75FF1075, &H50406A0C, &H5034418D, &HFFFE73E8, &H10C25DFF, &HEC8B5500, &H5370EC83, &H14758B56, &H71E85657, &H8D000046, &HF88B044E, &HD07D8951, &H4663E8, &H84E8D00, &H51F84589, &HE8CC4589, &H4654&, &H890C4E8D, &H8951F445, &H45E8C845, &H8B000046, &H89560875, &H4589F045, &H4636E8C4, &H4E8D0000, &HD4458904, &HC0458951, &H4627E8, &H84E8D00, &H51E84589, &HE8BC4589
    pvAppendBuffer &H4618&, &H890C4E8D, &H8951E445, &H9E8B845, &H8B000046, &H89560C75, &H4589E045, &H45FAE8B4, &H4E8D0000, &HFC458904, &HB0458951, &H45EBE8, &H84E8D00, &H51084589, &HE8AC4589, &H45DC&, &H890C4E8D, &H89511445, &HCDE8A845, &H8B000045, &H89561075, &H45890C45, &H45BEE8A4, &H4E8D0000, &H51D88B04, &HE8A05D89, &H45B0&, &H89084E8D, &H8951DC45, &HA1E89C45, &H8D000045, &H45890C4E
    pvAppendBuffer &H458951D8, &H4592E898, &H4D8B0000, &H8BF08BDC, &H7589D855, &H9045C794, &HA&, &H758B06EB, &H107D8BEC, &H8BD47D03, &HDF33FC45, &H310C3C1, &HFC4589C3, &HC1D44533, &HF8030CC0, &H7D89DF33, &HFC7D8B10, &H308C3C1, &HFC7D89FB, &H458BF833, &HE84503F8, &H4589C833, &H8458BF8, &H310C1C1, &H7C7C1C1, &H33084589, &HC0C1E845, &HF845010C, &HC1F84D33, &H4D0108C1, &HDC4D8908, &H33084D8B
    pvAppendBuffer &HF4458BC8, &H33E44503, &HF44589D0, &HC114458B, &HC20310C2, &H8907C1C1, &H45331445, &HCC0C1E4, &H33F44501, &HC2C1F455, &H14550108, &H89104D01, &H558BD855, &H8BD03314, &H4503F045, &H89F033E0, &H458BF045, &H10C6C10C, &HC2C1C603, &HC458907, &HC1E04533, &H45010CC0, &HF07533F0, &H108C6C1, &H75890C75, &HC758BEC, &H458BF033, &H104533EC, &H110C0C1, &H45891445, &H14458BEC, &HC6C1C133
    pvAppendBuffer &HCC0C107, &H8B104501, &H4D33EC4D, &H8C1C110, &H89144D01, &H4D8BEC4D, &H8BC83314, &HC1C10C45, &HE84D8907, &H3F84D8B, &HC1D933CA, &HC30310C3, &H330C4589, &HF4558BC2, &H30CC0C1, &H33C803D6, &HF84D89D9, &HC10C4D8B, &HCB0308C3, &H330C4D89, &HFC458BC8, &H8907C1C1, &H4D8BE44D, &HC1CA33DC, &HC10310C1, &H33FC4589, &HF0758BC6, &H30CC0C1, &H33D003F7, &HF45589CA, &HC1FC558B, &HD10308C1
    pvAppendBuffer &H33FC5589, &H8458BD0, &H8907C2C1, &H558BE055, &HC1D633D8, &HC20310C2, &H33084589, &H87D8BC7, &H30CC0C1, &H89D633F0, &HC2C1F075, &H89FA0308, &HF833087D, &H8307C7C1, &H8901906D, &H850FD47D, &HFFFFFE5A, &H3D0458B, &H45891045, &HCC458BD0, &H89F84503, &H458BCC45, &HF44503C8, &H8BC84589, &H4503BC45, &HBC4589E8, &H3B8458B, &H4589E445, &HB4458BB8, &H89E04503, &H458BB445, &H84503AC
    pvAppendBuffer &H89A05D01, &H458BAC45, &H144503A8, &H8B985501, &H5D8B1855, &HA84589D0, &H3A4458B, &H45890C45, &H94458BA4, &H88EC4503, &H9445891A, &HE8C1C38B, &H1428808, &HE8C1C38B, &H2428810, &HC1C47501, &H5A8818EB, &HCC5D8B03, &H5A88C38B, &H8E8C104, &H8B054288, &HC07D01C3, &H8810E8C1, &H7D8B0642, &HB07D01FC, &HC19C4D01, &H5A8818EB, &HC85D8B07, &H5A88C38B, &H8E8C108, &H8B094288, &H10E8C1C3
    pvAppendBuffer &HC10A4288, &H5A8818EB, &HC45D8B0B, &H5A88C38B, &H8E8C10C, &H8B0D4288, &H10E8C1C3, &HC10E4288, &H5A8818EB, &HC05D8B0F, &H5A88C38B, &H8E8C110, &H8B114288, &H10E8C1C3, &HC1124288, &H5A8818EB, &HBC5D8B13, &H5A88C38B, &H8E8C114, &H8B154288, &H10E8C1C3, &HC1164288, &H5A8818EB, &HB85D8B17, &H5A88C38B, &H8E8C118, &H8B194288, &H10E8C1C3, &HC11A4288, &H5A8818EB, &HB45D8B1B, &H5A88C38B
    pvAppendBuffer &H8E8C11C, &H8B1D4288, &H10E8C1C3, &HC11E4288, &H5A8818EB, &HB05D8B1F, &H5A88C38B, &H8E8C120, &H8B214288, &H10E8C1C3, &HC1224288, &H5A8818EB, &HAC5D8B23, &H5A88C38B, &H8E8C124, &H8B254288, &H10E8C1C3, &HC1264288, &H5A8818EB, &HA85D8B27, &H5A88C38B, &H8E8C128, &H8B294288, &H10E8C1C3, &HC12A4288, &H5A8818EB, &HA45D8B2B, &H5A88C38B, &H8E8C12C, &H8B2D4288, &H10E8C1C3, &HC12E4288
    pvAppendBuffer &H5A8818EB, &HA05D8B2F, &H5A88C38B, &H8E8C130, &H8D314288, &HC38B3C4A, &HC118EBC1, &H428810E8, &H335A8832, &H8B9C5D8B, &H345A88C3, &H8808E8C1, &HC38B3542, &H8810E8C1, &HEBC13642, &H375A8818, &H8B985D8B, &H385A88C3, &H8808E8C1, &HC38B3942, &H8810E8C1, &HEBC13A42, &H3B5A8818, &H8B94558B, &H8E8C1C2, &H41881188, &H5FC28B01, &HC110E8C1, &H885E18EA, &H51880241, &HE58B5B03, &H14C25D
    pvAppendBuffer &H56EC8B55, &H8B1075FF, &H75FF0875, &H22E8560C, &H6A000045, &H1475FF10, &H5020468D, &H2C67E8, &H18458B00, &H830CC483, &H89007466, &H5D5E7846, &H550014C2, &H8B56EC8B, &HFF570875, &H76FF0C75, &H207E8D30, &H10468D57, &HCAE85650, &H8BFFFFFA, &HC9337856, &H75010780, &HCA3B410B, &H4800674, &HF5740139, &HC25D5E5F, &H8B550008, &H10EC83EC, &H6AF0458D, &H2075FF10, &H2C0AE850, &HC4830000
    pvAppendBuffer &HF0458D0C, &HFF006A50, &H75FF2475, &H1875FF1C, &HFF1475FF, &H75FF1075, &H875FF0C, &H3ECBE8, &H5DE58B00, &H550020C2, &H75FFEC8B, &HFF016A24, &H75FF2075, &H1875FF1C, &HFF1475FF, &H75FF1075, &H875FF0C, &H3EA3E8, &H20C25D00, &HEC8B5500, &HFFD7B7E8, &H7886B9FF, &HE9810047, &H474000, &H4D8BC103, &HFF505108, &H18B1475, &HFF1075FF, &H30FF0C75, &H5028418D, &H5018418D, &HFFF887E8
    pvAppendBuffer &H10C25DFF, &HEC8B5500, &H8B084D8B, &H41890C45, &H10458B2C, &H5D304189, &H55000CC2, &H8B56EC8B, &H346A0875, &HE856006A, &H2B81&, &H830C4D8B, &H8B002C66, &H28668301, &H30468900, &H8910458B, &H468D0446, &HFF0E8908, &H1475FF31, &H2B36E850, &HC4830000, &HC25D5E18, &H8B550010, &H20EC81EC, &H53000004, &HDB335756, &HFD6085C7, &HDB41FFFF, &H706A0000, &HFD70858D, &H9D89FFFF, &HFFFFFD64
    pvAppendBuffer &H85C75053, &HFFFFFD68, &H1&, &HFD6C9D89, &H16E8FFFF, &H8B00002B, &H858D0C75, &HFFFFFF60, &H50561F6A, &H2ADFE8, &H1F468A00, &H8018C483, &HFFFF60A5, &H3F24F8FF, &H8588400C, &HFFFFFF7F, &HFBE0858D, &H75FFFFFF, &H70E85010, &H6A00004A, &H1E6A591E, &HFE609D89, &HB58DFFFF, &HFFFFFE60, &HFE649D89, &HBD8DFFFF, &HFFFFFE68, &H8959A5F3, &H758D805D, &H845D8980, &HF3887D8D, &H591E6AA5
    pvAppendBuffer &HFEE09D89, &HB58DFFFF, &HFFFFFEE0, &HFEE49D89, &HBD8DFFFF, &HFFFFFEE8, &H206AA5F3, &HE0B58D59, &H89FFFFFB, &HFFFE649D, &HE0BD8DFF, &H89FFFFFD, &HA5F3845D, &HFEBBFF33, &H47000000, &HFE60BD89, &H7D89FFFF, &HFC38B80, &HF8C1CBB6, &H7E18303, &H5B4B60F, &HFFFFFF60, &HFDE0858D, &HEED3FFFF, &H5056F723, &H5080458D, &H4244E8, &H858D5600, &HFFFFFE60, &HE0858D50, &H50FFFFFE, &H4230E8
    pvAppendBuffer &HE0858D00, &H50FFFFFE, &H5080458D, &HFCE0858D, &HE850FFFF, &HFFFFE3AA, &HFEE0858D, &H8D50FFFF, &H50508045, &H48B9E8, &H60858D00, &H50FFFFFE, &HFDE0858D, &H8D50FFFF, &HFFFEE085, &H7FE850FF, &H8DFFFFE3, &HFFFE6085, &H858D50FF, &HFFFFFDE0, &H8BE85050, &H8D000048, &HFFFCE085, &H858D50FF, &HFFFFFE60, &H4863E850, &H458D0000, &H858D5080, &HFFFFFC60, &H4853E850, &H458D0000, &H858D5080
    pvAppendBuffer &HFFFFFEE0, &H80458D50, &H364EE850, &H858D0000, &HFFFFFCE0, &HE0858D50, &H50FFFFFD, &HFEE0858D, &HE850FFFF, &H3634&, &HFEE0858D, &H8D50FFFF, &H8D508045, &HFFFCE085, &H3E850FF, &H8DFFFFE3, &HFFFEE085, &H458D50FF, &HE8505080, &H4812&, &H5080458D, &HFDE0858D, &HE850FFFF, &H47ED&, &HFC60858D, &H8D50FFFF, &HFFFE6085, &H858D50FF, &HFFFFFEE0, &H47E8E850, &H858D0000, &HFFFFFD60
    pvAppendBuffer &HE0858D50, &H50FFFFFE, &H5080458D, &H35CBE8, &H60858D00, &H50FFFFFE, &H5080458D, &HE2A0E850, &H458DFFFF, &H858D5080, &HFFFFFEE0, &HA9E85050, &H8D000035, &HFFFC6085, &H858D50FF, &HFFFFFE60, &H80458D50, &H3592E850, &H858D0000, &HFFFFFBE0, &HE0858D50, &H50FFFFFD, &HFE60858D, &HE850FFFF, &H3578&, &HFCE0858D, &H8D50FFFF, &HFFFDE085, &H56E850FF, &H56000047, &HFDE0858D, &H8D50FFFF
    pvAppendBuffer &HE8508045, &H40A9&, &H60858D56, &H50FFFFFE, &HFEE0858D, &HE850FFFF, &H4095&, &HF01EB83, &HFFFE2089, &HE0858DFF, &H50FFFFFE, &H266FE850, &H858D0000, &HFFFFFEE0, &H80458D50, &H19E85050, &H8D000035, &HFF508045, &H8AE80875, &H5F000037, &HE58B5B5E, &HCC25D, &H83EC8B55, &H6A5720EC, &HC0335907, &H9E045C6, &HF3E17D8D, &HAAAB66AB, &H50E0458D, &HFF0C75FF, &HDBE80875, &H5FFFFFFC
    pvAppendBuffer &HC25DE58B, &H8B550008, &H14EC81EC, &H53000001, &HC0335756, &H3308758B, &HE17D8DDB, &HABE05D88, &HAB66ABAB, &HE0458DAA, &H75FF5050, &H456FF0C, &HC247D83, &HC6A1E75, &H8D2075FF, &HE850F045, &H27C4&, &H660CC483, &H88FD5D89, &H45C6FC5D, &H32EB01FF, &H50E0458D, &HFEEC858D, &HE850FFFF, &H1E1A&, &H8D2475FF, &HFFFEEC85, &H2075FFFF, &H1CF1E850, &H458D0000, &H858D50F0, &HFFFFFEEC
    pvAppendBuffer &H1D7EE850, &H458D0000, &H858D50E0, &HFFFFFF3C, &H1DE8E850, &H75FF0000, &H3C858D1C, &HFFFFFFFF, &HE8501875, &H1CA1&, &H5D88C033, &HD17D8DD0, &H66ABABAB, &H458DAAAB, &H75FF50F0, &H8C458D0C, &HD4E85056, &H6AFFFFFB, &H8D0C6A04, &HE8508C45, &HFFFFFBB1, &H458D106A, &H8D5050D0, &HE8508C45, &HFFFFFB69, &H8D1475FF, &HFFFF3C85, &H1075FFFF, &H1C71E850, &H458D0000, &H858D50C0, &HFFFFFF3C
    pvAppendBuffer &H1CFEE850, &H758B0000, &HD0458D2C, &H458D5056, &HE85050C0, &H5706&, &H85C04D8D, &H8B1874F6, &HC18B2855, &H48AD02B, &HA01320A, &HEE8341D8, &H84F37501, &HFF1675DB, &H458D1475, &H3075FF8C, &H501075FF, &HFFFB04E8, &HEBF633FF, &H46F63303, &H7D8DC033, &H506AABE0, &HABAB006A, &HE0458AAB, &H33F07D8D, &HABABABC0, &HF0458AAB, &H33D07D8D, &HABABABC0, &HD0458AAB, &H33C07D8D, &HABABABC0
    pvAppendBuffer &HC0458AAB, &HFF3C858D, &HE850FFFF, &H269D&, &HFF3C8D8A, &H458DFFFF, &H6A346A8C, &H8AE85000, &H8A000026, &HC4838C4D, &H5FC68B18, &HE58B5B5E, &H2CC25D, &H81EC8B55, &H114EC, &H57565300, &H758BC033, &H8DDB3308, &H5D88E17D, &HABABABE0, &H8DAAAB66, &H5050E045, &HFF0C75FF, &H7D830456, &H1E750C24, &H75FF0C6A, &HF0458D20, &H2616E850, &HC4830000, &H5D89660C, &HFC5D88FD, &H1FF45C6
    pvAppendBuffer &H458D32EB, &H858D50E0, &HFFFFFEEC, &H1C6CE850, &H75FF0000, &HEC858D24, &HFFFFFFFE, &HE8502075, &H1B43&, &H50F0458D, &HFEEC858D, &HE850FFFF, &H1BD0&, &H50E0458D, &HFF3C858D, &HE850FFFF, &H1C3A&, &H8D1C75FF, &HFFFF3C85, &H1875FFFF, &H1AF3E850, &HC0330000, &H8DD05D88, &HABABD17D, &HAAAB66AB, &H50F0458D, &H8D0C75FF, &H50568C45, &HFFFA26E8, &H6A046AFF, &H8C458D0C, &HFA03E850
    pvAppendBuffer &H106AFFFF, &H50D0458D, &H8C458D50, &HF9BBE850, &H75FFFFFF, &H8C458D14, &HFF2875FF, &HE8501075, &HFFFFF9A9, &H8D1475FF, &HFFFF3C85, &H2875FFFF, &H1AB1E850, &HC0330000, &H8DC05D88, &HABABC17D, &HAAAB66AB, &H50C0458D, &HFF3C858D, &HE850FFFF, &H1B30&, &H8D3075FF, &H8D50D045, &HFF50C045, &H37E82C75, &H33000055, &HE07D8DC0, &HABABABAB, &H8DE0458A, &HC033F07D, &HABABABAB, &H8DF0458A
    pvAppendBuffer &HC033D07D, &H53506AAB, &H8AABABAB, &H7D8DD045, &HABC033C0, &H8AABABAB, &H858DC045, &HFFFFFF3C, &H2507E850, &H858A0000, &HFFFFFF3C, &H458D346A, &HE850538C, &H24F5&, &H838C458A, &H5E5F18C4, &H5DE58B5B, &H55002CC2, &H558BEC8B, &H104D8B0C, &H8758B56, &H233068B, &H468B0189, &H4423304, &H8B044189, &H42330846, &H8418908, &H330C468B, &H41890C42, &HC25D5E0C, &H8B55000C, &H8B5351EC
    pvAppendBuffer &H57560C5D, &H66087D8B, &HFC45C7, &H8B0F8BE1, &H83E8D1C1, &H38901E1, &H8B04578B, &H83E8D1C2, &HE1C101E2, &HC1C80B1F, &H4B891FE2, &H8778B04, &HE8D1C68B, &HB01E683, &H1FE6C1D0, &H8B085389, &HC18B0C4F, &HE183E8D1, &H5FF00B01, &HF0C7389, &HFC0D44B6, &H3118E0C1, &H8B5B5E03, &H8C25DE5, &HEC8B5500, &H8758B56, &H5E85657, &H8B000039, &H7890C7D, &H5004468D, &H38F7E8, &H4478900
    pvAppendBuffer &H5008468D, &H38EBE8, &H8478900, &H500C468D, &H38DFE8, &HC478900, &HC25D5E5F, &H8B550008, &H20EC83EC, &H56E0458D, &H6A106A57, &HFAE85000, &H6A000023, &HC75FF10, &H50F0458D, &H23C7E8, &H87D8B00, &H3318C483, &H83C68BF6, &H1F6A1FE0, &H8BC82B59, &H5F8C1C6, &HD387048B, &H7401A8E8, &H8BC93312, &H31F00D44, &H83E00D44, &HF98304C1, &H8DF07210, &H5050F045, &HFFFEFDE8, &HFE8146FF
    pvAppendBuffer &H80&, &H106AC37C, &H50E0458D, &HE81075FF, &H2374&, &H5F0CC483, &H5DE58B5E, &H55000CC2, &H458BEC8B, &H100F08, &H660C458B, &H1BE8700F, &HFD5280F, &H100FDD28, &H10458B00, &HC8700F66, &HC5280F1B, &H443A0F66, &HF6601C1, &H10D1443A, &HD0EF0F66, &H443A0F66, &HF6600D9, &H11E9443A, &H66C2280F, &H8DA730F, &HF8730F66, &HEF0F6608, &HEF0F66EA, &HE5280FD8, &H66C3280F, &H1FD4720F
    pvAppendBuffer &HD0720F66, &H720F661F, &H280F01F3, &H730F66C8, &HF6604FC, &H660CD873, &H4F9730F, &HCBEB0F66, &HF5720F66, &HD9280F01, &HE5EB0F66, &HF3720F66, &HEB0F661F, &HC1280FE0, &HF0720F66, &HEF0F661E, &HC1280FD8, &HF0720F66, &HEF0F6619, &HD3280FD8, &HDB730F66, &H730F6604, &HF660CFA, &H280FD1EF, &HC2280FCA, &HD1720F66, &H720F6602, &HF6601D0, &H280FC8EF, &H720F66C2, &HF6607D0, &HF66C8EF
    pvAppendBuffer &HF66CBEF, &HF66CAEF, &HF66CCEF, &HF1BC170, &HC25D0011, &H8B55000C, &H758B56EC, &H7D8B570C, &H37FF5608, &H519EE8, &H4468D00, &H477FF50, &H5192E8, &H8468D00, &H877FF50, &H5186E8, &HC468D00, &HC77FF50, &H517AE8, &H5D5E5F00, &H550008C2, &HEC83EC8B, &H758B5644, &HA8BE8308, &H0&, &HE8560674, &H3395&, &HB60FC933, &H880E84, &H44890000, &H8341BC8D, &HEE7210F9
    pvAppendBuffer &HFC6583, &H3302E856, &H458D0000, &HE85650BC, &H32AA&, &H330C558B, &H8E048AC9, &H41110488, &H7210F983, &HAC68F4, &H6A0000, &H21FBE856, &H68A0000, &H5E0CC483, &HC25DE58B, &H8B550008, &H758B56EC, &HAC6808, &H6A0000, &H21DBE856, &H4D8B0000, &HFCBA0C, &H106A0000, &HF1075FF, &H468901B6, &H41B60F44, &H48468901, &H241B60F, &HF4C4689, &H830341B6, &H46890FE0, &H41B60F50
    pvAppendBuffer &H89C22304, &HB60F5446, &H46890541, &H41B60F58, &H5C468906, &H741B60F, &H890FE083, &HB60F6046, &HC2230841, &HF644689, &H890941B6, &HB60F6846, &H46890A41, &H41B60F6C, &HFE0830B, &HF704689, &H230C41B6, &H744689C2, &HD41B60F, &HF784689, &H890E41B6, &HB60F7C46, &HA6830F41, &H84&, &HFE08300, &H808689, &H868D0000, &H88&, &H2112E850, &HC4830000, &HC25D5E18, &H8B55000C
    pvAppendBuffer &HCD06E8EC, &HBAB9FFFF, &H810047A7, &H474000E9, &H8BC10300, &H5051084D, &H8D1075FF, &HA881&, &HC75FF00, &H8D50106A, &H9881&, &HE85000, &H5DFFFFED, &H55000CC2, &HEC83EC8B, &H5756530C, &HFFCCC7E8, &H84D8BFF, &H47AE28BE, &H5A406A00, &H4000EE81, &HF0030047, &H8B64798D, &HE2F76041, &H703006A, &H515ADA8B, &H13F08B56, &H58406ADA, &H83084E8D, &HC12B3FE1, &H68525250, &H80&
    pvAppendBuffer &H8B57406A, &H478D087D, &H7E85020, &H8DFFFFEC, &HF50F445, &H5303F3A4, &H5603E6C1, &H4FCDE8, &H8D086A00, &H5750F445, &HC5E8&, &HC758B00, &HE837FF56, &H4F8F&, &H5004468D, &HE80477FF, &H4F83&, &H5008468D, &HE80877FF, &H4F77&, &H500C468D, &HE80C77FF, &H4F6B&, &H5010468D, &HE81077FF, &H4F5F&, &H5014468D, &HE81477FF, &H4F53&, &H5018468D, &HE81877FF, &H4F47&
    pvAppendBuffer &H501C468D, &HE81C77FF, &H4F3B&, &H6A686A, &H2017E857, &HC4830000, &H5B5E5F0C, &HC25DE58B, &H8B550008, &H758B56EC, &H6A686A08, &HFAE85600, &H8300001F, &H6C70CC4, &H6A09E667, &H850446C7, &HC7BB67AE, &HF3720846, &H46C73C6E, &H4FF53A0C, &H1046C7A5, &H510E527F, &H8C1446C7, &HC79B0568, &HD9AB1846, &H46C71F83, &HE0CD191C, &HC25D5E5B, &H8B550004, &HCB92E8EC, &H28B9FFFF, &H810047AE
    pvAppendBuffer &H474000E9, &H8BC10300, &H5051084D, &H8D1075FF, &H75FF6441, &H50406A0C, &H5020418D, &HFFEB92E8, &HCC25DFF, &HEC8B5500, &H8D40EC83, &HFF50C045, &HA7E80875, &H6A000000, &HC0458D30, &HC75FF50, &H1F43E8, &HCC48300, &HC25DE58B, &H8B550008, &H758B56EC, &HC86808, &H6A0000, &H1F4BE856, &HC4830000, &HD806C70C, &HC7C1059E, &H9D5D0446, &H46C7CBBB, &H7CD50708, &HC46C736, &H629A292A
    pvAppendBuffer &H171046C7, &HC73070DD, &H15A1446, &H46C79159, &HE593918, &H1C46C7F7, &H152FECD8, &H312046C7, &HC7FFC00B, &H26672446, &H46C76733, &H58151128, &H2C46C768, &H8EB44A87, &HA73046C7, &HC764F98F, &H2E0D3446, &H46C7DB0C, &HFA4FA438, &H3C46C7BE, &H47B5481D, &H4C25D5E, &H1B8E900, &H8B550000, &H8B5151EC, &H80B90845, &H53000000, &HB08D5756, &HC4&, &HC0808B, &HE1F70000, &HFA8BD88B
    pvAppendBuffer &HD7831E03, &HCA82E800, &H558BFFFF, &HAFCEB908, &H81520047, &H474000E9, &H8DC10300, &H8350104B, &H80B87FE1, &H2B000000, &H6A50C1, &H80B8006A, &H50000000, &H428D5650, &HD7E85040, &H8DFFFFE9, &H6A50F845, &HE8006A00, &H4DA2&, &H458D086A, &H75FF50F8, &H13CE808, &H458D0000, &HA40F50F8, &HC15703DF, &HE85303E3, &H4D82&, &H8D085D8B, &H86AF845, &H1BE85350, &H8B000001, &HFF560C75
    pvAppendBuffer &H33FF0473, &H4D65E8, &H8468D00, &HC73FF50, &HE80873FF, &H4D56&, &H5010468D, &HFF1473FF, &H47E81073, &H8D00004D, &HFF501846, &H73FF1C73, &H4D38E818, &H468D0000, &H73FF5020, &H2073FF24, &H4D29E8, &H28468D00, &H2C73FF50, &HE82873FF, &H4D1A&, &H5030468D, &HFF3473FF, &HBE83073, &H8D00004D, &HFF503846, &H73FF3C73, &H4CFCE838, &HC8680000, &H6A000000, &HAEE85300, &H8300001D
    pvAppendBuffer &H5E5F0CC4, &H5DE58B5B, &H550008C2, &H8B56EC8B, &HC8680875, &H6A000000, &H8EE85600, &H8300001D, &H6C70CC4, &HF3BCC908, &H670446C7, &HC76A09E6, &HA73B0846, &H46C784CA, &H67AE850C, &H1046C7BB, &HFE94F82B, &H721446C7, &HC73C6EF3, &H36F11846, &H46C75F1D, &H4FF53A1C, &H2046C7A5, &HADE682D1, &H7F2446C7, &HC7510E52, &H6C1F2846, &H46C72B3E, &H5688C2C, &H3046C79B, &HFB41BD6B, &HAB3446C7
    pvAppendBuffer &HC71F83D9, &H21793846, &H46C7137E, &HE0CD193C, &HC25D5E5B, &H8B550004, &HC8EEE8EC, &HCEB9FFFF, &H810047AF, &H474000E9, &H8BC10300, &H5051084D, &H8D1075FF, &HC481&, &HC75FF00, &H8068&, &H418D5000, &HE8E85040, &H5DFFFFE8, &H55000CC2, &H5756EC8B, &HFFC8A0E8, &H87D8BFF, &HC8D0F8B, &H48D&, &H10FF5100, &HF08B0F8B, &H48D0C8D, &H51000000, &H8DE85657, &H8300001C, &HC68B0CC4
    pvAppendBuffer &HC25D5E5F, &H8B550004, &H758B56EC, &HC75FF08, &H468D0E8B, &H76FF5008, &H451FF04, &H8B2C568B, &HD603304E, &H8504EB5E, &H490874C9, &H80A4480, &H5DF47401, &H550008C2, &H4D8BEC8B, &H40C03308, &HF7E0139, &H813C83, &HC830975, &H3B40FF81, &HFFF17C01, &HC25D810C, &H8B550004, &HC458BEC, &H33575653, &H1A788DDB, &H8918C083, &HB60F0C45, &H8B99FF47, &H8BF28BC8, &H6D830C45, &HB60F080C
    pvAppendBuffer &HA40F9900, &HF20B08C2, &HB08E0C1, &H7B60FC8, &H8CEA40F, &HF87F8D99, &HB08E1C1, &HFC80BF2, &HF0947B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0A47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0B47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0C47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0D47B6, &H9908CEA4, &HB08E1C1, &H8BC80BF2, &HC890845, &HD87489D8, &HFB834304, &H6B820F04
    pvAppendBuffer &H5FFFFFFF, &HC25D5B5E, &H8B550008, &HC458BEC, &H33575653, &H2A788DDB, &H8928C083, &HB60F0C45, &H8B99FF47, &H8BF28BC8, &H6D830C45, &HB60F080C, &HA40F9900, &HF20B08C2, &HB08E0C1, &H7B60FC8, &H8CEA40F, &HF87F8D99, &HB08E1C1, &HFC80BF2, &HF0947B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0A47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0B47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2
    pvAppendBuffer &HF0C47B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF0D47B6, &H9908CEA4, &HB08E1C1, &H8BC80BF2, &HC890845, &HD87489D8, &HFB834304, &H6B820F06, &H5FFFFFFF, &HC25D5B5E, &H8B550008, &H60EC83EC, &HFFE0458D, &HE8500C75, &HFFFFFE8E, &H458D046A, &HABE850E0, &H8500003B, &H330474C0, &H567FEBC0, &H458D046A, &H8EE850E0, &HBEFFFFC6, &H90&, &HE850C603, &H3B3B&, &H7401F883, &HE8046A14
    pvAppendBuffer &HFFFFC675, &H8D50C603, &H5050E045, &H4919E8, &H8D006A00, &HE850E045, &HFFFFC65D, &H5050C083, &H50A0458D, &HFFC9F4E8, &HA0458DFF, &HC98FE850, &HC085FFFF, &HC0330474, &H758B23EB, &HA04D8D08, &H14E8D51, &H510406C6, &HD2E8&, &HC04D8D00, &H214E8D51, &HC5E851, &HC0330000, &HE58B5E40, &H8C25D, &H81EC8B55, &H90EC&, &HD0458D00, &H500C75FF, &HFFFE91E8, &H8D066AFF, &HE850D045
    pvAppendBuffer &H3AFE&, &H774C085, &H8AE9C033, &H56000000, &H458D066A, &HDEE850D0, &HBEFFFFC5, &H170&, &HE850C603, &H3A8B&, &H7401F883, &HE8066A14, &HFFFFC5C5, &H8D50C603, &H5050D045, &H4869E8, &H8D006A00, &HE850D045, &HFFFFC5AD, &H11005, &H858D5000, &HFFFFFF70, &HCAEBE850, &H858DFFFF, &HFFFFFF70, &HC905E850, &HC085FFFF, &HC0330474, &H758B26EB, &H708D8D08, &H51FFFFFF, &HC6014E8D
    pvAppendBuffer &HE8510406, &HAE&, &H51A04D8D, &H51314E8D, &HA1E8&, &H40C03300, &H5DE58B5E, &H550008C2, &H458BEC8B, &H57565308, &H8D0C7D8B, &HF6331848, &H8D084D89, &H448A1A58, &H18807F7, &H448A28B1, &H438806F7, &HF7048BFF, &H4F7548B, &HFFD2D2E8, &HB10388FF, &HF7048B20, &H4F7548B, &HFFD2C2E8, &H14388FF, &H8BF85B8D, &H448BF70C, &HAC0F04F7, &HE8C118C1, &HA4B8818, &H8BF70C8B, &HF04F744
    pvAppendBuffer &HC110C1AC, &H4B8810E8, &HF70C8B0B, &H4F7448B, &H8C1AC0F, &H8808E8C1, &H48A0C4B, &H4D8B46F7, &H8E98308, &H890D4388, &HFE83084D, &H5F877204, &HC25D5B5E, &H8B550008, &H8458BEC, &H8B575653, &H488D0C7D, &H89F63328, &H588D084D, &HF7448A2A, &HB1018807, &HF7448A28, &HFF438806, &H8BF7048B, &HE804F754, &HFFFFD23B, &H20B10388, &H8BF7048B, &HE804F754, &HFFFFD22B, &H8D014388, &HC8BF85B
    pvAppendBuffer &HF7448BF7, &HC1AC0F04, &H18E8C118, &H8B0A4B88, &H448BF70C, &HAC0F04F7, &HE8C110C1, &HB4B8810, &H8BF70C8B, &HF04F744, &HC108C1AC, &H4B8808E8, &HF7048A0C, &H84D8B46, &H8808E983, &H4D890D43, &H6FE8308, &H5E5F8772, &H8C25D5B, &HEC8B5500, &H5320EC83, &H33085D8B, &H758B56C0, &H66A570C, &HE4458959, &HC7E87D8D, &H3E045, &HABF30000, &H5001468D, &HFBE0E853, &H3E80FFFF, &H8D0F7504
    pvAppendBuffer &H8D502146, &HE8502043, &HFFFFFBCE, &H46A76EB, &H207B8D53, &H42A5E857, &H46A0000, &HFFC3D8E8, &H10C083FF, &HE0458D50, &HE8575750, &H42B9&, &H5753046A, &H4258E857, &H46A0000, &HFFC3B8E8, &H10C083FF, &HC3AFE850, &HC083FFFF, &H57575030, &H3CEEE8, &H5DE85700, &H8A000018, &H240F8B06, &HC0B60F01, &H9901E183, &H675C83B, &HC23BC033, &H46A1274, &HC37FE857, &HC083FFFF, &HE8575010
    pvAppendBuffer &H4626&, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H5330EC83, &H33085D8B, &H758B56C0, &HA6A570C, &HD4458959, &HC7D87D8D, &H3D045, &HABF30000, &H5001468D, &HFBD0E853, &H3E80FFFF, &H8D0F7504, &H8D503146, &HE8503043, &HFFFFFBBE, &H66A7DEB, &H307B8D53, &H41E5E857, &H66A0000, &HFFC318E8, &HB005FF, &H8D500000, &H5750D045, &H41F7E857, &H66A0000, &HE8575753, &H4196&, &HF6E8066A
    pvAppendBuffer &HBBFFFFC2, &HB0&, &HE850C303, &HFFFFC2E9, &HE005&, &H57575000, &H3C26E8, &H29E85700, &H8A000018, &H240F8B06, &HC0B60F01, &H9901E183, &H675C83B, &HC23BC033, &H66A1174, &HC2B7E857, &HC303FFFF, &H5FE85750, &H5F000045, &HE58B5B5E, &H8C25D, &H81EC8B55, &HA0EC&, &H60858D00, &HFFFFFFFF, &HE8500875, &HFFFFFE61, &H8D0C75FF, &HE850E045, &HFFFFFA62, &H458D006A, &H858D50E0
    pvAppendBuffer &HFFFFFF60, &HA0458D50, &HC60FE850, &H458DFFFF, &H75FF50A0, &HFD05E810, &H458DFFFF, &H9EE850A0, &HF7FFFFC5, &H40C01BD8, &HC25DE58B, &H8B55000C, &HF0EC81EC, &H8D000000, &HFFFF1085, &H875FFFF, &HFEC7E850, &H75FFFFFF, &HD0458D0C, &HFAB8E850, &H6AFFFF, &H50D0458D, &HFF10858D, &H8D50FFFF, &HFFFF7085, &H5EE850FF, &H8DFFFFC7, &HFFFF7085, &H75FF50FF, &HFD3CE810, &H858DFFFF, &HFFFFFF70
    pvAppendBuffer &HC569E850, &HD8F7FFFF, &H8B40C01B, &HCC25DE5, &HEC8B5500, &H8D40EC83, &HFF56C045, &HE8500875, &HFFFFFDA9, &H8D0C758B, &H8D50C045, &H6C60146, &H68E85004, &H8DFFFFFC, &H8D50E045, &HE8502146, &HFFFFFC5B, &H5E40C033, &HC25DE58B, &H8B550008, &H60EC83EC, &H56A0458D, &H500875FF, &HFFFE2CE8, &HC758BFF, &H50A0458D, &HC601468D, &HE8500406, &HFFFFFCC2, &H50D0458D, &H5031468D, &HFFFCB5E8
    pvAppendBuffer &H40C033FF, &H5DE58B5E, &H550008C2, &HEC81EC8B, &H80&, &H7D8B5753, &H5B046A10, &H5FE85753, &H85000036, &H330774C0, &H115E9C0, &H53560000, &HC143E857, &H90BEFFFF, &H3000000, &HF0E850C6, &H83000035, &H107401F8, &HC12BE853, &HC603FFFF, &HE8575750, &H43D2&, &HE857006A, &HFFFFC119, &H5050C083, &H5080458D, &HFFC4B0E8, &H458D53FF, &H2E85080, &H3FFFFC1, &HB4E850C6, &H83000035
    pvAppendBuffer &H137401F8, &HC0EFE853, &HC603FFFF, &H80458D50, &H93E85050, &H53000043, &H5080458D, &H35E1E8, &H74C08500, &HE9C03307, &H96&, &H8D14758B, &H56508045, &HFFFB66E8, &H875FFFF, &H50C0458D, &HFFF895E8, &HC0AFE8FF, &H9005FFFF, &H50000000, &H50C0458D, &H5080458D, &H50E0458D, &H3C41E8, &HC75FF00, &H50C0458D, &HFFF86DE8, &H86E853FF, &H5FFFFC0, &H90&, &HE0458D50, &HC0458D50
    pvAppendBuffer &HE0458D50, &H39B9E850, &HE8530000, &HFFFFC069, &H90BB&, &H50C30300, &HE3E85757, &HE8000039, &HFFFFC055, &H4D8DC303, &H515750E0, &H3BF0E851, &H4D8D0000, &H4E8D51E0, &HE0E85120, &H33FFFFFA, &H5F5E40C0, &H5DE58B5B, &H550010C2, &HEC81EC8B, &HC0&, &H7D8B5753, &H5B066A10, &H1FE85753, &H85000035, &H330774C0, &H129E9C0, &H53560000, &HC003E857, &H70BEFFFF, &H3000001, &HB0E850C6
    pvAppendBuffer &H83000034, &H107401F8, &HBFEBE853, &HC603FFFF, &HE8575750, &H4292&, &HE857006A, &HFFFFBFD9, &H11005, &H858D5000, &HFFFFFF40, &HC517E850, &H8D53FFFF, &HFFFF4085, &HBAE850FF, &H3FFFFBF, &H6CE850C6, &H83000034, &H167401F8, &HBFA7E853, &HC603FFFF, &H40858D50, &H50FFFFFF, &H4248E850, &H8D530000, &HFFFF4085, &H93E850FF, &H85000034, &H330774C0, &H9CE9C0, &H758B0000, &H40858D14
    pvAppendBuffer &H50FFFFFF, &HFAACE856, &H75FFFFFF, &HA0458D08, &HF7F4E850, &H5EE8FFFF, &H5FFFFBF, &H170&, &HA0458D50, &H40858D50, &H50FFFFFF, &H50D0458D, &H3C54E8, &HC75FF00, &H50A0458D, &HFFF7C9E8, &H32E853FF, &H5FFFFBF, &H170&, &HD0458D50, &HA0458D50, &HD0458D50, &H3865E850, &HE8530000, &HFFFFBF15, &H170BB, &H50C30300, &H8FE85757, &HE8000038, &HFFFFBF01, &H4D8DC303, &H515750D0
    pvAppendBuffer &H3C03E851, &H4D8D0000, &H4E8D51D0, &H23E85130, &H33FFFFFA, &H5F5E40C0, &H5DE58B5B, &H550010C2, &HEC81EC8B, &H1B0&, &HFE50858D, &H5653FFFF, &H500875FF, &HFFFA94E8, &H10758BFF, &HFF30858D, &H5056FFFF, &HFFF691E8, &H20468DFF, &H90458D50, &HF684E850, &H46AFFFF, &H858D535B, &HFFFFFF30, &H339CE850, &HC0850000, &H36B850F, &H8D530000, &HE8509045, &H338A&, &H850FC085, &H359&
    pvAppendBuffer &H30858D53, &H50FFFFFF, &HFFBE6CE8, &H90BEFF, &HC6030000, &H3319E850, &HF8830000, &H36850F01, &H53000003, &H5090458D, &HFFBE4CE8, &H50C603FF, &H32FEE8, &H1F88300, &H31B850F, &H53570000, &HFFBE34E8, &H50C603FF, &H5090458D, &H50E0458D, &H37ADE8, &HC75FF00, &HFF50858D, &HE850FFFF, &HFFFFF5F6, &HFFBE10E8, &H50C603FF, &H50E0458D, &HFF50858D, &H5050FFFF, &H39A5E8, &HBDF7E800
    pvAppendBuffer &HC603FFFF, &HE0458D50, &H30858D50, &H50FFFFFF, &HFF10858D, &HE850FFFF, &H3986&, &H50858D53, &H50FFFFFE, &HFEB0858D, &HE850FFFF, &H3EBF&, &H70858D53, &H50FFFFFE, &HFED0858D, &HE850FFFF, &H3EAB&, &HBDAFE853, &HC083FFFF, &H858D5050, &HFFFFFF70, &H3E95E850, &HE8530000, &HFFFFBD99, &H5070C083, &H5090458D, &H3E82E8, &H86E85300, &H83FFFFBD, &H8D5010C0, &HFFFF7085, &H858D50FF
    pvAppendBuffer &HFFFFFEB0, &HE0458D50, &H3C5BE850, &H858D0000, &HFFFFFED0, &HB0858D50, &H50FFFFFE, &H5090458D, &HFF70858D, &HE850FFFF, &HFFFFC45A, &HBD47E853, &HC083FFFF, &H458D5010, &HE85050E0, &H36C2&, &H50E0458D, &HFED0858D, &H8D50FFFF, &HFFFEB085, &H9AE850FF, &H83FFFFCC, &HE800D065, &HFFFFBD19, &H8950C083, &H858DD445, &HFFFFFE50, &H8DD84589, &HFFFEB085, &HDC4589FF, &HFF10858D, &H5053FFFF
    pvAppendBuffer &H3D44E8, &H8DF08B00, &HFFFF5085, &HE85053FF, &H3D35&, &HFE3BF88B, &HFE8B0277, &H56FF778D, &HFF50858D, &HE850FFFF, &H3FF3&, &H574C20B, &HEB43DB33, &H56DB3302, &HFF10858D, &HE850FFFF, &H3FDB&, &H574C20B, &HEB5E026A, &H6AF63302, &H8DF30B04, &H535BB045, &HD0B5748B, &H8CE85056, &H5300003D, &H5020468D, &HFEF0858D, &HE850FFFF, &H3D7B&, &HE0458D53, &H3108E850, &H65830000
    pvAppendBuffer &H778D00E4, &HE045C7FE, &H1&, &H880FF685, &HD5&, &H50E0458D, &HFEF0858D, &H8D50FFFF, &HE850B045, &HFFFFBC75, &H50858D56, &H50FFFFFF, &H3F6AE8, &H74C20B00, &H47FF3305, &HFF3302EB, &H10858D56, &H50FFFFFF, &H3F52E8, &H74C20B00, &H58026A05, &HC03302EB, &H7C8BF80B, &HFF85D0BD, &H57537E74, &HFF70858D, &HE850FFFF, &H3CFF&, &H20478D53, &H90458D50, &H3CF1E850, &H458D0000
    pvAppendBuffer &H458D50E0, &H858D5090, &HFFFFFF70, &HCB5FE850, &HE853FFFF, &HFFFFBBE1, &H5010C083, &HFF70858D, &H8D50FFFF, &H8D50B045, &HFFFE9085, &HB6E850FF, &H8D00003A, &HFFFEF085, &H458D50FF, &H458D50B0, &H858D5090, &HFFFFFF70, &HC2B8E850, &H8D53FFFF, &HFFFE9085, &H458D50FF, &HE85050E0, &H3A32&, &HF01EE83, &HFFFF2B89, &H8AE853FF, &H83FFFFBB, &H8D5010C0, &H5050E045, &H3505E8, &HE0458D00
    pvAppendBuffer &HF0858D50, &H50FFFFFE, &H50B0458D, &HFFCAE0E8, &H458D53FF, &H5EE850B0, &HBEFFFFBB, &H90&, &HE850C603, &H300B&, &H1F8835F, &HE8531374, &HFFFFBB45, &H8D50C603, &H5050B045, &H3DE9E8, &H858D5300, &HFFFFFF30, &HB0458D50, &H2FE1E850, &HD8F70000, &HEB40C01B, &H5EC03302, &H5DE58B5B, &H55000CC2, &HEC81EC8B, &H280&, &HFD80858D, &H5653FFFF, &H500875FF, &HFFF790E8, &H10758BFF
    pvAppendBuffer &HFED0858D, &H5056FFFF, &HFFF37DE8, &H30468DFF, &H60858D50, &H50FFFFFF, &HFFF36DE8, &H5B066AFF, &HD0858D53, &H50FFFFFE, &H2FD5E8, &HFC08500, &H39185, &H858D5300, &HFFFFFF60, &H2FC0E850, &HC0850000, &H37C850F, &H8D530000, &HFFFED085, &HA2E850FF, &HBEFFFFBA, &H170&, &HE850C603, &H2F4F&, &HF01F883, &H35985, &H858D5300, &HFFFFFF60, &HBA7FE850, &HC603FFFF, &H2F31E850
    pvAppendBuffer &HF8830000, &H3B850F01, &H57000003, &HBA67E853, &HC603FFFF, &H60858D50, &H50FFFFFF, &H50C0458D, &H33DDE8, &HC75FF00, &HFF00858D, &HE850FFFF, &HFFFFF2D6, &HFFBA40E8, &H50C603FF, &H50C0458D, &HFF00858D, &H5050FFFF, &H373CE8, &HBA27E800, &HC603FFFF, &HC0458D50, &HD0858D50, &H50FFFFFE, &HFEA0858D, &HE850FFFF, &H371D&, &H80858D53, &H50FFFFFD, &HFE10858D, &HE850FFFF, &H3AEF&
    pvAppendBuffer &HB0858D53, &H50FFFFFD, &HFE40858D, &HE850FFFF, &H3ADB&, &HB9DFE853, &HC683FFFF, &H50C603A0, &HFF30858D, &HE850FFFF, &H3AC3&, &HB9C7E853, &H4005FFFF, &H50000001, &HFF60858D, &HE850FFFF, &H3AAB&, &HB9AFE853, &HB0BFFFFF, &H3000000, &H858D50C7, &HFFFFFF30, &H10858D50, &H50FFFFFE, &H50C0458D, &H3880E8, &H40858D00, &H50FFFFFE, &HFE10858D, &H8D50FFFF, &HFFFF6085, &H858D50FF
    pvAppendBuffer &HFFFFFF30, &HC18CE850, &HE853FFFF, &HFFFFB969, &H8D50C703, &H5050C045, &H32E5E8, &HC0458D00, &H40858D50, &H50FFFFFE, &HFE10858D, &HE850FFFF, &HFFFFC908, &HF06583, &HFFB93CE8, &H89C603FF, &H858DF445, &HFFFFFD80, &H8DF84589, &HFFFE1085, &HFC4589FF, &HFEA0858D, &H5053FFFF, &H3968E8, &H8DF08B00, &HFFFF0085, &HE85053FF, &H3959&, &HFE3BF88B, &HFE8B0277, &H56FF778D, &HFF00858D
    pvAppendBuffer &HE850FFFF, &H3C17&, &H574C20B, &HEB43DB33, &H56DB3302, &HFEA0858D, &HE850FFFF, &H3BFF&, &H574C20B, &HEB5E026A, &H6AF63302, &H8DF30B06, &H535B9045, &HF0B5748B, &HB0E85056, &H53000039, &H5030468D, &HFE70858D, &HE850FFFF, &H399F&, &HC0458D53, &H2D2CE850, &H65830000, &H778D00C4, &HC045C7FE, &H1&, &H880FF685, &HE4&, &H50C0458D, &HFE70858D, &H8D50FFFF, &HE8509045
    pvAppendBuffer &HFFFFBA21, &H858D56, &H50FFFFFF, &H3B8EE8, &H74C20B00, &H47FF3305, &HFF3302EB, &HA0858D56, &H50FFFFFE, &H3B76E8, &H74C20B00, &H58026A05, &HC03302EB, &H7C8BF80B, &HFF85F0BD, &H89840F, &H57530000, &HFF30858D, &HE850FFFF, &H391F&, &H30478D53, &H60858D50, &H50FFFFFF, &H390EE8, &HC0458D00, &H60858D50, &H50FFFFFF, &HFF30858D, &HE850FFFF, &HFFFFC7C4, &HB7FBE853, &HB005FFFF
    pvAppendBuffer &H50000000, &HFF30858D, &H8D50FFFF, &H8D509045, &HFFFDE085, &HCEE850FF, &H8D000036, &HFFFE7085, &H458D50FF, &H858D5090, &HFFFFFF60, &H30858D50, &H50FFFFFF, &HFFBFDDE8, &H858D53FF, &HFFFFFDE0, &HC0458D50, &H47E85050, &H83000036, &H890F01EE, &HFFFFFF1C, &HB79FE853, &HB005FFFF, &H50000000, &H50C0458D, &H3118E850, &H458D0000, &H858D50C0, &HFFFFFE70, &H90458D50, &HC73EE850, &H8D53FFFF
    pvAppendBuffer &HE8509045, &HFFFFB771, &H170BE, &H50C60300, &H2C1EE8, &HF8835F00, &H53137401, &HFFB758E8, &H50C603FF, &H5090458D, &H39FCE850, &H8D530000, &HFFFED085, &H458D50FF, &HF4E85090, &HF700002B, &H40C01BD8, &HC03302EB, &HE58B5B5E, &HCC25D, &H56EC8B55, &H8B08758B, &H85048D06, &H4&, &HC8E85650, &HE8000029, &HFFFFB70D, &H850FF56, &H4C25D5E, &HEC8B5500, &H8B084D8B, &H7E8C1C1
    pvAppendBuffer &H7F7FE181, &H125FF7F, &H3010101, &H1BC06BC9, &HC25DC133, &H8B550004, &HB6EEE8EC, &HD1B9FFFF, &H8100478A, &H474000E9, &H8BC10300, &H5051084D, &H8D1075FF, &H75FF3041, &H50106A0C, &H5020418D, &HFFD6EEE8, &HCC25DFF, &HEC8B5500, &H8B084D8B, &HFF501045, &H41010C75, &H51835138, &HB3E8003C, &H5DFFFFFF, &H55000CC2, &H8B56EC8B, &H7E830875, &HD750148, &H20E856, &H46C70000, &H248&
    pvAppendBuffer &H10458B00, &H50404601, &H830C75FF, &H56004456, &HFFFF81E8, &HC25D5EFF, &H8B55000C, &H758B56EC, &H304E8B08, &H2474C985, &H2B58106A, &H468D50C1, &H6AC10320, &H72E85000, &H8300000A, &H468D0CC4, &HE8565020, &H9&, &H306683, &H4C25D5E, &HEC8B5500, &H8D10EC83, &H5756F045, &HC75FF50, &HFFE5FCE8, &H87D8BFF, &H8DF0458D, &H56561077, &HE551E850, &H5756FFFF, &H4C57FF56, &HE58B5E5F
    pvAppendBuffer &H8C25D, &H51EC8B55, &H758B5651, &H487E8308, &H83067401, &H7502487E, &H7BE8560A, &H83FFFFFF, &H8B004866, &H458D384E, &H468B50F8, &HC8A40F3C, &H3E1C103, &H33E85150, &H6A000039, &HF8458D08, &HCFE85650, &H8BFFFFFE, &H458D404E, &H468B50F8, &HC8A40F44, &H3E1C103, &HFE85150, &H6A000039, &HF8458D08, &HABE85650, &HFFFFFFFE, &H468D0C75, &H1FE85010, &H5EFFFFE7, &HC25DE58B, &H8B550008
    pvAppendBuffer &H10EC83EC, &H8B575653, &H506A087D, &HE857006A, &H999&, &H570CC483, &HE80C75FF, &HFFFFE541, &HC933C033, &H47895340, &H8BA20F48, &H5D8D5BF3, &H890389F0, &H4B890473, &HC538908, &HFFB54BE8, &HF845F6FF, &H719FB902, &H5750047, &H47711EB9, &HE98100, &H3004740, &H4C4789C1, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H5614558B, &HEA83F633, &H8B307801, &H8B530C45, &H2B57085D, &H903C8DD8
    pvAppendBuffer &H8B104529, &HC0333B0C, &HC013CE03, &HD0830F03, &H1EA8300, &H458BF08B, &H380C8910, &H79FC7F8D, &H8B5B5FE2, &HC25D5EC6, &H8B550010, &H8B5653EC, &HC68B1075, &H1FE28399, &H23C8D57, &H8105FFC1, &H1FE6&, &H4E057980, &H46E0CE83, &H8B0C558B, &H59206ADE, &HC28BCE2B, &HDBF7E8D3, &HDB1BCE8B, &HD823E2D3, &H5D89C933, &H85D8B10, &H110758B, &H1304BB54, &H75F685C9, &H74C98504, &HBB448B2C
    pvAppendBuffer &H3D23308, &H3D213C1, &HBB4489C6, &HD28308, &H8503C783, &H8D1274D2, &HC033BB34, &H768D1601, &H8BC01304, &H75D285D0, &H5B5E5FF1, &HCC25D, &H83EC8B55, &H558B30EC, &H5D8B5314, &H56C38B0C, &H2BFF3357, &HEC4589C2, &H1EB880F, &H558B0000, &H8BC88BEC, &HE1C10845, &HFC4D8905, &H85B8348B, &H470C75F6, &H8920E983, &HBCE9FC4D, &H56000001, &HFFCA53E8, &H8BD08BFF, &HD45589CA, &HD285E6D3
    pvAppendBuffer &H478D177E, &H7DC33B01, &H8458B10, &H2B59206A, &HB8548BCA, &HBEAD304, &H205D8BF2, &H65F7C68B, &HFC458B1C, &H2BE1C383, &HF28BD445, &H7589D803, &HF45D89F0, &HFB831F79, &H748E0FE0, &HF7000001, &H33CB8BDB, &H89EED3DB, &H5D89F075, &HFF685F4, &H15E84, &H25C38B00, &H8000001F, &H83480579, &H8B40E0C8, &H45890C4D, &H99C38BDC, &H31FE283, &H14558BC2, &H2B05F8C1, &HDC458BC8, &HC085CA2B
    pvAppendBuffer &H45217575, &H45214AF0, &H89CA03DC, &H4D89E455, &HE845C7E0, &H1&, &H8C0FCF3B, &HEF&, &H8B08458B, &H48D105D, &HF8458988, &H479D285, &H3EBC033, &HF793048B, &HF04503E6, &HD283D0F7, &H89C93300, &H558BF055, &H130203F8, &HE84503C9, &H4D130289, &HE4558BDC, &H4F86D83, &HDC65834A, &HE84D8900, &H49E04D8B, &H89E45589, &HCF3BE04D, &H95E9BA7D, &H83000000, &H4A00E865, &H5589CA03
    pvAppendBuffer &HD44D89D8, &H8C0FCF3B, &H83&, &H2B5B206A, &HE445C7D8, &H1&, &H8308458B, &H8900E065, &H5D8BD05D, &H88048DDC, &H85F84589, &H330479D2, &H8B06EBC0, &H48B1045, &H8BE6F790, &H3F08BCB, &HD283E075, &H89C03300, &HD68BE055, &H550BE2D3, &HF84D8BE8, &H1103D2F7, &H5503C013, &HD083E4, &H8BE44589, &HD04D8BC1, &HD68B1089, &H83F0758B, &HEAD304E8, &H89D44D8B, &H558BE855, &H45894AD8
    pvAppendBuffer &H558949F8, &HD44D89D8, &H9F7DCF3B, &H83F45D8B, &H7400187D, &HFF56530A, &H8BE81875, &H8BFFFFFD, &H4D8BEC55, &H8458BFC, &H3B0C5D8B, &H298E0FFA, &H8BFFFFFE, &H558B0C5D, &H87D8B14, &HDB85C933, &H458B247E, &H8DD32B10, &H48D9034, &H79C0850A, &HEBC03304, &H39068B02, &H67728F04, &H83410877, &HCB3B04C6, &H6583E47C, &H538D000C, &H845C7FF, &H1&, &H3978D285, &H8B10458B, &H3F32BF2
    pvAppendBuffer &H1C8D1475, &H79F685B0, &HEBC93304, &H330B8B02, &H3D1F7C0, &HC013970C, &H89084D03, &H4513970C, &H65834E0C, &HEB83000C, &H1EA8304, &H79084589, &H187D83D4, &H6A0C7400, &HFF016A00, &HE7E81875, &H5FFFFFFC, &HE58B5B5E, &H1CC25D, &H83EC8B55, &H8B5328EC, &H5756145D, &HF32FB83, &H1B88E, &H18758B00, &H2B99C38B, &H8BF88BC2, &HFFD156C3, &H1075FF57, &H7D89C72B, &HC75FFF0, &HFFFC4589
    pvAppendBuffer &HC5E80875, &H8BFFFFFF, &H8B561045, &H8D56FC75, &H458BF814, &H5589520C, &HB80C8DD8, &H5108458B, &H8DDC4D89, &H8950B804, &H9DE8E445, &H8BFFFFFF, &H468D184D, &H6AF63301, &H81148D04, &H89047A8D, &HF4558932, &H8904518D, &HEC558937, &H558B3289, &HE07D89F0, &H855F3189, &H3467ED2, &H8DC22BC0, &H45898104, &HFC458BE8, &H8DC22B40, &H4D8B8104, &HF845890C, &H2B08458B, &HC5589C1, &H89E8558B
    pvAppendBuffer &H5D8B0845, &H8048BF8, &H89F87D01, &H3018B03, &H145D8BCF, &HD7030289, &H10C6D83, &H7508458B, &HFC75FFE1, &H50EC458B, &H50E475FF, &HFFFBB4E8, &H184D8BFF, &H89FC75FF, &HE0458B01, &HDC75FF50, &HFB9FE850, &H4D8BFFFF, &HFC558BF4, &H8B018942, &H48D184D, &HC4589D1, &HE0C1C28B, &H50C10304, &HC75FF52, &H51F475FF, &HFFFEDFE8, &H184D8BFF, &H8DF0558B, &H71890841, &HDC45890C, &H48D3089
    pvAppendBuffer &H4718912, &H45893189, &H7EC08508, &HFC458B1E, &H40105D8B, &H558BC22B, &HC10C8D08, &H46B3048B, &HCF030189, &HF47CF23B, &H8B145D8B, &H48DFC75, &H458B5036, &H75FF50DC, &H26E850D8, &H8BFFFFFB, &H189EC4D, &H8B01468D, &H348D184D, &HC458B00, &H50515056, &H1F2E8, &HFC4D8B00, &H1418D56, &H458BD82B, &H2BDB0310, &H98048DD9, &HC75FF50, &HFAEFE850, &HF08BFFFF, &H840FF685, &H87&
    pvAppendBuffer &H8310558B, &H148DFCC2, &H1C0339A, &H2BC01332, &H85F08BD7, &HEBF275F6, &H10558B6E, &H851B0C8D, &H33067EC9, &HF3FA8BC0, &H8458BAB, &H3B980C8D, &H8B5476C8, &H148D0C75, &HFCC283DA, &H5589046A, &H9E348D10, &H5F147589, &H186583, &HDA8BCF2B, &H760C753B, &H2B018B21, &H326F7F7, &HD2831845, &H83030100, &HDF2B00D2, &H3B185589, &HE5770C75, &H8B08458B, &H758B1055, &H89D72B18, &H14758B33
    pvAppendBuffer &H3B105589, &H5FC177C8, &HE58B5B5E, &H14C25D, &H51EC8B55, &H5D8B5351, &H83575614, &H8E0F32FB, &HA1&, &H8B0C7D8B, &H8758BC3, &H8BC22B99, &H8BC88BD3, &HF9D11845, &H4D89D12B, &HF85589FC, &H53D81C8D, &HC8048D52, &H8F048D50, &H8E048D50, &HFD76E850, &H4D8BFFFF, &H18458BFC, &H48D5153, &H458B5088, &H87048DF8, &HA1E85650, &H53FFFFFF, &H8BFC75FF, &H5753185D, &H8DF87D8B, &HE850BE04
    pvAppendBuffer &HFFFFFF8C, &H85FC758B, &H8B1B7EFF, &HC8D1045, &H77048DB0, &H8B83148D, &H4528D02, &H498D0189, &H1EF8304, &H5356F175, &H50B3048D, &HF9B3E853, &H458BFFFF, &H75FF5614, &H83048D10, &HA2E85350, &HEBFFFFF9, &H7EDB857D, &H107D8B09, &HC033CB8B, &H458BABF3, &H107D8B08, &H8D980C8D, &H75899F34, &H76C83B14, &HC458B5D, &H8B98048D, &HC083085D, &HFC4589FC, &H106583, &H8304E983, &H89001865
    pvAppendBuffer &H45890C75, &H76F73BF8, &H8BDE8B2F, &H83068BF0, &H21F704EB, &H3FC768D, &HD28303, &H89104503, &H18551303, &H186583, &H3B105589, &H8BDE77DF, &H5D8B1475, &HFC458B08, &H8904EE83, &HCB3B1475, &H5E5FB277, &H5DE58B5B, &H550014C2, &H8B51EC8B, &H65831455, &HEA8300FC, &H1445C701, &H1&, &H458B3978, &H8B565308, &H8B570C75, &H1C8D107D, &H2BF02B90, &H1E0C8BF8, &HD1F7C033, &HC0130B03
    pvAppendBuffer &H89144D03, &H5B8D1F0C, &HFC4513FC, &HFC6583, &H8901EA83, &HDD791445, &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &H80EC81, &H8B560000, &H6A570C75, &H7D8D5920, &HBEA5F380, &HFD&, &H5080458D, &H2087E850, &HFE830000, &H83127402, &HD7404FE, &H8D0C75FF, &H50508045, &HE7FE8, &H1EE8300, &H7D8BDA79, &H80758D08, &HF359206A, &H8B5E5FA5, &H8C25DE5, &HEC8B5500, &HFF575653, &H72E80875
    pvAppendBuffer &H8BFFFFF6, &H6AE853D8, &H8BFFFFF6, &H62E852D0, &H8BFFFFF6, &H87D33F8, &HC78BF78B, &HCFC1C333, &HC1F23308, &HCE8B08C0, &H3310C9C1, &H33C733C1, &HC3335FC6, &H5E084533, &H4C25D5B, &HEC8B5500, &H8758B56, &HAAE836FF, &HFFFFFFFF, &H6890476, &HFFFFA0E8, &H876FFFF, &HE8044689, &HFFFFFF95, &H890C76FF, &H8AE80846, &H89FFFFFF, &H5D5E0C46, &H550004C2, &H8B53EC8B, &H5756085D, &H77BB60F
    pvAppendBuffer &H243B60F, &HB73B60F, &HF53B60F, &HB08E7C1, &H4BB60FF8, &H43B60F03, &H8E7C10D, &HE6C1F80B, &H43B60F08, &H8E7C108, &HE2C1F80B, &H43B60F08, &HC1F00B06, &HB60F08E1, &HE6C10143, &HFF00B08, &HC10C43B6, &HF00B08E6, &HA43B60F, &HB60FD00B, &HE2C10543, &HFD00B08, &HE2C103B6, &HFD00B08, &HB0E43B6, &HC5389C8, &H943B60F, &HB08E1C1, &H87389C8, &H443B60F, &HC1047B89, &HB5F08E1
    pvAppendBuffer &HB895EC8, &H4C25D5B, &HEC8B5500, &H66E85756, &H8BFFFFAC, &H93BF0875, &H3000006, &H36FF50C7, &H1FD2E8, &HE8068900, &HFFFFAC4D, &HFF50C703, &HC0E80476, &H8900001F, &H3AE80446, &H3FFFFAC, &H76FF50C7, &H1FADE808, &H46890000, &HAC27E808, &HC703FFFF, &HC76FF50, &H1F9AE8, &H46895F00, &HC25D5E0C, &H8B550004, &HC7D83EC, &H56157400, &H8B0C75FF, &H6A0875, &H2FE856, &H68A0000
    pvAppendBuffer &H5E0CC483, &H8C25D, &H8BEC8B55, &H458B1055, &HF08B5608, &H1274D285, &HC7D8B57, &HC8AF82B, &H460E8837, &H7501EA83, &H5D5E5FF5, &HEC8B55C3, &H85104D8B, &HF1F74C9, &H560C45B6, &HC069F18B, &H1010101, &H87D8B57, &HF302E9C1, &H83CE8BAB, &HAAF303E1, &H458B5E5F, &H55C35D08, &H8B56EC8B, &HE8560875, &HFFFFF48D, &HCE8BD08B, &HC9C1D633, &H8C2C110, &H3308CEC1, &H33D633D1, &HC25D5EC2
    pvAppendBuffer &H8B550004, &H758B56EC, &HE836FF08, &HFFFFFFCB, &H890476FF, &HFFC1E806, &H76FFFFFF, &H4468908, &HFFFFB6E8, &HC76FFFF, &HE8084689, &HFFFFFFAB, &H5E0C4689, &H4C25D, &H83EC8B55, &H658340EC, &HC03300C4, &H53E44521, &H66A5756, &H8DDB3359, &H66AC87D, &HC05D8943, &H6A59ABF3, &HE87D8D04, &HF3E05D89, &H8D575FAB, &HE850C045, &HFFFFAB09, &H5010C083, &H50C0458D, &H1F0CE8, &H458D5700
    pvAppendBuffer &H42E850C0, &H8D00002B, &H26EBFF70, &H50E0458D, &H29AAE8, &H458D5600, &HE850C0, &HB00002E, &H570E74C2, &H8D0875FF, &H5050E045, &H2961E8, &H458D4E00, &H3B5057E0, &HFFD177F3, &HACE80875, &H5F00002B, &HE58B5B5E, &H4C25D, &H83EC8B55, &H658360EC, &HC03300A4, &H53D44521, &HA6A5756, &H8DDB3359, &HA6AA87D, &HA05D8943, &H6A59ABF3, &HD87D8D06, &HF3D05D89, &H8D575FAB, &HE850A045
    pvAppendBuffer &HFFFFAA75, &HB005&, &H458D5000, &H76E850A0, &H5700001E, &H50A0458D, &H2AACE8, &HFF708D00, &H458D26EB, &H14E850D0, &H56000029, &H50A0458D, &H2D6AE8, &H74C20B00, &H75FF570E, &HD0458D08, &HCBE85050, &H4E000028, &H57D0458D, &H77F33B50, &H875FFD1, &H2B16E8, &H5B5E5F00, &HC25DE58B, &H8B550004, &H14EC83EC, &HFF575653, &H5CE80C75, &HFFFFFFE1, &H45890875, &HE151E8FC, &HF08BFFFF
    pvAppendBuffer &HFFA9F4E8, &H79405FF, &HE8500000, &HFFFFE13F, &HE8084589, &HFFFFA9E1, &H79805, &H2CE85000, &H33FFFFE1, &H43F88BDB, &H89F85D89, &HC6E8EC5D, &HE9FFFFA9, &HA6&, &HFFA9BCE8, &H79405FF, &H56500000, &HFFBDEEE8, &HFC085FF, &H10684, &HE836FF00, &HC1F&, &H8BF44589, &H30FFFC45, &HC12E8, &H45895000, &HF4458BF0, &H75FF5650, &HBA3FE8FC, &H4D8BFFFF, &H761939F4, &H83018B0F
    pvAppendBuffer &H7500813C, &H1894807, &HF177C33B, &H39F04D8B, &H8B0F7619, &H813C8301, &H48077500, &HC33B0189, &H75FFF177, &HF226E8FC, &H458BFFFF, &HFC7589F4, &H5608758B, &HF075FF57, &H89F44589, &H73E8087D, &HF7FFFFBB, &HF88BEC5D, &HF202E856, &H75FFFFFF, &HF1FAE8F0, &H1EE8FFFF, &H8BFFFFA9, &H9805F475, &H50000007, &HBD4DE856, &HC085FFFF, &HFF46850F, &HE856FFFF, &HFFFFF1D8, &HE8FC75FF, &HFFFFF1D0
    pvAppendBuffer &HE80875FF, &HFFFFF1C8, &HEC7D83, &HBE8D0F, &H458B0000, &HE830FF0C, &HB5B&, &HC933F08B, &H7589C033, &H63940EC, &H988C0F, &H458B0000, &H4568D0C, &H7D89C72B, &HFC7529FC, &H8BF44589, &H1E3B0C75, &H7FEC758B, &HFC458B2E, &H8BF44503, &H45891004, &H5624EBF0, &HFFF173E8, &HFC75FFFF, &HFFF16BE8, &H875FFFF, &HFFF163E8, &H5DE857FF, &H33FFFFF1, &H835BEBC0, &H3B00F065, &H8B0B7F1F
    pvAppendBuffer &H48BFC45, &H8458910, &H658304EB, &H458B0008, &H8452BF0, &H2B0855F7, &H850289C1, &H3B0874C9, &HC91B0845, &H3907EB41, &HC91B0845, &HC085D9F7, &H74F8458B, &H89C38B05, &H8343F845, &H1E3B04C2, &HFF798E0F, &H8957FFFF, &HF102E806, &HFE8BFFFF, &H5E5FC78B, &H5DE58B5B, &H550008C2, &HEC83EC8B, &H5756531C, &H8B107D8B, &HC1DE8B37, &H5D8902E3, &HA803E8E8, &HFF53FFFF, &H33D08B10, &H105589C9
    pvAppendBuffer &HF7EF685, &HC12B078B, &H8987048B, &H3B418A04, &H8BF17CCE, &H4D8B087D, &H3B078B0C, &H8B027601, &H8D198BCF, &HC63B1B04, &HC68B0A7F, &H8BC22B99, &H43FBD1D8, &HE0C1C38B, &HE4458902, &HFFA7B8E8, &HC1CB8BFF, &HFF5102E1, &H89D38B10, &H172BFC45, &HC7ED285, &H8BFC7D8B, &HF3C033CA, &H87D8BAB, &HC933078B, &H1A7EC085, &H8DFC7D8B, &H7D8B9714, &H41C12B08, &H8987048B, &H4528D02, &HC83B078B
    pvAppendBuffer &H72E8EF7C, &H8BFFFFA7, &H2E1C1CB, &H8B10FF51, &HCB8B0C55, &H7D89F88B, &H890A2BF8, &HC985084D, &HC033077E, &H4D8BABF3, &H33028B08, &H7EC085FF, &HF8558B1A, &H8B8A0C8D, &HC72B0C55, &H82048B47, &H498D0189, &H3B028B04, &H8BEF7CF8, &H3E7C1FB, &HE8EC7D89, &HFFFFA721, &H5310FF57, &HE8084589, &H96F&, &HE7C1F88B, &HF07D8902, &HFFA708E8, &H10FF57FF, &H89107D8B, &H3F8B0C45, &HBD46E857
    pvAppendBuffer &H4589FFFF, &H74C085F4, &HD3C88B18, &H1FE83E7, &H206A0F7E, &H8BC82B59, &H408B1045, &HBE8D304, &HC75FFF8, &H875FF53, &HFFF875FF, &H19E8FC75, &HFFFFFFF5, &HDB03F475, &H1006E857, &H6A500000, &H75FF5600, &H75FF5310, &HF262E808, &HF33BFFFF, &HF38B027C, &H91DE856, &HD2330000, &HF685F88B, &H458B197E, &H8DDE2B08, &HF8B981C, &HCA2B038B, &H45B8D42, &H3B8F0489, &H83EF7CD6, &H1076013F
    pvAppendBuffer &H3C83078B, &H8750087, &H83078948, &HF07701F8, &H8BF075FF, &HE8530C5D, &H1913&, &HFFA658E8, &H50FF53FF, &HEC75FF08, &H53085D8B, &H18FEE8, &HA643E800, &HFF53FFFF, &H75FF0850, &H105D8BE8, &H18E9E853, &H2EE80000, &H53FFFFA6, &H8B0850FF, &H5D8BE475, &HE85356FC, &H18D3&, &HFFA618E8, &H50FF53FF, &HF85D8B08, &HC0E85356, &HE8000018, &HFFFFA605, &H850FF53, &H5E5FC78B, &H5DE58B5B
    pvAppendBuffer &H55000CC2, &HEC83EC8B, &H5D8B532C, &H4438D10, &H890100F6, &H1175E045, &HC75FF53, &HE80875FF, &H36A&, &H35EE9, &H53575600, &HE80875FF, &HFFFFB7B4, &HF08B1B8B, &H5D89CB8B, &H5E1C1DC, &HBC2EE851, &HF88BFFFF, &H1075FF57, &HFFFB91E8, &H1075FFFF, &H57084589, &HFD7DE856, &H8956FFFF, &H69E8F045, &HFFFFFFEE, &HE8571075, &HFFFFB778, &HFC458957, &HFFEE57E8, &HC1FB8BFF, &H7D8902E7
    pvAppendBuffer &HA573E8D8, &HFF57FFFF, &HEC458910, &H197EDB85, &H8DE04D8B, &HD003FC57, &H18BF38B, &H8904498D, &HFC528D02, &H7501EE83, &HA54BE8F1, &HFF57FFFF, &H33F08B10, &HF47589C9, &H2D7EDB85, &H8308558B, &HC283FCC7, &H8BF70304, &HF3B087D, &H28B047D, &HC03302EB, &H83410689, &HEE8304C2, &H7CCB3B04, &HF4758BE9, &HE7C1FB8B, &H8458B02, &HEDDEE850, &H2E8FFFF, &H57FFFFA5, &HD08B10FF, &H85E85589
    pvAppendBuffer &H330D7EDB, &H8BCB8BC0, &H8BABF3FA, &H2E7C1FB, &H52565653, &HFFF6B2E8, &H85C033FF, &H8B2D7EDB, &H4D8BE855, &HFCC283F0, &H304C183, &HF07D8BD7, &H47D073B, &H2EB318B, &H3289F633, &H4C18340, &H3B04EA83, &H8BE97CC3, &H2E7C1FB, &H51F04D8B, &HFFED7BE8, &HC1F38BFF, &H758903E6, &HA497E8E0, &HFF56FFFF, &H89F08B10, &H8AE80875, &H8BFFFFA4, &H3E1C1CB, &H3310FF51, &HF84589C9, &H2C7EDB85
    pvAppendBuffer &H8BFC7D8B, &H3E0C1C3, &H3FCC083, &H4578DF0, &H47D0F3B, &H2EB028B, &H689C033, &H4C28341, &H3B04EE83, &H8BE97CCB, &H2E7C1FB, &HE8FC75FF, &HFFFFED1C, &H699E853, &HF36B0000, &HC1F00303, &H758902E6, &HA42FE8D4, &HFF56FFFF, &HC4D8B10, &H4589D233, &H891F6AFC, &H18BF055, &H5EE44589, &H317EC085, &H3C8DD88B, &H8BC03399, &HE0D340CE, &H10750785, &H7901EE83, &H421F6A07, &H5E04EF83
    pvAppendBuffer &HE57CD33B, &H8BDC5D8B, &HE4458BFB, &H890C4D8B, &HE7C1F055, &HFD03B02, &H978D&, &HFF68500, &H8088&, &H84D8B00, &H8DFC75FF, &HFF530F04, &H5050F875, &HFFF213E8, &H75FF53FF, &HF475FFFC, &HFFEC75FF, &HE6E8F875, &H8B000003, &HCE8B0C45, &HC033108B, &H40F0552B, &H4D8BE0D3, &H9104850C, &H75FF2C74, &HF8458BFC, &H875FF53, &H75FFC703, &HD5E850E8, &H53FFFFF1, &HFFFC75FF, &H75FFF475
    pvAppendBuffer &H875FFEC, &H3A8E8, &H84D8B00, &H458B0CEB, &HF84D8B08, &H89084D89, &HEE83F845, &H8B897901, &H4D8BF055, &H1F6A420C, &H5EF05589, &H8C0F113B, &HFFFFFF69, &HFC75FF53, &HFFF475FF, &H75FFEC75, &H36BE808, &H458B0000, &HE830FF10, &H5A7&, &HF08BD233, &H1E7EDB85, &H308458B, &HC4589C7, &HE8BF88B, &HCA2B078B, &H47F8D42, &H3B8E0489, &H8BEF7CD3, &H3E83D87D, &H8B107601, &H863C8306
    pvAppendBuffer &H48087500, &HF8830689, &HFFF07701, &H5D8BD475, &H98E853FC, &HE8000015, &HFFFFA2DD, &H850FF53, &H8BE075FF, &HE853085D, &H1583&, &HFFA2C8E8, &H50FF53FF, &HE075FF08, &H53F85D8B, &H156EE8, &HA2B3E800, &HFF53FFFF, &H5D8B0850, &HE85357F4, &H155B&, &HFFA2A0E8, &H50FF53FF, &HEC5D8B08, &H48E85357, &HE8000015, &HFFFFA28D, &H850FF53, &H57E85D8B, &H1535E853, &H7AE80000, &H53FFFFA2
    pvAppendBuffer &H5F0850FF, &H5B5EC68B, &HC25DE58B, &H8B55000C, &H2CEC83EC, &H758B5653, &HFF565710, &H45E80875, &H8BFFFFB4, &HC1FB8B1E, &H458902E7, &HE45D89E0, &HE8D47D89, &HFFFFA241, &H8B10FF57, &H89C933D0, &HDB85F855, &H68B0F7E, &H48BC12B, &H8A048986, &H7CCB3B41, &HA21FE8F1, &HFF57FFFF, &HE0758B10, &HF88BD38B, &H2BEC7D89, &H7ED28516, &H33CA8B09, &H8BABF3C0, &H68BEC7D, &HC085C933, &H148D147E
    pvAppendBuffer &H41C12B97, &H8986048B, &H4528D02, &HC83B068B, &HFB8BEF7C, &H8903E7C1, &HDAE8DC7D, &H57FFFFA1, &HF08B10FF, &HE8F47589, &HFFFFA1CD, &H8910FF57, &H48D0845, &H7EC0851B, &H8BC88B0D, &HF3C033FE, &HC1FB8BAB, &HC75303E7, &H1FC3744, &HE8000000, &H3FF&, &HE6C1F08B, &HD8758902, &HFFA198E8, &H10FF56FF, &H330C4D8B, &HF04589F6, &H75891F6A, &H5F018BFC, &H267EC085, &H8B81148D, &H8BC033D9
    pvAppendBuffer &HE0D340CF, &H10750285, &H7901EF83, &H461F6A07, &H5F04EA83, &HE57C333B, &H89E45D8B, &H758BFC75, &H56368BF8, &HFFB79FE8, &HE84589FF, &H1874C085, &HE6D3C88B, &H7E01FB83, &H59206A0F, &H458BC82B, &H4408BF8, &HF00BE8D3, &HA76E856, &H4D8B0000, &HE445890C, &HE9FC458B, &H94&, &H880FFF85, &H85&, &HFFF4758B, &H48DF075, &H75FF539E, &HE8505008, &HFFFFEF54, &H8DE875FF, &H75FF1B04
    pvAppendBuffer &H53006AE4, &H50F875FF, &HE80875FF, &HFFFFECA0, &H8B0C458B, &H33108BCF, &HFC552BC0, &H8BE0D340, &H4850C4D, &HFF2F7491, &H458BF075, &HFF565308, &H48DEC75, &H11E85098, &HFFFFFFEF, &H48DE875, &HE475FF1B, &HFF53006A, &H5650F875, &HFFEC5FE8, &HC4D8BFF, &HC68B08EB, &H8908758B, &HEF830845, &H8B847901, &H7589FC45, &H1F6A40F4, &H5FFC4589, &H8C0F013B, &HFFFFFF64, &HFF10458B, &H2F5E830
    pvAppendBuffer &HD2330000, &HDB85F08B, &H458B177E, &H983C8DF4, &H78B0E8B, &H8D42CA2B, &H489047F, &H7CD33B8E, &H13E83EF, &H68B1076, &H863C83, &H89480875, &H1F88306, &H7D8BF077, &HF45D8BDC, &HECE85357, &HE8000012, &HFFFFA031, &H850FF53, &H8BD875FF, &HE853F05D, &H12D7&, &HFFA01CE8, &H50FF53FF, &H85D8B08, &HC4E85357, &HE8000012, &HFFFFA009, &H850FF53, &H8BD45D8B, &H5753F87D, &H12AEE8
    pvAppendBuffer &H9FF3E800, &HFF57FFFF, &H8B530850, &HE853EC5D, &H129B&, &HFF9FE0E8, &H50FF53FF, &HE0458B08, &HE8AAE850, &H8B5FFFFF, &H8B5B5EC6, &HCC25DE5, &HEC8B5500, &H5308458B, &H56145D8B, &H187D8B57, &H8D0CF76B, &H4589B804, &H56F30314, &H75FF5357, &H39E85010, &H56FFFFF0, &H8D18758B, &H5756BB3C, &H530C75FF, &HFFEDE3E8, &H85D8BFF, &H5036048D, &HE8535753, &HFFFFEA6D, &H3314558B, &H104589C9
    pvAppendBuffer &HF685F98B, &H48B167E, &H8D0289BB, &HC890452, &HFE3B47BB, &H558BF07C, &H10458B14, &H2A75C085, &H267EF685, &H8B0C5D8B, &H3B078BFA, &H8758B04, &H4C78341, &HF17CCE3B, &HE7DCE3B, &H8D087D8B, &H48B3104, &H8B043B87, &H52560B76, &H520C75FF, &HFFF0F6E8, &H5B5E5FFF, &H14C25D, &H81EC8B55, &H108EC, &HF8A58300, &HFFFFFE, &HFEFCA583, &H8B00FFFF, &H57560C45, &H8D593C6A, &HFFFEF8B5
    pvAppendBuffer &HF845C7FF, &H10&, &HFF00BD8D, &HA5F3FFFF, &HFEF8B58D, &HCE8BFFFF, &H2BFC7589, &HC4589C1, &H300C8B53, &H448BFE8B, &HDB330430, &H8910758B, &H4589F44D, &HFF5150F0, &HFF04DE74, &H1EE8DE34, &H1FFFFAC, &HF44D8B07, &H43045711, &H8DF0458B, &HFB83087F, &H8BDE7210, &H458BFC75, &H8C6830C, &H1F86D83, &H75FC7589, &H5BF633B8, &H266A006A, &H7CF5B4FF, &HFFFFFFFF, &HFF78F5B4, &HDEE8FFFF
    pvAppendBuffer &H1FFFFAB, &HFEF8F584, &H9411FFFF, &HFFFEFCF5, &HFE8346FF, &H8BD5720F, &HB58D087D, &HFFFFFEF8, &HFF59206A, &HA5F30875, &HFFB4F2E8, &H875FFFF, &HFFB4EAE8, &H8B5E5FFF, &HCC25DE5, &HEC8B5500, &H5310EC83, &HDB335756, &HC75FF53, &H1475FF53, &HFFAB8CE8, &H75FF53FF, &HF045890C, &HFF53F28B, &H7AE81875, &H53FFFFAB, &H891075FF, &HFA8BF445, &H1875FF53, &HFFAB68E8, &H75FF53FF, &HFC458910
    pvAppendBuffer &H1475FF53, &HE8F85589, &HFFFFAB55, &H458BD88B, &H6ADE03F4, &HD6135E00, &HD713D803, &HD77D73B, &HD83B0472, &H75010773, &HF85583FC, &H8458B01, &H4D0BC933, &H5FDE0BF0, &H3308895E, &HFC5503C9, &H13045889, &H5089F84D, &HC488908, &H5DE58B5B, &H550014C2, &H4D8BEC8B, &HF6335608, &HC18B0DEB, &HD1C22B99, &H41C82BF8, &H838E348D, &HEE7F32F9, &H5D5EC68B, &H550004C2, &H8B53EC8B, &H5756085D
    pvAppendBuffer &H49D348D, &HE8000000, &HFFFF9D6D, &H5610FF56, &H6AF88B, &HF193E857, &HC483FFFF, &H8B1F890C, &H5B5E5FC7, &H4C25D, &H83EC8B55, &H8B5330EC, &H5756085D, &H75FF066A, &H30E8530C, &H6A00001E, &HFF206A06, &H458D0C75, &H5AE850D0, &H6A000012, &H8DF08B06, &H458D084B, &H50FA8BD0, &HC3835151, &H1123E838, &HC6030000, &H75FF066A, &H1303890C, &H8458BD7, &H8910C083, &H50500453, &H1108E8
    pvAppendBuffer &H84D8B00, &H4189066A, &HD0458D40, &H89515150, &H93E84451, &H8B00001F, &HF003084D, &H518BFA13, &H8BD62B30, &HF71B3471, &H7234713B, &H3B05771D, &H16763051, &HEBFFCF83, &H85B8D03, &H7B113B01, &H23038B04, &HC73B0443, &H895FEF74, &H895E3471, &H8B5B3051, &H8C25DE5, &HEC8B5500, &H108EC81, &H56530000, &HC75FF57, &HFF78858D, &HE850FFFF, &H919&, &HFF78858D, &HE850FFFF, &HFFFFB31F
    pvAppendBuffer &HFF78858D, &HE850FFFF, &HFFFFB313, &HFF78858D, &HE850FFFF, &HFFFFB307, &HFEF89D8D, &H45C7FFFF, &H20C&, &H8BF63300, &HFFFF788D, &H7C858BFF, &H81FFFFFF, &HFFEDE9, &H1B086A00, &HF88D89C6, &H89FFFFFE, &HFFFEFC85, &H548B5FFF, &H448BF83B, &H8C8BFC3B, &HFFFF783D, &HF85589FF, &H10C2AC0F, &H7C3D848B, &H83FFFFFF, &H748901E2, &HCA2BFC3B, &HE981C61B, &HFFFF&, &HF83D8C89, &H1BFFFFFE
    pvAppendBuffer &H3D8489C6, &HFFFFFEFC, &HF845B70F, &HF83B4489, &H8308C783, &HB27278FF, &HFF688D8B, &H858BFFFF, &HFFFFFF6C, &HFF0558B, &HF10C1AC, &HFF6885B7, &HE183FFFF, &H68858901, &H2BFFFFFF, &H6CB589D1, &H8BFFFFFF, &HCE1BF44D, &H7FFFEA81, &H95890000, &HFFFFFF70, &HC033CE1B, &HFF748D89, &HF40FFFF, &H8310CAAC, &HF9C101E2, &H50C22B10, &HFEF8858D, &H8D50FFFF, &HFFFF7885, &HB2E850FF, &H83000007
    pvAppendBuffer &HF010C6D, &HFFFF1E85, &H8558BFF, &H78F5848A, &H8BFFFFFF, &HFF78F58C, &H488FFFF, &HF5848B72, &HFFFFFF7C, &H8C1AC0F, &H8808F8C1, &H4601724C, &H7210FE83, &H5B5E5FD7, &HC25DE58B, &H8B550008, &H84D8BEC, &H8B56D233, &H6A570C75, &H5FF12B11, &H30E048B, &HFD00301, &HEAC1C2B6, &H8D018908, &HEF830449, &H5FE97501, &H8C25D5E, &HEC8B5500, &HC75FF56, &H5608758B, &HFFFFC1E8, &H44468DFF
    pvAppendBuffer &H15E85650, &H5E000001, &H8C25D, &H83EC8B55, &H8B5344EC, &H5756085D, &H8B59116A, &HBC7D8DF3, &HD2E8A5F3, &H5FFFF9A, &H544&, &HBC458D50, &HFF88E850, &H458BFFFF, &H25D0F7FC, &H80&, &H8D5E116A, &HD0F7FF50, &HC11FEAC1, &HD0231FE8, &HF7BC458D, &H89C32BD2, &H7A8D0845, &H8BD7F701, &H428D180C, &H23032301, &H8BC80BCF, &HB890845, &H83045B8D, &HE77501EE, &H8B5B5E5F, &H4C25DE5
    pvAppendBuffer &HEC8B5500, &H8D44EC83, &H6A56BC45, &H56F63344, &HEE97E850, &H4D8BFFFF, &HCC48308, &HA8918B, &HD2850000, &HB60F1174, &H983184, &H44890000, &H3B46BCB5, &H8DEF72F2, &H44C7BC45, &H1BC95, &H51500000, &HFFFF24E8, &HE58B5EFF, &H4C25D, &H8BEC8B55, &H3356084D, &HD68B57F6, &H1403FE8B, &HC2B60FB9, &H8908EAC1, &H8347B904, &HEE7C10FF, &H8B405103, &H2EAC1C2, &H6B03E083, &H418905D2
    pvAppendBuffer &HB1140340, &HC1C2B60F, &H48908EA, &HFE8346B1, &H1EE7C10, &H5E5F4051, &H4C25D, &H83EC8B55, &H8B534CEC, &H458D0C5D, &H2BD233B4, &H895756D8, &HFF33FC5D, &HD285C933, &H48D1E78, &H85D8B93, &H3B4758D, &H8B048BF0, &H8D06AF0F, &HF803FC76, &H7ECA3B41, &HFC5D8BF0, &H8901728D, &HCE8BF875, &H7D11FE83, &H85D8B2D, &HC62BC28B, &H830C758B, &H348D44C6, &H8B048B86, &H8D06AF0F, &HC069FC76
    pvAppendBuffer &H140&, &H8341F803, &HE97C11F9, &H8BF8758B, &H7C89FC5D, &HD68BB495, &H7C11FA83, &HB4458D95, &HFF2AE850, &H7D8BFFFF, &HB4758D08, &HF359116A, &H5B5E5FA5, &HC25DE58B, &H8B550008, &HC558BEC, &H3344EC83, &H4B60FC9, &H8D448911, &HF98341BC, &H8DF27C10, &H45C7BC45, &H1FC&, &H75FF5000, &HFE07E808, &HE58BFFFF, &H8C25D, &H81EC8B55, &H178EC, &H57565300, &H7D8DC033, &H88DB33B1
    pvAppendBuffer &H6AABB05D, &HC75FF0C, &HAB66ABAB, &HB4458DAA, &HECF2E850, &HC483FFFF, &HD05D880C, &H7D8DC033, &H59076AD1, &H46AABF3, &H8DAAAB66, &H6A50B045, &H875FF20, &HFF34858D, &HE850FFFF, &HFFFFC040, &H458D206A, &H8D5050D0, &HFFFF3485, &H10E850FF, &H8DFFFFBB, &H8D50E045, &H8D50D045, &HFFFE8885, &HD7E850FF, &H6AFFFFCA, &HC0335908, &HF3D07D8D, &HD0458AAB, &H458D206A, &H8D5050D0, &HFFFF3485
    pvAppendBuffer &HDCE850FF, &H33FFFFBA, &HD07D8DC0, &HFF59086A, &HABF31475, &H8DD0458A, &H75FFF17D, &H88C03310, &HABABF05D, &HAAAB66AB, &HFE88858D, &HE850FFFF, &HFFFFCB4A, &HF714458B, &HFE083D8, &HF0458D50, &H88858D50, &H50FFFFFE, &HFFCB31E8, &H1C758BFF, &H7D8BC033, &H45894018, &H4539560C, &HFF167524, &H858D2075, &HFFFFFF34, &H74E85057, &H56FFFFBA, &HEB2075FF, &H858D5701, &HFFFFFE88, &HCAFCE850
    pvAppendBuffer &HC68BFFFF, &HE083D8F7, &H458D500F, &H858D50F0, &HFFFFFE88, &HCAE4E850, &H458DFFFF, &HFF5350F0, &HA0E81475, &H8D00001B, &H5350F845, &H1B95E856, &H106A0000, &H458D535B, &H858D50F0, &HFFFFFE88, &HCAB8E850, &H7D83FFFF, &H2C750124, &H8D2875FF, &HFFFE8885, &H78E850FF, &H6AFFFFC9, &H34858D7C, &H6AFFFFFF, &HC2E85000, &H8AFFFFEB, &HFFFF3485, &HCC483FF, &H73EBC033, &H50C0458D, &HFE88858D
    pvAppendBuffer &HE850FFFF, &HFFFFC94B, &H8D287D8B, &HC18BC04D, &HF82BD232, &H320F048A, &H41D00A01, &H7501EB83, &H187D8BF3, &H2075FF56, &H1275D284, &H34858D57, &H50FFFFFF, &HFFB9AAE8, &HC5D21FF, &H23E805EB, &H6AFFFFEB, &H34858D7C, &H6AFFFFFF, &H5AE85000, &H8AFFFFEB, &HFFFF348D, &HC07D8DFF, &H330CC483, &HABABABC0, &HC04D8AAB, &H5F0C458B, &HE58B5B5E, &H24C25D, &H8BEC8B55, &HB60F0855, &H4AB60F02
    pvAppendBuffer &H8E0C101, &HB60FC10B, &HE0C1024A, &HFC10B08, &HC1034AB6, &HC10B08E0, &H4C25D, &H8BEC8B55, &HB60F0855, &HB60F0342, &HE0C1024A, &HFC10B08, &HC1014AB6, &HC10B08E0, &HC10AB60F, &HC10B08E0, &H4C25D, &H53EC8B55, &H33085D8B, &HC1CB8BD2, &H564110E9, &H8D006A57, &HFFFFFF81, &H8BF1F77F, &H10E6C1F0, &HE3F7C68B, &HCE03C88B, &HD283D1F7, &H83C03300, &HD2F701C1, &HC203C013, &HE8C1E6F7
    pvAppendBuffer &H3FA8B1F, &H8BF80BFF, &H8BE3F7C7, &H3CA8BF0, &HBAF7&, &H13588000, &H72CA3BC8, &H33D3F718, &H13F303C0, &H1C683C0, &H4900D083, &H3B4FC803, &HEBEC73CA, &H3D23319, &HEBD013F3, &HC9334707, &HC813F303, &HFA81D103, &H80000000, &HC78BEF72, &H5D5B5E5F, &H550004C2, &H458BEC8B, &HC558B08, &H406A5756, &H104D2B59, &HFFA397E8, &H104D8BFF, &H458BF08B, &H8BFA8B08, &HA4E80C55, &HBFFFFA3
    pvAppendBuffer &H5FD70BC6, &HCC25D5E, &HEC8B5500, &H5324EC83, &H7D8B5756, &H75FF5708, &HAA95E80C, &HFF57FFFF, &H45891075, &HAA89E808, &HFF57FFFF, &H45891475, &HAA7DE80C, &HF78BFFFF, &HD1F04589, &H75FF56EE, &HAA6DE818, &HFF56FFFF, &HD88B1C75, &HE8EC5D89, &HFFFFAA5F, &H2075FF56, &HE8104589, &HFFFFAA53, &HE8458953, &HFFCCE6E8, &H458950FF, &HCD45E81C, &H758BFFFF, &HD4E85610, &H50FFFFCC, &HE8204589
    pvAppendBuffer &HFFFFCD33, &HFF1C75FF, &H59E80C75, &HFFFFFFA7, &H45892075, &HC75FFFC, &HFFA74BE8, &H75FF53FF, &HF84589FC, &HE80875FF, &HFFFFEF57, &HF875FF56, &H75FFD88B, &H145D8908, &HFFEF46E8, &H56F08BFF, &H18758953, &HFFA96EE8, &H79C085FF, &HEC75FF19, &HE853F38B, &HFFFFA52F, &H89D88B56, &HEDE81445, &H8BFFFFDD, &H53561875, &HFFAA9DE8, &H1075FFFF, &HFFF44589, &H27E8E875, &H56FFFFA7, &H89F475FF
    pvAppendBuffer &HE850E045, &HFFFFA72E, &H89F075FF, &HE850DC45, &HFFFFA6D4, &H85E44589, &H8B1A74FF, &HD88B2475, &HE853574F, &HFFFFA8CB, &H85460688, &H8BF175FF, &H758B145D, &H1C75FF18, &HFFDD93E8, &H2075FFFF, &HFFDD8BE8, &HFC75FFFF, &HFFDD83E8, &HF875FFFF, &HFFDD7BE8, &H75E853FF, &H56FFFFDD, &HFFDD6FE8, &HF475FFFF, &HFFDD67E8, &HE075FFFF, &HFFDD5FE8, &HDC75FFFF, &HFFDD57E8, &H875FFFF, &HFFDD4FE8
    pvAppendBuffer &HC75FFFF, &HFFDD47E8, &HF075FFFF, &HFFDD3FE8, &HEC75FFFF, &HFFDD37E8, &H1075FFFF, &HFFDD2FE8, &HE875FFFF, &HFFDD27E8, &HE475FFFF, &HFFDD1FE8, &H5B5E5FFF, &HC25DE58B, &H8B550020, &H8B5653EC, &H56570875, &HE80C75FF, &HFFFFA8E7, &H1075FF56, &H5D89D88B, &HA8D9E80C, &HFF56FFFF, &HF88B1475, &HE8087D89, &HFFFFA8CB, &H89535750, &H10E81045, &H8BFFFFEE, &H74F685D8, &H187D8B15, &HE853564E
    pvAppendBuffer &HFFFFA7EF, &H85470788, &H8BF175F6, &H75FF087D, &HDCBAE80C, &HE857FFFF, &HFFFFDCB4, &HE81075FF, &HFFFFDCAC, &HDCA6E853, &H5E5FFFFF, &H14C25D5B, &HEC8B5500, &H458B5151, &H5D8B5310, &HF7564808, &H1045C7D0, &H10&, &H7D8B5799, &H89DF2B0C, &H5589FC45, &H3B348BF8, &H43B548B, &H4F8B078B, &H23C63304, &HCA33FC45, &H33F84D23, &H89D133F0, &H54893B34, &H731043B, &H31087F8D, &H6D83FC4F
    pvAppendBuffer &HD1750110, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H8B0C558B, &HD12B084D, &H5E106A56, &H890A048B, &H8498D01, &HFC0A448B, &H83FC4189, &HEC7501EE, &H8C25D5E, &HEC8B5500, &H106A5653, &H105D395B, &H7D833674, &H5A752010, &H570C758B, &H53087D8B, &H31E85756, &H53FFFFE7, &H5010468D, &H5010478D, &HFFE723E8, &H18C483FF, &HFF930CE8, &H53105FF, &H47890000, &H2AEB5F30, &H5308758B, &H560C75FF
    pvAppendBuffer &HFFE703E8, &H75FF53FF, &H10468D0C, &HE6F6E850, &HC483FFFF, &H92DFE818, &H2005FFFF, &H89000005, &H5B5E3046, &HCC25D, &H83EC8B55, &H8B536CEC, &H5756085D, &HA0BFF633, &H8B000001, &H104B8B03, &H8BF84589, &H45890443, &H8438BF0, &H8BEC4589, &H45890C43, &H14438BE0, &H8BE84589, &H45891843, &H1C438BE4, &H89945D8D, &HDF2BF44D, &H89DC4589, &H7D89D875, &H10FE83D4, &H758B1773, &H71E8560C
    pvAppendBuffer &H83FFFFFB, &H458904C6, &H3B0489FC, &HEB0C7589, &H17E8D5D, &H8D0FE683, &HE083FD47, &H85548B0F, &H83C78B94, &H4C8B0FE0, &HC18B9485, &H890EC0C1, &HC18BFC45, &H3107C8C1, &HC28BFC45, &H3103E9C1, &HCA8BFC4D, &HC10DC0C1, &HC8330FC1, &H330AEAC1, &HF8478DCA, &H83FC558B, &H7D8B0FE0, &H3D103D4, &H3948554, &H8994B554, &H5489FC55, &H6E894B5, &H8BFFFF92, &HD68BF475, &HCAC1CE8B, &H7C1C10B
    pvAppendBuffer &HCE8BD133, &HF706C9C1, &HE47523D6, &HC8BD133, &H4C78338, &H3E8458B, &HF44523CA, &H33FC4D03, &H89F103F0, &H4D8BD47D, &H3D18BF8, &HC18BDC75, &HC10AC0C1, &HD0330DCA, &HC8C1C18B, &H8BD03302, &HC88BF045, &H33F84523, &H4D23F84D, &H8BC833EC, &H4589E445, &H8BD103DC, &H4D8BE845, &HE44589E0, &H458BCE03, &HE84589F4, &H89EC458B, &H458BE045, &HEC4589F0, &H89F8458B, &H48DF045, &HD8758B32
    pvAppendBuffer &HF44D8946, &H89F84589, &HFF81D875, &H2A0&, &HFEDF820F, &H5D8BFFFF, &H15E5F08, &HF0458B03, &H8B044301, &H4301EC45, &HE0458B08, &H8B0C4301, &H4301E845, &HE4458B14, &H1184301, &H458B104B, &H1C4301DC, &H5B6043FF, &HC25DE58B, &H8B550008, &HDCEC81EC, &H8B000000, &H56530845, &H4488B57, &H8BF84D89, &H108B0848, &H8BD84D89, &H4D890C48, &H10488BEC, &H8BD44D89, &H4D891448, &H18488BD0
    pvAppendBuffer &H8BB84D89, &H4D891C48, &H20488BB4, &H8BE84D89, &H4D892448, &H28488BFC, &H8BCC4D89, &H4D892C48, &H30488BC8, &H8BC44D89, &H4D893448, &H38488BC0, &H893C408B, &H858DAC45, &HFFFFFF24, &HB9B04D89, &H2A0&, &H33E05589, &H89C12BD2, &H4D89BC55, &HA44589DC, &H7310FA83, &HC7D8B36, &H5004478D, &HFFF98FE8, &HD88B57FF, &H85E8F633, &H8BFFFFF9, &HF00BDC4D, &H83A4458B, &H5D8908C7, &HF07589F4
    pvAppendBuffer &H890C7D89, &H7489081C, &HD6E90408, &H8D000000, &HE083FE42, &H8B3D6A0F, &HFF24C5BC, &HB48BFFFF, &HFFFF28C5, &H1428DFF, &H830FE083, &H57560FE2, &H8BA85589, &HFF24C58C, &H9C8BFFFF, &HFFFF28C5, &HE44D89FF, &HFFFA0EE8, &H56136AFF, &HF4458957, &HE8F05589, &HFFFFF9FF, &H33F44D8B, &HF0458BC8, &H6F7AC0F, &HC233086A, &H5306EEC1, &H33E475FF, &H89C633CF, &H4589F44D, &HF9D9E8F0, &H16AFFFF
    pvAppendBuffer &HE475FF53, &HFA8BF08B, &HFFF9CAE8, &HF44D8BFF, &H458BF033, &H8BFA33E4, &HAC0FF055, &HF03307D8, &H8B07EBC1, &HFB33BC45, &HD713CE03, &H83F9C083, &H8C030FE0, &HFFFF24C5, &HC59413FF, &HFFFFFF28, &H3A8458B, &HFF24C58C, &H4D89FFFF, &HC59413F4, &HFFFFFF28, &H24C58C89, &H89FFFFFF, &H9489F055, &HFFFF28C5, &HE8758BFF, &H75FF296A, &H64E856FC, &H6AFFFFF9, &HFC75FF12, &HFA8BD88B, &HF955E856
    pvAppendBuffer &HE6AFFFF, &H33FC75FF, &H56FA33D8, &HFFF946E8, &H33D833FF, &H8F5FE8FA, &H4D8BFFFF, &H8BD6F7DC, &HD2F7FC55, &H1C03276A, &HF875FF08, &H4087C13, &H23C47523, &H4D8BC055, &HE84D23CC, &H33C8458B, &HFC4523F1, &HDE03D033, &H13E0758B, &HF45D03FA, &HF07D1356, &H89B05D03, &H7D13A85D, &HE47D89AC, &HFFF8F2E8, &HFF226AFF, &HF88BF875, &HE856DA8B, &HFFFFF8E3, &H75FF1C6A, &H33F833F8, &HD4E856DA
    pvAppendBuffer &H8BFFFFF8, &HDA33D84D, &H33EC558B, &HF85533F8, &H7533F18B, &HD05523E0, &H23EC458B, &H7523F845, &H23D033D4, &H458BE04D, &H89F133C4, &HF703B045, &H89C0458B, &HDA13AC45, &H89CC458B, &H458BC445, &HA84D8BC8, &H8BB84D03, &H5513E455, &HA87503B4, &H8BC04589, &H5D13E845, &HCC4589E4, &H89FC458B, &H558BFC55, &HB85589D4, &H89D0558B, &H558BB455, &HD45589D8, &H89EC558B, &H558BD055, &HD85589E0
    pvAppendBuffer &H89F8558B, &H4D8BE84D, &HEC5589DC, &H8B08C183, &H8942BC55, &H7589C845, &HF85D89E0, &H89BC5589, &HF981DC4D, &H520&, &HFDA6820F, &H458BFFFF, &HD8558B08, &H5FD44D8B, &H758B3001, &H45811B4, &H8B085001, &H5011EC55, &H1048010C, &H11D04D8B, &H558B1448, &H185001B8, &H11E84D8B, &H48011C70, &HFC4D8B20, &H8B244811, &H4801CC4D, &HC84D8B28, &H8B2C4811, &H4801C44D, &HC04D8B30, &H8B344811
    pvAppendBuffer &H4801B04D, &HAC4D8B38, &HFF3C4811, &HC080&, &H8B5B5E00, &H8C25DE5, &HEC8B5500, &H85D8B53, &HB60F5756, &HB60F077B, &HB60F0A43, &HB60F0B73, &HE7C10F53, &HFF80B08, &HF034BB6, &HC10D43B6, &HF80B08E7, &HF08E6C1, &HE7C103B6, &HC1F80B08, &HB60F08E2, &HF00B0E43, &HF08E1C1, &HC10143B6, &HF00B08E6, &H443B60F, &HB08E6C1, &H43B60FF0, &HFD00B02, &HC10543B6, &HD00B08E2, &H843B60F
    pvAppendBuffer &HB08E2C1, &H43B60FD0, &H89C80B06, &HB60F047B, &HE1C10943, &H89C80B08, &HB60F0873, &HE1C10C43, &HC80B5F08, &H5E0C5389, &H5D5B0B89, &H550004C2, &H458BEC8B, &H74C08508, &HC4D8B10, &H974C985, &H400000C6, &H7501E983, &H8C25DF7, &HEC8B5500, &HFF0C75FF, &H75FF0C75, &HEDFEE808, &HC25DFFFF, &H8B550008, &H10558BEC, &H758B5653, &H7D8B570C, &H6AF22B08, &H5BFA2B10, &H2B160C8B, &H16448B0A
    pvAppendBuffer &H4421B04, &H8D170C89, &H44890852, &HEB83FC17, &H5FE57501, &HC25D5B5E, &H8B55000C, &HE85756EC, &HFFFF8CD5, &HBF08758B, &H588&, &HFF50C703, &H41E836, &H6890000, &HFF8CBCE8, &H50C703FF, &HE80476FF, &H2F&, &HE8044689, &HFFFF8CA9, &HFF50C703, &H1CE80876, &H89000000, &H96E80846, &H3FFFF8C, &H76FF50C7, &H9E80C, &H895F0000, &H5D5E0C46, &H550004C2, &H558BEC8B, &H5D8B530C
    pvAppendBuffer &HC1C38B08, &HCB8B18E8, &H8E9C156, &HFC9B60F, &H8B1034B6, &H10E8C1C3, &HFC0B60F, &HC1110CB6, &HB60F08E6, &HC60B1004, &HB08E0C1, &HCBB60FC1, &H5E08E0C1, &HCB60F5B, &H5DC10B11, &H550008C2, &H5653EC8B, &H87D8B57, &H458BDB33, &H44B60F0C, &H8B990158, &H8BF28BC8, &HA40F0C45, &HE1C108CE, &H4B60F08, &HC8039958, &H13DF0C89, &HDF7489F2, &HFB834304, &H81D37210, &H7FFF7867, &H67830000
    pvAppendBuffer &H5E5F007C, &H8C25D5B, &HEC8B5500, &H145D8B53, &HD233C933, &H6774DB85, &H560C458B, &H2B10758B, &H84529F0, &H10758957, &H306348B, &HC758930, &H8B10758B, &H748B0C7D, &H70130406, &H89F90304, &HF2130C7D, &H7D8B383B, &H3B057508, &H17740470, &H7704703B, &H8B07720E, &HC4D3908, &HC9330573, &H3302EB41, &H8BD233C9, &H74890C5D, &H758B0407, &H71C8910, &H8308C083, &H7501146D, &H8B5E5FAB
    pvAppendBuffer &HC25D5BC1, &H8B550010, &HC4D8BEC, &H1F74C985, &H7D8B5756, &HCD0C8D08, &HFFFFFFFB, &HE9C1F78B, &H278302, &H46783, &HF308C783, &H5D5E5FA5, &H550008C2, &H4D8BEC8B, &H1E98310, &H30785756, &H8B0C458B, &HF02B0875, &H8BC8148D, &H8B04167C, &H7A3B1604, &H72267704, &H77023B1F, &H47A3B20, &H4771672, &H1072023B, &H8308EA83, &HDB7901E9, &H5E5FC033, &HCC25D, &HEBFFC883, &H40C033F5
    pvAppendBuffer &H8B55F0EB, &H39C933EC, &H12760C4D, &H8B08558B, &H440BCA04, &HD7504CA, &HC4D3B41, &HC033F172, &H8C25D40, &HEBC03300, &HEC8B55F8, &H33535151, &H89D38BDB, &H5539FC55, &H8B5B7614, &H406A084D, &H10452B58, &H8BF84589, &HC12B0C45, &HC458956, &H8748B57, &H8BD68B04, &HC78B083C, &HE8104D8B, &HFFFF9820, &HB084D8B, &HFC550BC3, &HC78B0189, &H8B045189, &HF84D8BD6, &HFF9826E8, &H84D8BFF
    pvAppendBuffer &H458BD88B, &H8C1830C, &H1146D83, &H89FC5589, &HBD75084D, &HC38B5E5F, &H5DE58B5B, &H550010C2, &HEC83EC8B, &H8B565328, &H8B570C75, &H46A087D, &H38E85756, &H8B00000B, &HC9332C46, &H8BE44589, &H45893046, &H34468BE8, &H8BEC4589, &H45893846, &H3C468BF0, &H4589046A, &HD8458DF4, &H5050016A, &H89D84D89, &H4D89DC4D, &HFF3BE8E0, &H46AFFFF, &H458DD88B, &H575750D8, &HFFFE0CE8, &H384E8BFF
    pvAppendBuffer &H468BD803, &H3C568B30, &HE06583, &HF46583, &H33E44589, &H34460BC0, &H4589046A, &HD8458DE8, &H5050016A, &H89EC4D89, &HFAE8F055, &H6AFFFFFE, &H8DD80304, &H5750D845, &HFDCBE857, &HD803FFFF, &H8920468B, &H468BD845, &HDC458924, &H8928468B, &HC033E045, &H89E44589, &H4589E845, &H38468BEC, &H8BF04589, &H46A3C46, &H8DF44589, &H5750D845, &HFD93E857, &HD803FFFF, &H33244E8B, &H34568BC0
    pvAppendBuffer &H8928460B, &H468BDC45, &HF8458930, &H460BC033, &HE045892C, &H8938468B, &H468BE845, &HEC45893C, &H460BC033, &H89046A20, &H458DF445, &H4D8950D8, &H57CA8BD8, &HE44D8957, &HE8F05589, &HFFFFFD49, &H334568B, &H2C4E8BD8, &H460BC033, &HDC458930, &H8920468B, &HD233E055, &H4D89C20B, &H89046AD8, &HC933F045, &H8D284E0B, &H5750D845, &HE4558957, &H89E85589, &H4D89EC55, &HBACE8F4, &H568B0000
    pvAppendBuffer &H8BD82B24, &H20B13046, &HE86583, &HEC6583, &H8BD84589, &H45893446, &H38468BDC, &H8BE04589, &H45893C46, &H20468BE4, &HFF9676E8, &H2C560BFF, &H4589046A, &HD8458DF0, &H89575750, &H67E8F455, &H8B00000B, &HD82B344E, &H8BFC5D89, &H568B305E, &HBC03320, &H7E8B3846, &HDC458924, &H460BC033, &H89F6333C, &H558BE455, &HE045890C, &HB1D84D89, &H28428B20, &HE82C528B, &HFFFF9608, &HF06583
    pvAppendBuffer &H46AF80B, &H8DF45D89, &H535BD845, &HE87D8950, &H7D8BF20B, &H89575708, &HBE8EC75, &H8B00000B, &H758B0C4D, &HE06583FC, &H83F02B00, &H8B00F065, &H45893841, &H3C418BD8, &H8BDC4589, &H45892441, &H28418BE4, &H8BE84589, &H45892C41, &H34418BEC, &HF4458953, &H50D8458D, &HCBE85757, &H2B00000A, &H531E79F0, &HFF8810E8, &H10C083FF, &HE8575750, &HFFFFFC15, &HEB78F003, &H8B5B5E5F, &H8C25DE5
    pvAppendBuffer &H75F68500, &HE8575315, &HFFFF87ED, &H5010C083, &HFFFC9EE8, &H1F883FF, &HE853DE74, &HFFFF87D9, &H5010C083, &H7FE85757, &H2B00000A, &H55D2EBF0, &HEC83EC8B, &H758B5668, &H66A570C, &H57307E8D, &HFFFCBDE8, &HFC085FF, &H9385&, &H5D8D5300, &H83DE2B98, &H8D00FC65, &H65839845, &H66A00F8, &HFC20E850, &H66AFFFF, &H50C8458D, &HFFFC15E8, &H458D57FF, &H2DE85098, &H6AFFFFEA, &H3E85706
    pvAppendBuffer &H6AFFFFFC, &H5EC68B09, &H303148B, &H34C8B10, &H4481304, &H13FC5503, &H103BF84D, &H483B0575, &H3B1C7404, &HF770448, &H103B0472, &H45C70973, &H1FC&, &H8304EB00, &H8300FC65, &H8900F865, &H4488910, &H8308C083, &HBC7501EE, &HE857066A, &HFFFFFC2E, &H850C758B, &H74840FC0, &H5BFFFFFF, &H12E8066A, &HBFFFFF87, &HB0&, &H6E816EB, &H3FFFF87, &H565650C7, &H9ADE8, &HE8066A00
    pvAppendBuffer &HFFFF86F5, &H5650C703, &HFFFBA6E8, &H85066AFF, &H56DB7FC0, &HE80875FF, &H7D3&, &HE58B5E5F, &H8C25D, &H83EC8B55, &HFF04107D, &H75FF0C75, &HE8077508, &HFFFFFC67, &HF0E805EB, &H5DFFFFFE, &H55000CC2, &H75FFEC8B, &H1075FF18, &HFF0C75FF, &HB2E80875, &HBFFFFFA, &HFF1275C2, &H75FF1875, &H875FF14, &HFFFB4AE8, &H78C085FF, &H1875FF11, &HFF1475FF, &H75FF0875, &H92CE808, &HC25D0000
    pvAppendBuffer &H8B550014, &HC0EC81EC, &H56000000, &H5614758B, &HE80C75FF, &HFFFFFB6A, &H74C08556, &H875FF0D, &HFFFAE1E8, &H1F1E9FF, &H75FF0000, &H40858D0C, &H50FFFFFF, &H736E8, &H75FF5600, &H70858D10, &H50FFFFFF, &H726E8, &H458D5600, &HB3E850D0, &H83FFFFFA, &H8D00D465, &H5056A045, &H1D045C7, &HE8000000, &HFFFFFA9E, &H70858D56, &H50FFFFFF, &HFF40858D, &HE850FFFF, &HFFFFFAB7, &HC985C88B
    pvAppendBuffer &H188840F, &H57530000, &HFF40858B, &HDB33FFFF, &H3301E083, &H56C30BFF, &H858D0E75, &HFFFFFF40, &H684E850, &H77EB0000, &HFF70858B, &HE083FFFF, &HC88301, &H858D1175, &HFFFFFF70, &H668E850, &HECE90000, &H85000000, &H948E0FC9, &H8D000000, &HFFFF7085, &H858D50FF, &HFFFFFF40, &H47E85050, &H56000008, &HFF40858D, &HE850FFFF, &H63A&, &HA0458D56, &HD0458D50, &HFA35E850, &HC085FFFF
    pvAppendBuffer &HFF560E79, &H458D1075, &HE85050D0, &HFFFFF979, &HA0458D56, &HD0458D50, &HBE85050, &H8B000008, &HE083D045, &HC88301, &HFF561274, &H458D1075, &HE85050D0, &HFFFFF951, &HDA8BF88B, &HD0458D56, &H5E4E850, &HFB0B0000, &HA0840F, &H448B0000, &H4C81C8F5, &HCCF5&, &H44898000, &H8BE9C8F5, &H8D000000, &HFFFF4085, &H858D50FF, &HFFFFFF70, &HB3E85050, &H56000007, &HFF70858D, &HE850FFFF
    pvAppendBuffer &H5A6&, &HD0458D56, &HA0458D50, &HF9A1E850, &HC085FFFF, &HFF560E79, &H458D1075, &HE85050A0, &HFFFFF8E5, &HD0458D56, &HA0458D50, &H77E85050, &H8B000007, &HE083A045, &HC88301, &HFF561274, &H458D1075, &HE85050A0, &HFFFFF8BD, &HDA8BF88B, &HA0458D56, &H550E850, &HFB0B0000, &H448B1074, &H4C8198F5, &H9CF5&, &H44898000, &H8D5698F5, &HFFFF7085, &H858D50FF, &HFFFFFF40, &HF931E850
    pvAppendBuffer &HC88BFFFF, &H850FC985, &HFFFFFE7C, &H8D565B5F, &HFF50D045, &H54E80875, &H5E000005, &HC25DE58B, &H8B550010, &H80EC81EC, &H53000000, &H46A5756, &H75FF535B, &H48FE814, &HFF530000, &HF08B1075, &HFF80458D, &HE8500C75, &H34A&, &HA0458D53, &H473E850, &HF88B0000, &H874FF85, &H100C781, &HCEB0000, &H80458D53, &H45BE850, &HF88B0000, &H73FE3B53, &H80458D0C, &H875FF50, &HFAE9&
    pvAppendBuffer &HC0458D00, &HF87CE850, &H8D53FFFF, &HE850E045, &HFFFFF872, &HC62BC78B, &HEEC1F08B, &HE0835306, &H501A743F, &H8D1475FF, &H48DC045, &HF6E850F0, &H89FFFFF8, &H89E0F544, &HEBE4F554, &H1475FF0F, &H8DC0458D, &HE850F004, &H4A3&, &H85D8B53, &HF830E853, &H6383FFFF, &H3C70004, &H1&, &H815E046A, &H100FF, &H56117700, &H8D1475FF, &HE850C045, &HFFFFF83B, &H7978C085, &HA0458D56
    pvAppendBuffer &HE0458D50, &HF829E850, &HC085FFFF, &H40751478, &H80458D56, &HC0458D50, &HF815E850, &HC085FFFF, &H8D562E7F, &H8D50C045, &H50508045, &H5F9E8, &H74C20B00, &H8D53560C, &H5050A045, &H5E9E8, &H458D5600, &H458D50E0, &HE85050A0, &H5DA&, &H8DE0758B, &H46AE045, &H1FE6C150, &H3C9E8, &H8D046A00, &HE850C045, &H3BE&, &H4FDC7509, &HFFFF6BE9, &H458D56FF, &HE8535080, &H3EF&
    pvAppendBuffer &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &HC0EC81, &H56530000, &H5B066A57, &H1475FF53, &H328E8, &H75FF5300, &H8DF08B10, &HFFFF4085, &HC75FFFF, &H1E0E850, &H8D530000, &HFFFF7085, &H6E850FF, &H8B000003, &H74FF85F8, &H80C78108, &HEB000001, &H858D530F, &HFFFFFF40, &H2EBE850, &HF88B0000, &H73FE3B53, &H40858D0F, &H50FFFFFF, &HE90875FF, &H110&, &H50A0458D, &HFFF709E8, &H458D53FF
    pvAppendBuffer &HFFE850D0, &H8BFFFFF6, &H8BC62BC7, &H6EEC1F0, &H3FE08353, &HFF501A74, &H458D1475, &HF0048DA0, &HF783E850, &H4489FFFF, &H5489D0F5, &HFEBD4F5, &H8D1475FF, &H48DA045, &H30E850F0, &H53000003, &H53085D8B, &HFFF6BDE8, &H46383FF, &H103C700, &H6A000000, &HFF815E06, &H180&, &HFF561577, &H458D1475, &HC8E850A0, &H85FFFFF6, &H88880FC0, &H56000000, &HFF70858D, &H8D50FFFF, &HE850D045
    pvAppendBuffer &HFFFFF6AF, &H1778C085, &H8D564C75, &HFFFF4085, &H458D50FF, &H98E850A0, &H85FFFFF6, &H56377FC0, &H50A0458D, &HFF40858D, &H5050FFFF, &H479E8, &H74C20B00, &H8D53560F, &HFFFF7085, &HE85050FF, &H466&, &HD0458D56, &H70858D50, &H50FFFFFF, &H454E850, &H758B0000, &HD0458DD0, &HC150066A, &H43E81FE6, &H6A000002, &HA0458D06, &H238E850, &H75090000, &H58E94FCC, &H56FFFFFF, &HFF40858D
    pvAppendBuffer &H5350FFFF, &H266E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H60EC83EC, &HFFA0458D, &H75FF1475, &HC75FF10, &H6CE850, &H75FF0000, &HA0458D14, &H875FF50, &HFFFA6BE8, &H5DE58BFF, &H550010C2, &HEC83EC8B, &HA0458D60, &HFF1075FF, &HE8500C75, &H247&, &H8D1075FF, &HFF50A045, &H41E80875, &H8BFFFFFA, &HCC25DE5, &HEC8B5500, &HFF1875FF, &H75FF1075, &H875FF0C, &H3ADE8, &H74C20B00
    pvAppendBuffer &H1875FF11, &HFF1475FF, &H75FF0875, &HF4F7E808, &HC25DFFFF, &H8B550014, &H54EC83EC, &H3314558B, &H335653C0, &H21C933F6, &H2140E475, &H3357E875, &HDC7589FF, &H7D89C22B, &HF04D89E0, &H3EC4589, &H1BCA3BC1, &HFC6583DB, &HF8658300, &H23D3F700, &HFD93BD8, &HA387&, &H10758B00, &HC32BC18B, &H8BC6348D, &H7589E445, &HFDA3BF4, &H8383&, &H476FF00, &HFF0C458B, &HD874FF36, &HD834FF04
    pvAppendBuffer &H50BC458D, &HFFE234E8, &H8DF08BFF, &HEC83CC7D, &HA5A5A510, &H8BFC8BA5, &H10EC83F0, &HA5AC458D, &H8BA5A5A5, &HDC758DFC, &HA5A5A550, &H8E48E8A5, &HF08BFFFF, &HA5DC7D8D, &H8BA5A5A5, &H453BE845, &HE4458BD8, &H5720C77, &H73D4453B, &H41C93305, &HC93302EB, &H33F4758B, &HFC4D01D2, &H11F04D8B, &H8B43F855, &HEE831455, &HF4758908, &H860FD93B, &HFFFFFF75, &H8BE07D8B, &H3EBDC75, &H8BE4458B
    pvAppendBuffer &H3489085D, &H8BF08BCB, &H4589FC45, &HF8458BE4, &H4CB7C89, &HE87D8B41, &H8DE84589, &HFFFF5504, &HC83BFFFF, &H8BDC7589, &H7D89EC45, &HF04D89E0, &HFF09820F, &HD203FFFF, &HFCD37C89, &HD374895F, &H8B5B5EF8, &H10C25DE5, &HEC8B5500, &HC75FF56, &H5608758B, &H2EE8&, &H85C88B00, &H8B2374C9, &H57FCCE54, &HF8CE7C8B, &H7EBF633, &H1D7AC0F, &H8B46EAD1, &H75C20BC7, &H6E1C1F3, &HC0418D5F
    pvAppendBuffer &H5D5EC603, &H550008C2, &H4D8BEC8B, &H1E9830C, &H558B1178, &HCA048B08, &H4CA440B, &HE9830575, &H8DF27901, &HC25D0141, &H8B550008, &HC458BEC, &H8B575653, &H6583087D, &HDB330008, &HEBC7348D, &HFC4E8B23, &H8B08EE83, &HFC28B16, &HB01C8AC, &HE9D10845, &H689CB0B, &H4E89DA8B, &H1FE3C104, &H86583, &HD977F73B, &H5D5B5E5F, &H550008C2, &H558BEC8B, &H74D28510, &H84D8B1E, &HC758B56
    pvAppendBuffer &H48BF12B, &H8D01890E, &H448B0849, &H4189FC0E, &H1EA83FC, &H5D5EEC75, &H55000CC2, &HEC83EC8B, &H10558B5C, &H5653C033, &HC933F633, &H40DC7521, &H57E07521, &H7589FF33, &H89C22BD4, &H4D89D87D, &HE44589E8, &HCA3BC103, &H6583DB1B, &H658300FC, &HD3F700F8, &HD93BD823, &HFF870F, &H7D8B0000, &H2BC18B0C, &HC7348DC3, &H89DC458B, &HD18BEC75, &H5589D32B, &HFDA3BF0, &HD587&, &H476FF00
    pvAppendBuffer &HFFB4458D, &HDF74FF36, &HDF34FF04, &HE02BE850, &H7D8DFFFF, &HA5F08BC4, &H3BA5A5A5, &H4373F05D, &H8BD04D8B, &HC8558BC1, &HE8C1F28B, &HFC45011F, &H83CC458B, &H3300F855, &HC1A40FFF, &H1FEEC101, &HF90BC003, &H7D89F00B, &HC4458BF4, &H1C2A40F, &H3F07589, &HCC7589C0, &H89D07D89, &H5589C445, &H8B0CEBC8, &H4589D045, &HCC458BF4, &H83F04589, &H758D10EC, &H8DFC8BC4, &HEC83A445, &HA5A5A510
    pvAppendBuffer &H8DFC8BA5, &HA550D475, &HE8A5A5A5, &HFFFF8BEA, &H7D8DF08B, &HA5A5A5D4, &HE0458BA5, &H8BF4453B, &HC77DC45, &H453B0572, &H330573F0, &H2EB41C9, &H758BC933, &H1D233EC, &H4D8BFC4D, &HF85511E8, &HC7D8B43, &H8908EE83, &HD93BEC75, &HFF1C860F, &H7D8BFFFF, &HD4758BD8, &HEB10558B, &HDC458B03, &H89085D8B, &HF08BCB34, &H89FC458B, &H458BDC45, &HCB7C89F8, &H7D8B4104, &HE04589E0, &HFF55048D
    pvAppendBuffer &H3BFFFFFF, &HD47589C8, &H89E4458B, &H4D89D87D, &HAD820FE8, &H3FFFFFE, &HD37C89D2, &H74895FFC, &H5B5EF8D3, &HC25DE58B, &H8B55000C, &H5D8B53EC, &H33C93314, &H74DB85D2, &HC458B5F, &H29104529, &H57560845, &H8B107D8B, &H7342B30, &H8B0C7589, &H741B0470, &H7D8B0407, &H89F92B0C, &HF21B0C7D, &H7D8B383B, &H3B057508, &H17740470, &H7204703B, &H8B07770E, &HC4D3908, &HC9330576, &H3302EB41
    pvAppendBuffer &H8BD233C9, &H1C890C5D, &H7748907, &H8C08304, &H1146D83, &H5E5FAE75, &H5D5BC18B, &H550010C2, &H8B56EC8B, &HC0330C75, &H8340CE8B, &HD2333FE1, &HFF8A4FE8, &H84D8BFF, &H2306EEC1, &H5423F104, &H5D5E04F1, &H550008C2, &H558BEC8B, &H8BC28B08, &HE8C10C4D, &H8B018818, &H10E8C1C2, &H8B014188, &H8E8C1C2, &H88024188, &HC25D0351, &H8B550008, &H8558BEC, &H8B53CA8B, &HC38B0C5D, &H5618E8C1
    pvAppendBuffer &H8810758B, &HC1C38B06, &H468810E8, &HC1C38B01, &H468808E8, &HFC38B02, &H8818C1AC, &HE8C1035E, &H44E8818, &HCA8BC38B, &H10C1AC0F, &H8B10E8C1, &H54E88C2, &H8D8AC0F, &H88064688, &HEBC10756, &H5D5B5E08, &H55000CC2, &H558BEC8B, &H53CA8B08, &H8B0C5D8B, &HC1AC0FC3, &H758B5608, &H8E8C110, &H1688C38B, &H8B014E88, &HC1AC0FCA, &H10E8C110, &H4E88C38B, &HC2AC0F02, &H18E8C118, &H8B035688
    pvAppendBuffer &H45E88C3, &H8808E8C1, &HC38B0546, &HC110E8C1, &H468818EB, &H75E8806, &HC25D5B5E, &H8B55000C, &H14558BEC, &H1F74D285, &H56104D8B, &H570C758B, &H2B087D8B, &H8AF92BF1, &H1320E04, &H410F0488, &H7501EA83, &H5D5E5FF2, &H10C2&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 34125 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub

