Attribute VB_Name = "mdTlsThunks"
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

#Const ImplUseLibSodium = (ASYNCSOCKET_USE_LIBSODIUM <> 0)
#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)

'=========================================================================
' API
'=========================================================================

'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
'--- for CryptExportKey
Private Const PUBLICKEYBLOB                             As Long = 6
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_RSA_SSA_PSS                         As String = "1.2.840.113549.1.1.10"
Private Const szOID_ECC_PUBLIC_KEY                      As String = "1.2.840.10045.2.1"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
'--- Crypt32
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
#If ImplUseLibSodium Then
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    '--- libsodium
    Private Declare Function sodium_init Lib "libsodium" () As Long
    Private Declare Function randombytes_buf Lib "libsodium" (ByVal lpOut As Long, ByVal lSize As Long) As Long
    Private Declare Function crypto_scalarmult_curve25519 Lib "libsodium" (lpOut As Any, lpConstN As Any, lpConstP As Any) As Long
    Private Declare Function crypto_scalarmult_curve25519_base Lib "libsodium" (lpOut As Any, lpConstN As Any) As Long
    Private Declare Function crypto_hash_sha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha256_init Lib "libsodium" (lpState As Any) As Long
    Private Declare Function crypto_hash_sha256_update Lib "libsodium" (lpState As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha256_final Lib "libsodium" (lpState As Any, lpOut As Any) As Long
    Private Declare Function crypto_hash_sha512_init Lib "libsodium" (lpState As Any) As Long
    Private Declare Function crypto_hash_sha512_update Lib "libsodium" (lpState As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha512_final Lib "libsodium" (lpState As Any, lpOut As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_hash_sha512_statebytes Lib "libsodium" () As Long
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

Private Type CRYPT_PRIVATE_KEY_INFO
    Version             As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_BLOB_DATA
    pAttributes         As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_LIBSODIUM_SHA384_STATE As String = "2J4FwV2du8sH1Xw2KimaYhfdcDBaAVmROVkO99jsLxUxC8D/ZyYzZxEVWGiHSrSOp4/5ZA0uDNukT/q+HUi1Rw=="
'--- numeric
Private Const LNG_SHA256_HASHSZ         As Long = 32
Private Const LNG_SHA256_BLOCKSZ        As Long = 64
Private Const LNG_SHA384_HASHSZ         As Long = 48
Private Const LNG_SHA384_BLOCKSZ        As Long = 128
Private Const LNG_SHA384_CONTEXTSZ      As Long = 200
Private Const LNG_SHA512_HASHSZ         As Long = 64
Private Const LNG_HMAC_INNER_PAD        As Long = &H36
Private Const LNG_HMAC_OUTER_PAD        As Long = &H5C
Private Const LNG_FACILITY_WIN32        As Long = &H80070000
Private Const LNG_CHACHA20_KEYSZ        As Long = 32
Private Const LNG_CHACHA20POLY1305_IVSZ As Long = 12
Private Const LNG_CHACHA20POLY1305_TAGSZ As Long = 16
Private Const LNG_AES128_KEYSZ          As Long = 16
Private Const LNG_AES256_KEYSZ          As Long = 32
Private Const LNG_AESGCM_IVSZ           As Long = 12
Private Const LNG_AESGCM_TAGSZ          As Long = 16
Private Const LNG_LIBSODIUM_SHA512_CONTEXTSZ As Long = 64 + 16 + 128
Private Const LNG_OUT_OF_MEMORY         As Long = 8
Private Const STR_VL_ALERTS                             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|43|Unsupported certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown certificate authority|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_VL_STATE                              As String = "0|New|1|Closed|2|HandshakeStart|3|ExpectServerHello|4|ExpectExtensions|5|ExpectServerFinished|6|ExpectClientHello|7|ExpectClientFinished|8|PostHandshake|9|Shutdown"
Private Const STR_VL_HANDSHAKE_TYPE                     As String = "1|client_hello|2|server_hello|4|new_session_ticket|5|end_of_early_data|8|encrypted_extensions|11|certificate|12|server_key_exchange|13|certificate_request|14|server_hello_done|15|certificate_verify|16|client_key_exchange|20|finished|24|key_update|25|compressed_certificate|254|message_hash"
Private Const STR_VL_EXTENSION_TYPE                     As String = "0|server_name|1|max_fragment_length|2|client_certificate_url|3|trusted_ca_keys|4|truncated_hmac|5|status_request|6|user_mapping|7|client_authz|8|server_authz|9|cert_type|10|supported_groups|11|ec_point_formats|12|srp|13|signature_algorithms|14|use_srtp|15|heartbeat|16|application_layer_protocol_negotiation|17|status_request_v2|18|signed_certificate_timestamp|19|client_certificate_type|20|server_certificate_type|21|padding|22|encrypt_then_mac|23|extended_master_secret|24|token_binding|25|cached_info|26|tls_lts|27|compress_certificate|28|record_size_limit|29|pwd_protect|30|pwd_clear|31|password_salt|32|ticket_pinning|33|tls_cert_with_extern_psk|34|delegated_credentials|35|session_ticket|41|pre_shared_key|42|early_data|43|supported_versions|44|cookie|45|psk_key_exchange_modes|47|certificate_authorities|48|oid_filters|49|post_handshake_auth|" & _
                                                                    "50|signature_algorithms_cert|51|key_share|52|transparency_info|53|connection_id|55|external_id_hash|56|external_session_id"
Private Const STR_UNKNOWN                               As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT                          As String = "%1."
'--- numeric
Private Const LNG_AAD_SIZE                              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const LNG_LEGACY_AAD_SIZE                       As Long = 13    '--- for TLS 1.2
Private Const LNG_ANS1_TYPE_SEQUENCE                    As Long = &H30
Private Const LNG_ANS1_TYPE_INTEGER                     As Long = &H2
'Private Const LNG_FACILITY_WIN32                        As Long = &H80070000
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
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST    As Long = 13
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
Private Const TLS_EXTENSION_TYPE_EXTENDED_MASTER_SECRET As Long = 23
'Private Const TLS_EXTENSION_TYPE_COMPRESS_CERTIFICATE   As Long = 27
'Private Const TLS_EXTENSION_TYPE_PRE_SHARED_KEY         As Long = 41
'Private Const TLS_EXTENSION_TYPE_EARLY_DATA             As Long = 42
Private Const TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS     As Long = 43
Private Const TLS_EXTENSION_TYPE_COOKIE                 As Long = 44
'Private Const TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES As Long = 45
Private Const TLS_EXTENSION_TYPE_CERTIFICATE_AUTHORITIES As Long = 47
Private Const TLS_EXTENSION_TYPE_POST_HANDSHAKE_AUTH    As Long = 49
Private Const TLS_EXTENSION_TYPE_KEY_SHARE              As Long = 51
Private Const TLS_EXTENSION_TYPE_RENEGOTIATION_INFO     As Long = &HFF01
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
Private Const ERR_UNEXPECTED_MSG_TYPE                   As String = "Unexpected message type for %1 state (%2)"
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
Private Const ERR_NO_PRIVATE_KEY                        As String = "Missing server private key"
'Private Const ERR_NO_MATCHING_ALT_NAME                  As String = "No certificate subject name matches target host name"
'Private Const ERR_TRUST_IS_REVOKED                      As String = "Trust for this certificate or one of the certificates in the certificate chain has been revoked"
'Private Const ERR_TRUST_IS_PARTIAL_CHAIN                As String = "The certificate chain is not complete"
'Private Const ERR_TRUST_IS_UNTRUSTED_ROOT               As String = "The certificate or certificate chain is based on an untrusted root"
'Private Const ERR_TRUST_IS_NOT_TIME_VALID               As String = "The certificate has expired"
'Private Const ERR_TRUST_REVOCATION_STATUS_UNKNOWN       As String = "The revocation status of the certificate or one of the certificates in the certificate chain is unknown"
'Private Const STR_CHR1                                  As String = "" '--- CHAR(1)
'Private Const DEF_TIMEOUT                               As Long = 5000

Private m_baBuffer()                As Byte
Private m_lBuffIdx                  As Long
Private m_uData                     As UcsCryptoThunkData
Public g_oRequestSocket             As cTlsSocket

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 2
    ucsTlsSupportAll = ucsTlsSupportTls12 Or ucsTlsSupportTls13
End Enum

Public Enum UcsTlsStatesEnum '--- sync w/ STR_VL_STATE
    ucsTlsStateNew = 0
    ucsTlsStateClosed = 1
    ucsTlsStateHandshakeStart = 2
    ucsTlsStateExpectServerHello = 3
    ucsTlsStateExpectExtensions = 4
    ucsTlsStateExpectServerFinished = 5     '--- not used in TLS 1.3
    '--- server states
    ucsTlsStateExpectClientHello = 6
    ucsTlsStateExpectClientFinished = 7
    ucsTlsStatePostHandshake = 8
    ucsTlsStateShutdown = 9
End Enum

Public Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoExchX25519 = 1
    ucsTlsAlgoExchSecp256r1 = 2
    ucsTlsAlgoExchSecp384r1 = 3
    ucsTlsAlgoExchSecp521r1 = 4
    ucsTlsAlgoExchCertificate = 5
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

Public Type UcsTlsContext
    '--- config
    IsServer            As Boolean
    RemoteHostName      As String
    LocalFeatures       As UcsTlsLocalFeaturesEnum
    OnClientCertificate As Long
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
    LocalPrivateKey     As Collection
    LocalSignatureType  As Long
    LocalLegacyVerifyData() As Byte
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
    RemoteLegacyNextTrafficKey() As Byte
    RemoteLegacyNextTrafficIV() As Byte
    LocalTrafficSecret() As Byte
    LocalTrafficKey()   As Byte
    LocalTrafficIV()    As Byte
    LocalTrafficSeqNo   As Long
    '--- hello retry request
    HelloRetryRequest   As Boolean
    HelloRetryCipherSuite As Long
    HelloRetryExchGroup As Long
    HelloRetryCookie()  As Byte
    '--- client certificate request
    CertRequestContext() As Byte
    CertRequestSignatureType As Long
    CertRequestCaDn     As Collection
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
    ucsPfnRsaModExp
    [_ucsPfnMax]
End Enum

Private Type UcsCryptoThunkData
    Thunk               As Long
    Glob()              As Byte
    Pfn(1 To [_ucsPfnMax] - 1) As Long
    Ecc256KeySize       As Long
    Ecc384KeySize       As Long
#If ImplUseLibSodium Then
    HashCtx(0 To LNG_LIBSODIUM_SHA512_CONTEXTSZ - 1) As Byte
#Else
    HashCtx(0 To LNG_SHA384_CONTEXTSZ - 1) As Byte
#End If
    HashPad(0 To LNG_SHA384_BLOCKSZ - 1 + 1000) As Byte
    HashFinal(0 To LNG_SHA384_HASHSZ - 1 + 1000) As Byte
    hRandomProv         As Long
End Type

'=========================================================================
' Properties
'=========================================================================

Public Property Get TlsIsClosed(uCtx As UcsTlsContext) As Boolean
    TlsIsClosed = (uCtx.State = ucsTlsStateClosed)
End Property

Public Function TlsClose(uCtx As UcsTlsContext)
    uCtx.State = ucsTlsStateClosed
End Function

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
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsSetLastError uEmpty, vbNullString
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        .OnClientCertificate = ObjPtr(OnClientCertificate)
        pvTlsArrayRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
    End With
    uCtx = uEmpty
    '--- success
    TlsInitClient = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection) As Boolean
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
        Set .LocalPrivateKey = PrivateKey
        pvTlsArrayRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
    End With
    uCtx = uEmpty
    '--- success
    TlsInitServer = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lOutputPos
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
        TlsHandshake = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lOutputPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    Dim lPos            As Long
    
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baPlainText)
        End If
        If lSize = 0 Then
            '--- flush
            If .SendPos > 0 Then
                If lOutputPos = 0 Then
                    pvArraySwap .SendBuffer, .SendPos, baOutput, lOutputPos
                Else
                    lOutputPos = pvWriteBuffer(baOutput, lOutputPos, VarPtr(.SendBuffer(0)), .SendPos)
                End If
            End If
            '--- success
            TlsSend = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lOutputPos
        Do While lPos < lSize
            .SendPos = pvTlsBuildApplicationData(uCtx, .SendBuffer, .SendPos, baPlainText, lPos, Clamp(lSize - lPos, 0, TLS_MAX_PLAINTEXT_RECORD_SIZE), TLS_CONTENT_TYPE_APPDATA)
            lPos = lPos + TLS_MAX_PLAINTEXT_RECORD_SIZE
        Loop
        '--- success
        TlsSend = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lOutputPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If lSize = 0 Then
            '--- flush
            If .DecrPos > 0 Then
                If lPos = 0 Then
                    pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
                Else
                    lPos = pvWriteBuffer(baPlainText, lPos, VarPtr(.DecrBuffer(0)), .DecrPos)
                End If
            End If
            '--- success
            TlsReceive = True
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
        TlsReceive = True
QH:
        '--- swap-out
        pvArraySwap baPlainText, lPos, .DecrBuffer, .DecrPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
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
        TlsShutdown = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext) As String
    TlsGetLastError = uCtx.LastError
    If uCtx.LastAlertCode <> -1 Then
        TlsGetLastError = IIf(LenB(TlsGetLastError) <> 0, TlsGetLastError & ". ", vbNullString) & Replace(STR_FORMAT_ALERT, "%1", TlsGetLastAlert(uCtx))
    End If
End Function

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

Private Function pvTlsGetStateAsText(ByVal lState As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_STATE, "|")
    End If
    If lState <= UBound(vTexts) Then
        pvTlsGetStateAsText = vTexts(lState)
    End If
    If LenB(pvTlsGetStateAsText) = 0 Then
        pvTlsGetStateAsText = Replace(STR_UNKNOWN, "%1", lState)
    End If
End Function

Private Function pvTlsGetHandshakeType(ByVal lType As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_HANDSHAKE_TYPE, "|")
    End If
    If lType <= UBound(vTexts) Then
        pvTlsGetHandshakeType = vTexts(lType)
    End If
    If LenB(pvTlsGetHandshakeType) = 0 Then
        pvTlsGetHandshakeType = Replace(STR_UNKNOWN, "%1", lType)
    Else
        pvTlsGetHandshakeType = pvTlsGetHandshakeType & " (" & lType & ")"
    End If
End Function

#If ImplUseDebugLog Then
Private Function pvTlsGetExtensionType(ByVal lType As Long) As String
    Static vTexts       As Variant
    
    If IsEmpty(vTexts) Then
        vTexts = SplitOrReindex(STR_VL_EXTENSION_TYPE, "|")
    End If
    If lType <= UBound(vTexts) Then
        pvTlsGetExtensionType = vTexts(lType)
    ElseIf lType = &HFF01& Then
        pvTlsGetExtensionType = "renegotiation_info"
    End If
    If LenB(pvTlsGetExtensionType) = 0 Then
        pvTlsGetExtensionType = Replace(STR_UNKNOWN, "%1", lType)
    Else
        pvTlsGetExtensionType = pvTlsGetExtensionType & " (" & lType & ")"
    End If
End Function
#End If

Private Function pvTlsBuildClientHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim vElem           As Variant
    Dim baTemp()        As Byte
    
    With uCtx
        If (.LocalFeatures And ucsTlsSupportTls13) <> 0 And .ExchGroup = 0 Then
            '--- populate preferred .ExchGroup and .LocalExchPublic
            pvTlsSetupExchEccGroup uCtx, TLS_GROUP_X25519
        End If
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE, uCtx)
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
                        pvArrayByte baTemp, 0, TLS_EXTENSION_TYPE_EXTENDED_MASTER_SECRET, 0, 0
                        lPos = pvWriteArray(baOutput, lPos, baTemp)     '--- supported
                        '--- Extension - Renegotiation Info
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_RENEGOTIATION_INFO, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                lPos = pvWriteArray(baOutput, lPos, .LocalLegacyVerifyData)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Signature Algorithms
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, Size:=2)
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512, Size:=2)
                            End If
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
                        '--- Extension - Post Handshake Auth
                        pvArrayByte baTemp, 0, TLS_EXTENSION_TYPE_POST_HANDSHAKE_AUTH, 0, 0
                        lPos = pvWriteArray(baOutput, lPos, baTemp)     '--- supported
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
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildClientHello = lPos
QH:
End Function

Private Function pvTlsBuildClientLegacyKeyExchange(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim baSignature()   As Byte
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim lLocalTrafficSeqNo As Long
    
    With uCtx
        lLocalTrafficSeqNo = .LocalTrafficSeqNo
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE, uCtx)
            If .CertRequestSignatureType <> 0 Then
                '--- Client Certificate
                lMessagePos = lPos
                lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                        For lIdx = 1 To pvCollectionCount(.LocalCertificates)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                                baCert = .LocalCertificates.Item(lIdx)
                                lPos = pvWriteArray(baOutput, lPos, baCert)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        Next
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
            End If
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
            If .CertRequestSignatureType > 0 Then
                '--- Client Certificate Verify
                lMessagePos = lPos
                lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                    lPos = pvWriteLong(baOutput, lPos, .CertRequestSignatureType, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        pvTlsSignatureSign .LocalPrivateKey, .CertRequestSignatureType, .HandshakeMessages, baSignature
                        lPos = pvWriteArray(baOutput, lPos, baSignature)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
            End If
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
        '--- Legacy Change Cipher Spec
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, uCtx)
            lPos = pvWriteLong(baOutput, lPos, 1)
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
        pvTlsDeriveLegacySecrets uCtx
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE, uCtx)
            '--- Client Handshake Finished
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                pvTlsKdfLegacyPrf baVerifyData, .DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, 12
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
                .LocalLegacyVerifyData = baVerifyData
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lMessageSize = lPos - lMessagePos
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
QH:
    pvTlsBuildClientLegacyKeyExchange = lPos
End Function

Private Function pvTlsBuildClientHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lHandshakePos   As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim baSignature()   As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim baTemp()        As Byte
    Dim baEmpty()       As Byte
    
    With uCtx
        If .CertRequestSignatureType <> 0 Then
            '--- Record Header
            lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA, uCtx)
                '--- Client Certificate
                lHandshakePos = lPos
                lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                    '--- certificate request context
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteArray(baOutput, lPos, .CertRequestContext)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                        For lIdx = 1 To pvCollectionCount(.LocalCertificates)
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
                If .CertRequestSignatureType > 0 Then
                    '--- Client Certificate Verify
                    lHandshakePos = lPos
                    lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                        lPos = pvWriteLong(baOutput, lPos, .CertRequestSignatureType, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                            lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, client CertificateVerify" & Chr$(0))
                            lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                            pvTlsSignatureSign .LocalPrivateKey, .CertRequestSignatureType, baVerifyData, baSignature
                            lPos = pvWriteArray(baOutput, lPos, baSignature)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
                End If
                '--- Record Type
                lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
        End If
        '--- Legacy Change Cipher Spec
        pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
        lPos = pvWriteArray(baOutput, lPos, baTemp)
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA, uCtx)
            '--- Client Handshake Finished
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .LocalTrafficSecret, "finished", baEmpty, .DigestSize
                pvTlsHkdfExtract baVerifyData, .DigestAlgo, baTemp, baHandshakeHash
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            '--- Record Type
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildClientHandshakeFinished = lPos
QH:
End Function

Private Function pvTlsBuildServerHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    
    With uCtx
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE, uCtx)
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
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildServerHello = lPos
End Function

Private Function pvTlsBuildServerHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
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
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA, uCtx)
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
                    For lIdx = 1 To pvCollectionCount(.LocalCertificates)
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
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildServerHandshakeFinished = lPos
End Function

Private Function pvTlsBuildApplicationData(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, ByVal lDataPos As Long, ByVal lSize As Long, ByVal lContentType As Long) As Long
    With uCtx
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA, uCtx)
            If lSize > 0 Then
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baData(lDataPos)), lSize)
            End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                '--- Record Type
                lPos = pvWriteLong(baOutput, lPos, lContentType)
            End If
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildApplicationData = lPos
End Function

Private Function pvTlsBuildAlert(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, ByVal eAlertDesc As UcsTlsAlertDescriptionsEnum, ByVal lAlertLevel As Long) As Long
    Dim baTemp()        As Byte
    
    With uCtx
        '--- for TLS 1.3 -> tunnel alert through application data encryption
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            pvArrayByte baTemp, lAlertLevel, eAlertDesc
            pvTlsBuildAlert = pvTlsBuildApplicationData(uCtx, baOutput, lPos, baTemp, 0, UBound(baTemp) + 1, TLS_CONTENT_TYPE_ALERT)
            GoTo QH
        End If
        '--- Record Header
        lPos = pvWriteBeginOfRecord(baOutput, lPos, TLS_CONTENT_TYPE_ALERT, uCtx)
            lPos = pvWriteLong(baOutput, lPos, lAlertLevel)
            lPos = pvWriteLong(baOutput, lPos, eAlertDesc)
        lPos = pvWriteEndOfRecord(baOutput, lPos, uCtx)
    End With
    pvTlsBuildAlert = lPos
QH:
End Function

Private Function pvWriteBeginOfRecord(baOutput() As Byte, ByVal lPos As Long, ByVal lContentType As Long, uCtx As UcsTlsContext) As Long
    Dim lRecordPos      As Long
    Dim baLocalIV()     As Byte
    
    With uCtx
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, lContentType)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If pvArraySize(.LocalTrafficKey) > 0 Then
                pvArrayXor baLocalIV, .LocalTrafficIV, .LocalTrafficSeqNo
                If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                    lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
                End If
                .BlocksStack.Add Array(lRecordPos, lPos, baLocalIV), Before:=1
                '--- to be continued in end-of-record. . .
            End If
    End With
    pvWriteBeginOfRecord = lPos
End Function

Private Function pvWriteEndOfRecord(baOutput() As Byte, ByVal lPos As Long, uCtx As UcsTlsContext) As Long
    Const FUNC_NAME     As String = "pvWriteEndOfRecord"
    Dim vRecordData     As Variant
    Dim lRecordPos      As Long
    Dim baLocalIV()     As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    
    With uCtx
        If pvArraySize(.LocalTrafficKey) > 0 Then
                '--- . . . continues from begin-of-record
                vRecordData = .BlocksStack.Item(1)
                .BlocksStack.Remove 1
                lRecordPos = vRecordData(0)
                lMessagePos = vRecordData(1)
                baLocalIV = vRecordData(2)
                lMessageSize = lPos - lMessagePos
                lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
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
        Else
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        End If
    End With
    pvWriteEndOfRecord = lPos
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
    Const FUNC_NAME     As String = "pvTlsParseRecord"
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim lRecordProtocol As Long
    Dim baRemoteIV()    As Byte
    Dim lPos            As Long
    Dim lEnd            As Long
    Dim baAad()         As Byte
    Dim lAadPos         As Long
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
            '--- try to decrypt record
            If pvArraySize(.RemoteTrafficKey) > 0 And lRecordSize > .TagSize Then
                lEnd = lPos + lRecordSize - .TagSize
                bResult = False
                pvArrayXor baRemoteIV, .RemoteTrafficIV, .RemoteTrafficSeqNo
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baInput, lRecordPos, LNG_AAD_SIZE, baInput, lPos, lRecordSize)
                    '--- trim zero padding at the end of decrypted record
                    Do While lEnd > lPos
                        lEnd = lEnd - 1
                        If baInput(lEnd) <> 0 Then
                            Exit Do
                        End If
                    Loop
                    lRecordType = baInput(lEnd)
                ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
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
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                End If
                If Not bResult Then
                    sError = ERR_DECRYPTION_FAILED
                    eAlertCode = uscTlsAlertBadRecordMac
                    GoTo QH
                End If
                .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
            Else
                lEnd = lPos + lRecordSize
            End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                If lPos + 1 <> lEnd Then
                    sError = ERR_RECORD_TOO_BIG
                    eAlertCode = uscTlsAlertDecodeError
                    GoTo QH
                End If
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    .RemoteTrafficKey = .RemoteLegacyNextTrafficKey
                    .RemoteTrafficIV = .RemoteLegacyNextTrafficIV
                    .RemoteTrafficSeqNo = 0
                End If
            Case TLS_CONTENT_TYPE_ALERT
                If lPos + 2 <> lEnd Then
                    sError = ERR_RECORD_TOO_BIG
                    eAlertCode = uscTlsAlertDecodeError
                    GoTo QH
                End If
                Select Case baInput(lPos)
                Case TLS_ALERT_LEVEL_FATAL
                    sError = ERR_FATAL_ALERT
                    eAlertCode = baInput(lPos + 1)
                    GoTo QH
                Case TLS_ALERT_LEVEL_WARNING
                    .LastAlertCode = baInput(lPos + 1)
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, TlsGetLastAlert(uCtx) & " (TLS_ALERT_LEVEL_WARNING)"
                    #End If
                    If .LastAlertCode = uscTlsAlertCloseNotify Then
                        .State = ucsTlsStateClosed
                    End If
                End Select
            Case TLS_CONTENT_TYPE_HANDSHAKE
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
            Case TLS_CONTENT_TYPE_APPDATA
                .DecrPos = pvWriteBuffer(.DecrBuffer, .DecrPos, VarPtr(baInput(lPos)), lEnd - lPos)
            Case Else
                sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
                eAlertCode = uscTlsAlertHandshakeFailure
                GoTo QH
            End Select
            '--- note: skip AEAD's authentication tag or zero padding
            lPos = lRecordPos + lRecordSize + 5
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
    Const FUNC_NAME     As String = "pvTlsParseHandshake"
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
            #If ImplUseDebugLog Then
'                DebugLog MODULE_NAME, FUNC_NAME, ".State=" & pvTlsGetStateAsText(.State) & ", lMessageType=" & pvTlsGetHandshakeType(lMessageType)
            #End If
            Select Case .State
            Case ucsTlsStateExpectServerHello
                Select Case lMessageType
                Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                    If Not pvTlsParseHandshakeServerHello(uCtx, baInput, lPos, lPos + lMessageSize, lRecordProtocol, sError, eAlertCode) Then
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
                    GoTo UnexpectedMessageType
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
                            lPos = lPos + lCertSize '--- skip RemoteCertReqContext
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
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        .State = ucsTlsStateExpectServerFinished
                        lPos = lPos + lMessageSize
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, baInput, lPos, sError, eAlertCode) Then
                        GoTo QH
                    End If
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
                            eAlertCode = uscTlsAlertCertificateUnknown
                            GoTo QH
                        End If
                        pvTlsSetupExchRsaCertificate uCtx, baCert
                    End If
                    .SendPos = pvTlsBuildClientLegacyKeyExchange(uCtx, .SendBuffer, .SendPos)
                End If
                If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                    .SendPos = pvTlsBuildClientHandshakeFinished(uCtx, .SendBuffer, .SendPos)
                    pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
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
                        GoTo UnexpectedMessageType
                    End If
                Case Else
                    GoTo UnexpectedMessageType
                End Select
            Case ucsTlsStateExpectClientHello
                Select Case lMessageType
                Case TLS_HANDSHAKE_TYPE_CLIENT_HELLO
                    If Not pvTlsParseHandshakeClientHello(uCtx, baInput, lPos, lPos + lMessageSize, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    .State = ucsTlsStateExpectClientFinished
                Case Else
                    GoTo UnexpectedMessageType
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
                    GoTo UnexpectedMessageType
                End Select
                '--- post-process ucsTlsStateExpectClientFinished
                If .State = ucsTlsStatePostHandshake Then
                    pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
                    pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
                    '--- not used past handshake
                    Erase .HandshakeMessages
                    Set .RemoteTickets = New Collection
                End If
            Case ucsTlsStatePostHandshake
                Select Case lMessageType
                Case 0 '--- Hello Request
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        Debug.Assert lMessageSize = 0
                        #If ImplUseDebugLog Then
                            DebugLog MODULE_NAME, FUNC_NAME, "Received empty message (Hello Request). Will renegotiate"
                        #End If
                        .SendPos = pvTlsBuildClientHello(uCtx, .SendBuffer, .SendPos)
                        .State = ucsTlsStateExpectServerHello
                        '--- renegotiate ephemeral keys too
                        .ExchGroup = 0
                        .CipherSuite = 0
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                    lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                    If Not .RemoteTickets Is Nothing Then
                        .RemoteTickets.Add baMessage
                    End If
                Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Received TLS_HANDSHAKE_TYPE_KEY_UPDATE"
                    #End If
                    If lMessageSize = 1 Then
                        lRequestUpdate = baInput(lPos)
                    Else
                        lRequestUpdate = -1
                    End If
                    pvTlsDeriveKeyUpdate uCtx, lRequestUpdate <> 0
                    If lRequestUpdate <> 0 Then
                        '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                        pvArrayByte baTemp, TLS_HANDSHAKE_TYPE_KEY_UPDATE, 0, 0, 1, 0
                        pvTlsBuildApplicationData uCtx, baMessage, 0, baTemp, 0, UBound(baTemp) + 1, TLS_CONTENT_TYPE_APPDATA
                        .SendPos = pvWriteArray(.SendBuffer, .SendPos, baMessage)
                    End If
                    lPos = lPos + lMessageSize
                Case TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, baInput, lPos, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    .SendPos = pvTlsBuildClientHandshakeFinished(uCtx, .SendBuffer, .SendPos)
                    '--- not used past handshake
                    Erase .HandshakeMessages
                Case Else
UnexpectedMessageType:
                    sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", pvTlsGetStateAsText(.State)), "%2", pvTlsGetHandshakeType(lMessageType))
                    eAlertCode = uscTlsAlertUnexpectedMessage
                    GoTo QH
                End Select
            Case Else
                sError = Replace(ERR_INVALID_STATE_HANDSHAKE, "%1", pvTlsGetStateAsText(.State))
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

Private Function pvTlsParseHandshakeServerHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshakeServerHello"
    Static baHelloRetryRandom() As Byte
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExchGroup      As Long
    Dim lPublicSize     As Long
    
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
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lBlockSize)
            lPos = pvReadArray(baInput, lPos, .RemoteSessionID, lBlockSize)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        lPos = pvReadLong(baInput, lPos, lCipherSuite, Size:=2)
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        #If ImplUseDebugLog Then
            DebugLog MODULE_NAME, FUNC_NAME, "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName
        #End If
        If .HelloRetryRequest Then
            .HelloRetryCipherSuite = lCipherSuite
        End If
        lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        Debug.Assert lLegacyCompress = 0
        Set .RemoteExtensions = Nothing
        If lPos < lEnd Then
            Set .RemoteExtensions = New Collection
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                lBlockEnd = lPos + lBlockSize
                Do While lPos < lBlockEnd
                    lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "lExtType=" & pvTlsGetExtensionType(lExtType)
                    #End If
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
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lPublicSize)
                                    lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lPublicSize)
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
                        If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                            .RemoteExtensions.Add lExtType, "#" & lExtType
                        End If
                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                Loop
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        End If
    End With
    '--- success
    pvTlsParseHandshakeServerHello = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeClientHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshakeClientHello"
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
    
    On Error GoTo EH
    Set cCipherPrefs = New Collection
    For Each vElem In pvTlsPrepareCipherSuitsOrder(ucsTlsSupportTls13)
        cCipherPrefs.Add cCipherPrefs.Count, "#" & vElem
    Next
    lCipherPref = 1000
    With uCtx
        If SearchCollection(.LocalCertificates, 1, RetVal:=baCert) Then
            If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
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
        #If ImplUseDebugLog Then
            DebugLog MODULE_NAME, FUNC_NAME, "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName
        #End If
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        Debug.Assert lLegacyCompress = 0
        '--- extensions
        Set .RemoteExtensions = Nothing
        If lPos < lInputEnd Then
            Set .RemoteExtensions = New Collection
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
                lEnd = lPos + lSize
                Do While lPos < lEnd
                    lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "lExtType=" & pvTlsGetExtensionType(lExtType)
                    #End If
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
                                    If pvMatchSignatureType(uCtx, lSignatureType, uCertInfo) Then
                                        .LocalSignatureType = lSignatureType
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
        End If
    End With
    '--- success
    pvTlsParseHandshakeClientHello = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeCertificateRequest(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lSignatureType  As Long
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim uCertInfo       As UcsKeyInfo
    Dim baDName()       As Byte
    Dim lDnSize         As Long
    Dim baCert()        As Byte
    Dim baSignatureTypes() As Byte
    Dim lSigPos         As Long
    Dim oCallback       As Object
    Dim bConfirmed      As Boolean
    
    On Error GoTo EH
    With uCtx
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=1, BlockSize:=lSize)
                lPos = pvReadArray(baInput, lPos, .CertRequestContext, lSize)
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
                lEnd = lPos + lSize
                Do While lPos < lEnd
                    lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lExtSize)
                        Select Case lExtType
                        Case TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                lPos = pvReadArray(baInput, lPos, baSignatureTypes, lBlockSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        Case TLS_EXTENSION_TYPE_CERTIFICATE_AUTHORITIES
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                lBlockEnd = lPos + lBlockSize
                                Set .CertRequestCaDn = New Collection
                                Do While lPos < lBlockEnd
                                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lDnSize)
                                        lPos = pvReadArray(baInput, lPos, baDName, lDnSize)
                                        .CertRequestCaDn.Add baDName
                                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                Loop
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        Case Else
                            lPos = lPos + lExtSize
                        End Select
                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                Loop
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        End If
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=1, BlockSize:=lSize)
                lPos = lPos + lSize '--- skip certificate_types
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
                lPos = pvReadArray(baInput, lPos, baSignatureTypes, lSize)
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
                lEnd = lPos + lSize
                Set .CertRequestCaDn = New Collection
                Do While lPos < lEnd
                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lDnSize)
                        lPos = pvReadArray(baInput, lPos, baDName, lDnSize)
                        .CertRequestCaDn.Add baDName
                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                Loop
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        End If
        Do
            If SearchCollection(.LocalPrivateKey, 1, RetVal:=baCert) Then
                If Not pvAsn1DecodePrivateKey(baCert, uCertInfo) Then
                    sError = ERR_UNSUPPORTED_CERTIFICATE
                    eAlertCode = uscTlsAlertHandshakeFailure
                    GoTo QH
                End If
            End If
            .CertRequestSignatureType = -1
            lSigPos = 0
            Do While lSigPos < pvArraySize(baSignatureTypes)
                lSigPos = pvReadLong(baSignatureTypes, lSigPos, lSignatureType, Size:=2)
                If pvMatchSignatureType(uCtx, lSignatureType, uCertInfo) Then
                    .CertRequestSignatureType = lSignatureType
                    Exit Do
                End If
            Loop
            bConfirmed = False
            If .CertRequestSignatureType = -1 And .OnClientCertificate <> 0 Then
                Call vbaObjSetAddref(oCallback, .OnClientCertificate)
                bConfirmed = oCallback.FireOnClientCertificate(.CertRequestCaDn)
            End If
        Loop While bConfirmed
    End With
    '--- success
    pvTlsParseHandshakeCertificateRequest = True
QH:
    Exit Function
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvMatchSignatureType(uCtx As UcsTlsContext, ByVal lSignatureType As Long, uCertInfo As UcsKeyInfo) As Boolean
    Dim lHashSize       As Long
    
    lHashSize = pvTlsSignatureHashSize(lSignatureType)
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        If (uCtx.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
            If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha1) Then
                pvMatchSignatureType = True
            End If
        End If
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If (uCtx.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
            If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                pvMatchSignatureType = True
            End If
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        '--- PSS w/ SHA512 fails on short key lengths (min PSS size is 2 + lHashSize + lSaltSize w/ lSaltSize=lHashSize)
        If (uCertInfo.BitLen + 7) \ 8 > 2 + 2 * lHashSize Then
            If uCertInfo.AlgoObjId = szOID_RSA_RSA And pvCryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                pvMatchSignatureType = True
            End If
        End If
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If (uCertInfo.BitLen + 7) \ 8 > 2 + 2 * lHashSize Then
            If uCertInfo.AlgoObjId = szOID_RSA_SSA_PSS And pvCryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                pvMatchSignatureType = True
            End If
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        If uCertInfo.AlgoObjId = szOID_ECC_PUBLIC_KEY Then
            pvMatchSignatureType = True
        ElseIf uCertInfo.AlgoObjId = szOID_ECC_CURVE_P256 And lSignatureType = TLS_SIGNATURE_ECDSA_SECP256R1_SHA256 Then
            pvMatchSignatureType = True
        ElseIf uCertInfo.AlgoObjId = szOID_ECC_CURVE_P384 And lSignatureType = TLS_SIGNATURE_ECDSA_SECP384R1_SHA384 Then
            pvMatchSignatureType = True
        ElseIf uCertInfo.AlgoObjId = szOID_ECC_CURVE_P521 And lSignatureType = TLS_SIGNATURE_ECDSA_SECP521R1_SHA512 And pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1) Then
            pvMatchSignatureType = True
        End If
    End Select
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
        If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
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

Private Sub pvTlsDeriveApplicationSecrets(uCtx As UcsTlsContext, baHandshakeHash() As Byte)
    Const FUNC_NAME     As String = "pvTlsDeriveApplicationSecrets"
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baEmpty()       As Byte
    Dim baEmpty2()       As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
        End If
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
    Dim baHandshakeHash() As Byte
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
        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_TYPE_EXTENDED_MASTER_SECRET) Then
            pvTlsArrayHash baHandshakeHash, .DigestAlgo, .HandshakeMessages, 0
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "extended master secret", baHandshakeHash, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2   '--- always 48
        Else
            lPos = pvWriteArray(baRandom, 0, .LocalExchRandom)
            lPos = pvWriteArray(baRandom, lPos, .RemoteExchRandom)
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "master secret", baRandom, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2   '--- always 48
        End If
        lPos = pvWriteArray(baRandom, 0, .RemoteExchRandom)
        lPos = pvWriteArray(baRandom, lPos, .LocalExchRandom)
        pvTlsKdfLegacyPrf baExpanded, .DigestAlgo, .MasterSecret, "key expansion", baRandom, 2 * (.MacSize + .KeySize + .IvSize)
        lPos = pvReadArray(baExpanded, 0, baEmpty, .MacSize) '--- LocalMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, baEmpty, .MacSize) '--- RemoteMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .RemoteLegacyNextTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficIV, .IvSize - .IvDynamicSize)
        pvTlsArrayRandom baRandom, .IvDynamicSize
        pvWriteArray .LocalTrafficIV, .IvSize - .IvDynamicSize, baRandom
        lPos = pvReadArray(baExpanded, lPos, .RemoteLegacyNextTrafficIV, .IvSize - .IvDynamicSize)
        pvTlsArrayRandom baRandom, .IvDynamicSize
        pvWriteArray .RemoteLegacyNextTrafficIV, .IvSize - .IvDynamicSize, baRandom
        .LocalTrafficSeqNo = 0
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

Private Sub pvTlsSignatureSign(cPrivKey As Collection, ByVal lSignatureType As Long, baVerifyData() As Byte, baSignature() As Byte)
    Const FUNC_NAME     As String = "pvTlsSignatureSign"
    Dim uKeyInfo        As UcsKeyInfo
    Dim lHashSize       As Long
    Dim baEnc()         As Byte
    Dim baVerifyHash()  As Byte
    Dim baTemp()        As Byte
        
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, "Signing with " & pvTlsSignatureTypeName(lSignatureType) & " signature"
    #End If
    If Not SearchCollection(cPrivKey, 1, baTemp) Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_NO_PRIVATE_KEY
    End If
    If Not pvAsn1DecodePrivateKey(baTemp, uKeyInfo) Then
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
    If Not pvAsn1DecodeCertificate(baCert, uCertInfo) Then
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
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, IIf(pvTlsSignatureVerify, IIf(bSkip, "Skipping ", IIf(bDeprecated, "Deprecated ", "Valid ")), "Invalid ") & pvTlsSignatureTypeName(lSignatureType) & " signature" & IIf(bDeprecated, " (lCurveSize=" & lCurveSize & " from server's public key)", vbNullString)
    #End If
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

'=========================================================================
' Crypto
'=========================================================================

Private Function pvCryptoIsSupported(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As Boolean
    Const PREF          As Long = &H1000
    
    Select Case eAlgo
    Case ucsTlsAlgoExchSecp521r1
        '--- not supported
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
        #Else
            pvCryptoIsSupported = True
        #End If
    Case ucsTlsAlgoSignaturePkcsSha1
        '--- not supported
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
    If Not pvAsn1EncodePkcs1SignatureHash(baHash, baDerHash) Then
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
    If Not pvAsn1EncodePkcs1SignatureHash(baHash, baDerHash) Then
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

Private Function pvAsn1DecodePrivateKey(baPrivKey() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "pvAsn1DecodePrivateKey"
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

Private Function pvAsn1DecodeCertificate(baCert() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "pvAsn1DecodeCertificate"
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
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvAsn1EncodePkcs1SignatureHash(baHash() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvAsn1EncodePkcs1SignatureHash"
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
    pvAsn1EncodePkcs1SignatureHash = True
QH:
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

'=========================================================================
' Crypto
'=========================================================================

Public Function CryptoInit() As Boolean
    Const FUNC_NAME     As String = "CryptoInit"
    Dim lOffset         As Long
    Dim lIdx            As Long
    Dim baThunk()       As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    With m_uData
        #If ImplUseLibSodium Then
            If GetModuleHandle("libsodium.dll") = 0 Then
                If LoadLibrary(App.Path & "\libsodium.dll") = 0 Then
                    Call LoadLibrary(App.Path & "\..\..\lib\libsodium.dll")
                End If
                If sodium_init() < 0 Then
                    hResult = LNG_OUT_OF_MEMORY
                    sApiSource = "sodium_init"
                    GoTo QH
                End If
            End If
        #Else
            If .hRandomProv = 0 Then
                If CryptAcquireContext(.hRandomProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CryptAcquireContext"
                    GoTo QH
                End If
            End If
        #End If
        If m_uData.Thunk = 0 Then
            .Ecc256KeySize = 32
            .Ecc384KeySize = 48
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
                Call CopyMemory(lOffset, ByVal UnsignedAdd(.Thunk, 4 * lIdx), 4)
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
            Call pvPatchTrampoline(AddressOf pvCallRsaModExp)
            '--- init thunk's first 4 bytes -> global data in C/C++
            Call CopyMemory(ByVal .Thunk, VarPtr(.Glob(0)), 4)
            Call CopyMemory(.Glob(0), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc"), 4)
            Call CopyMemory(.Glob(4), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemRealloc"), 4)
            Call CopyMemory(.Glob(8), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree"), 4)
        End If
    End With
    '--- success
    CryptoInit = True
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Sub CryptoTerminate()
    With m_uData
        #If Not ImplUseLibSodium Then
            If .hRandomProv <> 0 Then
                Call CryptReleaseContext(.hRandomProv, 0)
                .hRandomProv = 0
            End If
        #End If
    End With
End Sub

Public Function CryptoEccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccCurve25519MakeKey"
    
    pvArrayAllocate baPrivate, m_uData.Ecc256KeySize, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, m_uData.Ecc256KeySize, FUNC_NAME & ".baPublic"
    CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
    '--- fix issues w/ specific privkeys
    baPrivate(0) = baPrivate(0) And 248
    baPrivate(UBound(baPrivate)) = (baPrivate(UBound(baPrivate)) And 127) Or 64
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519_base(baPublic(0), baPrivate(0))
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519MulBase)
        pvCallCurve25519MulBase m_uData.Pfn(ucsPfnCurve25519ScalarMultBase), baPublic(0), baPrivate(0)
    #End If
    '--- success
    CryptoEccCurve25519MakeKey = True
End Function

Public Function CryptoEccCurve25519SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccCurve25519SharedSecret"
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc256KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc256KeySize - 1
    pvArrayAllocate baRetVal, m_uData.Ecc256KeySize, FUNC_NAME & ".baRetVal"
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519(baRetVal(0), baPrivate(0), baPublic(0))
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519Multiply)
        pvCallCurve25519Multiply m_uData.Pfn(ucsPfnCurve25519ScalarMultiply), baRetVal(0), baPrivate(0), baPublic(0)
    #End If
    '--- success
    CryptoEccCurve25519SharedSecret = True
End Function

Public Function CryptoEccSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp256r1MakeKey"
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    pvArrayAllocate baPrivate, m_uData.Ecc256KeySize, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, 2 * m_uData.Ecc256KeySize + 1, FUNC_NAME & ".baPublic"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp256r1MakeKey = True
QH:
End Function

Public Function CryptoEccSecp256r1SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp256r1SharedSecret"
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc256KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc256KeySize
    pvArrayAllocate baRetVal, m_uData.Ecc256KeySize, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
    If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp256r1SharedSecret = True
QH:
End Function

Public Function CryptoEccSecp256r1UncompressKey(baRetVal() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp256r1UncompressKey"
    
    pvArrayAllocate baRetVal, 1 + 2 * m_uData.Ecc256KeySize, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp256r1UncompressKey = True
QH:
End Function

Public Function CryptoEccSecp256r1Sign(baRetVal() As Byte, baPrivKey() As Byte, baHash() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp256r1Sign"
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim lIdx            As Long
    
    pvArrayAllocate baRandom, m_uData.Ecc256KeySize, FUNC_NAME & ".baRandom"
    pvArrayAllocate baRetVal, 2 * m_uData.Ecc256KeySize, FUNC_NAME & ".baRetVal"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc256KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp256r1Sign), baPrivKey(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp256r1Sign = True
QH:
End Function

Public Function CryptoEccSecp256r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    CryptoEccSecp256r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoEccSecp384r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp384r1MakeKey"
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
        
    pvArrayAllocate baPrivate, m_uData.Ecc384KeySize, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, 2 * m_uData.Ecc384KeySize + 1, FUNC_NAME & ".baPublic"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc384KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp384r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp384r1MakeKey = True
QH:
End Function

Public Function CryptoEccSecp384r1SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp384r1SharedSecret"
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc384KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc384KeySize
    pvArrayAllocate baRetVal, m_uData.Ecc384KeySize, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
    If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp384r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp384r1SharedSecret = True
QH:
End Function

Public Function CryptoEccSecp384r1UncompressKey(baRetVal() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp384r1UncompressKey"
    
    pvArrayAllocate baRetVal, 1 + 2 * m_uData.Ecc384KeySize, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp384r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp384r1UncompressKey = True
QH:
End Function

Public Function CryptoEccSecp384r1Sign(baRetVal() As Byte, baPrivKey() As Byte, baHash() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccSecp384r1Sign"
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim lIdx            As Long
    
    pvArrayAllocate baRandom, m_uData.Ecc384KeySize, FUNC_NAME & ".baRandom"
    pvArrayAllocate baRetVal, 2 * m_uData.Ecc384KeySize, FUNC_NAME & ".baRetVal"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc384KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp384r1Sign), baPrivKey(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    '--- success
    CryptoEccSecp384r1Sign = True
QH:
End Function

Public Function CryptoEccSecp384r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    CryptoEccSecp384r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp384r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoHashSha256(baRetVal() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha256"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA256_HASHSZ, FUNC_NAME & ".baRetVal"
    #If ImplUseLibSodium Then
        Call crypto_hash_sha256(baRetVal(0), ByVal lPtr, Size)
    #Else
        With m_uData
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            lCtxPtr = VarPtr(.HashCtx(0))
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
        End With
    #End If
    '--- success
    CryptoHashSha256 = True
End Function

Public Function CryptoHashSha384(baRetVal() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha384"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA384_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha384_init(.HashCtx)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA384_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    '--- success
    CryptoHashSha384 = True
End Function

Public Function CryptoHashSha512(baRetVal() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHashSha512"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    pvArrayAllocate baRetVal, LNG_SHA512_HASHSZ, FUNC_NAME & ".baRetVal"
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha512_init(ByVal lCtxPtr)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA512_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            pvCallSha2Init .Pfn(ucsPfnSha512Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha512Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha512Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    '--- success
    CryptoHashSha512 = True
End Function

Public Function CryptoHmacSha256(baRetVal() As Byte, baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmacSha256"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA256_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        pvArrayAllocate baRetVal, LNG_SHA256_HASHSZ, FUNC_NAME & ".baRetVal"
        #If ImplUseLibSodium Then
            '-- inner hash
            Call crypto_hash_sha256_init(ByVal lCtxPtr)
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA256_BLOCKSZ)
            Call crypto_hash_sha256_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha256_final(ByVal lCtxPtr, .HashFinal(0))
            '-- outer hash
            Call crypto_hash_sha256_init(ByVal lCtxPtr)
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA256_BLOCKSZ)
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashFinal(0), LNG_SHA256_HASHSZ)
            Call crypto_hash_sha256_final(ByVal lCtxPtr, baRetVal(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            '-- inner hash
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA256_HASHSZ
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    '--- success
    CryptoHmacSha256 = True
End Function

Public Function CryptoHmacSha384(baRetVal() As Byte, baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmacSha384"
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA384_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        pvArrayAllocate baRetVal, LNG_SHA384_HASHSZ, FUNC_NAME & ".baRetVal"
        #If ImplUseLibSodium Then
            '-- inner hash
            Call crypto_hash_sha384_init(.HashCtx)
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA384_BLOCKSZ)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            '-- outer hash
            Call crypto_hash_sha384_init(.HashCtx)
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA384_BLOCKSZ)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashFinal(0), LNG_SHA384_HASHSZ)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA384_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            '-- inner hash
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA384_HASHSZ
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    '--- success
    CryptoHmacSha384 = True
End Function

Public Function CryptoAeadChacha20Poly1305Encrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_CHACHA20POLY1305_TAGSZ
    If lSize > 0 Then
        If lAdSize > 0 Then
            lAdPtr = VarPtr(baAad(lAadPos))
        End If
        #If ImplUseLibSodium Then
            Call crypto_aead_chacha20poly1305_ietf_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Encrypt)
            Call pvCallChacha20Poly1305Encrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Encrypt), _
                    baKey(0), baNonce(0), _
                    lAdPtr, lAdSize, _
                    baBuffer(lPos), lSize, _
                    baBuffer(lPos), baBuffer(lPos + lSize))
        #End If
    End If
    '--- success
    CryptoAeadChacha20Poly1305Encrypt = True
End Function

Public Function CryptoAeadChacha20Poly1305Decrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    #If ImplUseLibSodium Then
        If crypto_aead_chacha20poly1305_ietf_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
            '--- success
            CryptoAeadChacha20Poly1305Decrypt = True
        End If
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Decrypt)
        If pvCallChacha20Poly1305Decrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Decrypt), _
                baKey(0), baNonce(0), _
                baAad(lAadPos), lAdSize, _
                baBuffer(lPos), lSize - LNG_CHACHA20POLY1305_TAGSZ, _
                baBuffer(lPos + lSize - LNG_CHACHA20POLY1305_TAGSZ), baBuffer(lPos)) = 0 Then
            '--- success
            CryptoAeadChacha20Poly1305Decrypt = True
        End If
    #End If
End Function

Public Function CryptoAeadAesGcmEncrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    #If ImplUseLibSodium Then
        Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    #Else
        Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    #End If
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_AESGCM_TAGSZ
    If lSize > 0 Then
        If lAdSize > 0 Then
            lAdPtr = VarPtr(baAad(lAadPos))
        End If
        #If ImplUseLibSodium Then
            Call crypto_aead_aes256gcm_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmEncrypt)
            Call pvCallAesGcmEncrypt(m_uData.Pfn(ucsPfnAesGcmEncrypt), _
                    baBuffer(lPos), baBuffer(lPos + lSize), _
                    baBuffer(lPos), lSize, _
                    lAdPtr, lAdSize, _
                    baNonce(0), baKey(0), UBound(baKey) + 1)
        #End If
    End If
    '--- success
    CryptoAeadAesGcmEncrypt = True
End Function

Public Function CryptoAeadAesGcmDecrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    #If ImplUseLibSodium Then
        Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    #Else
        Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    #End If
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    #If ImplUseLibSodium Then
        If crypto_aead_aes256gcm_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
            '--- success
            CryptoAeadAesGcmDecrypt = True
        End If
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmDecrypt)
        If pvCallAesGcmDecrypt(m_uData.Pfn(ucsPfnAesGcmDecrypt), _
                baBuffer(lPos), _
                baBuffer(lPos), lSize - LNG_AESGCM_TAGSZ, _
                baBuffer(lPos + lSize - LNG_AESGCM_TAGSZ), _
                baAad(lAadPos), lAdSize, _
                baNonce(0), baKey(0), UBound(baKey) + 1) = 0 Then
            '--- success
            CryptoAeadAesGcmDecrypt = True
        End If
    #End If
End Function

Public Sub CryptoRandomBytes(ByVal lPtr As Long, ByVal lSize As Long)
    #If ImplUseLibSodium Then
        Call randombytes_buf(lPtr, lSize)
    #Else
        Call CryptGenRandom(m_uData.hRandomProv, lSize, lPtr)
    #End If
End Sub

Public Function CryptoRsaModExp(baBase() As Byte, baExp() As Byte, baModulus() As Byte, baRetVal() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaModExp"
    
    pvArrayAllocate baRetVal, UBound(baBase) + 1, FUNC_NAME & ".baRetVal"
    Debug.Assert pvPatchTrampoline(AddressOf pvCallRsaModExp)
    Call pvCallRsaModExp(m_uData.Pfn(ucsPfnRsaModExp), UBound(baBase) + 1, baBase(0), baExp(0), baModulus(0), baRetVal(0))
    '--- success
    CryptoRsaModExp = True
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
        Call CopyMemory(Pfn, ByVal UnsignedAdd(Pfn, &H16), 4)
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
    bValue = True
    pvSetTrue = True
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
#End If

#If ImplUseLibSodium Then
Private Sub crypto_hash_sha384_init(baCtx() As Byte)
    Static baSha384State() As Byte
    
    If pvArraySize(baSha384State) = 0 Then
        baSha384State = FromBase64Array(STR_LIBSODIUM_SHA384_STATE)
    End If
    Debug.Assert pvArraySize(baCtx) >= crypto_hash_sha512_statebytes()
    Call crypto_hash_sha512_init(baCtx(0))
    Call CopyMemory(baCtx(0), baSha384State(0), UBound(baSha384State) + 1)
End Sub
#End If

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

Private Function pvCallRsaModExp(ByVal Pfn As Long, ByVal lSize As Long, pBasePtr As Byte, pExpPtr As Byte, pModuloPtr As Byte, pResultPtr As Byte) As Long
    ' void rsa_modexp(uint32_t maxbytes, const uint8_t *b, const uint8_t *e, const uint8_t *m, uint8_t *r)
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
    pvAppendBuffer &H77817160, &H77881BC0, &H778173C0, &H0&, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &H0&, &H0&, &H0&, &H1&, &HFFFFFFFF, &H27D2604B, &H3BCE3C3E, &HCC53B0F6, &H651D06B0, &H769886BC, &HB3EBBD55, &HAA3A93E7, &H5AC635D8, &HD898C296, &HF4A13945, &H2DEB33A0, &H77037D81, &H63A440F2, &HF8BCE6E5, &HE12C4247, &H6B17D1F2, &H37BF51F5, &HCBB64068, &H6B315ECE, &H2BCE3357
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
    ReDim m_baBuffer(0 To 33152 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &H80238, &H264C&, &H2944&, &H3684&, &H3A9E&, &H3B5B&, &H3BD5&, &H3E69&, &H372E&, &H3AF8&, &H3B98&, &H3D15&, &H4236&, &H318F&, &H31DF&, &H30A4&, &H323B&, &H32C6&, &H3212&, &H33F8&, &H3483&, &H32CB&, &H2599&, &H255C&, &H1CBF&, &H1C70&, &H679C&, &HC25DE58B, &HE80004, &H58000000, &H740772D, &H40000500
    pvAppendBuffer &H8B0007, &HE8C3&, &H2D580000, &H7408A, &H7400005, &H8B55C300, &H48EC83EC, &H105D8B53, &H5E046A56, &H53E85356, &H85000070, &H63850FC0, &H57000001, &HC75FF56, &H50D8458D, &H7A38E8, &H87D8B00, &H56D8458D, &H458D5750, &HF9E850B8, &H56000079, &H50D8458D, &H7A1BE850, &H53560000, &HFF0C75FF, &HE1E80C75, &H56000079, &H6E85353, &H5600007A, &HFFFF79E8, &H10C083FF, &H57575350
    pvAppendBuffer &H7467E8, &H67E85600, &H83FFFFFF, &H535010C0, &H55E85353, &H56000074, &HFFFF55E8, &H10C083FF, &H53575350, &H79FAE8, &H57535600, &H799AE857, &HE8560000, &HFFFFFF3A, &H5010C083, &HE8535757, &H7428&, &HFF28E856, &HC083FFFF, &H57535010, &H7416E857, &H6A0000, &H7E6EE857, &HC20B0000, &HE8257456, &HFFFFFF0A, &H5010C083, &H78E85757, &H6A00006E, &HF08B5704, &H7B8CE8, &H1FE6C100
    pvAppendBuffer &H6A1C7709, &H6EB5E04, &H7B7BE857, &H57560000, &H7963E853, &HE8560000, &HFFFFFED6, &H5010C083, &H50B8458D, &H78E85353, &H56000079, &HFFFEC1E8, &H10C083FF, &HB8458D50, &HE8535350, &H7963&, &HFEACE856, &HC083FFFF, &H8D535010, &H5050B845, &H794EE8, &H458D5600, &H575750B8, &H78EBE8, &H8BE85600, &H83FFFFFE, &H8D5010C0, &H5750D845, &H792DE850, &H53560000, &H7B61E857, &HFF560000
    pvAppendBuffer &HE8530C75, &H7B57&, &HD8458D56, &HC75FF50, &H7B4AE8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H68EC83EC, &H105D8B53, &H5E066A56, &HCBE85356, &H8500006E, &H77850FC0, &H57000001, &HC75FF56, &H50C8458D, &H78B0E8, &H87D8B00, &H56C8458D, &H458D5750, &H71E85098, &H56000078, &H50C8458D, &H7893E850, &H53560000, &HFF0C75FF, &H59E80C75, &H56000078, &H7EE85353, &H56000078, &HFFFDF1E8
    pvAppendBuffer &HB005FF, &H53500000, &HDDE85757, &H56000072, &HFFFDDDE8, &HB005FF, &H53500000, &HC9E85353, &H56000072, &HFFFDC9E8, &HB005FF, &H53500000, &H6CE85357, &H56000078, &HE8575753, &H780C&, &HFDACE856, &HB005FFFF, &H50000000, &HE8535757, &H7298&, &HFD98E856, &HB005FFFF, &H50000000, &HE8575753, &H7284&, &HE857006A, &H7CDC&, &H7456C20B, &HFD78E827, &HB005FFFF, &H50000000
    pvAppendBuffer &HE4E85757, &H6A00006C, &HF08B5706, &H79F8E8, &H1FE6C100, &H6A2C7709, &H6EB5E06, &H79E7E857, &H57560000, &H77CFE853, &HE8560000, &HFFFFFD42, &HB005&, &H458D5000, &H53535098, &H77E2E8, &H2BE85600, &H5FFFFFD, &HB0&, &H98458D50, &HE8535350, &H77CB&, &HFD14E856, &HB005FFFF, &H50000000, &H98458D53, &HB4E85050, &H56000077, &H5098458D, &H51E85757, &H56000077, &HFFFCF1E8
    pvAppendBuffer &HB005FF, &H8D500000, &H5750C845, &H7791E850, &H53560000, &H79C5E857, &HFF560000, &HE8530C75, &H79BB&, &HC8458D56, &HC75FF50, &H79AEE8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H758B56EC, &H56046A08, &H6D35E8, &H74C08500, &H8D046A14, &HE8502046, &H6D26&, &H574C085, &HEB40C033, &H5EC03302, &H4C25D, &H56EC8B55, &H6A08758B, &H7E85606, &H8500006D, &H6A1474C0, &H30468D06
    pvAppendBuffer &H6CF8E850, &HC0850000, &HC0330574, &H3302EB40, &HC25D5EC0, &H8B550004, &HA8EC81EC, &H53000000, &H8D0C5D8B, &H5756B845, &H5053046A, &H792EE8, &H20438D00, &H8950046A, &H858DF845, &HFFFFFF78, &H7919E850, &H75FF0000, &H58858D14, &H50FFFFFF, &H5098458D, &HFF78858D, &H8D50FFFF, &HE850B845, &H882&, &H75FF046A, &H783AE810, &HF6330000, &H4602E883, &H85144589, &H50577EC0, &HE81075FF
    pvAppendBuffer &H7B48&, &H475C20B, &H2EBCE8B, &HE1C1C933, &H589D8D05, &H3FFFFFF, &H98458DD9, &HB58DC103, &HFFFFFF78, &H89D9F753, &H350FC45, &HB87D8DF1, &H5756F903, &H4F3E8, &H53575600, &HE8FC75FF, &H2BB&, &H4814458B, &H6A144589, &HC0855E01, &H6AA97F, &HE81075FF, &H7AF0&, &H274C20B, &HE6C1F633, &H589D8D05, &H3FFFFFF, &H107589DE, &H98458D53, &HBD8DC603, &HFFFFFF78, &H758DFE2B
    pvAppendBuffer &H10752BB8, &HE8565750, &H4A0&, &H5FE8046A, &H83FFFFFB, &H8D5010C0, &H8D509845, &H8D50B845, &HE850D845, &H75FB&, &H8D57046A, &H5050D845, &H7597E8, &HFF046A00, &H458D0C75, &HE85050D8, &H7588&, &H27E8046A, &H83FFFFFB, &H8D5010C0, &H5050D845, &H7052E8, &HFF046A00, &H458DF875, &HE85050D8, &H7564&, &H8D56046A, &H5050D845, &H7557E8, &H8D565700, &H45039845, &HE8505310
    pvAppendBuffer &H1FB&, &H50D8458D, &HFF58858D, &H8D50FFFF, &HE8509845, &HA60&, &H8D08758B, &H46A9845, &HB8E85650, &H6A000077, &H58858D04, &H50FFFFFF, &H5020468D, &H77A6E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &HF8EC81EC, &H53000000, &H8D0C5D8B, &H57569845, &H5053066A, &H7782E8, &H30438D00, &H8950066A, &H858DF845, &HFFFFFF38, &H776DE850, &H75FF0000, &H8858D14, &H50FFFFFF, &HFF68858D
    pvAppendBuffer &H8D50FFFF, &HFFFF3885, &H458D50FF, &H4DE85098, &H6A000007, &H1075FF06, &H768BE8, &H83DB3300, &H894302E8, &HC0851445, &HFF505A7E, &H99E81075, &HB000079, &H8B0475C2, &H3302EBC3, &H30C06BC0, &HFF089D8D, &H8D8DFFFF, &HFFFFFF68, &HFF38B58D, &H7D8DFFFF, &H3D80398, &HD8F753C8, &H51FC4D89, &HF803F003, &HCCE85756, &H56000004, &H75FF5357, &H219E8FC, &H458B0000, &H45894814, &H5B016A14
    pvAppendBuffer &HA67FC085, &H75FF006A, &H793EE810, &HC20B0000, &HDB330274, &H8D30C36B, &HFFFF688D, &H89D8DFF, &H8DFFFFFF, &HFFFF38BD, &H98758DFF, &HC803D803, &HF82B5153, &H2B104D89, &HE85657F0, &H477&, &HABE8066A, &H5FFFFF9, &HB0&, &H68858D50, &H50FFFFFF, &H5098458D, &H50C8458D, &H7442E8, &H57066A00, &H50C8458D, &H73DEE850, &H66A0000, &H8D0C75FF, &H5050C845, &H73CFE8, &HE8066A00
    pvAppendBuffer &HFFFFF96E, &HB005&, &H458D5000, &HE85050C8, &H6E97&, &H75FF066A, &HC8458DF8, &HA9E85050, &H6A000073, &H458D5606, &HE85050C8, &H739C&, &HFF535657, &H54E81075, &H8D000001, &H8D50C845, &HFFFF0885, &H858D50FF, &HFFFFFF68, &H8F1E850, &H758B0000, &H68858D08, &H6AFFFFFF, &HE8565006, &H75FB&, &H858D066A, &HFFFFFF08, &H30468D50, &H75E9E850, &H5E5F0000, &H5DE58B5B, &H550010C2
    pvAppendBuffer &HEC83EC8B, &H57565320, &H565E046A, &HFFF8E1E8, &H85D8BFF, &H5010C083, &H1075FF53, &H50E0458D, &H737EE8, &H458D5600, &HE85050E0, &H7349&, &HE0458D56, &HE8535350, &H7310&, &HE0458D56, &H1075FF50, &HE81075FF, &H7300&, &HF8A0E856, &H758BFFFF, &H10C0830C, &H50147D8B, &HE8575756, &H733F&, &H8D57046A, &HE850E045, &H7309&, &H7BE8046A, &H83FFFFF8, &H535010C0, &H50E0458D
    pvAppendBuffer &H731DE850, &H46A0000, &HFFF865E8, &H10C083FF, &H1075FF50, &H50E0458D, &H7305E850, &H46A0000, &HFFF84DE8, &H10C083FF, &H75FF5350, &H1075FF10, &H72EEE8, &HFF046A00, &H56561075, &H728BE8, &HE8046A00, &HFFFFF82A, &H5010C083, &H50E0458D, &H1075FF53, &H72CAE8, &HFF046A00, &H57571075, &H7267E8, &HE8046A00, &HFFFFF806, &H5010C083, &HE8575756, &H72AB&, &H458D046A, &H75FF50E0
    pvAppendBuffer &H74D9E810, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC83EC8B, &H57565330, &H575F066A, &HFFF7D1E8, &H85D8BFF, &HB0BE&, &H50C60300, &H1075FF53, &H50D0458D, &H726AE8, &H458D5700, &HE85050D0, &H7235&, &HD0458D57, &HE8535350, &H71FC&, &HD0458D57, &H1075FF50, &HE81075FF, &H71EC&, &HF78CE857, &H7D8BFFFF, &H8BC60314, &H56500C75, &H2CE85757, &H6A000072, &H458D5706, &HF6E850D0
    pvAppendBuffer &H6A000071, &HF768E806, &HB005FFFF, &H50000000, &HD0458D53, &H8E85050, &H6A000072, &HF750E806, &HB005FFFF, &H50000000, &H8D1075FF, &H5050D045, &H71EEE8, &HE8066A00, &HFFFFF736, &HB005&, &HFF535000, &H75FF1075, &H71D5E810, &H66A0000, &H561075FF, &H7172E856, &H66A0000, &HFFF711E8, &HB005FF, &H8D500000, &H5350D045, &HE81075FF, &H71AF&, &H75FF066A, &HE8575710, &H714C&
    pvAppendBuffer &HEBE8066A, &H5FFFFF6, &HB0&, &H57575650, &H718EE8, &H8D066A00, &HFF50D045, &HBCE81075, &H5F000073, &HE58B5B5E, &H10C25D, &H83EC8B55, &H565360EC, &H5B046A57, &HF6B4E853, &H7D8BFFFF, &H10C08310, &H875FF50, &H57C0458D, &H7151E850, &H8D530000, &H5050C045, &H711CE8, &H458D5300, &H75FF50C0, &H875FF08, &H70DFE8, &H458D5300, &H575750C0, &H70D3E8, &H73E85300, &H8BFFFFF6
    pvAppendBuffer &HC0830C5D, &H14758B10, &H8D565350, &HE850C045, &H6B58&, &H57E8046A, &H83FFFFF6, &H535010C0, &HFCE85656, &H6A000070, &HF644E804, &HC083FFFF, &H75FF5010, &HE0458D08, &HE4E85057, &H6A000070, &HE0458D04, &HE8535350, &H7080&, &H1FE8046A, &H83FFFFF6, &H575010C0, &H8D0875FF, &HE850E045, &H6B08&, &H5756046A, &H708CE8, &HE8046A00, &HFFFFF5FE, &H5010C083, &H50E0458D, &HA0E85757
    pvAppendBuffer &H6A000070, &HF5E8E804, &HC083FFFF, &HFF575010, &H458D0875, &H88E850A0, &H6A000070, &HA0458D04, &HE8565650, &H7024&, &HC3E8046A, &H6AFFFFF5, &HC7035F10, &H56565350, &H7066E8, &H5E046A00, &HC0458D56, &HA0458D50, &H702BE850, &HE8560000, &HFFFFF59E, &H8D50C703, &H8D50E045, &H5050A045, &H703EE8, &H87E85600, &H3FFFFF5, &H75FF50C7, &HA0458D08, &HE0458D50, &H7025E850, &H8D560000
    pvAppendBuffer &H8D50C045, &H5050E045, &H6FBFE8, &H5FE85600, &H3FFFFF5, &H8D5350C7, &H5350E045, &H7002E8, &H458D5600, &H75FF50A0, &H7231E808, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC81EC8B, &H90&, &H6A575653, &HE8565E06, &HFFFFF526, &HBB107D8B, &HB0&, &HFF50C303, &H458D0875, &HE85057A0, &H6FBF&, &HA0458D56, &H8AE85050, &H5600006F, &H50A0458D, &HFF0875FF, &H4DE80875, &H5600006F
    pvAppendBuffer &H50A0458D, &H41E85757, &H5600006F, &HFFF4E1E8, &H14758BFF, &H5D8BC303, &H5653500C, &H50A0458D, &H69C7E8, &HE8066A00, &HFFFFF4C6, &HB005&, &H56535000, &H6F69E856, &H66A0000, &HFFF4B1E8, &HB005FF, &HFF500000, &H458D0875, &HE85057D0, &H6F4F&, &H458D066A, &H535350D0, &H6EEBE8, &HE8066A00, &HFFFFF48A, &HB005&, &HFF575000, &H458D0875, &H71E850D0, &H6A000069, &HE8575606
    pvAppendBuffer &H6EF5&, &H67E8066A, &H5FFFFF4, &HB0&, &HD0458D50, &HE8575750, &H6F07&, &H4FE8066A, &H5FFFFF4, &HB0&, &H75FF5750, &H70858D08, &H50FFFFFF, &H6EEAE8, &H8D066A00, &HFFFF7085, &H565650FF, &H6E83E8, &HE8066A00, &HFFFFF422, &HB0BF&, &H50C70300, &HE8565653, &H6EC3&, &H565E066A, &H50A0458D, &HFF70858D, &HE850FFFF, &H6E85&, &HF3F8E856, &HC703FFFF, &HD0458D50
    pvAppendBuffer &H70858D50, &H50FFFFFF, &H6E95E850, &HE8560000, &HFFFFF3DE, &HFF50C703, &H858D0875, &HFFFFFF70, &HD0458D50, &H6E79E850, &H8D560000, &H8D50A045, &H5050D045, &H6E13E8, &HB3E85600, &H3FFFFF3, &H8D5350C7, &H5350D045, &H6E56E8, &H858D5600, &HFFFFFF70, &H875FF50, &H7082E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H20EC83EC, &H75FF046A, &H1075FF08, &H7066E8, &HFF046A00, &H75FF0C75
    pvAppendBuffer &H7059E814, &H46A0000, &H50E0458D, &H6371E8, &HE4658300, &H187D8300, &HE045C700, &H1&, &H46A0B74, &H501875FF, &H7032E8, &HE0458D00, &HC75FF50, &HE80875FF, &H2BC&, &H50E0458D, &HFF0C75FF, &H4BE80875, &H8DFFFFF3, &HFF50E045, &H75FF1475, &H29EE810, &HE58B0000, &H14C25D, &H83EC8B55, &H66A30EC, &HFF0875FF, &HECE81075, &H6A00006F, &HC75FF06, &HE81475FF, &H6FDF&
    pvAppendBuffer &H458D066A, &HF7E850D0, &H83000062, &H8300D465, &HC700187D, &H1D045, &HB740000, &H75FF066A, &HB8E85018, &H8D00006F, &HFF50D045, &H75FF0C75, &H28DE808, &H458D0000, &H75FF50D0, &H875FF0C, &HFFF459E8, &HD0458DFF, &H1475FF50, &HE81075FF, &H26F&, &HC25DE58B, &H8B530014, &H8B0C2444, &HF710244C, &H8BD88BE1, &HF7082444, &H3142464, &H24448BD8, &H3E1F708, &H10C25BD3, &H40F98000
    pvAppendBuffer &HF9801573, &HF067320, &HE0D3C2A5, &H33D08BC3, &H1FE180C0, &H33C3E2D3, &HC3D233C0, &H7340F980, &H20F98015, &HAD0F0673, &HC3EAD3D0, &HD233C28B, &HD31FE180, &HC033C3E8, &H55C3D233, &H558BEC8B, &H8B565310, &H8B570C75, &HF22B087D, &HFA2B106A, &H160C8B5B, &H448B0A03, &H42130416, &H170C8904, &H8908528D, &H83FC1744, &HE57501EB, &H5D5B5E5F, &H55000CC2, &H5151EC8B, &H8B1C558B, &H8B562045
    pvAppendBuffer &H8B570875, &HD7030C7D, &H89104513, &H4468916, &H7710453B, &H3B04720D, &H330773D7, &HC93340C0, &H570F0EEB, &H130F66C0, &H4D8BF845, &HF8458BFC, &H5F244503, &H3284D13, &H46891445, &H13C68B08, &H4E89184D, &HE58B5E0C, &H24C25D, &H8BEC8B55, &H4D8B0C55, &H31028B08, &H4428B01, &H8B044131, &H41310842, &HC428B08, &H5D0C4131, &H550008C2, &HEC83EC8B, &H84D8B0C, &H758B5653, &H83018B10
    pvAppendBuffer &HEEC104C1, &HFF335702, &H89107589, &H48DFC4D, &H485&, &HF8458900, &H2C74F685, &H8B0C4D29, &H87589D9, &H8B0C758B, &H48D087D, &H54E8501E, &H89000057, &H45B8D03, &H7501EF83, &H10758BED, &H4D8BFE8B, &HF8458BFC, &H45C7DB33, &H108&, &HFF83B00, &H9683&, &H2BC78B00, &H81048DC6, &H8B104589, &H89FCB954, &HDE3B0C55, &H45FF0575, &H85DB3308, &HE83475DB, &HFFFFF0F6, &H58805
    pvAppendBuffer &H458B5000, &H8C0C10C, &H5FCBE850, &H45890000, &HF0DCE80C, &H4D8BFFFF, &HC558B08, &H884B60F, &H688&, &H3318E0C1, &H8320EBD0, &H1E7606FE, &H7504FB83, &HF0B8E819, &H8805FFFF, &H50000005, &H500C458B, &H5F90E8, &H8BD08B00, &H88B1045, &H33FC458B, &HB80C89CA, &H10458B47, &H83FC4D8B, &H894304C0, &H7D3B1045, &H74820FF8, &H5FFFFFFF, &HE58B5B5E, &HCC25D, &H83EC8B55, &H458D20EC
    pvAppendBuffer &HFF046AE0, &HE8501075, &H6AED&, &H458D046A, &H75FF50E0, &H875FF08, &H6AAFE8, &HFF046A00, &H458D1075, &HE85050E0, &H6AA0&, &H458D046A, &H75FF50E0, &HC75FF0C, &H6A8FE8, &H5DE58B00, &H55000CC2, &HEC83EC8B, &HD0458D30, &H75FF066A, &HA2E85010, &H6A00006A, &HD0458D06, &H875FF50, &HE80875FF, &H6A64&, &H75FF066A, &HD0458D10, &H55E85050, &H6A00006A, &HD0458D06, &HC75FF50
    pvAppendBuffer &HE80C75FF, &H6A44&, &HC25DE58B, &H8B55000C, &H10EC83EC, &H8B575653, &H378B0C7D, &HE3C1DE8B, &HF4758902, &HE8F05D89, &HFFFFEFC6, &H8B10FF53, &H89C933D0, &HF6850C55, &H78B0F7E, &H48BC12B, &H8A048987, &H7CCE3B41, &H85D8BF1, &HDE3B1B8B, &H5E8D037F, &HC1FB8B01, &H7D8902E7, &HEF90E8F8, &HFF57FFFF, &H89D08B10, &HDB85FC55, &HCF8B0B7E, &HE9C1FA8B, &HF3C03302, &H84D8BAB, &H3947FF33
    pvAppendBuffer &H8B1C7C39, &HFCC283C3, &H8B02E0C1, &H8BD003F7, &H8946B104, &HFC528D02, &HF37E313B, &H8BF4758B, &H8B0C45, &H8458950, &H416E8, &H8558B00, &H85F44589, &H8B1A74C0, &H3BE2D3C8, &H6A127EF7, &HC82B5920, &H8B0C458B, &HE8D30440, &H458BD00B, &HE85250F4, &H557F&, &H1475FF50, &HC758B56, &H75FF5356, &H379CE8FC, &H4D8B0000, &H74C98510, &H7C393924, &HFC458B20, &HFCC0834B, &H85F84503
    pvAppendBuffer &H8B0478DB, &H3302EB10, &HB91489D2, &H4704E883, &H7E393B4B, &HF075FFEA, &H5CF3E856, &HCFE80000, &H56FFFFEE, &HFF0850FF, &H758BF875, &HDEE856FC, &HE800005C, &HFFFFEEBA, &H850FF56, &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &HC758B56, &HE836FF57, &H4E36&, &HF88B006A, &H75FF5657, &HFEACE808, &H3F83FFFF, &H8B107601, &H873C8307, &H48087500, &HF8830789, &H8BF07701, &H5D5E5FC7, &H550008C2
    pvAppendBuffer &HEC83EC8B, &H8458B1C, &H758B5653, &H57008B0C, &H8BE84589, &H8BC13B0E, &HEC4D89F9, &H8BF84F0F, &HE3C157DF, &H4DBCE802, &HD8030000, &H8902E3C1, &H3BE8E45D, &H53FFFFEE, &HD88B10FF, &H5D89D233, &H7EFF85FC, &HBB0C8D4D, &H5D8BC78B, &HF85D8908, &H8BF87529, &H4D89FC5D, &HBE0C8DF4, &H3B08758B, &H8B087F06, &H348BF875, &H3302EB0E, &H933489F6, &H3B0C758B, &H8B047F06, &H3302EB31, &HF45D8BF6
    pvAppendBuffer &HF4458342, &H4E98304, &H8B338948, &HD73BFC5D, &HC78BCA7C, &H304E0C1, &H458950C3, &HFB048D08, &H48D5057, &HE85350BB, &H3921&, &H8BE84D8B, &H3411045, &H4D89EC4D, &H74C0850C, &H3B008B0C, &H8D067FC8, &H4D890148, &H3BE8510C, &H8B00004D, &H42D233F0, &HC28BC933, &H297C1639, &H8D085D8B, &HC3833F14, &H7FC23BFC, &HEB3B8B04, &H85FF3302, &H863C89FF, &H40C8450F, &H3B04EB83, &H8BE67E06
    pvAppendBuffer &HD233FC5D, &H107D8342, &HF0E8900, &H8384&, &HC0570F00, &H45130F66, &HC5539F0, &H5D8B767C, &H4468D10, &H8B085D89, &H7529F45D, &HEC5D8908, &H89F05D8B, &H5D8BF85D, &HE8458910, &H77F163B, &H7D89388B, &H8304EBF4, &H3B00F465, &H8B0B7F13, &H48B087D, &H10458907, &H658304EB, &H7D8B0010, &H3C03310, &HC013F47D, &H13F87D03, &H6583EC45, &HFF8500EC, &H8BF84589, &H3889E845, &H740C7D8B
    pvAppendBuffer &HFD13B05, &H8342CA4F, &H458904C0, &H7ED73BE8, &HFC5D8BAB, &H89E475FF, &HF2E8530E, &HE800005A, &HFFFFECCE, &H850FF53, &H5EC68B5F, &H5DE58B5B, &H55000CC2, &H4D8BEC8B, &HC985560C, &H758B3278, &HC1068B08, &HC83B02E0, &HC18B267D, &H3E28399, &HF8C1C203, &H3E18102, &H79800000, &HC9834905, &H448B41FC, &HE1C10486, &HFE8D303, &H2EBC0B6, &H5D5EC033, &H550008C2, &H5653EC8B, &H570C758B
    pvAppendBuffer &H8B087D8B, &H830F8B16, &H87501F9, &H4739C033, &HC8440F04, &H7501FA83, &H39C03308, &H440F0446, &H8BCA3BD0, &HC14F0FC2, &H3174C085, &H2B861C8D, &H7EC13BFE, &H8658306, &H8B06EB00, &H75891F34, &H7EC23B08, &HEBF63304, &H39338B07, &H1B720875, &H77087539, &H4EB8311, &H7501E883, &H5FC033D4, &HC25D5B5E, &HC0330008, &H83F4EB40, &HEFEBFFC8, &H53EC8B55, &H560C5D8B, &H3438D57, &H3E28399
    pvAppendBuffer &HC1023C8D, &HE85702FF, &H4B7E&, &HFF83F08B, &H8B097C01, &H8DC033CF, &HABF3047E, &H3A74DB85, &HE7C1FB8B, &H84D8B03, &H8EF834B, &H8941018A, &HD38B084D, &H379D285, &HC103C283, &HCF8B02FA, &H1FE181, &H5798000, &HE0C98349, &HC0B60F41, &H4409E0D3, &HDB850496, &H3E83CB75, &H8B107601, &H863C8306, &H48087500, &HF8830689, &H5FF07701, &H5B5EC68B, &H8C25D, &H8BEC8B55, &HC9850C4D
    pvAppendBuffer &H8B563D78, &H68B0875, &H3B05E0C1, &H8B2F7DC8, &HE28399C1, &HC1C2031F, &HE18105F8, &H8000001F, &H83490579, &H3341E0C9, &HE2D342D2, &H107D83, &H54090674, &H6EB0486, &H5421D2F7, &H5D5E0486, &H55000CC2, &H5756EC8B, &H33087D8B, &H5A106AF6, &H2B59206A, &HD3C78BCA, &H75C085E8, &HD3CA8B06, &HD1F203E7, &H5FE975FA, &H5D5EC68B, &H550004C2, &H458BEC8B, &HE2839908, &HC1C2031F, &H405605F8
    pvAppendBuffer &H4A84E850, &H16A0000, &H8B0875FF, &H65E856F0, &H8BFFFFFF, &HC25D5EC6, &H8B550004, &H84D8BEC, &H531CEC83, &HDB335756, &HD90481, &H8B000100, &H5483D914, &H8B0004D9, &HF04D944, &HC110C2AC, &H558910F8, &HF84589F4, &H750FFB83, &H41C93309, &HFC6583, &H570F1CEB, &H130F66C0, &H458BEC45, &HEC4D8BF0, &H45130F66, &HE4558BE4, &H8BFC4589, &HFB83E845, &H17B8D0F, &HF61B006A, &HAF0FDEF7
    pvAppendBuffer &H6AD12BF7, &HFC451B25, &HCBE85250, &H3FFFFF7, &H4D8BF445, &HF8551308, &H8301E883, &H40100DA, &HF8458BF1, &H4F15411, &HFF4558B, &HC110D0A4, &H142910E2, &HD94419D9, &H83DF8B04, &H820F10FB, &HFFFFFF64, &H8B5B5E5F, &H4C25DE5, &HEC8B5500, &H5610EC83, &H570C758B, &H5029E856, &H45890000, &H4468DF0, &H501DE850, &H45890000, &H8468DF4, &H5011E850, &H45890000, &HC468DF8, &H5005E850
    pvAppendBuffer &H4D8B0000, &HFC458908, &H798D318B, &HC1C68B04, &HF80304E0, &H57F0458D, &HF82EE850, &H21EBFFFF, &H39E0E8, &HF0458D00, &H3A6DE850, &H8D570000, &HE850F045, &HFFFFF814, &H50F0458D, &H398EE8, &H10EF8300, &H50F0458D, &H7501EE83, &H39B3E8D3, &H458D0000, &H40E850F0, &H5700003A, &H50F0458D, &HFFF7E7E8, &H10758BFF, &HF075FF56, &H6912E8, &H4468D00, &HF475FF50, &H6906E8, &H8468D00
    pvAppendBuffer &HF875FF50, &H68FAE8, &HC468D00, &HFC75FF50, &H68EEE8, &H8B5E5F00, &HCC25DE5, &HEC8B5500, &H5310EC83, &HC758B56, &H58E85657, &H8900004F, &H468DF045, &H4CE85004, &H8900004F, &H468DF445, &H40E85008, &H8900004F, &H468DF845, &H34E8500C, &H8B00004F, &H4589085D, &HF0458DFC, &H5604738D, &HF766E850, &HFF33FFFF, &H4710C683, &H2D763B39, &H50F0458D, &H5783E8, &HF0458D00, &H5679E850
    pvAppendBuffer &H458D0000, &H89E850F0, &H5600003A, &H50F0458D, &HFFF737E8, &H10C683FF, &H723B3B47, &HF0458DD3, &H5756E850, &H458D0000, &H4CE850F0, &H56000056, &H50F0458D, &HFFF713E8, &H10758BFF, &HF075FF56, &H683EE8, &H4468D00, &HF475FF50, &H6832E8, &H8468D00, &HF875FF50, &H6826E8, &HC468D00, &HFC75FF50, &H681AE8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H758B56EC, &HF46808, &H6A0000
    pvAppendBuffer &H39B9E856, &HC4830000, &H107D830C, &H83207410, &H7418107D, &H107D8310, &HC7257520, &HE06&, &HEB206A00, &HC06C712, &H6A000000, &HC708EB18, &HA06&, &HFF106A00, &HE8560C75, &HFFFFF6AF, &HCC25D5E, &HDC8B5300, &HE4835151, &H4C483F0, &H46B8B55, &H4246C89, &HEC83EC8B, &H84B8B2C, &H560C438B, &H918B106A, &H200&, &H8B38100F, &H280F5E01, &H66D60302, &HFF8EF0F, &H83F07D29
    pvAppendBuffer &H3C740AE8, &H1E88348, &H83481E74, &H850F01E8, &H95&, &H30A280F, &H380F66D6, &H280FF9DE, &H66D60302, &HF8DE380F, &H30A280F, &H380F66D6, &H280FF9DE, &H66D60302, &HF8DE380F, &HF07D290F, &H33A280F, &H32280FD6, &H290FD603, &H280FE075, &HF66F075, &HFF7DE38, &HD6032A28, &HF075290F, &HF07D280F, &HDE380F66, &H280FE07D, &H66D60322, &HFDDE380F, &HDE380F66, &H1A280FFC, &HF66D603
    pvAppendBuffer &HFFBDE38, &HD6031228, &HDE380F66, &HA280FFA, &HF66D603, &H66F9DE38, &H3ADE380F, &HDE380F66, &HF66323C, &H327CDF38, &H10438B10, &H38110F5E, &H8B5DE58B, &HCC25BE3, &HDC8B5300, &HE4835151, &H4C483F0, &H46B8B55, &H4246C89, &HEC83EC8B, &H84B8B2C, &H560C438B, &H918B106A, &H1FC&, &H8B38100F, &H280F5E01, &H66D60302, &HFF8EF0F, &H83F07D29, &H3C740AE8, &H1E88348, &H83481E74
    pvAppendBuffer &H850F01E8, &H95&, &H30A280F, &H380F66D6, &H280FF9DC, &H66D60302, &HF8DC380F, &H30A280F, &H380F66D6, &H280FF9DC, &H66D60302, &HF8DC380F, &HF07D290F, &H33A280F, &H32280FD6, &H290FD603, &H280FE075, &HF66F075, &HFF7DC38, &HD6032A28, &HF075290F, &HF07D280F, &HDC380F66, &H280FE07D, &H66D60322, &HFDDC380F, &HDC380F66, &H1A280FFC, &HF66D603, &HFFBDC38, &HD6031228, &HDC380F66
    pvAppendBuffer &HA280FFA, &HF66D603, &H66F9DC38, &H3ADC380F, &HDC380F66, &HF66323C, &H327CDD38, &H10438B10, &H38110F5E, &H8B5DE58B, &HCC25BE3, &HEC8B5500, &H10CEC81, &H458B0000, &H56C9330C, &HF445C757, &H8040201, &H788DF18B, &HF845C706, &H80402010, &H4BD148D, &H66000000, &H1BFC45C7, &HE47D8936, &H85E85589, &HE7840FD2, &H8B000000, &H558D0C7D, &HF05589F4, &HFEF4958D, &HE0C1FFFF, &H8BD02B02
    pvAppendBuffer &H55890845, &HF4958DEC, &H2BFFFFFE, &H84589C2, &H1B73F73B, &H8DB0048D, &HFFFEF48D, &H50C103FF, &H4C0AE8, &HEC558B00, &H86E9C933, &H8B000000, &HFEF0B584, &HD233FFFF, &H8B0C4589, &H83F7F7C6, &HA7508FF, &H724C68B, &H1B0043C, &HC18A0274, &H3275D285, &H4D6E0F66, &HFC0330C, &HF66D257, &H458BC06E, &H620F66F0, &H620F66C8, &H3A0F66D1, &HF00C2DF, &HF6600B6, &H1C2163A, &H45FFD033
    pvAppendBuffer &HC5589F0, &HC08427EB, &HF662374, &H330C4D6E, &HD2570FC0, &HC06E0F66, &HC8620F66, &HD1620F66, &HDF3A0F66, &HF6600C2, &HC45163A, &HEC558B00, &H4533028B, &HB584890C, &HFFFFFEF4, &H8B04C283, &H89460845, &H753BEC55, &H41820FE8, &H8BFFFFFF, &H558BE47D, &HF8858D10, &H8BFFFFFE, &H8DD62BF0, &H75890177, &H10758B08, &H1086D83, &H586E0F66, &H6E0F6608, &HF660440, &HF66086E, &H66FC506E
    pvAppendBuffer &H66D0620F, &H66CB620F, &HFD1620F, &H8D021429, &HD4751040, &HF068&, &H858D5100, &HFFFFFEF4, &H3619E850, &HCF8B0000, &HC10CC483, &H4D0304E1, &HFD23314, &HD2850628, &HC78B0B74, &H574C22B, &HDB380F66, &H1290FC0, &H10C68342, &H3B10E983, &H5FE076D7, &H5DE58B5E, &H550010C2, &HEC83EC8B, &H57565310, &H7D8DC033, &HC93340F0, &H8BA20F53, &H7895BF3, &H89047789, &H5789084F, &HF845F70C
    pvAppendBuffer &H2000000, &H45F75F74, &H80000F8, &H8B567400, &HC78B107D, &H8B08758B, &HD1F799CE, &HC103E283, &HC20302EF, &H8302F8C1, &H68906C0, &HF704468D, &HE8C140D0, &H3E08302, &HE9C14140, &H3E18302, &H8D40C183, &H96898614, &H1FC&, &H518E0C8D, &H75FF5752, &H8E890C, &HE8000002, &HFFFFFDCD, &HEB40C033, &H5FC03302, &HE58B5B5E, &HCC25D, &H81EC8B55, &H210EC, &H2875FF00, &HFDF0858D
    pvAppendBuffer &H75FFFFFF, &H8D505024, &H8D502845, &HE850F445, &H7A&, &H8D0875FF, &H106AF445, &H6A1475FF, &H2075FF0C, &HFF1C75FF, &H75FF1875, &HC75FF10, &H502875FF, &HCB9E8, &H5DE58B00, &H550024C2, &HEC81EC8B, &H210&, &H8D2875FF, &HFFFDF085, &H2475FFFF, &H458D5050, &H458D5028, &H2BE850F4, &H6A000000, &HC75FF10, &HFFF4458D, &HC6A0875, &HFF2075FF, &H75FF1C75, &H1475FF18, &HFF1075FF
    pvAppendBuffer &HE8502875, &HE1C&, &HC25DE58B, &H8B550024, &H575653EC, &H8B1C75FF, &H75FF107D, &HB8E85718, &H85FFFFFE, &HE81F74C0, &HFFFFE359, &H75939BE, &H4000BB00, &HF32B0007, &H46E8F003, &HB9FFFFE3, &H75845, &H75FF2CEB, &H147D8B1C, &H571875FF, &HFFFA9DE8, &HE32BE8FF, &H1DBEFFFF, &HBB000757, &H74000, &HF003F32B, &HFFE318E8, &H564DB9FF, &HCB2B0007, &H4D8BC103, &H8418908, &H890C458B
    pvAppendBuffer &H1C70471, &H10&, &H5E5F3889, &H18C25D5B, &HEC8B5500, &H7D8B5751, &H89C03318, &HFF85FC45, &H8B536374, &H8B560C5D, &H10758B0B, &H4D89F12B, &HFFE3B18, &HC085F742, &HB60F1A75, &H50561445, &H308458B, &HF4E850C1, &H8B000033, &HC483184D, &HFC458B0C, &HC75C985, &H3B42D233, &H440F1075, &HFC4589C2, &H3B0E048D, &HE751045, &HFF0875FF, &H55FF2075, &H23831C, &H330102EB, &H2BFC458B
    pvAppendBuffer &H5EA475FE, &HE58B5F5B, &H1CC25D, &H56EC8B55, &H8B20758B, &HE883C6, &H86840F, &H75FF0000, &H2475FF28, &HE883016A, &H83617401, &H458D01E8, &H75FF5014, &HC75FF10, &H740875FF, &H67E848, &H75FF0000, &H184D8B28, &H382475FF, &H20741C4D, &H50FE468D, &H1075FF51, &HFF0C75FF, &H2EE80875, &HFFFFFFFF, &H458D2875, &H2475FF1C, &H25EB016A, &H50FF468D, &H1075FF51, &HFF0C75FF, &HEE80875
    pvAppendBuffer &HEBFFFFFF, &H1FE81F, &HD7EB0000, &H301C458A, &H458D1445, &H75FF5014, &HC75FF10, &HE80875FF, &H5&, &H24C25D5E, &HEC8B5500, &HFF2075FF, &H75FF1C75, &H1875FF1C, &HFF1475FF, &H75FF1075, &H875FF0C, &H4E8&, &H1CC25D00, &HEC8B5500, &H530C4D8B, &H56145D8B, &H57003983, &H74187D8B, &H74FF854A, &H10458B46, &H12BF78B, &H420FC73B, &H3018BF0, &H53560845, &H329CE850, &H4D8B0000
    pvAppendBuffer &H2BDE030C, &HCC483FE, &H758B3101, &H75313910, &H875FF33, &H852475FF, &HFF0575FF, &H3EB2055, &H8B1C55FF, &H20830C45, &H8B19EB00, &H14EB1075, &H2475FF53, &H575FE3B, &HEB2055FF, &H1C55FF03, &HFE2BDE03, &HE873FE3B, &H2E74FF85, &H8B0C458B, &H2BC68B08, &H3BF78BC1, &HF0420FC7, &H5608458B, &H5053C103, &H3231E8, &HC458B00, &HC483DE03, &H2B30010C, &H10758BFE, &H5E5FD575, &H20C25D5B
    pvAppendBuffer &HEC8B5500, &H1C7D8B57, &H4474FF85, &HC5D8B53, &H3B8356, &H75FF1075, &H2475FF08, &H8B2055FF, &H3891045, &H458B03EB, &H8B032B10, &HF3B39F7, &H45033342, &HFF505608, &H75FF1475, &H6118E818, &H33290000, &H1187501, &HFE2B1475, &H5B5EC375, &H20C25D5F, &HEC8B5500, &HFFE0A8E8, &H6525B9FF, &HE9810007, &H74000, &H4D8BC103, &HFF505108, &H418D1475, &H1075FF74, &H6A0C75FF, &H418D5040
    pvAppendBuffer &H7AE85034, &H5DFFFFFF, &H550010C2, &HEC83EC8B, &H8B565370, &H56571475, &H4692E8, &H44E8D00, &H8951F88B, &H84E8D07D, &H8D000046, &H4589084E, &H458951F8, &H4675E8CC, &H4E8D0000, &HF445890C, &HC8458951, &H4666E8, &H8758B00, &HF0458956, &HE8C44589, &H4657&, &H89044E8D, &H8951D445, &H48E8C045, &H8D000046, &H4589084E, &H458951E8, &H4639E8BC, &H4E8D0000, &HE445890C, &HB8458951
    pvAppendBuffer &H462AE8, &HC758B00, &HE0458956, &HE8B44589, &H461B&, &H89044E8D, &H8951FC45, &HCE8B045, &H8D000046, &H4589084E, &H45895108, &H45FDE8AC, &H4E8D0000, &H1445890C, &HA8458951, &H45EEE8, &H10758B00, &HC458956, &HE8A44589, &H45DF&, &H8B044E8D, &H5D8951D8, &H45D1E8A0, &H4E8D0000, &HDC458908, &H9C458951, &H45C2E8, &HC4E8D00, &H51D84589, &HE8984589, &H45B3&, &H8BDC4D8B
    pvAppendBuffer &HD8558BF0, &HC7947589, &HA9045, &H6EB0000, &H8BEC758B, &H7D03107D, &HFC458BD4, &HC3C1DF33, &H89C30310, &H4533FC45, &HCC0C1D4, &HDF33F803, &H8B107D89, &HC3C1FC7D, &H89FB0308, &HF833FC7D, &H3F8458B, &HC833E845, &H8BF84589, &HC1C10845, &HC1C10310, &H458907C7, &HE8453308, &H10CC0C1, &H4D33F845, &H8C1C1F8, &H89084D01, &H4D8BDC4D, &H8BC83308, &H4503F445, &H89D033E4, &H458BF445
    pvAppendBuffer &H10C2C114, &HC1C1C203, &H14458907, &HC1E44533, &H45010CC0, &HF45533F4, &H108C2C1, &H4D011455, &HD8558910, &H3314558B, &HF0458BD0, &H33E04503, &HF04589F0, &HC10C458B, &HC60310C6, &H8907C2C1, &H45330C45, &HCC0C1E0, &H33F04501, &HC6C1F075, &HC750108, &H8BEC7589, &HF0330C75, &H33EC458B, &HC0C11045, &H14450110, &H8BEC4589, &HC1331445, &HC107C6C1, &H45010CC0, &HEC4D8B10, &HC1104D33
    pvAppendBuffer &H4D0108C1, &HEC4D8914, &H33144D8B, &HC458BC8, &H8907C1C1, &H4D8BE84D, &H33CA03F8, &H10C3C1D9, &H4589C303, &H8BC2330C, &HC0C1F455, &H3D6030C, &H89D933C8, &H4D8BF84D, &H8C3C10C, &H4D89CB03, &H8BC8330C, &HC1C1FC45, &HE44D8907, &H33DC4D8B, &H10C1C1CA, &H4589C103, &H8BC633FC, &HC0C1F075, &H3F7030C, &H89CA33D0, &H558BF455, &H8C1C1FC, &H5589D103, &H8BD033FC, &HC2C10845, &HE0558907
    pvAppendBuffer &H33D8558B, &H10C2C1D6, &H4589C203, &H8BC73308, &HC0C1087D, &H33F0030C, &HF07589D6, &H308C2C1, &H87D89FA, &HC7C1F833, &H906D8307, &HD47D8901, &HFE5A850F, &H458BFFFF, &H104503D0, &H8BD04589, &H4503CC45, &HCC4589F8, &H3C8458B, &H4589F445, &HBC458BC8, &H89E84503, &H458BBC45, &HE44503B8, &H8BB84589, &H4503B445, &HB44589E0, &H3AC458B, &H5D010845, &HAC4589A0, &H3A8458B, &H55011445
    pvAppendBuffer &H18558B98, &H89D05D8B, &H458BA845, &HC4503A4, &H8BA44589, &H45039445, &H891A88EC, &HC38B9445, &H8808E8C1, &HC38B0142, &H8810E8C1, &H75010242, &H18EBC1C4, &H8B035A88, &HC38BCC5D, &HC1045A88, &H428808E8, &H1C38B05, &HE8C1C07D, &H6428810, &H1FC7D8B, &H4D01B07D, &H18EBC19C, &H8B075A88, &HC38BC85D, &HC1085A88, &H428808E8, &HC1C38B09, &H428810E8, &H18EBC10A, &H8B0B5A88, &HC38BC45D
    pvAppendBuffer &HC10C5A88, &H428808E8, &HC1C38B0D, &H428810E8, &H18EBC10E, &H8B0F5A88, &HC38BC05D, &HC1105A88, &H428808E8, &HC1C38B11, &H428810E8, &H18EBC112, &H8B135A88, &HC38BBC5D, &HC1145A88, &H428808E8, &HC1C38B15, &H428810E8, &H18EBC116, &H8B175A88, &HC38BB85D, &HC1185A88, &H428808E8, &HC1C38B19, &H428810E8, &H18EBC11A, &H8B1B5A88, &HC38BB45D, &HC11C5A88, &H428808E8, &HC1C38B1D, &H428810E8
    pvAppendBuffer &H18EBC11E, &H8B1F5A88, &HC38BB05D, &HC1205A88, &H428808E8, &HC1C38B21, &H428810E8, &H18EBC122, &H8B235A88, &HC38BAC5D, &HC1245A88, &H428808E8, &HC1C38B25, &H428810E8, &H18EBC126, &H8B275A88, &HC38BA85D, &HC1285A88, &H428808E8, &HC1C38B29, &H428810E8, &H18EBC12A, &H8B2B5A88, &HC38BA45D, &HC12C5A88, &H428808E8, &HC1C38B2D, &H428810E8, &H18EBC12E, &H8B2F5A88, &HC38BA05D, &HC1305A88
    pvAppendBuffer &H428808E8, &H3C4A8D31, &HEBC1C38B, &H10E8C118, &H88324288, &H5D8B335A, &H88C38B9C, &HE8C1345A, &H35428808, &HE8C1C38B, &H36428810, &H8818EBC1, &H5D8B375A, &H88C38B98, &HE8C1385A, &H39428808, &HE8C1C38B, &H3A428810, &H8818EBC1, &H558B3B5A, &HC1C28B94, &H118808E8, &H8B014188, &HE8C15FC2, &H18EAC110, &H241885E, &H5B035188, &HC25DE58B, &H8B550014, &H75FF56EC, &H8758B10, &H560C75FF
    pvAppendBuffer &H4392E8, &HFF106A00, &H468D1475, &H87E85020, &H8B00002C, &HC4831845, &H7466830C, &H78468900, &H14C25D5E, &HEC8B5500, &H8758B56, &HC75FF57, &H8D3076FF, &H8D57207E, &H56501046, &HFFFACAE8, &H78568BFF, &H780C933, &H410B7501, &H674CA3B, &H1390480, &H5E5FF574, &H8C25D, &H83EC8B55, &H458D10EC, &HFF106AF0, &HE8502075, &H2C2A&, &H8D0CC483, &H6A50F045, &H2475FF00, &HFF1C75FF
    pvAppendBuffer &H75FF1875, &H1075FF14, &HFF0C75FF, &HECE80875, &H8B00003E, &H20C25DE5, &HEC8B5500, &H6A2475FF, &H2075FF01, &HFF1C75FF, &H75FF1875, &H1075FF14, &HFF0C75FF, &HC4E80875, &H5D00003E, &H550020C2, &HBEE8EC8B, &HB9FFFFDA, &H774F3, &H4000E981, &HC1030007, &H51084D8B, &H1475FF50, &H75FF018B, &HC75FF10, &H418D30FF, &H418D5028, &H8EE85018, &H5DFFFFF9, &H550010C2, &H4D8BEC8B, &HC458B08
    pvAppendBuffer &H8B2C4189, &H41891045, &HCC25D30, &HEC8B5500, &H8758B56, &H6A346A, &H2BA1E856, &H4D8B0000, &H2C66830C, &H83018B00, &H89002866, &H458B3046, &H4468910, &H8908468D, &HFF31FF0E, &HE8501475, &H2B56&, &H5E18C483, &H10C25D, &H81EC8B55, &H420EC, &H57565300, &H85C7DB33, &HFFFFFD60, &HDB41&, &H858D706A, &HFFFFFD70, &HFD649D89, &H5053FFFF, &HFD6885C7, &H1FFFF, &H9D890000
    pvAppendBuffer &HFFFFFD6C, &H2B36E8, &HC758B00, &HFF60858D, &H1F6AFFFF, &HFFE85056, &H8A00002A, &HC4831F46, &H60A58018, &HF8FFFFFF, &H400C3F24, &HFF7F8588, &H858DFFFF, &HFFFFFBE0, &H501075FF, &H48E0E8, &H591E6A00, &H8DC0570F, &HFFFE60B5, &H661E6AFF, &H6085130F, &H8DFFFFFE, &HFFFE68BD, &H59A5F3FF, &H45130F66, &H80758D80, &H66887D8D, &HE085130F, &HF3FFFFFE, &H591E6AA5, &HFEE0B58D, &H9D89FFFF
    pvAppendBuffer &HFFFFFE64, &HFEE8BD8D, &H5D89FFFF, &H6AA5F384, &HB58D5920, &HFFFFFBE0, &HFEBB&, &HE0BD8D00, &HF3FFFFFD, &H47FF33A5, &HFE60BD89, &H7D89FFFF, &HFC38B80, &HF8C1CBB6, &H7E18303, &H5B4B60F, &HFFFFFF60, &HFDE0858D, &HEED3FFFF, &H5056F723, &H5080458D, &H40BAE8, &H858D5600, &HFFFFFE60, &HE0858D50, &H50FFFFFE, &H40A6E8, &HE0858D00, &H50FFFFFE, &H5080458D, &HFCE0858D, &HE850FFFF
    pvAppendBuffer &HFFFFE6B7, &HFEE0858D, &H8D50FFFF, &H50508045, &H472FE8, &H60858D00, &H50FFFFFE, &HFDE0858D, &H8D50FFFF, &HFFFEE085, &H8CE850FF, &H8DFFFFE6, &HFFFE6085, &H858D50FF, &HFFFFFDE0, &H1E85050, &H8D000047, &HFFFCE085, &H858D50FF, &HFFFFFE60, &H46D9E850, &H458D0000, &H858D5080, &HFFFFFC60, &H46C9E850, &H458D0000, &H858D5080, &HFFFFFEE0, &H80458D50, &H3678E850, &H858D0000, &HFFFFFCE0
    pvAppendBuffer &HE0858D50, &H50FFFFFD, &HFEE0858D, &HE850FFFF, &H365E&, &HFEE0858D, &H8D50FFFF, &H8D508045, &HFFFCE085, &H10E850FF, &H8DFFFFE6, &HFFFEE085, &H458D50FF, &HE8505080, &H4688&, &H5080458D, &HFDE0858D, &HE850FFFF, &H4663&, &HFC60858D, &H8D50FFFF, &HFFFE6085, &H858D50FF, &HFFFFFEE0, &H465EE850, &H858D0000, &HFFFFFD60, &HE0858D50, &H50FFFFFE, &H5080458D, &H35F5E8, &H60858D00
    pvAppendBuffer &H50FFFFFE, &H5080458D, &HE5ADE850, &H458DFFFF, &H858D5080, &HFFFFFEE0, &HD3E85050, &H8D000035, &HFFFC6085, &H858D50FF, &HFFFFFE60, &H80458D50, &H35BCE850, &H858D0000, &HFFFFFBE0, &HE0858D50, &H50FFFFFD, &HFE60858D, &HE850FFFF, &H35A2&, &HFCE0858D, &H8D50FFFF, &HFFFDE085, &HCCE850FF, &H56000045, &HFDE0858D, &H8D50FFFF, &HE8508045, &H3F1F&, &H60858D56, &H50FFFFFE, &HFEE0858D
    pvAppendBuffer &HE850FFFF, &H3F0B&, &HF01EB83, &HFFFE2089, &HE0858DFF, &H50FFFFFE, &H2695E850, &H858D0000, &HFFFFFEE0, &H80458D50, &H43E85050, &H8D000035, &HFF508045, &HB1E80875, &H5F000037, &HE58B5B5E, &HCC25D, &H83EC8B55, &H6A5720EC, &HC0335907, &H9E045C6, &HF3E17D8D, &HAAAB66AB, &H50E0458D, &HFF0C75FF, &HE1E80875, &H5FFFFFFC, &HC25DE58B, &H8B550008, &H14EC81EC, &H53000001, &HC0335756
    pvAppendBuffer &H3308758B, &HE17D8DDB, &HABE05D88, &HAB66ABAB, &HE0458DAA, &H75FF5050, &H456FF0C, &HC247D83, &HC6A1E75, &H8D2075FF, &HE850F045, &H27EA&, &H660CC483, &H88FD5D89, &H45C6FC5D, &H32EB01FF, &H50E0458D, &HFEEC858D, &HE850FFFF, &H1E04&, &H8D2475FF, &HFFFEEC85, &H2075FFFF, &H1CDBE850, &H458D0000, &H858D50F0, &HFFFFFEEC, &H1D68E850, &H458D0000, &H858D50E0, &HFFFFFF3C, &H1DD2E850
    pvAppendBuffer &H75FF0000, &H3C858D1C, &HFFFFFFFF, &HE8501875, &H1C8B&, &H5D88C033, &HD17D8DD0, &H66ABABAB, &H458DAAAB, &H75FF50F0, &H8C458D0C, &HDAE85056, &H6AFFFFFB, &H8D0C6A04, &HE8508C45, &HFFFFFBB7, &H458D106A, &H8D5050D0, &HE8508C45, &HFFFFFB6F, &H8D1475FF, &HFFFF3C85, &H1075FFFF, &H1C5BE850, &H458D0000, &H858D50C0, &HFFFFFF3C, &H1CE8E850, &H558B0000, &H85CB8B2C, &H8A0D74D2, &H30D00D44
    pvAppendBuffer &H41C00D44, &HF372CA3B, &H85C04D8D, &H8B1874D2, &HC18B2875, &H48AF02B, &HA01320E, &HEA8341D8, &H84F37501, &HFF1675DB, &H458D1475, &H3075FF8C, &H501075FF, &HFFFB06E8, &HEBF633FF, &H46F63303, &H7D8DC033, &H506AABE0, &HABAB006A, &HE0458AAB, &H33F07D8D, &HABABABC0, &HF0458AAB, &H33D07D8D, &HABABABC0, &HD0458AAB, &H33C07D8D, &HABABABC0, &HC0458AAB, &HFF3C858D, &HE850FFFF, &H26BF&
    pvAppendBuffer &HFF3C8D8A, &H458DFFFF, &H6A346A8C, &HACE85000, &H8A000026, &HC4838C4D, &H5FC68B18, &HE58B5B5E, &H2CC25D, &H81EC8B55, &H114EC, &H57565300, &H758BC033, &H8DDB3308, &H5D88E17D, &HABABABE0, &H8DAAAB66, &H5050E045, &HFF0C75FF, &H7D830456, &H1E750C24, &H75FF0C6A, &HF0458D20, &H2638E850, &HC4830000, &H5D89660C, &HFC5D88FD, &H1FF45C6, &H458D32EB, &H858D50E0, &HFFFFFEEC, &H1C52E850
    pvAppendBuffer &H75FF0000, &HEC858D24, &HFFFFFFFE, &HE8502075, &H1B29&, &H50F0458D, &HFEEC858D, &HE850FFFF, &H1BB6&, &H50E0458D, &HFF3C858D, &HE850FFFF, &H1C20&, &H8D1C75FF, &HFFFF3C85, &H1875FFFF, &H1AD9E850, &HC0330000, &H8DD05D88, &HABABD17D, &HAAAB66AB, &H50F0458D, &H8D0C75FF, &H50568C45, &HFFFA28E8, &H6A046AFF, &H8C458D0C, &HFA05E850, &H106AFFFF, &H50D0458D, &H8C458D50, &HF9BDE850
    pvAppendBuffer &H75FFFFFF, &H8C458D14, &HFF2875FF, &HE8501075, &HFFFFF9AB, &H8D1475FF, &HFFFF3C85, &H2875FFFF, &H1A97E850, &HC0330000, &H8DC05D88, &HABABC17D, &HAAAB66AB, &H50C0458D, &HFF3C858D, &HE850FFFF, &H1B16&, &H8D3075FF, &H8D50D045, &HFF50C045, &H83E82C75, &H33000054, &HE07D8DC0, &HABABABAB, &H8DE0458A, &HC033F07D, &HABABABAB, &H8DF0458A, &HC033D07D, &H53506AAB, &H8AABABAB, &H7D8DD045
    pvAppendBuffer &HABC033C0, &H8AABABAB, &H858DC045, &HFFFFFF3C, &H2529E850, &H858A0000, &HFFFFFF3C, &H458D346A, &HE850538C, &H2517&, &H838C458A, &H5E5F18C4, &H5DE58B5B, &H55002CC2, &H558BEC8B, &H104D8B0C, &H8758B56, &H233068B, &H468B0189, &H4423304, &H8B044189, &H42330846, &H8418908, &H330C468B, &H41890C42, &HC25D5E0C, &H8B55000C, &H8B5351EC, &H57560C5D, &H66087D8B, &HFC45C7, &H8B0F8BE1
    pvAppendBuffer &H83E8D1C1, &H38901E1, &H8B04578B, &H83E8D1C2, &HE1C101E2, &HC1C80B1F, &H4B891FE2, &H8778B04, &HE8D1C68B, &HB01E683, &H1FE6C1D0, &H8B085389, &HC18B0C4F, &HE183E8D1, &H5FF00B01, &HF0C7389, &HFC0D44B6, &H3118E0C1, &H8B5B5E03, &H8C25DE5, &HEC8B5500, &H8758B56, &H28E85657, &H8B000039, &H7890C7D, &H5004468D, &H391AE8, &H4478900, &H5008468D, &H390EE8, &H8478900, &H500C468D
    pvAppendBuffer &H3902E8, &HC478900, &HC25D5E5F, &H8B550008, &H20EC83EC, &H106A5756, &H458DF633, &HE85056E0, &H241B&, &H75FF106A, &HF0458D0C, &H23E8E850, &H7D8B0000, &H18C48308, &HE04D100F, &HE083C68B, &H591F6A1F, &HC68BC82B, &H8B05F8C1, &HE8D38704, &HC7401A8, &HF045100F, &HC8EF0F66, &HE04D110F, &H50F0458D, &HFF00E850, &H8146FFFF, &H80FE&, &H6AC97C00, &HE0458D10, &H1075FF50, &H2399E8
    pvAppendBuffer &HCC48300, &HE58B5E5F, &HCC25D, &H8BEC8B55, &H100F0845, &HC458B00, &HE8700F66, &HD5280F1B, &HFDD280F, &H458B0010, &H700F6610, &H280F1BC8, &H3A0F66C5, &H6601C144, &HD1443A0F, &HEF0F6610, &H3A0F66D0, &H6600D944, &HE9443A0F, &HC2280F11, &HDA730F66, &H730F6608, &HF6608F8, &HF66EAEF, &H280FD8EF, &HC3280FE5, &HD4720F66, &H720F661F, &HF661FD0, &HF01F372, &HF66C828, &H6604FC73
    pvAppendBuffer &HCD8730F, &HF9730F66, &HEB0F6604, &H720F66CB, &H280F01F5, &HEB0F66D9, &H720F66E5, &HF661FF3, &H280FE0EB, &H720F66C1, &HF661EF0, &H280FD8EF, &H720F66C1, &HF6619F0, &H280FD8EF, &H730F66D3, &HF6604DB, &H660CFA73, &HFD1EF0F, &H280FCA28, &H720F66C2, &HF6602D1, &H6601D072, &HFC8EF0F, &HF66C228, &H6607D072, &H66C8EF0F, &H66CBEF0F, &H66CAEF0F, &H66CCEF0F, &H1BC1700F, &H5D00110F
    pvAppendBuffer &H55000CC2, &H8B56EC8B, &H8B570C75, &HFF56087D, &H50EDE837, &H468D0000, &H77FF5004, &H50E1E804, &H468D0000, &H77FF5008, &H50D5E808, &H468D0000, &H77FF500C, &H50C9E80C, &H5E5F0000, &H8C25D, &H83EC8B55, &H8B5644EC, &HBE830875, &HA8&, &H56067400, &H33BBE8, &HFC93300, &H880E84B6, &H89000000, &H41BC8D44, &H7210F983, &HFC6583EE, &H28E85600, &H8D000033, &H5650BC45, &H32D0E8
    pvAppendBuffer &HC558B00, &H48AC933, &H1104888E, &H10F98341, &HAC68F472, &H6A000000, &H20E85600, &H8A000022, &HCC48306, &H5DE58B5E, &H550008C2, &H8B56EC8B, &HAC680875, &H6A000000, &HE85600, &H8B000022, &HFCBA0C4D, &H6A000000, &H1075FF10, &H8901B60F, &HB60F4446, &H46890141, &H41B60F48, &H4C468902, &H341B60F, &H890FE083, &HB60F5046, &HC2230441, &HF544689, &H890541B6, &HB60F5846, &H46890641
    pvAppendBuffer &H41B60F5C, &HFE08307, &HF604689, &H230841B6, &H644689C2, &H941B60F, &HF684689, &H890A41B6, &HB60F6C46, &HE0830B41, &H7046890F, &HC41B60F, &H4689C223, &H41B60F74, &H7846890D, &HE41B60F, &HF7C4689, &H830F41B6, &H84A6&, &HE0830000, &H8086890F, &H8D000000, &H8886&, &H37E85000, &H83000021, &H5D5E18C4, &H55000CC2, &H12E8EC8B, &HB9FFFFD0, &H7A44D, &H4000E981, &HC1030007
    pvAppendBuffer &H51084D8B, &H1075FF50, &HA8818D, &H75FF0000, &H50106A0C, &H98818D, &HE8500000, &HFFFFEE09, &HCC25D, &H83EC8B55, &H56530CEC, &HCFD3E857, &H4D8BFFFF, &HA90ABE08, &H406A0007, &HEE815A, &H3000740, &H64798DF0, &HF760418B, &H3006AE2, &H5ADA8B07, &HF08B5651, &H406ADA13, &H84E8D58, &H2B3FE183, &H525250C1, &H8068&, &H57406A00, &H8D087D8B, &HE8502047, &HFFFFED10, &H50F4458D
    pvAppendBuffer &H3F3A40F, &H3E6C153, &H4F1CE856, &H86A0000, &H50F4458D, &HC5E857, &H758B0000, &H37FF560C, &H4EDEE8, &H4468D00, &H477FF50, &H4ED2E8, &H8468D00, &H877FF50, &H4EC6E8, &HC468D00, &HC77FF50, &H4EBAE8, &H10468D00, &H1077FF50, &H4EAEE8, &H14468D00, &H1477FF50, &H4EA2E8, &H18468D00, &H1877FF50, &H4E96E8, &H1C468D00, &H1C77FF50, &H4E8AE8, &H6A686A00, &H3CE85700
    pvAppendBuffer &H83000020, &H5E5F0CC4, &H5DE58B5B, &H550008C2, &H8B56EC8B, &H686A0875, &HE856006A, &H201F&, &HC70CC483, &H9E66706, &H446C76A, &HBB67AE85, &H720846C7, &HC73C6EF3, &HF53A0C46, &H46C7A54F, &HE527F10, &H1446C751, &H9B05688C, &HAB1846C7, &HC71F83D9, &HCD191C46, &H5D5E5BE0, &H550004C2, &H9EE8EC8B, &HB9FFFFCE, &H7A90A, &H4000E981, &HC1030007, &H51084D8B, &H1075FF50, &HFF64418D
    pvAppendBuffer &H406A0C75, &H20418D50, &HEC9BE850, &HC25DFFFF, &H8B55000C, &H40EC83EC, &H50C0458D, &HE80875FF, &HA7&, &H458D306A, &H75FF50C0, &H1F68E80C, &HC4830000, &H5DE58B0C, &H550008C2, &H8B56EC8B, &HC8680875, &H6A000000, &H70E85600, &H8300001F, &H6C70CC4, &HC1059ED8, &H5D0446C7, &HC7CBBB9D, &HD5070846, &H46C7367C, &H9A292A0C, &H1046C762, &H3070DD17, &H5A1446C7, &HC7915901, &H59391846
    pvAppendBuffer &H46C7F70E, &H2FECD81C, &H2046C715, &HFFC00B31, &H672446C7, &HC7673326, &H15112846, &H46C76858, &HB44A872C, &H3046C78E, &H64F98FA7, &HD3446C7, &HC7DB0C2E, &H4FA43846, &H46C7BEFA, &HB5481D3C, &HC25D5E47, &HB8E90004, &H55000001, &H5151EC8B, &HB908458B, &H80&, &H8D575653, &HC4B0&, &HC0808B00, &HF7000000, &H8BD88BE1, &H831E03FA, &H8EE800D7, &H8BFFFFCD, &HB0B90855, &H520007AA
    pvAppendBuffer &H4000E981, &HC1030007, &H50104B8D, &HB87FE183, &H80&, &H6A50C12B, &HB8006A00, &H80&, &H8D565050, &HE8504042, &HFFFFEAE0, &H50F8458D, &H6A006A, &H4CF1E8, &H8D086A00, &HFF50F845, &H3CE80875, &H8D000001, &HF50F845, &H5703DFA4, &H5303E3C1, &H4CD1E8, &H85D8B00, &H6AF8458D, &HE8535008, &H11B&, &H560C758B, &HFF0473FF, &H4CB4E833, &H468D0000, &H73FF5008, &H873FF0C
    pvAppendBuffer &H4CA5E8, &H10468D00, &H1473FF50, &HE81073FF, &H4C96&, &H5018468D, &HFF1C73FF, &H87E81873, &H8D00004C, &HFF502046, &H73FF2473, &H4C78E820, &H468D0000, &H73FF5028, &H2873FF2C, &H4C69E8, &H30468D00, &H3473FF50, &HE83073FF, &H4C5A&, &H5038468D, &HFF3C73FF, &H4BE83873, &H6800004C, &HC8&, &HE853006A, &H1DD3&, &H5F0CC483, &HE58B5B5E, &H8C25D, &H56EC8B55, &H6808758B
    pvAppendBuffer &HC8&, &HE856006A, &H1DB3&, &HC70CC483, &HBCC90806, &H446C7F3, &H6A09E667, &H3B0846C7, &HC784CAA7, &HAE850C46, &H46C7BB67, &H94F82B10, &H1446C7FE, &H3C6EF372, &HF11846C7, &HC75F1D36, &HF53A1C46, &H46C7A54F, &HE682D120, &H2446C7AD, &H510E527F, &H1F2846C7, &HC72B3E6C, &H688C2C46, &H46C79B05, &H41BD6B30, &H3446C7FB, &H1F83D9AB, &H793846C7, &HC7137E21, &HCD193C46, &H5D5E5BE0
    pvAppendBuffer &H550004C2, &HFAE8EC8B, &HB9FFFFCB, &H7AAB0, &H4000E981, &HC1030007, &H51084D8B, &H1075FF50, &HC4818D, &H75FF0000, &H80680C, &H8D500000, &HE8504041, &HFFFFE9F1, &HCC25D, &H56EC8B55, &HCBACE857, &H7D8BFFFF, &H8D0F8B08, &H48D0C, &HFF510000, &H8B0F8B10, &H8D0C8DF0, &H4&, &HE8565751, &H1CB2&, &H8B0CC483, &H5D5E5FC6, &H550004C2, &H8B56EC8B, &H75FF0875, &H8D0E8B0C
    pvAppendBuffer &HFF500846, &H51FF0476, &H2C568B04, &H3304E8B, &H4EB5ED6, &H874C985, &HA448049, &HF4740108, &H8C25D, &H8BEC8B55, &H56530C45, &H8DDB3357, &HC0831A78, &HC458918, &HFF47B60F, &H8BC88B99, &HC458BF2, &H80C6D83, &H9900B60F, &H8C2A40F, &HE0C1F20B, &HFC80B08, &HA40F07B6, &H8D9908CE, &HE1C1F87F, &HBF20B08, &H47B60FC8, &HCEA40F09, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0A
    pvAppendBuffer &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0B, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0C, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0D, &HE1C19908, &HBF20B08, &H8458BC8, &H89D80C89, &H4304D874, &HF04FB83, &HFFFF6B82, &H5B5E5FFF, &H8C25D, &H8BEC8B55, &H56530C45, &H8DDB3357, &HC0832A78, &HC458928, &HFF47B60F, &H8BC88B99, &HC458BF2, &H80C6D83, &H9900B60F, &H8C2A40F
    pvAppendBuffer &HE0C1F20B, &HFC80B08, &HA40F07B6, &H8D9908CE, &HE1C1F87F, &HBF20B08, &H47B60FC8, &HCEA40F09, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0A, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0B, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0C, &HE1C19908, &HBF20B08, &H47B60FC8, &HCEA40F0D, &HE1C19908, &HBF20B08, &H8458BC8, &H89D80C89, &H4304D874, &HF06FB83, &HFFFF6B82, &H5B5E5FFF
    pvAppendBuffer &H8C25D, &H83EC8B55, &H458D60EC, &HC75FFE0, &HFE8EE850, &H46AFFFF, &H50E0458D, &H3A5DE8, &H74C08500, &HEBC03304, &H46A567F, &H50E0458D, &HFFC9BDE8, &H90BEFF, &HC6030000, &H39EDE850, &HF8830000, &H6A147401, &HC9A4E804, &HC603FFFF, &HE0458D50, &H62E85050, &H6A000048, &HE0458D00, &HC98CE850, &HC083FFFF, &H458D5050, &H23E850A0, &H8DFFFFCD, &HE850A045, &HFFFFCCBE, &H474C085
    pvAppendBuffer &H23EBC033, &H8D08758B, &H8D51A04D, &H6C6014E, &HD2E85104, &H8D000000, &H8D51C04D, &HE851214E, &HC5&, &H5E40C033, &HC25DE58B, &H8B550008, &H90EC81EC, &H8D000000, &H75FFD045, &H91E8500C, &H6AFFFFFE, &HD0458D06, &H39B0E850, &HC0850000, &HC0330774, &H8AE9&, &H66A5600, &H50D0458D, &HFFC90DE8, &H170BEFF, &HC6030000, &H393DE850, &HF8830000, &H6A147401, &HC8F4E806, &HC603FFFF
    pvAppendBuffer &HD0458D50, &HB2E85050, &H6A000047, &HD0458D00, &HC8DCE850, &H1005FFFF, &H50000001, &HFF70858D, &HE850FFFF, &HFFFFCE1A, &HFF70858D, &HE850FFFF, &HFFFFCC34, &H474C085, &H26EBC033, &H8D08758B, &HFFFF708D, &H4E8D51FF, &H406C601, &HAEE851, &H4D8D0000, &H4E8D51A0, &HA1E85131, &H33000000, &H8B5E40C0, &H8C25DE5, &HEC8B5500, &H5308458B, &H7D8B5756, &H18488D0C, &H4D89F633, &H1A588D08
    pvAppendBuffer &H7F7448A, &H28B10188, &H6F7448A, &H8BFF4388, &H548BF704, &H1E804F7, &H88FFFFD6, &H8B20B103, &H548BF704, &HF1E804F7, &H88FFFFD5, &H5B8D0143, &HF70C8BF8, &H4F7448B, &H18C1AC0F, &H8818E8C1, &HC8B0A4B, &HF7448BF7, &HC1AC0F04, &H10E8C110, &H8B0B4B88, &H448BF70C, &HAC0F04F7, &HE8C108C1, &HC4B8808, &H46F7048A, &H83084D8B, &H438808E9, &H84D890D, &H7204FE83, &H5B5E5F87, &H8C25D
    pvAppendBuffer &H8BEC8B55, &H56530845, &HC7D8B57, &H3328488D, &H84D89F6, &H8A2A588D, &H8807F744, &H8A28B101, &H8806F744, &H48BFF43, &HF7548BF7, &HD56AE804, &H388FFFF, &H48B20B1, &HF7548BF7, &HD55AE804, &H4388FFFF, &HF85B8D01, &H8BF70C8B, &HF04F744, &HC118C1AC, &H4B8818E8, &HF70C8B0A, &H4F7448B, &H10C1AC0F, &H8810E8C1, &HC8B0B4B, &HF7448BF7, &HC1AC0F04, &H8E8C108, &H8A0C4B88, &H8B46F704
    pvAppendBuffer &HE983084D, &HD438808, &H83084D89, &H877206FE, &H5D5B5E5F, &H550008C2, &HEC83EC8B, &H5D8B5320, &H56C03308, &H570C758B, &H8959066A, &H7D8DE445, &HE045C7E8, &H3&, &H468DABF3, &HE8535001, &HFFFFFBE0, &H75043E80, &H21468D0F, &H20438D50, &HFBCEE850, &H76EBFFFF, &H8D53046A, &HE857207B, &H4195&, &H7E8046A, &H83FFFFC7, &H8D5010C0, &H5750E045, &H41A9E857, &H46A0000, &HE8575753
    pvAppendBuffer &H4148&, &HE7E8046A, &H83FFFFC6, &HE85010C0, &HFFFFC6DE, &H5030C083, &HCDE85757, &H5700003B, &H18A5E8, &H8B068A00, &HF01240F, &HE183C0B6, &HC83B9901, &HC0330675, &H1274C23B, &HE857046A, &HFFFFC6AE, &H5010C083, &H456FE857, &H5E5F0000, &H5DE58B5B, &H550008C2, &HEC83EC8B, &H5D8B5330, &H56C03308, &H570C758B, &H89590A6A, &H7D8DD445, &HD045C7D8, &H3&, &H468DABF3, &HE8535001
    pvAppendBuffer &HFFFFFBD0, &H75043E80, &H31468D0F, &H30438D50, &HFBBEE850, &H7DEBFFFF, &H8D53066A, &HE857307B, &H40D5&, &H47E8066A, &H5FFFFC6, &HB0&, &HD0458D50, &HE8575750, &H40E7&, &H5753066A, &H4086E857, &H66A0000, &HFFC625E8, &HB0BBFF, &HC3030000, &HC618E850, &HE005FFFF, &H50000000, &H5E85757, &H5700003B, &H1871E8, &H8B068A00, &HF01240F, &HE183C0B6, &HC83B9901, &HC0330675
    pvAppendBuffer &H1174C23B, &HE857066A, &HFFFFC5E6, &H5750C303, &H44A8E8, &H5B5E5F00, &HC25DE58B, &H8B550008, &HA0EC81EC, &H8D000000, &HFFFF6085, &H875FFFF, &HFE61E850, &H75FFFFFF, &HE0458D0C, &HFA62E850, &H6AFFFF, &H50E0458D, &HFF60858D, &H8D50FFFF, &HE850A045, &HFFFFC93E, &H50A0458D, &HE81075FF, &HFFFFFD05, &H50A0458D, &HFFC8CDE8, &H1BD8F7FF, &HE58B40C0, &HCC25D, &H81EC8B55, &HF0EC&
    pvAppendBuffer &H10858D00, &HFFFFFFFF, &HE8500875, &HFFFFFEC7, &H8D0C75FF, &HE850D045, &HFFFFFAB8, &H458D006A, &H858D50D0, &HFFFFFF10, &H70858D50, &H50FFFFFF, &HFFCA8DE8, &H70858DFF, &H50FFFFFF, &HE81075FF, &HFFFFFD3C, &HFF70858D, &HE850FFFF, &HFFFFC898, &HC01BD8F7, &H5DE58B40, &H55000CC2, &HEC83EC8B, &HC0458D40, &H875FF56, &HFDA9E850, &H758BFFFF, &HC0458D0C, &H1468D50, &H500406C6, &HFFFC68E8
    pvAppendBuffer &HE0458DFF, &H21468D50, &HFC5BE850, &HC033FFFF, &HE58B5E40, &H8C25D, &H83EC8B55, &H458D60EC, &H75FF56A0, &H2CE85008, &H8BFFFFFE, &H458D0C75, &H468D50A0, &H406C601, &HFCC2E850, &H458DFFFF, &H468D50D0, &HB5E85031, &H33FFFFFC, &H8B5E40C0, &H8C25DE5, &HEC8B5500, &H80EC81, &H57530000, &H6A107D8B, &H57535B04, &H3511E8, &H74C08500, &HE9C03307, &H115&, &HE8575356, &HFFFFC472
    pvAppendBuffer &H90BE&, &H50C60300, &H34A2E8, &H1F88300, &HE8531074, &HFFFFC45A, &H5750C603, &H431BE857, &H6A0000, &HC448E857, &HC083FFFF, &H458D5050, &HDFE85080, &H53FFFFC7, &H5080458D, &HFFC431E8, &H50C603FF, &H3466E8, &H1F88300, &HE8531374, &HFFFFC41E, &H8D50C603, &H50508045, &H42DCE8, &H458D5300, &H93E85080, &H85000034, &H330774C0, &H96E9C0, &H758B0000, &H80458D14, &H66E85650
    pvAppendBuffer &HFFFFFFFB, &H458D0875, &H95E850C0, &HE8FFFFF8, &HFFFFC3DE, &H9005&, &H458D5000, &H458D50C0, &H458D5080, &H31E850E0, &HFF00003B, &H458D0C75, &H6DE850C0, &H53FFFFF8, &HFFC3B5E8, &H9005FF, &H8D500000, &H8D50E045, &H8D50C045, &HE850E045, &H3898&, &HC398E853, &H90BBFFFF, &H3000000, &H575750C3, &H38C2E8, &HC384E800, &HC303FFFF, &H50E04D8D, &HE8515157, &H3AE0&, &H51E04D8D
    pvAppendBuffer &H51204E8D, &HFFFAE0E8, &H40C033FF, &H8B5B5F5E, &H10C25DE5, &HEC8B5500, &HC0EC81, &H57530000, &H6A107D8B, &H57535B06, &H33D1E8, &H74C08500, &HE9C03307, &H129&, &HE8575356, &HFFFFC332, &H170BE, &H50C60300, &H3362E8, &H1F88300, &HE8531074, &HFFFFC31A, &H5750C603, &H41DBE857, &H6A0000, &HC308E857, &H1005FFFF, &H50000001, &HFF40858D, &HE850FFFF, &HFFFFC846, &H40858D53
    pvAppendBuffer &H50FFFFFF, &HFFC2E9E8, &H50C603FF, &H331EE8, &H1F88300, &HE8531674, &HFFFFC2D6, &H8D50C603, &HFFFF4085, &HE85050FF, &H4191&, &H40858D53, &H50FFFFFF, &H3345E8, &H74C08500, &HE9C03307, &H9C&, &H8D14758B, &HFFFF4085, &HE85650FF, &HFFFFFAAC, &H8D0875FF, &HE850A045, &HFFFFF7F4, &HFFC28DE8, &H17005FF, &H8D500000, &H8D50A045, &HFFFF4085, &H458D50FF, &H44E850D0, &HFF00003B
    pvAppendBuffer &H458D0C75, &HC9E850A0, &H53FFFFF7, &HFFC261E8, &H17005FF, &H8D500000, &H8D50D045, &H8D50A045, &HE850D045, &H3744&, &HC244E853, &H70BBFFFF, &H3000001, &H575750C3, &H376EE8, &HC230E800, &HC303FFFF, &H50D04D8D, &HE8515157, &H3AF3&, &H51D04D8D, &H51304E8D, &HFFFA23E8, &H40C033FF, &H8B5B5F5E, &H10C25DE5, &HEC8B5500, &H1B0EC81, &H858D0000, &HFFFFFE50, &H75FF5653, &H94E85008
    pvAppendBuffer &H8BFFFFFA, &H858D1075, &HFFFFFF30, &H91E85056, &H8DFFFFF6, &H8D502046, &HE8509045, &HFFFFF684, &H565E046A, &HFF30858D, &HE850FFFF, &H324E&, &H850FC085, &H374&, &H90458D56, &H323CE850, &HC0850000, &H362850F, &H8D560000, &HFFFF3085, &H9BE850FF, &HBBFFFFC1, &H90&, &HE850C303, &H31CB&, &HF01F883, &H33F85, &H458D5600, &H7BE85090, &H3FFFFC1, &HB0E850C3, &H83000031
    pvAppendBuffer &H850F01F8, &H324&, &H63E85657, &H3FFFFC1, &H458D50C3, &H458D5090, &H8CE850E0, &HFF000036, &H858D0C75, &HFFFFFF50, &HF5F6E850, &H3FE8FFFF, &H3FFFFC1, &H458D50C3, &H858D50E0, &HFFFFFF50, &H95E85050, &HE8000038, &HFFFFC126, &H8D50C303, &H8D50E045, &HFFFF3085, &H858D50FF, &HFFFFFF10, &H3876E850, &H8D560000, &HFFFE5085, &H858D50FF, &HFFFFFEB0, &H3DE5E850, &H8D560000, &HFFFE7085
    pvAppendBuffer &H858D50FF, &HFFFFFED0, &H3DD1E850, &HE8560000, &HFFFFC0DE, &H5050C083, &HFF70858D, &HE850FFFF, &H3DBB&, &HC0C8E856, &HC083FFFF, &H458D5070, &HA8E85090, &H5600003D, &HFFC0B5E8, &H10C083FF, &H70858D50, &H50FFFFFF, &HFEB0858D, &H8D50FFFF, &HE850E045, &H3B4B&, &HFED0858D, &H8D50FFFF, &HFFFEB085, &H458D50FF, &H858D5090, &HFFFFFF70, &HC789E850, &HE856FFFF, &HFFFFC076, &H5010C083
    pvAppendBuffer &H50E0458D, &H35A1E850, &H458D0000, &H858D50E0, &HFFFFFED0, &HB0858D50, &H50FFFFFE, &HFFCFD7E8, &HD06583FF, &HC048E800, &HC083FFFF, &HD4458950, &HFE50858D, &H4589FFFF, &HB0858DD8, &H89FFFFFE, &H858DDC45, &HFFFFFF10, &H59E85056, &H8B00003C, &H50858DD8, &H56FFFFFF, &H3C4AE850, &HC33B0000, &H8DD8470F, &HFFFF5085, &HFF738DFF, &H59E85056, &HB00003F, &H330574C2, &H2EB47FF, &H8D56FF33
    pvAppendBuffer &HFFFF1085, &H41E850FF, &HB00003F, &H6A0574C2, &H2EB5E02, &HF70BF633, &H6AB0458D, &HB5748B04, &HE85056D0, &H3CB7&, &H468D046A, &H858D5020, &HFFFFFEF0, &H3CA5E850, &H46A0000, &H50E0458D, &H2FBDE8, &HE4658300, &HFE738D00, &H1E045C7, &H85000000, &HDD880FF6, &H8D000000, &H8D50E045, &HFFFEF085, &H458D50FF, &HA7E850B0, &H56FFFFBF, &HFF50858D, &HE850FFFF, &H3ED0&, &H574C20B
    pvAppendBuffer &HEB47FF33, &H56FF3302, &HFF10858D, &HE850FFFF, &H3EB8&, &H574C20B, &HEB58026A, &HBC03302, &HBD7C8BF8, &HFFF85D0, &H8284&, &H57046A00, &HFF70858D, &HE850FFFF, &H3C23&, &H478D046A, &H458D5020, &H14E85090, &H8D00003C, &H8D50E045, &H8D509045, &HFFFF7085, &H99E850FF, &H6AFFFFCE, &HBF0CE804, &HC083FFFF, &H858D5010, &HFFFFFF70, &HB0458D50, &H90858D50, &H50FFFFFE, &H39A2E8
    pvAppendBuffer &HF0858D00, &H50FFFFFE, &H50B0458D, &H5090458D, &HFF70858D, &HE850FFFF, &HFFFFC5E3, &H858D046A, &HFFFFFE90, &HE0458D50, &H1DE85050, &H83000039, &H890F01EE, &HFFFFFF23, &HB3E8046A, &H83FFFFBE, &H8D5010C0, &H5050E045, &H33DEE8, &HE0458D00, &HF0858D50, &H50FFFFFE, &H50B0458D, &HFFCE17E8, &H8D046AFF, &HE850B045, &HFFFFBE86, &H90BE&, &H50C60300, &H2EB6E8, &HF8835F00, &H6A147401
    pvAppendBuffer &HBE6CE804, &HC603FFFF, &HB0458D50, &H2AE85050, &H6A00003D, &H30858D04, &H50FFFFFF, &H50B0458D, &H2E8AE8, &H1BD8F700, &H2EB40C0, &H5B5EC033, &HC25DE58B, &H8B55000C, &H80EC81EC, &H8D000002, &HFFFD8085, &H575653FF, &H500875FF, &HFFF786E8, &H10758BFF, &HFED0858D, &H5056FFFF, &HFFF373E8, &H30468DFF, &H60858D50, &H50FFFFFF, &HFFF363E8, &H5F066AFF, &HD0858D57, &H50FFFFFE, &H2E7DE8
    pvAppendBuffer &HFC08500, &H39285, &H858D5700, &HFFFFFF60, &H2E68E850, &HC0850000, &H37D850F, &H8D570000, &HFFFED085, &HC7E850FF, &HBBFFFFBD, &H170&, &HE850C303, &H2DF7&, &HF01F883, &H35A85, &H858D5700, &HFFFFFF60, &HBDA4E850, &HC303FFFF, &H2DD9E850, &HF8830000, &H3C850F01, &H57000003, &HFFBD8DE8, &H50C303FF, &HFF60858D, &H8D50FFFF, &HE850C045, &H32B3&, &H8D0C75FF, &HFFFF0085
    pvAppendBuffer &HCDE850FF, &HE8FFFFF2, &HFFFFBD66, &H8D50C303, &H8D50C045, &HFFFF0085, &HE85050FF, &H3623&, &HFFBD4DE8, &H50C303FF, &H50C0458D, &HFED0858D, &H8D50FFFF, &HFFFEA085, &H4E850FF, &H57000036, &HFD80858D, &H8D50FFFF, &HFFFE1085, &HCE850FF, &H5700003A, &HFDB0858D, &H8D50FFFF, &HFFFE4085, &HF8E850FF, &H57000039, &HFFBD05E8, &HA0738DFF, &H8D50C603, &HFFFF3085, &HE0E850FF, &H57000039
    pvAppendBuffer &HFFBCEDE8, &H14005FF, &H8D500000, &HFFFF6085, &HC8E850FF, &H57000039, &HFFBCD5E8, &HA05E8DFF, &H8D50C303, &HFFFF3085, &H858D50FF, &HFFFFFE10, &HC0458D50, &H3769E850, &H858D0000, &HFFFFFE40, &H10858D50, &H50FFFFFE, &HFF60858D, &H8D50FFFF, &HFFFF3085, &HB4E850FF, &H57FFFFC4, &HFFBC91E8, &H50C303FF, &H50C0458D, &H31BDE850, &H458D0000, &H858D50C0, &HFFFFFE40, &H10858D50, &H50FFFFFE
    pvAppendBuffer &HFFCC3EE8, &HF06583FF, &HBC64E800, &HC603FFFF, &H8DF44589, &HFFFD8085, &HF84589FF, &HFE10858D, &H4589FFFF, &HA0858DFC, &H57FFFFFE, &H3876E850, &HD88B0000, &HFF00858D, &H5057FFFF, &H3867E8, &HFC33B00, &H858DD847, &HFFFFFF00, &H56FF738D, &H3B76E850, &HC20B0000, &HFF330574, &H3302EB47, &H858D56FF, &HFFFFFEA0, &H3B5EE850, &HC20B0000, &H26A0574, &H3302EB5E, &H8DF70BF6, &H66A9045
    pvAppendBuffer &HF0B5748B, &HD4E85056, &H6A000038, &H30468D06, &H70858D50, &H50FFFFFE, &H38C2E8, &H8D066A00, &HE850C045, &H2BDA&, &HC46583, &HC7FE738D, &H1C045, &HF6850000, &HE8880F, &H458D0000, &H858D50C0, &HFFFFFE70, &H90458D50, &HBD4CE850, &H8D56FFFF, &HFFFF0085, &HEDE850FF, &HB00003A, &H330574C2, &H2EB47FF, &H8D56FF33, &HFFFEA085, &HD5E850FF, &HB00003A, &H6A0574C2, &H2EB5802
    pvAppendBuffer &HF80BC033, &HF0BD7C8B, &H840FFF85, &H8D&, &H8D57066A, &HFFFF3085, &H40E850FF, &H6A000038, &H30478D06, &H60858D50, &H50FFFFFF, &H382EE8, &HC0458D00, &H60858D50, &H50FFFFFF, &HFF30858D, &HE850FFFF, &HFFFFCAFB, &H23E8066A, &H5FFFFBB, &HB0&, &H30858D50, &H50FFFFFF, &H5090458D, &HFDE0858D, &HE850FFFF, &H35B7&, &HFE70858D, &H8D50FFFF, &H8D509045, &HFFFF6085, &H858D50FF
    pvAppendBuffer &HFFFFFF30, &HC305E850, &H66AFFFF, &HFDE0858D, &H8D50FFFF, &H5050C045, &H352FE8, &H1EE8300, &HFF18890F, &H66AFFFF, &HFFBAC5E8, &HB005FF, &H8D500000, &H5050C045, &H2FEEE8, &HC0458D00, &H70858D50, &H50FFFFFE, &H5090458D, &HFFCA72E8, &H8D066AFF, &HE8509045, &HFFFFBA96, &H170BE, &H50C60300, &H2AC6E8, &H1F88300, &H66A1474, &HFFBA7DE8, &H50C603FF, &H5090458D, &H393BE850
    pvAppendBuffer &H66A0000, &HFED0858D, &H8D50FFFF, &HE8509045, &H2A9B&, &HC01BD8F7, &H3302EB40, &H5B5E5FC0, &HC25DE58B, &H8B55000C, &H758B56EC, &H8D068B08, &H48504, &H56500000, &H2854E8, &HBA30E800, &HFF56FFFF, &H5D5E0850, &H550004C2, &H4D8BEC8B, &HC1C18B08, &HE18107E8, &HFF7F7F7F, &H1010125, &H6BC90301, &HC1331BC0, &H4C25D, &HE8EC8B55, &HFFFFBA11, &H78727B9, &HE98100, &H3000740
    pvAppendBuffer &H84D8BC1, &H75FF5051, &H30418D10, &H6A0C75FF, &H418D5010, &HEE85020, &H5DFFFFD8, &H55000CC2, &H4D8BEC8B, &H10458B08, &HC75FF50, &H51384101, &H3C5183, &HFFFFB3E8, &HCC25DFF, &HEC8B5500, &H8758B56, &H1487E83, &HE8560D75, &H20&, &H24846C7, &H8B000000, &H46011045, &H75FF5040, &H4456830C, &H81E85600, &H5EFFFFFF, &HCC25D, &H56EC8B55, &H8B08758B, &HC985304E, &H106A2474
    pvAppendBuffer &H50C12B58, &H320468D, &H50006AC1, &HAAEE8, &HCC48300, &H5020468D, &H9E856, &H66830000, &H5D5E0030, &H550004C2, &HEC83EC8B, &HF0458D10, &HFF505756, &H16E80C75, &H8BFFFFE6, &H458D087D, &H10778DF0, &HE8505656, &HFFFFE56B, &HFF565756, &H5E5F4C57, &HC25DE58B, &H8B550008, &H565151EC, &H8308758B, &H7401487E, &H487E8306, &H560A7502, &HFFFF7BE8, &H486683FF, &H384E8B00, &H50F8458D
    pvAppendBuffer &HF3C468B, &HC103C8A4, &H515003E1, &H3899E8, &H8D086A00, &H5650F845, &HFFFECFE8, &H404E8BFF, &H50F8458D, &HF44468B, &HC103C8A4, &H515003E1, &H3875E8, &H8D086A00, &H5650F845, &HFFFEABE8, &HC75FFFF, &H5010468D, &HFFE736E8, &HE58B5EFF, &H8C25D, &H83EC8B55, &H565310EC, &H87D8B57, &H6A506A, &H9D5E857, &HC4830000, &H75FF570C, &HE55BE80C, &HC033FFFF, &H5340C933, &HF484789
    pvAppendBuffer &H5BF38BA2, &H89F05D8D, &H4738903, &H89084B89, &H6EE80C53, &HF6FFFFB8, &HB902F845, &H76E0C, &H8EB90575, &H8100076D, &H74000E9, &H89C10300, &H5E5F4C47, &H5DE58B5B, &H550008C2, &H558BEC8B, &HF6335614, &H7801EA83, &HC458B30, &H85D8B53, &H8DD82B57, &H4529903C, &H3B0C8B10, &HCE03C033, &HF03C013, &H8300D083, &HF08B01EA, &H8910458B, &H7F8D380C, &H5FE279FC, &H5EC68B5B, &H10C25D
    pvAppendBuffer &H53EC8B55, &H10758B56, &H8399C68B, &H8D571FE2, &HFFC1023C, &H1FE68105, &H79800000, &HCE834E05, &H558B46E0, &H6ADE8B0C, &HCE2B5920, &HE8D3C28B, &HCE8BDBF7, &HE2D3DB1B, &HC933D823, &H8B105D89, &H758B085D, &HBB540110, &H85C91304, &H850475F6, &H8B2C74C9, &H3308BB44, &H13C103D2, &H89C603D2, &H8308BB44, &HC78300D2, &H74D28503, &HBB348D12, &H1601C033, &H1304768D, &H85D08BC0, &H5FF175D2
    pvAppendBuffer &HC25D5B5E, &H8B55000C, &H38EC83EC, &H5314558B, &H8B0C5D8B, &H335756C3, &H89C22BFF, &H880FEC45, &H216&, &H8BEC558B, &H8458BC8, &H8905E1C1, &H348BFC4D, &H75F685B8, &HE983470C, &HFC4D8920, &H1E7E9, &HF8E85600, &H8BFFFFCB, &H89CA8BD0, &HE6D3D055, &H177ED285, &H3B01478D, &H8B107DC3, &H206A0845, &H8BCA2B59, &HD304B854, &H8BF20BEA, &HC68B205D, &H8B1C65F7, &HC383FC45, &HD0452BE1
    pvAppendBuffer &HD803F28B, &H89F07589, &H1F79F45D, &HFE0FB83, &H19F8E, &H8BDBF700, &HD3DB33CB, &HF07589EE, &H85F45D89, &H89840FF6, &H8B000001, &H1F25C3, &H5798000, &HE0C88348, &HC4D8B40, &H8BD84589, &HE28399C3, &H8BC2031F, &HF8C11455, &H8BC82B05, &HCA2BD845, &H850FC085, &H89&, &H4AD84521, &H5589CA03, &HE44D89E8, &H66C0570F, &HC845130F, &H1F045C7, &H3B000000, &H118C0FCF, &H8B000001
    pvAppendBuffer &H5D8B0845, &H88048D10, &H89C84D8B, &H458BF845, &HE04589CC, &H479D285, &H3EBC033, &HF793048B, &HF7C103E6, &HE05513D0, &H5589C933, &HF8558BD0, &HC9130203, &H89F04503, &HD84D1302, &H83E8558B, &H4A04F86D, &HE06583, &HD86583, &H8BF04D89, &H8949E44D, &H4D89E855, &H8BCF3BE4, &HB47DD04D, &HA8E9&, &HE8658300, &HCA034A00, &H89D45589, &HCF3BD04D, &H968C0F, &H206A0000, &HC7D82B5B
    pvAppendBuffer &H1E445, &H458B0000, &HC0570F08, &H45130F66, &HCC5D89DC, &H8DD85D8B, &H4D8B8804, &HF84589E0, &H89DC458B, &HD285E045, &HC0330479, &H458B06EB, &H90048B10, &HF08BE6F7, &H13E07503, &H89CB8BD1, &HC033E055, &HE2D3D68B, &H8BE8550B, &HD2F7F84D, &HC0131103, &H83E45503, &H658300D0, &H458900D8, &H8BC18BE4, &H1089CC4D, &H758BD68B, &H4E883F0, &H4D8BEAD3, &HE85589D0, &H4AD4558B, &H49F84589
    pvAppendBuffer &H89D45589, &HCF3BD04D, &H7DD84D8B, &HF45D8B99, &H187D83, &H56530A74, &HE81875FF, &HFFFFFD60, &H8BEC558B, &H458BFC4D, &HC5D8B08, &H8E0FFA3B, &HFFFFFDFE, &H8B0C5D8B, &H7D8B1455, &H85C93308, &H8B247EDB, &HD32B1045, &H8D90348D, &HC0850A04, &HC0330479, &H68B02EB, &H728F0439, &H41087767, &H3B04C683, &H83E47CCB, &H8D000C65, &H45C7FF53, &H108&, &H78D28500, &H10458B39, &HF32BF28B
    pvAppendBuffer &H8D147503, &HF685B01C, &HC9330479, &HB8B02EB, &HD1F7C033, &H13970C03, &H84D03C0, &H13970C89, &H834E0C45, &H83000C65, &HEA8304EB, &H8458901, &H7D83D479, &HC740018, &H16A006A, &HE81875FF, &HFFFFFCBC, &H8B5B5E5F, &H1CC25DE5, &HEC8B5500, &H5328EC83, &H56145D8B, &H32FB8357, &H1B88E0F, &H758B0000, &H99C38B18, &HF88BC22B, &HD156C38B, &H75FF57FF, &H89C72B10, &H75FFF07D, &HFC45890C
    pvAppendBuffer &HE80875FF, &HFFFFFFC5, &H5610458B, &H56FC758B, &H8BF8148D, &H89520C45, &HC8DD855, &H8458BB8, &HDC4D8951, &H50B8048D, &HE8E44589, &HFFFFFF9D, &H8D184D8B, &HF6330146, &H148D046A, &H47A8D81, &H55893289, &H4518DF4, &H55893789, &H8B3289EC, &H7D89F055, &H5F3189E0, &H467ED285, &HC22BC003, &H8981048D, &H458BE845, &HC22B40FC, &H8B81048D, &H45890C4D, &H8458BF8, &H5589C12B, &HE8558B0C
    pvAppendBuffer &H8B084589, &H48BF85D, &HF87D0108, &H18B0389, &H5D8BCF03, &H3028914, &HC6D83D7, &H8458B01, &H75FFE175, &HEC458BFC, &HE475FF50, &HFB89E850, &H4D8BFFFF, &HFC75FF18, &H458B0189, &H75FF50E0, &H74E850DC, &H8BFFFFFB, &H558BF44D, &H18942FC, &H8D184D8B, &H4589D104, &HC1C28B0C, &HC10304E0, &H75FF5250, &HF475FF0C, &HFEDFE851, &H4D8BFFFF, &HF0558B18, &H8908418D, &H45890C71, &H8D3089DC
    pvAppendBuffer &H71891204, &H89318904, &HC0850845, &H458B1E7E, &H105D8BFC, &H8BC22B40, &HC8D0855, &HB3048BC1, &H3018946, &H7CF23BCF, &H145D8BF4, &H8DFC758B, &H8B503604, &HFF50DC45, &HE850D875, &HFFFFFAFB, &H89EC4D8B, &H1468D01, &H8D184D8B, &H458B0034, &H5150560C, &H203E850, &H4D8B0000, &H418D56FC, &H8BD82B01, &HDB031045, &H48DD92B, &H75FF5098, &HC4E8500C, &H8BFFFFFA, &HFF685F0, &H8784&
    pvAppendBuffer &H10558B00, &H8DFCC283, &HC0339A14, &HC0133201, &HF08BD72B, &HF275F685, &H558B6EEB, &H1B0C8D10, &H67EC985, &HFA8BC033, &H458BABF3, &H980C8D08, &H5476C83B, &H8D0C758B, &HC283DA14, &H89046AFC, &H348D1055, &H1475899E, &H1865835F, &H8BCF2B00, &HC753BDA, &H18B2176, &H26F7F72B, &H83184503, &H30100D2, &H2B00D283, &H185589DF, &H770C753B, &H8458BE5, &H8B10558B, &HD72B1875, &H758B3389
    pvAppendBuffer &H10558914, &HC177C83B, &H8B5B5E5F, &H14C25DE5, &HEC8B5500, &H5310EC83, &H56145D8B, &H32FB8357, &HA48E0F, &H7D8B0000, &H8BC38B0C, &H2B990875, &H8BD38BC2, &H18458BC8, &HD12BF9D1, &H89FC4D89, &H1C8DF855, &H8D5253D8, &H8D50C804, &H8D508F04, &HE8508E04, &HFFFFFD75, &H8BFC4D8B, &H51531845, &H5088048D, &H8DF8458B, &H56508704, &HFFFFA0E8, &H75FF53FF, &H185D8BFC, &H7D8B5753, &HBE048DF8
    pvAppendBuffer &HFF8BE850, &H758BFFFF, &H7EFF85FC, &H10458B1B, &H8DB00C8D, &H148D7704, &H8D028B83, &H1890452, &H8304498D, &HF17501EF, &H48D5356, &HE85350B3, &HFFFFF987, &H5614458B, &H8D1075FF, &H53508304, &HFFF976E8, &H8AE9FF, &HDB850000, &H7D8B097E, &H33CB8B10, &H8BABF3C0, &H7D8B0845, &H980C8D10, &H899F348D, &HC83B1475, &H458B6A76, &H98048D0C, &H83085D8B, &H4589FCC0, &H4E983FC, &H890C7589
    pvAppendBuffer &H570FF845, &H130F66C0, &HF73BF045, &H458B3C76, &H8BDE8BF4, &H4589F875, &HF0458B18, &H8B104589, &H4EB8306, &H768D21F7, &H830303FC, &H450300D2, &H13038910, &H65831855, &H55890018, &H77DF3B10, &H14758BDE, &H8B085D8B, &HEE83FC45, &H14758904, &HA577CB3B, &H8B5B5E5F, &H14C25DE5, &HEC8B5500, &H14558B51, &HFC6583, &HC701EA83, &H11445, &H39780000, &H5308458B, &HC758B56, &H107D8B57
    pvAppendBuffer &H2B901C8D, &H8BF82BF0, &HC0331E0C, &HB03D1F7, &H4D03C013, &H1F0C8914, &H13FC5B8D, &H6583FC45, &HEA8300FC, &H14458901, &H5E5FDD79, &H5DE58B5B, &H550010C2, &HEC81EC8B, &H80&, &HC758B56, &H59206A57, &HF3807D8D, &HFDBEA5, &H458D0000, &HE8505080, &H1ED7&, &H7402FE83, &H4FE8312, &H75FF0D74, &H80458D0C, &H83E85050, &H8300000E, &HDA7901EE, &H8D087D8B, &H206A8075, &H5FA5F359
    pvAppendBuffer &H5DE58B5E, &H550008C2, &H5653EC8B, &H875FF57, &HFFF636E8, &H53D88BFF, &HFFF62EE8, &H52D08BFF, &HFFF626E8, &H33F88BFF, &HF78B087D, &HC333C78B, &H3308CFC1, &H8C0C1F2, &HC9C1CE8B, &H33C13310, &H5FC633C7, &H4533C333, &H5D5B5E08, &H550004C2, &H8B56EC8B, &H36FF0875, &HFFFFAAE8, &H476FFFF, &HA0E80689, &HFFFFFFFF, &H46890876, &HFF95E804, &H76FFFFFF, &H846890C, &HFFFF8AE8, &HC4689FF
    pvAppendBuffer &H4C25D5E, &HEC8B5500, &H85D8B53, &HB60F5756, &HB60F077B, &HB60F0243, &HB60F0B73, &HE7C10F53, &HFF80B08, &HF034BB6, &HC10D43B6, &HF80B08E7, &HF08E6C1, &HC10843B6, &HF80B08E7, &HF08E2C1, &HB0643B6, &H8E1C1F0, &H143B60F, &HB08E6C1, &H43B60FF0, &H8E6C10C, &HB60FF00B, &HD00B0A43, &H543B60F, &HB08E2C1, &H3B60FD0, &HB08E2C1, &H43B60FD0, &H89C80B0E, &HB60F0C53, &HE1C10943
    pvAppendBuffer &H89C80B08, &HB60F0873, &H7B890443, &H8E1C104, &H5EC80B5F, &H5D5B0B89, &H550004C2, &H5756EC8B, &HFFAF4DE8, &H8758BFF, &H693BF, &H50C70300, &H22E836FF, &H8900001E, &HAF34E806, &HC703FFFF, &H476FF50, &H1E10E8, &H4468900, &HFFAF21E8, &H50C703FF, &HE80876FF, &H1DFD&, &HE8084689, &HFFFFAF0E, &HFF50C703, &HEAE80C76, &H5F00001D, &H5E0C4689, &H4C25D, &H83EC8B55, &H74000C7D
    pvAppendBuffer &H75FF5615, &H8758B0C, &HE856006A, &H2F&, &HC483068A, &HC25D5E0C, &H8B550008, &H10558BEC, &H5608458B, &HD285F08B, &H8B571274, &HF82B0C7D, &H88370C8A, &HEA83460E, &H5FF57501, &H55C35D5E, &H4D8BEC8B, &H74C98510, &H45B60F1F, &HF18B560C, &H101C069, &H8B570101, &HE9C1087D, &H8BABF302, &H3E183CE, &H5E5FAAF3, &H5D08458B, &HEC8B55C3, &H8758B56, &HF451E856, &HD08BFFFF, &HD633CE8B
    pvAppendBuffer &HC110C9C1, &HCEC108C2, &H33D13308, &H5EC233D6, &H4C25D, &H56EC8B55, &HFF08758B, &HFFCBE836, &H76FFFFFF, &HE8068904, &HFFFFFFC1, &H890876FF, &HB6E80446, &HFFFFFFFF, &H46890C76, &HFFABE808, &H4689FFFF, &HC25D5E0C, &H8B550004, &H40EC83EC, &HC46583, &H4521C033, &H575653E4, &H3359066A, &HC87D8DDB, &H8943066A, &HABF3C05D, &H8D046A59, &H5D89E87D, &H5FABF3E0, &HC0458D57, &HADF0E850
    pvAppendBuffer &HC083FFFF, &H458D5010, &H5CE850C0, &H5700001D, &H50C0458D, &H2A0FE8, &HFF708D00, &H458D26EB, &H52E850E0, &H56000028, &H50C0458D, &H2D1BE8, &H74C20B00, &H75FF570E, &HE0458D08, &H9E85050, &H4E000028, &H57E0458D, &H77F33B50, &H875FFD1, &H2A8AE8, &H5B5E5F00, &HC25DE58B, &H8B550004, &H60EC83EC, &HA46583, &H4521C033, &H575653D4, &H33590A6A, &HA87D8DDB, &H89430A6A, &HABF3A05D
    pvAppendBuffer &H8D066A59, &H5D89D87D, &H5FABF3D0, &HA0458D57, &HAD5CE850, &HB005FFFF, &H50000000, &H50A0458D, &H1CC6E8, &H458D5700, &H79E850A0, &H8D000029, &H26EBFF70, &H50D0458D, &H27BCE8, &H458D5600, &H85E850A0, &HB00002C, &H570E74C2, &H8D0875FF, &H5050D045, &H2773E8, &H458D4E00, &H3B5057D0, &HFFD177F3, &HF4E80875, &H5F000029, &HE58B5B5E, &H4C25D, &H83EC8B55, &H565310EC, &HC75FF57
    pvAppendBuffer &HFFE137E8, &H875FFFF, &HE8FC4589, &HFFFFE12C, &HDBE8F08B, &H5FFFFAC, &H794&, &HE11AE850, &H4589FFFF, &HACC8E808, &H9805FFFF, &H50000007, &HFFE107E8, &H8BFF33FF, &H7D8947D8, &HACB0E8F0, &HA6E9FFFF, &HE8000000, &HFFFFACA6, &H79405, &HE8565000, &HFFFFC01F, &H840FC085, &H10B&, &H23E836FF, &H8900000C, &H458BF845, &HE830FFFC, &HC16&, &HF4458950, &H50F8458B, &HFC75FF56
    pvAppendBuffer &HFFBC89E8, &HF84D8BFF, &HF763939, &H3C83018B, &H7750081, &H3B018948, &H8BF177C7, &H3939F44D, &H18B0F76, &H813C83, &H89480775, &H77C73B01, &HFC75FFF1, &HFFF1EDE8, &HF8458BFF, &H8BFC7589, &H53560875, &H89F475FF, &H5D89F845, &HBDADE808, &H5DF7FFFF, &H56D88BF0, &HFFF1C9E8, &HF475FFFF, &HFFF1C1E8, &HAC08E8FF, &H758BFFFF, &H79805F8, &H56500000, &HFFBF7EE8, &HFC085FF, &HFFFF4685
    pvAppendBuffer &H9FE856FF, &HFFFFFFF1, &H97E8FC75, &HFFFFFFF1, &H8FE80875, &H83FFFFF1, &HF00F07D, &HB58D&, &HC458B00, &H5FE830FF, &H8B00000B, &HFC7D89F0, &H7589C933, &HF3E39F8, &H8F8C&, &HC458B00, &H8904568D, &HC78BF045, &H89F05D29, &H7529085D, &HC758B08, &H758B063B, &H8B307F08, &HC603F045, &H8910048B, &H458BF445, &H5624EBFC, &HFFF135E8, &HFC75FFFF, &HFFF12DE8, &H875FFFF, &HFFF125E8
    pvAppendBuffer &H1FE853FF, &H33FFFFF1, &H834DEBC0, &H3B00F465, &H8B057F03, &H2EB1604, &H758BC033, &HF7F02BF4, &H89F12BD0, &H74C98532, &H1BF03B07, &H6EB41C9, &HC91BC63B, &H458BD9F7, &H8BF685FC, &H450FF875, &HC28340F8, &HFC458904, &H857E063B, &HE83E8953, &HFFFFF0D2, &HC38BDE8B, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H5320EC83, &H10758B56, &H8B1E8B57, &H2E7C1FB, &HE8E47D89, &HFFFFAAF6, &H8B10FF57
    pvAppendBuffer &H89C933D0, &HDB851055, &H68B0F7E, &H48BC12B, &H8A048986, &H7CCB3B41, &H87D8BF1, &H8B0C558B, &H3B078BCA, &HCF470F02, &H48D318B, &H7FC33B36, &H99C38B0A, &HF08BC22B, &H8B46FED1, &H2E0C1C6, &HE8E04589, &HFFFFAAAA, &HE1C1CE8B, &H10FF5102, &H4589D68B, &H85172BFC, &H8B0C7ED2, &HCA8BFC7D, &HABF3C033, &H8B087D8B, &H85C93307, &H8B1A7EC0, &H148DFC7D, &H87D8B97, &H8B41C12B, &H2898704
    pvAppendBuffer &H8B04528D, &H7CC83B07, &HAA64E8EF, &HCE8BFFFF, &H5102E1C1, &H558B10FF, &H8BCE8B0C, &HF87D89F8, &H4D890A2B, &H7EC98508, &HF3C03307, &H84D8BAB, &HFF33028B, &H1A7EC085, &H8DF8558B, &H558B8A0C, &H47C72B0C, &H8982048B, &H4498D01, &HF83B028B, &HFE8BEF7C, &H8903E7C1, &H13E8E87D, &H57FFFFAA, &H895610FF, &H7BE80845, &H8B000009, &H2E7C1F8, &HE8EC7D89, &HFFFFA9FA, &H8B10FF57, &H4589107D
    pvAppendBuffer &H573F8B0C, &HFFBEBAE8, &HF44589FF, &H1874C085, &HE7D3C88B, &H7E01FB83, &H59206A0F, &H458BC82B, &H4408B10, &HF80BE8D3, &H560C75FF, &HFF0875FF, &H75FFF875, &HF513E8FC, &H75FFFFFF, &H57F603F4, &HE8F07589, &H100F&, &H53006A50, &H8B1075FF, &H75FF57FE, &HF22CE808, &HDF3BFFFF, &H4C0FC78B, &H458950C3, &H920E8F0, &H5D8B0000, &H8BD233F0, &H7EDB85F0, &H8458B19, &H3C8DFB2B, &H8B0E8BB8
    pvAppendBuffer &H42CA2B07, &H89047F8D, &HD33B8E04, &H3E83EF7C, &H8B107601, &H863C8306, &H48087500, &HF8830689, &HFFF07701, &H5D8BEC75, &H62E8530C, &HE8000017, &HFFFFA93E, &H850FF53, &H8BE875FF, &HE853085D, &H174D&, &HFFA929E8, &H50FF53FF, &HE475FF08, &H53105D8B, &H1738E8, &HA914E800, &HFF53FFFF, &H7D8B0850, &HFC5D8BE0, &H22E85357, &HE8000017, &HFFFFA8FE, &H850FF53, &H57F85D8B, &H170FE853
    pvAppendBuffer &HEBE80000, &H53FFFFA8, &H5F0850FF, &H5B5EC68B, &HC25DE58B, &H8B55000C, &H2CEC83EC, &H105D8B53, &HF604438D, &H45890100, &H531175E0, &HFF0C75FF, &H6DE80875, &HE9000003, &H361&, &HFF535756, &HFEE80875, &H8BFFFFB9, &H8BF08B1B, &HDC5D89CB, &H5105E1C1, &HFFBD96E8, &H57F88BFF, &HE81075FF, &HFFFFFB90, &H891075FF, &H56570845, &HFFFD70E8, &H458956FF, &HEE2CE8F0, &H75FFFFFF, &HC2E85710
    pvAppendBuffer &H57FFFFB9, &HE8FC4589, &HFFFFEE1A, &HE7C1FB8B, &HD87D8902, &HFFA859E8, &H10FF57FF, &H85EC4589, &H8B197EDB, &H578DE04D, &H8BD003FC, &H8D018BF3, &H2890449, &H83FC528D, &HF17501EE, &HFFA831E8, &H10FF57FF, &HC933F08B, &H85F47589, &H8B2D7EDB, &HC7830855, &H4C283FC, &H7D8BF703, &H7D0F3B08, &HEB028B04, &H89C03302, &HC2834106, &H4EE8304, &HE97CCB3B, &H8BF4758B, &H2E7C1FB, &H5008458B
    pvAppendBuffer &HFFEDA1E8, &HA7E8E8FF, &HFF57FFFF, &H89D08B10, &HDB85E855, &HCF8B107E, &HE9C1C033, &HF3FA8B02, &HC1FB8BAB, &H565302E7, &HAEE85256, &H33FFFFF6, &H7EDB85C0, &HE8558B2D, &H83F04D8B, &HC183FCC2, &H8BD70304, &H73BF07D, &H318B047D, &HF63302EB, &H83403289, &HEA8304C1, &H7CC33B04, &HC1FB8BE9, &H4D8B02E7, &H3BE851F0, &H8BFFFFED, &H3E6C1F3, &HE8E07589, &HFFFFA77A, &H8B10FF56, &H87589F0
    pvAppendBuffer &HFFA76DE8, &HC1CB8BFF, &HFF5103E1, &H89C93310, &HDB85F845, &H7D8B2C7E, &HC1C38BFC, &HC08303E0, &H8DF003FC, &HF3B0457, &H28B047D, &HC03302EB, &H83410689, &HEE8304C2, &H7CCB3B04, &HC1FB8BE9, &H75FF02E7, &HECDCE8FC, &HE853FFFF, &H696&, &H303F36B, &H2E6C1F0, &HE8D47589, &HFFFFA712, &H8B10FF56, &HD2330C4D, &H6AFC4589, &HF055891F, &H4589018B, &HC0855EE4, &HD88B317E, &H33993C8D
    pvAppendBuffer &H40CE8BC0, &H785E0D3, &HEE831075, &H6A077901, &HEF83421F, &HD33B5E04, &H5D8BE57C, &H8BFB8BDC, &H4D8BE445, &HF055890C, &H3B02E7C1, &H978D0FD0, &H85000000, &H80880FF6, &H8B000000, &H75FF084D, &HF048DFC, &HF875FF53, &HFEE85050, &H53FFFFF1, &HFFFC75FF, &H75FFF475, &HF875FFEC, &H3E6E8, &HC458B00, &H108BCE8B, &H552BC033, &HE0D340F0, &H850C4D8B, &H2C749104, &H8BFC75FF, &HFF53F845
    pvAppendBuffer &HC7030875, &H50E875FF, &HFFF1C0E8, &H75FF53FF, &HF475FFFC, &HFFEC75FF, &HA8E80875, &H8B000003, &HCEB084D, &H8B08458B, &H4D89F84D, &HF8458908, &H7901EE83, &HF0558B89, &H420C4D8B, &H55891F6A, &H113B5EF0, &HFF698C0F, &HFF53FFFF, &H75FFFC75, &HEC75FFF4, &HE80875FF, &H36B&, &HFF10458B, &H5A4E830, &HD2330000, &HDB85F08B, &H458B1E7E, &H89C70308, &HF88B0C45, &H78B0E8B, &H8D42CA2B
    pvAppendBuffer &H489047F, &H7CD33B8E, &HD87D8BEF, &H76013E83, &H83068B10, &H7500863C, &H6894808, &H7701F883, &HD475FFF0, &H53FC5D8B, &H13E4E8, &HA5C0E800, &HFF53FFFF, &H75FF0850, &H85D8BE0, &H13CFE853, &HABE80000, &H53FFFFA5, &HFF0850FF, &H5D8BE075, &HBAE853F8, &HE8000013, &HFFFFA596, &H850FF53, &H57F45D8B, &H13A7E853, &H83E80000, &H53FFFFA5, &H8B0850FF, &H5357EC5D, &H1394E8, &HA570E800
    pvAppendBuffer &HFF53FFFF, &H5D8B0850, &HE85357E8, &H1381&, &HFFA55DE8, &H50FF53FF, &HC68B5F08, &HE58B5B5E, &HCC25D, &H83EC8B55, &H56532CEC, &H5710758B, &H875FF56, &HFFB68CE8, &H8B1E8BFF, &H2E7C1FB, &H89E04589, &H7D89E45D, &HA524E8D4, &HFF57FFFF, &H33D08B10, &HF85589C9, &HF7EDB85, &HC12B068B, &H8986048B, &H3B418A04, &HE8F17CCB, &HFFFFA502, &H8B10FF57, &HD38BE075, &H7D89F88B, &H85162BEC
    pvAppendBuffer &H8B097ED2, &HF3C033CA, &HEC7D8BAB, &HC933068B, &H147EC085, &H2B97148D, &H48B41C1, &H8D028986, &H68B0452, &HEF7CC83B, &HE7C1FB8B, &HDC7D8903, &HFFA4BDE8, &H10FF57FF, &H7589F08B, &HA4B0E8F4, &HFF57FFFF, &H8458910, &H851B048D, &H8B0D7EC0, &H33FE8BC8, &H8BABF3C0, &H3E7C1FB, &H3744C753, &H1FC&, &H3FCE800, &HF08B0000, &H8902E6C1, &H7BE8D875, &H56FFFFA4, &H4D8B10FF, &H89F6330C
    pvAppendBuffer &H1F6AF045, &H8BFC7589, &HC0855F01, &H148D267E, &H33D98B81, &H40CF8BC0, &H285E0D3, &HEF831075, &H6A077901, &HEA83461F, &H333B5F04, &H5D8BE57C, &HFC7589E4, &H8BF8758B, &H4E85636, &H89FFFFB9, &HC085E845, &HC88B1874, &HFB83E6D3, &H6A0F7E01, &HC82B5920, &H8BF8458B, &HE8D30440, &HE856F00B, &HA73&, &H890C4D8B, &H458BE445, &H94E9FC, &HFF850000, &H85880F, &H758B0000, &HF075FFF4
    pvAppendBuffer &H539E048D, &H500875FF, &HEF3FE850, &H75FFFFFF, &H1B048DE8, &H6AE475FF, &H75FF5300, &H75FF50F8, &HEC60E808, &H458BFFFF, &H8BCF8B0C, &H2BC03310, &HD340FC55, &HC4D8BE0, &H74910485, &HF075FF2F, &H5308458B, &HEC75FF56, &H5098048D, &HFFEEFCE8, &HE875FFFF, &HFF1B048D, &H6AE475, &HF875FF53, &H1FE85650, &H8BFFFFEC, &H8EB0C4D, &H758BC68B, &H8458908, &H7901EF83, &HFC458B84, &H40F47589
    pvAppendBuffer &H45891F6A, &H13B5FFC, &HFF648C0F, &H458BFFFF, &HE830FF10, &H2F2&, &HF08BD233, &H177EDB85, &H8DF4458B, &HE8B983C, &HCA2B078B, &H47F8D42, &H3B8E0489, &H83EF7CD3, &H1076013E, &H3C83068B, &H8750086, &H83068948, &HF07701F8, &H8BDC7D8B, &H5357F45D, &H1138E8, &HA314E800, &HFF53FFFF, &H75FF0850, &HF05D8BD8, &H1123E853, &HFFE80000, &H53FFFFA2, &H8B0850FF, &H5357085D, &H1110E8
    pvAppendBuffer &HA2ECE800, &HFF53FFFF, &H5D8B0850, &HF87D8BD4, &HFAE85753, &HE8000010, &HFFFFA2D6, &H850FF57, &HEC5D8B53, &H10E7E853, &HC3E80000, &H53FFFFA2, &H8B0850FF, &HE850E045, &HFFFFE86A, &H5EC68B5F, &H5DE58B5B, &H55000CC2, &H458BEC8B, &H5D8B5308, &H8B575614, &HF76B187D, &HB8048D0C, &H3144589, &H535756F3, &H501075FF, &HFFF024E8, &H758B56FF, &HBB3C8D18, &H75FF5756, &HCEE8530C, &H8BFFFFED
    pvAppendBuffer &H48D085D, &H57535036, &HEA2DE853, &H558BFFFF, &H89C93314, &HF98B1045, &H167EF685, &H89BB048B, &H4528D02, &H47BB0C89, &HF07CFE3B, &H8B14558B, &HC0851045, &HF6852A75, &H5D8B267E, &H8BFA8B0C, &H8B043B07, &H83410875, &HCE3B04C7, &HCE3BF17C, &H7D8B0E7D, &H31048D08, &H3B87048B, &HB768B04, &H75FF5256, &HF2E8520C, &H5FFFFFF0, &HC25D5B5E, &H8B550014, &H8EC81EC, &H8B000001, &H570F0C45
    pvAppendBuffer &H6A5756C0, &HB58D593C, &HFFFFFEF8, &H85130F66, &HFFFFFEF8, &HFF00BD8D, &H45C7FFFF, &H10F8&, &H8DA5F300, &HFFFEF8B5, &H89CE8BFF, &HC12BFC75, &H530C4589, &H8B300C8B, &H30448BFE, &H8BDB3304, &H4D891075, &HF04589F4, &H74FF5150, &H34FF04DE, &HAF04E8DE, &H701FFFF, &H11F44D8B, &H8B430457, &H7F8DF045, &H10FB8308, &H758BDE72, &HC458BFC, &H8308C683, &H8901F86D, &HB875FC75, &H6A5BF633
    pvAppendBuffer &HFF266A00, &HFF7CF5B4, &HB4FFFFFF, &HFFFF78F5, &HAEC4E8FF, &H8401FFFF, &HFFFEF8F5, &HF59411FF, &HFFFFFEFC, &HFFE8346, &H7D8BD572, &HF8B58D08, &H6AFFFFFE, &H75FF5920, &HE8A5F308, &HFFFFB65A, &HE80875FF, &HFFFFB652, &HE58B5E5F, &HCC25D, &H83EC8B55, &H565310EC, &H53DB3357, &H530C75FF, &HE81475FF, &HFFFFAE72, &HC75FF53, &H8BF04589, &H75FF53F2, &HAE60E818, &HFF53FFFF, &H45891075
    pvAppendBuffer &H53FA8BF4, &HE81875FF, &HFFFFAE4E, &H1075FF53, &H53FC4589, &H891475FF, &H3BE8F855, &H8BFFFFAE, &HF4458BD8, &H6ADE03, &H3D6135E, &H3BD713D8, &H720D77D7, &H73D83B04, &HFC750107, &H1F85583, &H3308458B, &HF04D0BC9, &H5E5FDE0B, &HC9330889, &H89FC5503, &H4D130458, &H85089F8, &H5B0C4889, &HC25DE58B, &H8B550014, &H84D8BEC, &HEBF63356, &H99C18B0D, &HF8D1C22B, &H8D41C82B, &HF9838E34
    pvAppendBuffer &H8BEE7F32, &HC25D5EC6, &H8B550004, &H5D8B53EC, &H8D575608, &H49D34, &H53E80000, &H56FFFFA0, &H8B5610FF, &H57006AF8, &HFFF192E8, &HCC483FF, &HC78B1F89, &H5D5B5E5F, &H550004C2, &HEC83EC8B, &H5D8B5330, &H6A575608, &HC75FF06, &H1D0DE853, &H66A0000, &H75FF206A, &HD0458D0C, &H10C3E850, &H66A0000, &H4B8DF08B, &HD0458D08, &H5150FA8B, &H38C38351, &HF72E8, &H6AC60300, &HC75FF06
    pvAppendBuffer &HD7130389, &H8308458B, &H538910C0, &HE8505004, &HF57&, &H6A084D8B, &H40418906, &H50D0458D, &H51895151, &H1E93E844, &H4D8B0000, &H13F00308, &H30518BFA, &H718BD62B, &H3BF71B34, &H1D723471, &H513B0577, &H83167630, &H3EBFFCF, &H1085B8D, &H47B113B, &H4323038B, &H74C73B04, &H71895FEF, &H51895E34, &HE58B5B30, &H8C25D, &H81EC8B55, &H108EC, &H57565300, &H8D0C75FF, &HFFFF7885
    pvAppendBuffer &H68E850FF, &H8D000007, &HFFFF7885, &H87E850FF, &H8DFFFFB4, &HFFFF7885, &H7BE850FF, &H8DFFFFB4, &HFFFF7885, &H6FE850FF, &H8DFFFFB4, &HFFFEF89D, &HC45C7FF, &H2&, &H8D8BF633, &HFFFFFF78, &HFF7C858B, &HE981FFFF, &HFFED&, &HC61B086A, &HFEF88D89, &H8589FFFF, &HFFFFFEFC, &H3B548B5F, &H3B448BF8, &H3D8C8BFC, &HFFFFFF78, &HFF85589, &H8B10C2AC, &HFF7C3D84, &HE283FFFF, &H3B748901
    pvAppendBuffer &H1BCA2BFC, &HFFE981C6, &H890000FF, &HFEF83D8C, &HC61BFFFF, &HFC3D8489, &HFFFFFFE, &H89F845B7, &H83F83B44, &HFF8308C7, &H8BB27278, &HFFFF688D, &H6C858BFF, &H8BFFFFFF, &HAC0FF055, &HB70F10C1, &HFFFF6885, &H1E183FF, &HFF688589, &HD12BFFFF, &HFF6CB589, &H4D8BFFFF, &H81CE1BF4, &H7FFFEA, &H70958900, &H1BFFFFFF, &H89C033CE, &HFFFF748D, &HAC0F40FF, &HE28310CA, &H10F9C101, &H8D50C22B
    pvAppendBuffer &HFFFEF885, &H858D50FF, &HFFFFFF78, &H601E850, &H6D830000, &H850F010C, &HFFFFFF1E, &H8A08558B, &HFF78F584, &H8C8BFFFF, &HFFFF78F5, &H720488FF, &H7CF5848B, &HFFFFFFF, &HC108C1AC, &H4C8808F8, &H83460172, &HD77210FE, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H33084D8B, &H758B56D2, &H116A570C, &H8B5FF12B, &H1030E04, &HB60FD003, &H8EAC1C2, &H498D0189, &H1EF8304, &H5E5FE975, &H8C25D
    pvAppendBuffer &H56EC8B55, &H8B0C75FF, &HE8560875, &HFFFFFFC1, &H5044468D, &H115E856, &H5D5E0000, &H550008C2, &HEC83EC8B, &H5D8B5344, &H6A575608, &HF38B5911, &HF3BC7D8D, &H9DB8E8A5, &H4405FFFF, &H50000005, &H50BC458D, &HFFFF88E8, &HFC458BFF, &H8025D0F7, &H6A000000, &H508D5E11, &HC1D0F7FF, &HE8C11FEA, &H8DD0231F, &HD2F7BC45, &H4589C32B, &H17A8D08, &HC8BD7F7, &H1428D18, &HCF230323, &H458BC80B
    pvAppendBuffer &H8D0B8908, &HEE83045B, &H5FE77501, &HE58B5B5E, &H4C25D, &H83EC8B55, &H458D44EC, &H446A56BC, &H5056F633, &HFFEE96E8, &H84D8BFF, &H8B0CC483, &HA891&, &H74D28500, &H84B60F11, &H9831&, &HB5448900, &HF23B46BC, &H458DEF72, &H9544C7BC, &H1BC&, &HE8515000, &HFFFFFF24, &H5DE58B5E, &H550004C2, &H4D8BEC8B, &HF6335608, &H8BD68B57, &HB91403FE, &HC1C2B60F, &H48908EA, &HFF8347B9
    pvAppendBuffer &H3EE7C10, &HC28B4051, &H8302EAC1, &HD26B03E0, &H40418905, &HFB11403, &HEAC1C2B6, &HB1048908, &H10FE8346, &H5101EE7C, &H5D5E5F40, &H550004C2, &HEC83EC8B, &H5D8B534C, &HB4458D0C, &HD82BD233, &H5D895756, &H33FF33FC, &H78D285C9, &H93048D1E, &H8D085D8B, &HF003B475, &HF8B048B, &H768D06AF, &H41F803FC, &HF07ECA3B, &H8DFC5D8B, &H75890172, &H83CE8BF8, &H2D7D11FE, &H8B085D8B, &H8BC62BC2
    pvAppendBuffer &HC6830C75, &H86348D44, &HF8B048B, &H768D06AF, &H40C069FC, &H3000001, &HF98341F8, &H8BE97C11, &H5D8BF875, &H957C89FC, &H83D68BB4, &H957C11FA, &H50B4458D, &HFFFF2AE8, &H87D8BFF, &H6AB4758D, &HA5F35911, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H830C558B, &HC93344EC, &H1104B60F, &HBC8D4489, &H10F98341, &H458DF27C, &HFC45C7BC, &H1&, &H875FF50, &HFFFE07E8, &H5DE58BFF, &H550008C2
    pvAppendBuffer &HEC81EC8B, &H178&, &H33575653, &HB17D8DC0, &H5D88DB33, &HC6AABB0, &HAB0C75FF, &HAAAB66AB, &H50B4458D, &HFFECF1E8, &HCC483FF, &H33D05D88, &HD17D8DC0, &HF359076A, &H66046AAB, &H458DAAAB, &H206A50B0, &H8D0875FF, &HFFFF3485, &H1FE850FF, &H6AFFFFC0, &HD0458D20, &H858D5050, &HFFFFFF34, &HBAEFE850, &H458DFFFF, &H458D50E0, &H858D50D0, &HFFFFFE88, &HCAB1E850, &H86AFFFF, &H8DC03359
    pvAppendBuffer &HABF3D07D, &H6AD0458A, &HD0458D20, &H858D5050, &HFFFFFF34, &HBABBE850, &HC033FFFF, &H6AD07D8D, &H75FF5908, &H8AABF314, &H7D8DD045, &H1075FFF1, &H5D88C033, &HABABABF0, &H8DAAAB66, &HFFFE8885, &H24E850FF, &H8BFFFFCB, &HD8F71445, &H500FE083, &H50F0458D, &HFE88858D, &HE850FFFF, &HFFFFCB0B, &H331C758B, &H187D8BC0, &HC458940, &H24453956, &H75FF1675, &H34858D20, &H57FFFFFF, &HBA53E850
    pvAppendBuffer &HFF56FFFF, &H1EB2075, &H88858D57, &H50FFFFFE, &HFFCAD6E8, &HF7C68BFF, &HFE083D8, &HF0458D50, &H88858D50, &H50FFFFFE, &HFFCABEE8, &HF0458DFF, &H75FF5350, &H1AC9E814, &H458D0000, &H565350F8, &H1ABEE8, &H5B106A00, &HF0458D53, &H88858D50, &H50FFFFFE, &HFFCA92E8, &H247D83FF, &HFF2C7501, &H858D2875, &HFFFFFE88, &HC952E850, &H7C6AFFFF, &HFF34858D, &H6AFFFF, &HEBC1E850, &H858AFFFF
    pvAppendBuffer &HFFFFFF34, &H330CC483, &H8D73EBC0, &H8D50C045, &HFFFE8885, &H25E850FF, &H8BFFFFC9, &H4D8D287D, &H32C18BC0, &H8AF82BD2, &H1320F04, &H8341D00A, &HF37501EB, &H56187D8B, &H842075FF, &H571275D2, &HFF34858D, &HE850FFFF, &HFFFFB989, &HEB0C5D21, &HEB22E805, &H7C6AFFFF, &HFF34858D, &H6AFFFF, &HEB59E850, &H8D8AFFFF, &HFFFFFF34, &H83C07D8D, &HC0330CC4, &HABABABAB, &H8BC04D8A, &H5E5F0C45
    pvAppendBuffer &H5DE58B5B, &H550024C2, &H558BEC8B, &H2B60F08, &H14AB60F, &HB08E0C1, &H4AB60FC1, &H8E0C102, &HB60FC10B, &HE0C1034A, &H5DC10B08, &H550004C2, &H558BEC8B, &H42B60F08, &H4AB60F03, &H8E0C102, &HB60FC10B, &HE0C1014A, &HFC10B08, &HE0C10AB6, &H5DC10B08, &H550004C2, &H8B53EC8B, &HD233085D, &HE9C1CB8B, &H57564110, &H818D006A, &H7FFFFFFF, &HF08BF1F7, &H8B10E6C1, &H8BE3F7C6, &HF7CE03C8
    pvAppendBuffer &HD283D1, &HC183C033, &H13D2F701, &HF7C203C0, &H1FE8C1E6, &HFF03FA8B, &HC78BF80B, &HF08BE3F7, &HF703CA8B, &HBA&, &HC8135880, &H1872CA3B, &HC033D3F7, &HC013F303, &H8301C683, &H34900D0, &HCA3B4FC8, &H19EBEC73, &HF303D233, &H7EBD013, &H3C93347, &H3C813F3, &HFA81D1, &H72800000, &H5FC78BEF, &HC25D5B5E, &H8B550004, &H8458BEC, &H560C558B, &H59406A57, &HE8104D2B, &HFFFFA67D
    pvAppendBuffer &H8B104D8B, &H8458BF0, &H558BFA8B, &HA68AE80C, &HC60BFFFF, &H5E5FD70B, &HCC25D, &H53EC8B55, &H8758B56, &H75FF5657, &HACC2E80C, &HFF56FFFF, &HD88B1075, &HE80C5D89, &HFFFFACB4, &H1475FF56, &H7D89F88B, &HACA6E808, &H5750FFFF, &H10458953, &HFFEFC1E8, &H85D88BFF, &H8B1574F6, &H564E187D, &HABCDE853, &H788FFFF, &H75F68547, &H87D8BF1, &HE80C75FF, &HFFFFDE2E, &HDE28E857, &H75FFFFFF
    pvAppendBuffer &HDE20E810, &HE853FFFF, &HFFFFDE1A, &H5D5B5E5F, &H550014C2, &H5151EC8B, &H5310458B, &H48085D8B, &HC7D0F756, &H101045, &H57990000, &H2B0C7D8B, &HFC4589DF, &H8BF85589, &H548B3B34, &H78B043B, &H33044F8B, &HFC4523C6, &H4D23CA33, &H33F033F8, &H3B3489D1, &H43B5489, &H7F8D0731, &HFC4F3108, &H1106D83, &H5E5FD175, &H5DE58B5B, &H55000CC2, &H558BEC8B, &H84D8B0C, &H6A56D12B, &H48B5E10
    pvAppendBuffer &H8D01890A, &H448B0849, &H4189FC0A, &H1EE83FC, &H5D5EEC75, &H550008C2, &H5653EC8B, &H395B106A, &H3674105D, &H20107D83, &H758B5A75, &H7D8B570C, &H57565308, &HFFE8E1E8, &H468D53FF, &H478D5010, &HD3E85010, &H83FFFFE8, &HA3E818C4, &H5FFFF97, &H531&, &H5F304789, &H758B2AEB, &H75FF5308, &HB3E8560C, &H53FFFFE8, &H8D0C75FF, &HE8501046, &HFFFFE8A6, &HE818C483, &HFFFF9776, &H52005
    pvAppendBuffer &H30468900, &HC25D5B5E, &H8B55000C, &H6CEC83EC, &H85D8B53, &HF6335756, &H1A0BF, &H8B038B00, &H4589104B, &H4438BF8, &H8BF04589, &H45890843, &HC438BEC, &H8BE04589, &H45891443, &H18438BE8, &H8BE44589, &H5D8D1C43, &HF44D8994, &H4589DF2B, &HD87589DC, &H83D47D89, &H177310FE, &H560C758B, &HFFFD22E8, &H4C683FF, &H89FC4589, &H75893B04, &H8D5DEB0C, &HE683017E, &HFD478D0F, &H8B0FE083
    pvAppendBuffer &H8B948554, &HFE083C7, &H94854C8B, &HC0C1C18B, &HFC45890E, &HC8C1C18B, &HFC453107, &HE9C1C28B, &HFC4D3103, &HC0C1CA8B, &HFC1C10D, &HEAC1C833, &H8DCA330A, &H558BF847, &HFE083FC, &H3D47D8B, &H855403D1, &HB5540394, &HFC558994, &H94B55489, &HFF969DE8, &HF4758BFF, &HCE8BD68B, &HC10BCAC1, &HD13307C1, &HC9C1CE8B, &H23D6F706, &HD133E475, &H83380C8B, &H458B04C7, &H23CA03E8, &H4D03F445
    pvAppendBuffer &H3F033FC, &HD47D89F1, &H8BF84D8B, &HDC7503D1, &HC0C1C18B, &HDCAC10A, &HC18BD033, &H3302C8C1, &HF0458BD0, &H4523C88B, &HF84D33F8, &H33EC4D23, &HE4458BC8, &H3DC4589, &HE8458BD1, &H89E04D8B, &HCE03E445, &H89F4458B, &H458BE845, &HE04589EC, &H89F0458B, &H458BEC45, &HF04589F8, &H8B32048D, &H8946D875, &H4589F44D, &HD87589F8, &H2A0FF81, &H820F0000, &HFFFFFEDF, &H5F085D8B, &H8B03015E
    pvAppendBuffer &H4301F045, &HEC458B04, &H8B084301, &H4301E045, &HE8458B0C, &H8B144301, &H4301E445, &H104B0118, &H1DC458B, &H43FF1C43, &HE58B5B60, &H8C25D, &H81EC8B55, &HDCEC&, &H8458B00, &H8B575653, &H4D890448, &H8488BF8, &H4D89108B, &HC488BD8, &H8BEC4D89, &H4D891048, &H14488BD4, &H8BD04D89, &H4D891848, &H1C488BB8, &H8BB44D89, &H4D892048, &H24488BE8, &H8BFC4D89, &H4D892848, &H2C488BCC
    pvAppendBuffer &H8BC84D89, &H4D893048, &H34488BC4, &H8BC04D89, &H408B3848, &HAC45893C, &HFF24858D, &H4D89FFFF, &H2A0B9B0, &H55890000, &H2BD233E0, &HBC5589C1, &H89DC4D89, &HFA83A445, &H8B367310, &H478D0C7D, &H40E85004, &H57FFFFFB, &HF633D88B, &HFFFB36E8, &HDC4D8BFF, &H458BF00B, &H8C783A4, &H89F45D89, &H7D89F075, &H81C890C, &H4087489, &HD6E9&, &HFE428D00, &H6A0FE083, &HC5BC8B3D, &HFFFFFF24
    pvAppendBuffer &H28C5B48B, &H8DFFFFFF, &HE0830142, &HFE2830F, &H55895756, &HC58C8BA8, &HFFFFFF24, &H28C59C8B, &H89FFFFFF, &HBFE8E44D, &H6AFFFFFB, &H89575613, &H5589F445, &HFBB0E8F0, &H4D8BFFFF, &H8BC833F4, &HAC0FF045, &H86A06F7, &HEEC1C233, &H75FF5306, &H33CF33E4, &HF44D89C6, &HE8F04589, &HFFFFFB8A, &HFF53016A, &HF08BE475, &H7BE8FA8B, &H8BFFFFFB, &HF033F44D, &H33E4458B, &HF0558BFA, &H7D8AC0F
    pvAppendBuffer &HEBC1F033, &HBC458B07, &HCE03FB33, &HC083D713, &HFE083F9, &H24C58C03, &H13FFFFFF, &HFF28C594, &H458BFFFF, &HC58C03A8, &HFFFFFF24, &H13F44D89, &HFF28C594, &H8C89FFFF, &HFFFF24C5, &HF05589FF, &H28C59489, &H8BFFFFFF, &H296AE875, &H56FC75FF, &HFFFB15E8, &HFF126AFF, &HD88BFC75, &HE856FA8B, &HFFFFFB06, &H75FF0E6A, &H33D833FC, &HF7E856FA, &H33FFFFFA, &HE8FA33D8, &HFFFF93F6, &HF7DC4D8B
    pvAppendBuffer &HFC558BD6, &H276AD2F7, &HFF081C03, &H7C13F875, &H75230408, &HC05523C4, &H23CC4D8B, &H458BE84D, &H23F133C8, &HD033FC45, &H758BDE03, &H3FA13E0, &H1356F45D, &H5D03F07D, &HA85D89B0, &H89AC7D13, &HA3E8E47D, &H6AFFFFFA, &HF875FF22, &HDA8BF88B, &HFA94E856, &H1C6AFFFF, &H33F875FF, &H56DA33F8, &HFFFA85E8, &HD84D8BFF, &H558BDA33, &H33F833EC, &HF18BF855, &H23E07533, &H458BD055, &HF84523EC
    pvAppendBuffer &H33D47523, &HE04D23D0, &H33C4458B, &HB04589F1, &H458BF703, &HAC4589C0, &H458BDA13, &HC44589CC, &H8BC8458B, &H4D03A84D, &HE4558BB8, &H3B45513, &H4589A875, &HE8458BC0, &H89E45D13, &H458BCC45, &HFC5589FC, &H89D4558B, &H558BB855, &HB45589D0, &H89D8558B, &H558BD455, &HD05589EC, &H89E0558B, &H558BD855, &HE84D89F8, &H89DC4D8B, &HC183EC55, &HBC558B08, &HC8458942, &H89E07589, &H5589F85D
    pvAppendBuffer &HDC4D89BC, &H520F981, &H820F0000, &HFFFFFDA6, &H8B08458B, &H4D8BD855, &H30015FD4, &H11B4758B, &H50010458, &HEC558B08, &H10C5011, &H4D8B1048, &H144811D0, &H1B8558B, &H4D8B1850, &H1C7011E8, &H8B204801, &H4811FC4D, &HCC4D8B24, &H8B284801, &H4811C84D, &HC44D8B2C, &H8B304801, &H4811C04D, &HB04D8B34, &H8B384801, &H4811AC4D, &HC080FF3C, &H5E000000, &H5DE58B5B, &H550008C2, &H8B53EC8B
    pvAppendBuffer &H5756085D, &H77BB60F, &HA43B60F, &HB73B60F, &HF53B60F, &HB08E7C1, &H4BB60FF8, &H43B60F03, &H8E7C10D, &HE6C1F80B, &H3B60F08, &HB08E7C1, &H8E2C1F8, &HE43B60F, &HE1C1F00B, &H43B60F08, &H8E6C101, &HB60FF00B, &HE6C10443, &HFF00B08, &HB0243B6, &H43B60FD0, &H8E2C105, &HB60FD00B, &HE2C10843, &HFD00B08, &HB0643B6, &H47B89C8, &H943B60F, &HB08E1C1, &H87389C8, &HC43B60F
    pvAppendBuffer &H5F08E1C1, &H5389C80B, &HB895E0C, &H4C25D5B, &HEC8B5500, &H8508458B, &H8B1074C0, &HC9850C4D, &HC60974, &HE9834000, &H5DF77501, &H550008C2, &H75FFEC8B, &HC75FF0C, &HE80875FF, &HFFFFEFB2, &H8C25D, &H8BEC8B55, &H56531055, &H570C758B, &H2B087D8B, &H2B106AF2, &HC8B5BFA, &H8B0A2B16, &H1B041644, &HC890442, &H8528D17, &HFC174489, &H7501EB83, &H5B5E5FE5, &HCC25D, &H56EC8B55
    pvAppendBuffer &H916CE857, &H758BFFFF, &H588BF08, &HC7030000, &HE836FF50, &H41&, &H53E80689, &H3FFFF91, &H76FF50C7, &H2FE804, &H46890000, &H9140E804, &HC703FFFF, &H876FF50, &H1CE8&, &H8468900, &HFF912DE8, &H50C703FF, &HE80C76FF, &H9&, &HC46895F, &H4C25D5E, &HEC8B5500, &H530C558B, &H8B085D8B, &H18E8C1C3, &HC156CB8B, &HB60F08E9, &H34B60FC9, &HC1C38B10, &HB60F10E8, &HCB60FC0
    pvAppendBuffer &H8E6C111, &H1004B60F, &HE0C1C60B, &HFC10B08, &HE0C1CBB6, &HF5B5E08, &HB110CB6, &H8C25DC1, &HEC8B5500, &H8B575653, &HDB33087D, &HF0C458B, &H15844B6, &H8BC88B99, &HC458BF2, &H8CEA40F, &HF08E1C1, &H995804B6, &HC89C803, &H89F213DF, &H4304DF74, &H7210FB83, &H786781D3, &H7FFF&, &H7C6783, &H5D5B5E5F, &H550008C2, &H5151EC8B, &H57C0570F, &H66147D8B, &HF845130F, &H8BFC558B
    pvAppendBuffer &HFF85F84D, &H458B7374, &H8B56530C, &HF02B1075, &H89084529, &H348B1075, &H89300306, &H758B0C75, &HC5D8B10, &H406748B, &H3047013, &HC5D89D9, &H183BF213, &H75085D8B, &H4703B05, &H703B2374, &H72107704, &H39088B07, &H7730C4D, &H3341C933, &HF0EEBD2, &HF66C057, &H8BF84513, &H4D8BFC55, &HC7D8BF8, &H4037489, &H8910758B, &HC083033C, &H146D8308, &H5E9F7501, &H5FC18B5B, &HC25DE58B
    pvAppendBuffer &H8B550010, &HC4D8BEC, &H1F74C985, &H7D8B5756, &HCD0C8D08, &HFFFFFFF8, &HE9C1F78B, &H278302, &H46783, &HF308C783, &H5D5E5FA5, &H550008C2, &H4D8BEC8B, &H1E98310, &H30785756, &H8B0C458B, &HF02B0875, &H8BC8148D, &H8B04167C, &H7A3B1604, &H72267704, &H77023B1F, &H47A3B20, &H4771672, &H1072023B, &H8308EA83, &HDB7901E9, &H5E5FC033, &HCC25D, &HEBFFC883, &H40C033F5, &H8B55F0EB
    pvAppendBuffer &H39C933EC, &H12760C4D, &H8B08558B, &H440BCA04, &HD7504CA, &HC4D3B41, &HC033F172, &H8C25D40, &HEBC03300, &HEC8B55F8, &H830CEC83, &HF00147D, &HF66C057, &H53F84513, &H76FC5D8B, &HC4D8B63, &H6A08558B, &H452B5840, &H89CA2B10, &H458BF445, &H4D8956F8, &HFC45890C, &H113C8B57, &H748BC78B, &HD68B0411, &HE8104D8B, &HFFFF9C91, &HB084D8B, &HFC450BD3, &HC78B0189, &H8B045189, &HF44D8BD6
    pvAppendBuffer &HFF9C97E8, &HC4D8BFF, &H558BDA8B, &H8C28308, &H83FC4589, &H8901146D, &HBD750855, &H3EB5E5F, &H8BF8458B, &HE58B5BD3, &H10C25D, &H83EC8B55, &H565328EC, &H570C758B, &H6A087D8B, &HE8575604, &HB9B&, &HF2C468B, &H6583C057, &H458900E0, &H30468BE4, &H8BE84589, &H45893446, &H38468BEC, &H8BF04589, &H46A3C46, &H8DF44589, &H16AD845, &HF665050, &HE8D84513, &HFFFFFF29, &HD88B046A
    pvAppendBuffer &H50D8458D, &HE0E85757, &H8BFFFFFD, &HD803384E, &H8B30468B, &H65833C56, &H658300E0, &H458900F4, &HBC033E4, &H46A3446, &H8DE84589, &H16AD845, &H4D895050, &HF05589EC, &HFFFEE8E8, &H3046AFF, &HD8458DD8, &HE8575750, &HFFFFFD9F, &HE46583, &H468BD803, &HC0570F20, &H8BD84589, &H45892446, &H28468BDC, &H8BE04589, &H45893846, &H3C468BF0, &H4589046A, &HD8458DF4, &H66575750, &HE845130F
    pvAppendBuffer &HFFFD66E8, &H8BD803FF, &HC033244E, &H8B28460B, &H45893456, &H30468BDC, &H33F84589, &H2C460BC0, &H8BE04589, &H45893846, &H3C468BE8, &H33EC4589, &H20460BC0, &H4589046A, &HD8458DF4, &HD84D8950, &H5757CA8B, &H89E44D89, &H1CE8F055, &H8BFFFFFD, &HD8032C4E, &H3334568B, &H30460BC0, &H83C0570F, &H8900E465, &H468BDC45, &H89046A20, &H458DF045, &HD84D89D8, &H4E0BC933, &H57575028, &H66E05589
    pvAppendBuffer &HE845130F, &HE8F44D89, &HC31&, &H2B24568B, &H30468BD8, &H89C0570F, &H20B1D845, &H8934468B, &H468BDC45, &HE0458938, &H893C468B, &H468BE445, &H130F6620, &HE1E8E845, &HBFFFF9A, &H46A2C56, &H8DF04589, &H5750D845, &HF4558957, &HBECE8, &H344E8B00, &HC033D82B, &HBFC5D89, &H5E8B3846, &HD84D8930, &H8BDC4589, &HC0332056, &HB13C460B, &H247E8B20, &H5589F633, &HC558BE4, &H8BE04589
    pvAppendBuffer &H528B2842, &H9A73E82C, &H6583FFFF, &HF80B00F0, &H5D89046A, &HD8458DF4, &H8950535B, &HF20BE87D, &H57087D8B, &HEC758957, &HB90E8, &HC4D8B00, &H83FC758B, &H2B00E065, &HF06583F0, &H38418B00, &H8BD84589, &H45893C41, &H24418BDC, &H8BE44589, &H45892841, &H2C418BE8, &H8BEC4589, &H89533441, &H458DF445, &H575750D8, &HB50E8, &H79F02B00, &H7BE8531E, &H83FFFF8C, &H575010C0, &HFBE9E857
    pvAppendBuffer &HF003FFFF, &H5E5FEB78, &H5DE58B5B, &H850008C2, &H531575F6, &H8C58E857, &HC083FFFF, &H8CE85010, &H83FFFFFC, &HDE7401F8, &H8C44E853, &HC083FFFF, &H57575010, &HB04E8, &HEBF02B00, &HEC8B55D2, &H5674EC83, &H570C758B, &H7E8D066A, &H7D895730, &HFCA8E8EC, &HC085FFFF, &HAB850F, &H8D530000, &HDE2B8C5D, &H458D066A, &HC0570F8C, &H130F6650, &HBE8F045, &H6AFFFFFC, &HBC458D06, &HFC00E850
    pvAppendBuffer &H8D57FFFF, &HE8508C45, &HFFFFEBAF, &HE857066A, &HFFFFFBEE, &H8BF44D8B, &HF84D89C6, &H8BF04D8B, &HF8758BF9, &H9FC45C7, &H8B000000, &H10030314, &H4034C8B, &H3044813, &H3BCE13D7, &H3B057510, &H20740448, &H7704483B, &H3B04720D, &H33077310, &HF63347FF, &H570F0EEB, &H130F66C0, &H758BF045, &HF07D8BF4, &H48891089, &H8C08304, &H1FC6D83, &H7D8BB975, &H57066AEC, &HFFFC01E8, &HC758BFF
    pvAppendBuffer &H840FC085, &HFFFFFF5C, &HE8066A5B, &HFFFF8B62, &HB0BF&, &HE816EB00, &HFFFF8B56, &H5650C703, &HA17E856, &H66A0000, &HFF8B45E8, &H50C703FF, &HFB79E856, &H66AFFFF, &HDB7FC085, &H875FF56, &H81AE8, &H8B5E5F00, &H8C25DE5, &HEC8B5500, &H4107D83, &HFF0C75FF, &H7750875, &HFFFC4BE8, &HE805EBFF, &HFFFFFED5, &HCC25D, &HFFEC8B55, &H75FF1875, &HC75FF10, &HE80875FF, &HFFFFFA6B
    pvAppendBuffer &H1275C20B, &HFF1875FF, &H75FF1475, &HFB1DE808, &HC085FFFF, &H75FF1178, &H1475FF18, &HFF0875FF, &H96E80875, &H5D000009, &H550014C2, &HEC81EC8B, &HC8&, &H14758B56, &HC75FF56, &HFFFB3DE8, &H74C085FF, &H75FF560E, &HFAB4E808, &H2E9FFFF, &H53000002, &H75FF5657, &H38858D0C, &H50FFFFFF, &H77AE8, &H105D8B00, &HFF68858D, &H5356FFFF, &H769E850, &H8D560000, &HE850C845, &HFFFFFA82
    pvAppendBuffer &HCC6583, &H3398458D, &H504756FF, &HE8C87D89, &HFFFFFA6E, &H68858D56, &H50FFFFFF, &HFF38858D, &HE850FFFF, &HFFFFFA87, &H18FE9, &H38858B00, &HFFFFFFF, &HC723C057, &H45130F66, &HC883F8, &H858D0E75, &HFFFFFF38, &H6BBE850, &H74EB0000, &HFF68858B, &HC723FFFF, &H7500C883, &H68858D11, &H50FFFFFF, &H6A0E8, &HF2E900, &HC9850000, &H9C8E0F, &H858D0000, &HFFFFFF68, &H38858D50
    pvAppendBuffer &H50FFFFFF, &H8B3E850, &H8D560000, &HFFFF3885, &H72E850FF, &H56000006, &H5098458D, &H50C8458D, &HFFFA0AE8, &H79C085FF, &H8D53560C, &H5050C845, &HFFF936E8, &H458D56FF, &H458D5098, &HE85050C8, &H879&, &H23C8458B, &HC883C7, &H53561374, &H50C8458D, &HF911E850, &HF88BFFFF, &HEB0C5589, &HFC458B09, &H89F87D8B, &H8D560C45, &HE850C845, &H615&, &HF0C7D0B, &HA884&, &HF5448B00
    pvAppendBuffer &HF54C81C0, &HC4&, &HF5448980, &H93E9C0, &H858D0000, &HFFFFFF38, &H68858D50, &H50FFFFFF, &H817E850, &H8D560000, &HFFFF6885, &HD6E850FF, &H56000005, &H50C8458D, &H5098458D, &HFFF96EE8, &H79C085FF, &H8D53560C, &H50509845, &HFFF89AE8, &H458D56FF, &H458D50C8, &HE8505098, &H7DD&, &H2398458B, &HC883C7, &H53561374, &H5098458D, &HF875E850, &HF88BFFFF, &HEB0C5589, &HFC458B09
    pvAppendBuffer &H89F87D8B, &H8D560C45, &HE8509845, &H579&, &H740C7D0B, &HF5448B10, &HF54C8190, &H94&, &HF5448980, &H858D5690, &HFFFFFF68, &H38858D50, &H50FFFFFF, &HFFF8F6E8, &H5F016AFF, &H8556C88B, &H66850FC9, &H8DFFFFFE, &HFF50C845, &H8CE80875, &H5F000005, &HE58B5E5B, &H10C25D, &H81EC8B55, &H80EC&, &H57565300, &H535B046A, &HE81475FF, &H4B4&, &H1075FF53, &H458DF08B, &HC75FF80
    pvAppendBuffer &H34AE850, &H8D530000, &HE850A045, &H498&, &HFF85F88B, &HC7810874, &H100&, &H8D530CEB, &HE8508045, &H480&, &H3B53F88B, &H8D0C73FE, &HFF508045, &HFAE90875, &H8D000000, &HE850C045, &HFFFFF83E, &HE0458D53, &HF834E850, &HC78BFFFF, &HF08BC62B, &H5306EEC1, &H743FE083, &H75FF501A, &HC0458D14, &H50F0048D, &HFFF8B8E8, &HF54489FF, &HF55489E0, &HFF0FEBE4, &H458D1475, &HF0048DC0
    pvAppendBuffer &H4D9E850, &H8B530000, &HE853085D, &HFFFFF7F2, &H46383, &H103C7, &H46A0000, &HFF815E, &H77000001, &H75FF5611, &HC0458D14, &HF7FDE850, &HC085FFFF, &H8D567978, &H8D50A045, &HE850E045, &HFFFFF7EB, &H1478C085, &H8D564075, &H8D508045, &HE850C045, &HFFFFF7D7, &H2E7FC085, &HC0458D56, &H80458D50, &H52E85050, &HB000006, &H560C74C2, &HA0458D53, &H42E85050, &H56000006, &H50E0458D
    pvAppendBuffer &H50A0458D, &H633E850, &H758B0000, &HE0458DE0, &HC150046A, &HEEE81FE6, &H6A000003, &HC0458D04, &H3E3E850, &H75090000, &H6BE94FDC, &H56FFFFFF, &H5080458D, &H425E853, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC81EC8B, &HC0&, &H6A575653, &HFF535B06, &H4DE81475, &H53000003, &H8B1075FF, &H40858DF0, &HFFFFFFFF, &HE8500C75, &H1E0&, &H70858D53, &H50FFFFFF, &H32BE8, &H85F88B00
    pvAppendBuffer &H810874FF, &H180C7, &H530FEB00, &HFF40858D, &HE850FFFF, &H310&, &H3B53F88B, &H8D0F73FE, &HFFFF4085, &H75FF50FF, &H110E908, &H458D0000, &HCBE850A0, &H53FFFFF6, &H50D0458D, &HFFF6C1E8, &H2BC78BFF, &HC1F08BC6, &H835306EE, &H1A743FE0, &H1475FF50, &H8DA0458D, &HE850F004, &HFFFFF745, &HD0F54489, &HD4F55489, &H75FF0FEB, &HA0458D14, &H50F0048D, &H366E8, &H5D8B5300, &H7FE85308
    pvAppendBuffer &H83FFFFF6, &HC7000463, &H103&, &H5E066A00, &H180FF81, &H15770000, &H1475FF56, &H50A0458D, &HFFF68AE8, &HFC085FF, &H8888&, &H858D5600, &HFFFFFF70, &HD0458D50, &HF671E850, &HC085FFFF, &H4C751778, &H40858D56, &H50FFFFFF, &H50A0458D, &HFFF65AE8, &H7FC085FF, &H458D5637, &H858D50A0, &HFFFFFF40, &HD2E85050, &HB000004, &H560F74C2, &H70858D53, &H50FFFFFF, &H4BFE850, &H8D560000
    pvAppendBuffer &H8D50D045, &HFFFF7085, &HE85050FF, &H4AD&, &H8DD0758B, &H66AD045, &H1FE6C150, &H268E8, &H8D066A00, &HE850A045, &H25D&, &H4FCC7509, &HFFFF58E9, &H858D56FF, &HFFFFFF40, &H9CE85350, &H5F000002, &HE58B5B5E, &H10C25D, &H83EC8B55, &H458D60EC, &H1475FFA0, &HFF1075FF, &HE8500C75, &H6C&, &H8D1475FF, &HFF50A045, &H5AE80875, &H8BFFFFFA, &H10C25DE5, &HEC8B5500, &H8D60EC83
    pvAppendBuffer &H75FFA045, &HC75FF10, &H27DE850, &H75FF0000, &HA0458D10, &H875FF50, &HFFFA30E8, &H5DE58BFF, &H55000CC2, &H75FFEC8B, &H1075FF18, &HFF0C75FF, &H6E80875, &HB000004, &HFF1174C2, &H75FF1875, &H875FF14, &HE80875FF, &HFFFFF49F, &H14C25D, &H83EC8B55, &H4D8B64EC, &HC0570F14, &H33575653, &H130F66FF, &H758BCC45, &H4D048DCC, &HFFFFFFFF, &HE8458947, &H33D0458B, &H66F92BDB, &HD445130F
    pvAppendBuffer &H8DE47D89, &H570F1F14, &H66FF33C0, &HF845130F, &H420FD93B, &HFC7D8BD7, &H3BF05589, &HBD870FD3, &H8B000000, &HC38B1075, &H7D89C22B, &HC6348DF4, &H89F8458B, &H4589EC75, &HFD13BFC, &H9783&, &H476FF00, &HFF0C458B, &HD074FF36, &HD034FF04, &H50AC458D, &HFFE377E8, &H8DF08BFF, &HEC83BC7D, &HA5A5A510, &H8BFC8BA5, &H10EC83F0, &HA59C458D, &H8BA5A5A5, &HCC758DFC, &HA5A5A550, &H9271E8A5
    pvAppendBuffer &HF08BFFFF, &HA5CC7D8D, &H8BA5A5A5, &H453BD845, &H721177C8, &HD4458B08, &H73C4453B, &H40C03307, &HEEBC933, &H66C0570F, &HDC45130F, &H8BE04D8B, &H4501DC45, &HF47D8BFC, &H13F0558B, &HEC758BF9, &H144D8B42, &H8908EE83, &H5589F47D, &HEC7589F0, &H860FD33B, &HFFFFFF61, &H8BD0458B, &H6EBCC75, &H89F8558B, &H558BFC55, &HFC4D8B08, &H8BDA3489, &H4489D475, &H8B4304DA, &H4D89D845, &H144D8BD4
    pvAppendBuffer &H8BD87D89, &H7589E47D, &HD04589CC, &HFE85D3B, &HFFFEEE82, &H5FC903FF, &HF8CA7489, &HCA44895E, &HE58B5BFC, &H10C25D, &H56EC8B55, &H8B0C75FF, &HE8560875, &H2E&, &HC985C88B, &H548B2374, &H8B57FCCE, &H33F8CE7C, &HF07EBF6, &HD101D7AC, &HC78B46EA, &HF375C20B, &H5F06E1C1, &H3C0418D, &HC25D5EC6, &H8B550008, &HC4D8BEC, &H7801E983, &H8558B11, &HBCA048B, &H7504CA44, &H1E98305
    pvAppendBuffer &H418DF279, &H8C25D01, &HEC8B5500, &H458B5151, &HC0570F0C, &H7D8B5756, &H130F6608, &H348DF845, &H76F73BC7, &HF8458B32, &HFC5D8B53, &H8B084589, &HEE83FC4E, &H8B168B08, &HC8AC0FC2, &H8450B01, &H86583, &HCB0BE9D1, &HDA8B0689, &HC1044E89, &HF73B1FE3, &H5F5BD977, &H5DE58B5E, &H550008C2, &H558BEC8B, &H74D28510, &H84D8B1E, &HC758B56, &H48BF12B, &H8D01890E, &H448B0849, &H4189FC0E
    pvAppendBuffer &H1EA83FC, &H5D5EEC75, &H55000CC2, &HEC83EC8B, &H10458B64, &HF53D233, &H5642C057, &HFF450C8D, &H66FFFFFF, &HCC45130F, &H33E04D89, &H66D02BC9, &HD445130F, &HD07D8B57, &H89E44D89, &H1C8DDC55, &HC0570F0A, &HF66D233, &H8BEC4513, &HC83BEC75, &H3BDA420F, &H15870FD9, &H8B000001, &HC18B0C55, &H7589C32B, &HC23C8DFC, &H89F0558B, &H5589E87D, &H2BC18BF8, &HF04589C3, &H870FD83B, &HEB&
    pvAppendBuffer &H8B0477FF, &H37FF0C45, &H4D874FF, &H8DD834FF, &HE850AC45, &HFFFFE138, &H8BBC7D8D, &HA5A5A5F0, &HF05D3BA5, &H4D8B4373, &H8BC18BC8, &HF28BC055, &H11FE8C1, &H458BFC45, &HF85583C4, &HFFF3300, &HC101C1A4, &HC0031FEE, &HF00BF90B, &H8BF47D89, &HA40FBC45, &H758901C2, &H89C003F0, &H7D89C475, &HBC4589C8, &HEBC05589, &HC8458B0C, &H8BF44589, &H4589C445, &H10EC83F0, &H8BBC758D, &H9C458DFC
    pvAppendBuffer &HA510EC83, &H8BA5A5A5, &HCC758DFC, &HA5A5A550, &H8FDDE8A5, &HF08BFFFF, &HA5CC7D8D, &H8BA5A5A5, &H453BD845, &H721177F4, &HD4458B08, &H73F0453B, &H40C03307, &HEEBC933, &H66C0570F, &HEC45130F, &H8BF04D8B, &H758BEC45, &HF8558BFC, &H7D8BF003, &H89D113E8, &H4D8BFC75, &HEF8343E4, &HF8558908, &H3BE87D89, &H6860FD9, &H8BFFFFFF, &H3EBD07D, &H8BF0558B, &H5D8B0845, &HC81C89CC, &H458BD88B
    pvAppendBuffer &HCB7C8910, &H7D8B4104, &HD45D8BD8, &H8BD85589, &H5D89DC55, &HD07D89CC, &H89D47589, &H4D3BE44D, &H97820FE0, &H8BFFFFFE, &HC003084D, &HFCC17C89, &H5C895E5F, &H8B5BF8C1, &HCC25DE5, &HEC8B5500, &H570F5151, &H7D8B57C0, &H130F6614, &H558BF845, &HF84D8BFC, &H6B74FF85, &H290C458B, &H45291045, &H8B565308, &H308B105D, &H8903342B, &H708B0C75, &H3741B04, &HC5D8B04, &H5D89D92B, &H3BF21B0C
    pvAppendBuffer &H85D8B18, &H703B0575, &H3B237404, &H10720470, &H88B0777, &H760C4D39, &H41C93307, &HEEBD233, &H66C0570F, &HF845130F, &H8BFC558B, &H7D8BF84D, &H33C890C, &H4037489, &H8308C083, &H7501146D, &H8B5B5EA2, &HE58B5FC1, &H10C25D, &H8BEC8B55, &HD233084D, &H7D8B5756, &H8BF6330C, &H3FE083C7, &H83C6AB0F, &H430F20F8, &H83F233D6, &H430F40F8, &H6EFC1D6, &H23F93423, &H8B04F954, &H5D5E5FC6
    pvAppendBuffer &H550008C2, &H558BEC8B, &H8BC28B08, &HE8C10C4D, &H8B018818, &H10E8C1C2, &H8B014188, &H8E8C1C2, &H88024188, &HC25D0351, &H8B550008, &H8558BEC, &H8B53CA8B, &HC38B0C5D, &H5618E8C1, &H8810758B, &HC1C38B06, &H468810E8, &HC1C38B01, &H468808E8, &HFC38B02, &H8818C1AC, &HE8C1035E, &H44E8818, &HCA8BC38B, &H10C1AC0F, &H8B10E8C1, &H54E88C2, &H8D8AC0F, &H88064688, &HEBC10756, &H5D5B5E08
    pvAppendBuffer &H55000CC2, &H558BEC8B, &H53CA8B08, &H8B0C5D8B, &HC1AC0FC3, &H758B5608, &H8E8C110, &H1688C38B, &H8B014E88, &HC1AC0FCA, &H10E8C110, &H4E88C38B, &HC2AC0F02, &H18E8C118, &H8B035688, &H45E88C3, &H8808E8C1, &HC38B0546, &HC110E8C1, &H468818EB, &H75E8806, &HC25D5B5E, &H8B55000C, &H14558BEC, &H1F74D285, &H56104D8B, &H570C758B, &H2B087D8B, &H8AF92BF1, &H1320E04, &H410F0488, &H7501EA83
    pvAppendBuffer &H5D5E5FF2, &H10C2&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 33033 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub
