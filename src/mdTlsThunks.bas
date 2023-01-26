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
    pvAppendBuffer &H76F299D0, &H76F9CF80, &H76F8B7C0, &H0&, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &H0&, &H0&, &H0&, &H1&, &HFFFFFFFF, &H27D2604B, &H3BCE3C3E, &HCC53B0F6, &H651D06B0, &H769886BC, &HB3EBBD55, &HAA3A93E7, &H5AC635D8, &HD898C296, &HF4A13945, &H2DEB33A0, &H77037D81, &H63A440F2, &HF8BCE6E5, &HE12C4247, &H6B17D1F2, &H37BF51F5, &HCBB64068, &H6B315ECE, &H2BCE3357
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
    ReDim m_baBuffer(0 To 34432 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &HCB0238, &H299E&, &H2C96&, &H39F5&, &H3E0F&, &H3ECC&, &H3F46&, &H41DA&, &H3A9F&, &H3E69&, &H3F09&, &H4086&, &H45A7&, &H34DD&, &H352D&, &H33F2&, &H3589&, &H3614&, &H3560&, &H3746&, &H37D1&, &H3619&, &H28EB&, &H28AE&, &H1F8E&, &H1F3F&, &H1EEC&, &H1E99&, &H6CBE&, &H6B0D&, &H85848BB8, &HFFFFFF78
    pvAppendBuffer &HC25DE58B, &HE80004, &H58000000, &HCA408B2D, &H40000500, &H8B00CA, &HE8C3&, &H2D580000, &HCA409E, &HCA400005, &H8B55C300, &H48EC83EC, &H105D8B53, &H5E046A56, &H61E85356, &H85000075, &H63850FC0, &H57000001, &HC75FF56, &H50D8458D, &H7F46E8, &H87D8B00, &H56D8458D, &H458D5750, &H7E850B8, &H5600007F, &H50D8458D, &H7F29E850, &H53560000, &HFF0C75FF, &HEFE80C75, &H5600007E
    pvAppendBuffer &H14E85353, &H5600007F, &HFFFF79E8, &H10C083FF, &H57575350, &H7975E8, &H67E85600, &H83FFFFFF, &H535010C0, &H63E85353, &H56000079, &HFFFF55E8, &H10C083FF, &H53575350, &H7F08E8, &H57535600, &H7EA8E857, &HE8560000, &HFFFFFF3A, &H5010C083, &HE8535757, &H7936&, &HFF28E856, &HC083FFFF, &H57535010, &H7924E857, &H6A0000, &H837CE857, &HC20B0000, &HE8257456, &HFFFFFF0A, &H5010C083
    pvAppendBuffer &H86E85757, &H6A000073, &HF08B5704, &H809AE8, &H1FE6C100, &H6A1C7709, &H6EB5E04, &H8089E857, &H57560000, &H7E71E853, &HE8560000, &HFFFFFED6, &H5010C083, &H50B8458D, &H86E85353, &H5600007E, &HFFFEC1E8, &H10C083FF, &HB8458D50, &HE8535350, &H7E71&, &HFEACE856, &HC083FFFF, &H8D535010, &H5050B845, &H7E5CE8, &H458D5600, &H575750B8, &H7DF9E8, &H8BE85600, &H83FFFFFE, &H8D5010C0
    pvAppendBuffer &H5750D845, &H7E3BE850, &H53560000, &H806FE857, &HFF560000, &HE8530C75, &H8065&, &HD8458D56, &HC75FF50, &H8058E8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H68EC83EC, &H105D8B53, &H5E066A56, &HD9E85356, &H85000073, &H77850FC0, &H57000001, &HC75FF56, &H50C8458D, &H7DBEE8, &H87D8B00, &H56C8458D, &H458D5750, &H7FE85098, &H5600007D, &H50C8458D, &H7DA1E850, &H53560000, &HFF0C75FF
    pvAppendBuffer &H67E80C75, &H5600007D, &H8CE85353, &H5600007D, &HFFFDF1E8, &HB005FF, &H53500000, &HEBE85757, &H56000077, &HFFFDDDE8, &HB005FF, &H53500000, &HD7E85353, &H56000077, &HFFFDC9E8, &HB005FF, &H53500000, &H7AE85357, &H5600007D, &HE8575753, &H7D1A&, &HFDACE856, &HB005FFFF, &H50000000, &HE8535757, &H77A6&, &HFD98E856, &HB005FFFF, &H50000000, &HE8575753, &H7792&, &HE857006A
    pvAppendBuffer &H81EA&, &H7456C20B, &HFD78E827, &HB005FFFF, &H50000000, &HF2E85757, &H6A000071, &HF08B5706, &H7F06E8, &H1FE6C100, &H6A2C7709, &H6EB5E06, &H7EF5E857, &H57560000, &H7CDDE853, &HE8560000, &HFFFFFD42, &HB005&, &H458D5000, &H53535098, &H7CF0E8, &H2BE85600, &H5FFFFFD, &HB0&, &H98458D50, &HE8535350, &H7CD9&, &HFD14E856, &HB005FFFF, &H50000000, &H98458D53, &HC2E85050
    pvAppendBuffer &H5600007C, &H5098458D, &H5FE85757, &H5600007C, &HFFFCF1E8, &HB005FF, &H8D500000, &H5750C845, &H7C9FE850, &H53560000, &H7ED3E857, &HFF560000, &HE8530C75, &H7EC9&, &HC8458D56, &HC75FF50, &H7EBCE8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H758B56EC, &H56046A08, &H7243E8, &H74C08500, &H8D046A14, &HE8502046, &H7234&, &H574C085, &HEB40C033, &H5EC03302, &H4C25D, &H56EC8B55
    pvAppendBuffer &H6A08758B, &H15E85606, &H85000072, &H6A1474C0, &H30468D06, &H7206E850, &HC0850000, &HC0330574, &H3302EB40, &HC25D5EC0, &H8B550004, &HA8EC81EC, &H53000000, &H8D0C5D8B, &H5756B845, &H5053046A, &H7E3CE8, &H20438D00, &H8950046A, &H858DF845, &HFFFFFF78, &H7E27E850, &H75FF0000, &H58858D14, &H50FFFFFF, &H5098458D, &HFF78858D, &H8D50FFFF, &HE850B845, &H882&, &H75FF046A, &H7D48E810
    pvAppendBuffer &HF6330000, &H4602E883, &H85144589, &H50577EC0, &HE81075FF, &H8056&, &H475C20B, &H2EBCE8B, &HE1C1C933, &H589D8D05, &H3FFFFFF, &H98458DD9, &HB58DC103, &HFFFFFF78, &H89D9F753, &H350FC45, &HB87D8DF1, &H5756F903, &H4F3E8, &H53575600, &HE8FC75FF, &H2BB&, &H4814458B, &H6A144589, &HC0855E01, &H6AA97F, &HE81075FF, &H7FFE&, &H274C20B, &HE6C1F633, &H589D8D05, &H3FFFFFF
    pvAppendBuffer &H107589DE, &H98458D53, &HBD8DC603, &HFFFFFF78, &H758DFE2B, &H10752BB8, &HE8565750, &H4A0&, &H5FE8046A, &H83FFFFFB, &H8D5010C0, &H8D509845, &H8D50B845, &HE850D845, &H7B09&, &H8D57046A, &H5050D845, &H7AA5E8, &HFF046A00, &H458D0C75, &HE85050D8, &H7A96&, &H27E8046A, &H83FFFFFB, &H8D5010C0, &H5050D845, &H7560E8, &HFF046A00, &H458DF875, &HE85050D8, &H7A72&, &H8D56046A
    pvAppendBuffer &H5050D845, &H7A65E8, &H8D565700, &H45039845, &HE8505310, &H1FB&, &H50D8458D, &HFF58858D, &H8D50FFFF, &HE8509845, &HA60&, &H8D08758B, &H46A9845, &HC6E85650, &H6A00007C, &H58858D04, &H50FFFFFF, &H5020468D, &H7CB4E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &HF8EC81EC, &H53000000, &H8D0C5D8B, &H57569845, &H5053066A, &H7C90E8, &H30438D00, &H8950066A, &H858DF845, &HFFFFFF38
    pvAppendBuffer &H7C7BE850, &H75FF0000, &H8858D14, &H50FFFFFF, &HFF68858D, &H8D50FFFF, &HFFFF3885, &H458D50FF, &H4DE85098, &H6A000007, &H1075FF06, &H7B99E8, &H83DB3300, &H894302E8, &HC0851445, &HFF505A7E, &HA7E81075, &HB00007E, &H8B0475C2, &H3302EBC3, &H30C06BC0, &HFF089D8D, &H8D8DFFFF, &HFFFFFF68, &HFF38B58D, &H7D8DFFFF, &H3D80398, &HD8F753C8, &H51FC4D89, &HF803F003, &HCCE85756, &H56000004
    pvAppendBuffer &H75FF5357, &H219E8FC, &H458B0000, &H45894814, &H5B016A14, &HA67FC085, &H75FF006A, &H7E4CE810, &HC20B0000, &HDB330274, &H8D30C36B, &HFFFF688D, &H89D8DFF, &H8DFFFFFF, &HFFFF38BD, &H98758DFF, &HC803D803, &HF82B5153, &H2B104D89, &HE85657F0, &H477&, &HABE8066A, &H5FFFFF9, &HB0&, &H68858D50, &H50FFFFFF, &H5098458D, &H50C8458D, &H7950E8, &H57066A00, &H50C8458D, &H78ECE850
    pvAppendBuffer &H66A0000, &H8D0C75FF, &H5050C845, &H78DDE8, &HE8066A00, &HFFFFF96E, &HB005&, &H458D5000, &HE85050C8, &H73A5&, &H75FF066A, &HC8458DF8, &HB7E85050, &H6A000078, &H458D5606, &HE85050C8, &H78AA&, &HFF535657, &H54E81075, &H8D000001, &H8D50C845, &HFFFF0885, &H858D50FF, &HFFFFFF68, &H8F1E850, &H758B0000, &H68858D08, &H6AFFFFFF, &HE8565006, &H7B09&, &H858D066A, &HFFFFFF08
    pvAppendBuffer &H30468D50, &H7AF7E850, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC83EC8B, &H57565320, &H565E046A, &HFFF8E1E8, &H85D8BFF, &H5010C083, &H1075FF53, &H50E0458D, &H788CE8, &H458D5600, &HE85050E0, &H7857&, &HE0458D56, &HE8535350, &H781E&, &HE0458D56, &H1075FF50, &HE81075FF, &H780E&, &HF8A0E856, &H758BFFFF, &H10C0830C, &H50147D8B, &HE8575756, &H784D&, &H8D57046A, &HE850E045
    pvAppendBuffer &H7817&, &H7BE8046A, &H83FFFFF8, &H535010C0, &H50E0458D, &H782BE850, &H46A0000, &HFFF865E8, &H10C083FF, &H1075FF50, &H50E0458D, &H7813E850, &H46A0000, &HFFF84DE8, &H10C083FF, &H75FF5350, &H1075FF10, &H77FCE8, &HFF046A00, &H56561075, &H7799E8, &HE8046A00, &HFFFFF82A, &H5010C083, &H50E0458D, &H1075FF53, &H77D8E8, &HFF046A00, &H57571075, &H7775E8, &HE8046A00, &HFFFFF806
    pvAppendBuffer &H5010C083, &HE8575756, &H77B9&, &H458D046A, &H75FF50E0, &H79E7E810, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC83EC8B, &H57565330, &H575F066A, &HFFF7D1E8, &H85D8BFF, &HB0BE&, &H50C60300, &H1075FF53, &H50D0458D, &H7778E8, &H458D5700, &HE85050D0, &H7743&, &HD0458D57, &HE8535350, &H770A&, &HD0458D57, &H1075FF50, &HE81075FF, &H76FA&, &HF78CE857, &H7D8BFFFF, &H8BC60314
    pvAppendBuffer &H56500C75, &H3AE85757, &H6A000077, &H458D5706, &H4E850D0, &H6A000077, &HF768E806, &HB005FFFF, &H50000000, &HD0458D53, &H16E85050, &H6A000077, &HF750E806, &HB005FFFF, &H50000000, &H8D1075FF, &H5050D045, &H76FCE8, &HE8066A00, &HFFFFF736, &HB005&, &HFF535000, &H75FF1075, &H76E3E810, &H66A0000, &H561075FF, &H7680E856, &H66A0000, &HFFF711E8, &HB005FF, &H8D500000, &H5350D045
    pvAppendBuffer &HE81075FF, &H76BD&, &H75FF066A, &HE8575710, &H765A&, &HEBE8066A, &H5FFFFF6, &HB0&, &H57575650, &H769CE8, &H8D066A00, &HFF50D045, &HCAE81075, &H5F000078, &HE58B5B5E, &H10C25D, &H83EC8B55, &H565360EC, &H5B046A57, &HF6B4E853, &H7D8BFFFF, &H10C08310, &H875FF50, &H57C0458D, &H765FE850, &H8D530000, &H5050C045, &H762AE8, &H458D5300, &H75FF50C0, &H875FF08, &H75EDE8
    pvAppendBuffer &H458D5300, &H575750C0, &H75E1E8, &H73E85300, &H8BFFFFF6, &HC0830C5D, &H14758B10, &H8D565350, &HE850C045, &H7066&, &H57E8046A, &H83FFFFF6, &H535010C0, &HAE85656, &H6A000076, &HF644E804, &HC083FFFF, &H75FF5010, &HE0458D08, &HF2E85057, &H6A000075, &HE0458D04, &HE8535350, &H758E&, &H1FE8046A, &H83FFFFF6, &H575010C0, &H8D0875FF, &HE850E045, &H7016&, &H5756046A, &H759AE8
    pvAppendBuffer &HE8046A00, &HFFFFF5FE, &H5010C083, &H50E0458D, &HAEE85757, &H6A000075, &HF5E8E804, &HC083FFFF, &HFF575010, &H458D0875, &H96E850A0, &H6A000075, &HA0458D04, &HE8565650, &H7532&, &HC3E8046A, &H6AFFFFF5, &HC7035F10, &H56565350, &H7574E8, &H5E046A00, &HC0458D56, &HA0458D50, &H7539E850, &HE8560000, &HFFFFF59E, &H8D50C703, &H8D50E045, &H5050A045, &H754CE8, &H87E85600, &H3FFFFF5
    pvAppendBuffer &H75FF50C7, &HA0458D08, &HE0458D50, &H7533E850, &H8D560000, &H8D50C045, &H5050E045, &H74CDE8, &H5FE85600, &H3FFFFF5, &H8D5350C7, &H5350E045, &H7510E8, &H458D5600, &H75FF50A0, &H773FE808, &H5E5F0000, &H5DE58B5B, &H550010C2, &HEC81EC8B, &H90&, &H6A575653, &HE8565E06, &HFFFFF526, &HBB107D8B, &HB0&, &HFF50C303, &H458D0875, &HE85057A0, &H74CD&, &HA0458D56, &H98E85050
    pvAppendBuffer &H56000074, &H50A0458D, &HFF0875FF, &H5BE80875, &H56000074, &H50A0458D, &H4FE85757, &H56000074, &HFFF4E1E8, &H14758BFF, &H5D8BC303, &H5653500C, &H50A0458D, &H6ED5E8, &HE8066A00, &HFFFFF4C6, &HB005&, &H56535000, &H7477E856, &H66A0000, &HFFF4B1E8, &HB005FF, &HFF500000, &H458D0875, &HE85057D0, &H745D&, &H458D066A, &H535350D0, &H73F9E8, &HE8066A00, &HFFFFF48A, &HB005&
    pvAppendBuffer &HFF575000, &H458D0875, &H7FE850D0, &H6A00006E, &HE8575606, &H7403&, &H67E8066A, &H5FFFFF4, &HB0&, &HD0458D50, &HE8575750, &H7415&, &H4FE8066A, &H5FFFFF4, &HB0&, &H75FF5750, &H70858D08, &H50FFFFFF, &H73F8E8, &H8D066A00, &HFFFF7085, &H565650FF, &H7391E8, &HE8066A00, &HFFFFF422, &HB0BF&, &H50C70300, &HE8565653, &H73D1&, &H565E066A, &H50A0458D, &HFF70858D
    pvAppendBuffer &HE850FFFF, &H7393&, &HF3F8E856, &HC703FFFF, &HD0458D50, &H70858D50, &H50FFFFFF, &H73A3E850, &HE8560000, &HFFFFF3DE, &HFF50C703, &H858D0875, &HFFFFFF70, &HD0458D50, &H7387E850, &H8D560000, &H8D50A045, &H5050D045, &H7321E8, &HB3E85600, &H3FFFFF3, &H8D5350C7, &H5350D045, &H7364E8, &H858D5600, &HFFFFFF70, &H875FF50, &H7590E8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H20EC83EC
    pvAppendBuffer &H75FF046A, &H1075FF08, &H7574E8, &HFF046A00, &H75FF0C75, &H7567E814, &H46A0000, &H50E0458D, &H687FE8, &HE4658300, &H187D8300, &HE045C700, &H1&, &H46A0B74, &H501875FF, &H7540E8, &HE0458D00, &HC75FF50, &HE80875FF, &H2BC&, &H50E0458D, &HFF0C75FF, &H4BE80875, &H8DFFFFF3, &HFF50E045, &H75FF1475, &H29EE810, &HE58B0000, &H14C25D, &H83EC8B55, &H66A30EC, &HFF0875FF
    pvAppendBuffer &HFAE81075, &H6A000074, &HC75FF06, &HE81475FF, &H74ED&, &H458D066A, &H5E850D0, &H83000068, &H8300D465, &HC700187D, &H1D045, &HB740000, &H75FF066A, &HC6E85018, &H8D000074, &HFF50D045, &H75FF0C75, &H28DE808, &H458D0000, &H75FF50D0, &H875FF0C, &HFFF459E8, &HD0458DFF, &H1475FF50, &HE81075FF, &H26F&, &HC25DE58B, &H8B530014, &H8B0C2444, &HF710244C, &H8BD88BE1, &HF7082444
    pvAppendBuffer &H3142464, &H24448BD8, &H3E1F708, &H10C25BD3, &H40F98000, &HF9801573, &HF067320, &HE0D3C2A5, &H33D08BC3, &H1FE180C0, &H33C3E2D3, &HC3D233C0, &H7340F980, &H20F98015, &HAD0F0673, &HC3EAD3D0, &HD233C28B, &HD31FE180, &HC033C3E8, &H55C3D233, &H558BEC8B, &H8B565310, &H8B570C75, &HF22B087D, &HFA2B106A, &H160C8B5B, &H448B0A03, &H42130416, &H170C8904, &H8908528D, &H83FC1744, &HE57501EB
    pvAppendBuffer &H5D5B5E5F, &H55000CC2, &H5151EC8B, &H8B1C558B, &H8B562045, &H8B570875, &HD7030C7D, &H89104513, &H4468916, &H7710453B, &H3B04720D, &H330773D7, &HC93340C0, &H570F0EEB, &H130F66C0, &H4D8BF845, &HF8458BFC, &H5F244503, &H3284D13, &H46891445, &H13C68B08, &H4E89184D, &HE58B5E0C, &H24C25D, &H8BEC8B55, &H4D8B0C55, &H31028B08, &H4428B01, &H8B044131, &H41310842, &HC428B08, &H5D0C4131
    pvAppendBuffer &H550008C2, &HEC83EC8B, &H84D8B0C, &H758B5653, &H83018B10, &HEEC104C1, &HFF335702, &H89107589, &H48DFC4D, &H485&, &HF8458900, &H2C74F685, &H8B0C4D29, &H87589D9, &H8B0C758B, &H48D087D, &HB1E8501E, &H8900005A, &H45B8D03, &H7501EF83, &H10758BED, &H4D8BFE8B, &HF8458BFC, &H45C7DB33, &H108&, &HFF83B00, &H9683&, &H2BC78B00, &H81048DC6, &H8B104589, &H89FCB954, &HDE3B0C55
    pvAppendBuffer &H45FF0575, &H85DB3308, &HE83475DB, &HFFFFF0F6, &H58805, &H458B5000, &H8C0C10C, &H64D9E850, &H45890000, &HF0DCE80C, &H4D8BFFFF, &HC558B08, &H884B60F, &H688&, &H3318E0C1, &H8320EBD0, &H1E7606FE, &H7504FB83, &HF0B8E819, &H8805FFFF, &H50000005, &H500C458B, &H649EE8, &H8BD08B00, &H88B1045, &H33FC458B, &HB80C89CA, &H10458B47, &H83FC4D8B, &H894304C0, &H7D3B1045, &H74820FF8
    pvAppendBuffer &H5FFFFFFF, &HE58B5B5E, &HCC25D, &H83EC8B55, &H458D20EC, &HFF046AE0, &HE8501075, &H6FFB&, &H458D046A, &H75FF50E0, &H875FF08, &H6FBDE8, &HFF046A00, &H458D1075, &HE85050E0, &H6FAE&, &H458D046A, &H75FF50E0, &HC75FF0C, &H6F9DE8, &H5DE58B00, &H55000CC2, &HEC83EC8B, &HD0458D30, &H75FF066A, &HB0E85010, &H6A00006F, &HD0458D06, &H875FF50, &HE80875FF, &H6F72&, &H75FF066A
    pvAppendBuffer &HD0458D10, &H63E85050, &H6A00006F, &HD0458D06, &HC75FF50, &HE80C75FF, &H6F52&, &HC25DE58B, &H8B55000C, &H24EC83EC, &H5308458B, &H7D8B5756, &H39378B0C, &H304F0F30, &HBCE85646, &H33000052, &H41D88BC9, &H33DC5D89, &HC0570FD2, &H45130F66, &HFF13BE4, &H868C&, &H85D8B00, &H8904478D, &H458BF445, &HFC5D89E8, &H29DC5D8B, &H4589FC7D, &HE4458BE0, &H29F85D89, &H4589F87D, &HF4458BE8
    pvAppendBuffer &H3B085D8B, &H8B0B7F0B, &H1C8BFC5D, &HEC5D8903, &H658304EB, &HF3B00EC, &H8B077F, &HEBF04589, &HF0658304, &HF05D8B00, &H5D03C033, &HF87D8BEC, &H5D03C013, &HE04513E8, &HE06583, &H8BE84589, &H1C89F445, &HC7D8B07, &H574DB85, &H4F0FCA3B, &HC08341D1, &HF4458904, &HA47ECE3B, &H5FDC5D8B, &H8B13895E, &HE58B5BC3, &H8C25D, &H83EC8B55, &H565310EC, &HC7D8B57, &HDE8B378B, &H8902E3C1
    pvAppendBuffer &H5D89F475, &HEEFCE8F0, &HFF53FFFF, &H33D08B10, &HC5589C9, &HF7EF685, &HC12B078B, &H8987048B, &H3B418A04, &H8BF17CCE, &H1B8B085D, &H37FDE3B, &H8B015E8D, &H2E7C1FB, &HE8F87D89, &HFFFFEEC6, &H8B10FF57, &HFC5589D0, &HB7EDB85, &HFA8BCF8B, &H3302E9C1, &H8BABF3C0, &HFF33084D, &H7C393947, &H83C38B1C, &HE0C1FCC2, &H3F78B02, &HB1048BD0, &H8D028946, &H313BFC52, &H758BF37E, &HC458BF4
    pvAppendBuffer &H8950008B, &HDEE80845, &H8B000004, &H45890855, &H74C085F4, &HD3C88B1A, &H7EF73BE2, &H59206A12, &H458BC82B, &H4408B0C, &HD00BE8D3, &H50F4458B, &H5812E852, &HFF500000, &H8B561475, &H53560C75, &HE8FC75FF, &H3A2F&, &H85104D8B, &H392474C9, &H8B207C39, &H834BFC45, &H4503FCC0, &H78DB85F8, &HEB108B04, &H89D23302, &HE883B914, &H3B4B4704, &HFFEA7E39, &HE856F075, &H6137&, &HFFEE05E8
    pvAppendBuffer &H50FF56FF, &HF875FF08, &H56FC758B, &H6122E8, &HEDF0E800, &HFF56FFFF, &H5E5F0850, &H5DE58B5B, &H550010C2, &H8B56EC8B, &HFF570C75, &H50C9E836, &H6A0000, &H5657F88B, &HE80875FF, &HFFFFFEAC, &H76013F83, &H83078B10, &H7500873C, &H7894808, &H7701F883, &H5FC78BF0, &H8C25D5E, &HEC8B5500, &H75FF006A, &H875FF0C, &H4E8&, &H8C25D00, &HEC8B5500, &H8B1CEC83, &H56530845, &H8B0C758B
    pvAppendBuffer &H45895700, &H3B0E8BE8, &H89F98BC1, &H4F0FEC4D, &H57DF8BF8, &HE802E3C1, &H503B&, &HE3C1D803, &HE45D8902, &HFFED5DE8, &H10FF53FF, &HD233D88B, &H85FC5D89, &H8D4D7EFF, &HC78BBB0C, &H89085D8B, &H7529F85D, &HFC5D8BF8, &H8DF44D89, &H758BBE0C, &H7F063B08, &HF8758B08, &HEB0E348B, &H89F63302, &H758B9334, &H7F063B0C, &HEB318B04, &H8BF63302, &H8342F45D, &H8304F445, &H894804E9, &HFC5D8B33
    pvAppendBuffer &HCA7CD73B, &HE0C1C78B, &H50C30304, &H8D084589, &H5057FB04, &H50BB048D, &H3BA0E853, &H4D8B0000, &H10458BE8, &HEC4D0341, &H850C4D89, &H8B0C74C0, &H7FC83B00, &H1488D06, &H510C4D89, &H4FBAE8, &H33F08B00, &HC93342D2, &H1639C28B, &H5D8B297C, &H3F148D08, &H3BFCC383, &H8B047FC2, &H3302EB3B, &H89FF85FF, &H450F863C, &HEB8340C8, &H7E063B04, &HFC5D8BE6, &H8342D233, &H8900107D, &H83840F0E
    pvAppendBuffer &HF000000, &HF66C057, &H39F04513, &H767C0C55, &H8D105D8B, &H5D890446, &HF45D8B08, &H89087529, &H5D8BEC5D, &HF85D89F0, &H89105D8B, &H163BE845, &H388B077F, &HEBF47D89, &HF4658304, &H7F133B00, &H87D8B0B, &H8907048B, &H4EB1045, &H106583, &H33107D8B, &HF47D03C0, &H7D03C013, &HEC4513F8, &HEC6583, &H4589FF85, &HE8458BF8, &H7D8B3889, &H3B05740C, &HCA4F0FD1, &H4C08342, &H3BE84589
    pvAppendBuffer &H8BAB7ED7, &H75FFFC5D, &H530E89E4, &H5F22E8, &HEBF0E800, &HFF53FFFF, &H8B5F0850, &H8B5B5EC6, &HCC25DE5, &HEC8B5500, &H560C4D8B, &H3278C985, &H8B08758B, &H2E0C106, &H267DC83B, &H8399C18B, &HC20303E2, &H8102F8C1, &H3E1&, &H49057980, &H41FCC983, &H486448B, &HD303E1C1, &HC0B60FE8, &HC03302EB, &H8C25D5E, &HEC8B5500, &H758B5653, &H7D8B570C, &H8B168B08, &H1F9830F, &HC0330875
    pvAppendBuffer &HF044739, &HFA83C844, &H33087501, &H44639C0, &H3BD0440F, &HFC28BCA, &HC085C14F, &H1C8D3174, &H3BFE2B86, &H83067EC1, &HEB000865, &H1F348B06, &H3B087589, &H33047EC2, &H8B07EBF6, &H8753933, &H75391B72, &H83117708, &HE88304EB, &H33D47501, &H5B5E5FC0, &H8C25D, &HEB40C033, &HFFC883F4, &H8B55EFEB, &H5D8B53EC, &H8D57560C, &H83990343, &H3C8D03E2, &H2FFC102, &H4DFDE857, &HF08B0000
    pvAppendBuffer &H7C01FF83, &H33CF8B09, &H47E8DC0, &HDB85ABF3, &HFB8B3A74, &H8B03E7C1, &H834B084D, &H18A08EF, &H84D8941, &HD285D38B, &HC2830379, &H2FAC103, &HE181CF8B, &H8000001F, &H83490579, &HF41E0C9, &HE0D3C0B6, &H4964409, &HCB75DB85, &H76013E83, &H83068B10, &H7500863C, &H6894808, &H7701F883, &HC68B5FF0, &HC25D5B5E, &H8B550008, &HC4D8BEC, &H3D78C985, &H8758B56, &HE0C1068B, &H7DC83B05
    pvAppendBuffer &H99C18B2F, &H31FE283, &H5F8C1C2, &H1FE181, &H5798000, &HE0C98349, &H42D23341, &H7D83E2D3, &H6740010, &H4865409, &HD2F706EB, &H4865421, &HCC25D5E, &HEC8B5500, &H8B18EC83, &H56530845, &HC7D8B57, &H18391F8B, &H53184F0F, &HE8EC5D89, &H4D27&, &HF633D08B, &H5589C033, &H458940E8, &H3BC88BF8, &H8B677CD8, &H5F8D0845, &H8BC72B04, &HC558BFA, &H4589FA2B, &HF07D89F4, &H458B03EB
    pvAppendBuffer &H87D8BF4, &H87F0F3B, &H8918048B, &H4EBFC45, &HFC6583, &H47F0A3B, &H2EB3B8B, &H558BFF33, &HF7C033F0, &HFC7D03D7, &H7D03C013, &H1A3C89F8, &H830C558B, &H458900D0, &H74FF85F8, &HFCE3B05, &H8341F14F, &H4D3B04C3, &H8BB37EEC, &H3289E855, &H855B5E5F, &H520A75C0, &H32BAE8, &HEBC03300, &H8BC28B02, &H8C25DE5, &HEC8B5500, &H7D8B5756, &H6AF63308, &H206A5A10, &H8BCA2B59, &H85E8D3C7
    pvAppendBuffer &H8B0675C0, &H3E7D3CA, &H75FAD1F2, &HC68B5FE9, &H4C25D5E, &HEC8B5500, &H9908458B, &H31FE283, &H5F8C1C2, &HE8504056, &H4C4F&, &H75FF016A, &H56F08B08, &HFFFEB1E8, &H5EC68BFF, &H4C25D, &H8BEC8B55, &HEC83084D, &H5756531C, &H481DB33, &H10000D9, &HD9148B00, &H4D95483, &HD9448B00, &HC2AC0F04, &H10F8C110, &H89F45589, &HFB83F845, &H3309750F, &H658341C9, &H1CEB00FC, &H66C0570F
    pvAppendBuffer &HEC45130F, &H8BF0458B, &HF66EC4D, &H8BE44513, &H4589E455, &HE8458BFC, &H8D0FFB83, &H6A017B, &HDEF7F61B, &H2BF7AF0F, &H1B256AD1, &H5250FC45, &HFFF639E8, &HF44503FF, &H13084D8B, &HE883F855, &HDA8301, &H8BF10401, &H5411F845, &H558B04F1, &HD0A40FF4, &H10E2C110, &H19D91429, &H8B04D944, &H10FB83DF, &HFF64820F, &H5E5FFFFF, &H5DE58B5B, &H550004C2, &HEC83EC8B, &H758B5610, &HE856570C
    pvAppendBuffer &H51F4&, &H8DF04589, &HE8500446, &H51E8&, &H8DF44589, &HE8500846, &H51DC&, &H8DF84589, &HE8500C46, &H51D0&, &H89084D8B, &H318BFC45, &H8B04798D, &H4E0C1C6, &H458DF803, &HE85057F0, &HFFFFF69C, &HABE821EB, &H8D00003B, &HE850F045, &H3C38&, &HF0458D57, &HF682E850, &H458DFFFF, &H59E850F0, &H8300003B, &H458D10EF, &HEE8350F0, &HE8D37501, &H3B7E&, &H50F0458D, &H3C0BE8
    pvAppendBuffer &H458D5700, &H55E850F0, &H8BFFFFF6, &HFF561075, &H8EE8F075, &H8D00006C, &HFF500446, &H82E8F475, &H8D00006C, &HFF500846, &H76E8F875, &H8D00006C, &HFF500C46, &H6AE8FC75, &H5F00006C, &H5DE58B5E, &H55000CC2, &HEC83EC8B, &H8B565310, &H56570C75, &H5123E8, &HF0458900, &H5004468D, &H5117E8, &HF4458900, &H5008468D, &H510BE8, &HF8458900, &H500C468D, &H50FFE8, &H85D8B00, &H8DFC4589
    pvAppendBuffer &H738DF045, &HE8505604, &HFFFFF5D4, &HC683FF33, &H3B394710, &H458D2D76, &HFFE850F0, &H8D00005A, &HE850F045, &H59F5&, &H50F0458D, &H3C54E8, &H458D5600, &HA5E850F0, &H83FFFFF5, &H3B4710C6, &H8DD3723B, &HE850F045, &H5AD2&, &H50F0458D, &H59C8E8, &H458D5600, &H81E850F0, &H8BFFFFF5, &HFF561075, &HBAE8F075, &H8D00006B, &HFF500446, &HAEE8F475, &H8D00006B, &HFF500846, &HA2E8F875
    pvAppendBuffer &H8D00006B, &HFF500C46, &H96E8FC75, &H5F00006B, &HE58B5B5E, &HCC25D, &H53EC8B55, &H75FF5756, &H107D8B1C, &H571875FF, &H457E8, &H74C08500, &HE6E3E81F, &H62BEFFFF, &HBB00CA5B, &HCA4000, &HF003F32B, &HFFE6D0E8, &H5A6EB9FF, &H2CEB00CA, &H8B1C75FF, &H75FF147D, &H3CE85718, &HE8000000, &HFFFFE6B5, &HCA58C3BE, &H4000BB00, &HF32B00CA, &HA2E8F003, &HB9FFFFE6, &HCA57F3, &HC103CB2B
    pvAppendBuffer &H89084D8B, &H458B0841, &H471890C, &H1001C7, &H38890000, &H5D5B5E5F, &H550018C2, &H8B56EC8B, &HF4680875, &H6A000000, &H1E85600, &H8300003B, &H7D830CC4, &H20741010, &H18107D83, &H7D831074, &H25752010, &HE06C7, &H206A0000, &H6C712EB, &HC&, &H8EB186A, &HA06C7, &H106A0000, &H560C75FF, &HFFF49AE8, &HC25D5EFF, &H8B53000C, &H835151DC, &HC483F0E4, &H6B8B5504, &H246C8904
    pvAppendBuffer &H83EC8B04, &H4B8B2CEC, &HC438B08, &H8B106A56, &H20091, &H38100F00, &HF5E018B, &HD6030228, &HF8EF0F66, &HF07D290F, &H740AE883, &HE883483C, &H481E7401, &HF01E883, &H9585&, &HA280F00, &HF66D603, &HFF9DE38, &HD6030228, &HDE380F66, &HA280FF8, &HF66D603, &HFF9DE38, &HD6030228, &HDE380F66, &H7D290FF8, &H3A280FF0, &H280FD603, &HFD60332, &HFE07529, &H66F07528, &HF7DE380F
    pvAppendBuffer &H32A280F, &H75290FD6, &H7D280FF0, &H380F66F0, &HFE07DDE, &HD6032228, &HDE380F66, &H380F66FD, &H280FFCDE, &H66D6031A, &HFBDE380F, &H312280F, &H380F66D6, &H280FFADE, &H66D6030A, &HF9DE380F, &HDE380F66, &H380F663A, &H66323CDE, &H7CDF380F, &H438B1032, &H110F5E10, &H5DE58B38, &HC25BE38B, &H8B53000C, &H835151DC, &HC483F0E4, &H6B8B5504, &H246C8904, &H83EC8B04, &H4B8B2CEC, &HC438B08
    pvAppendBuffer &H8B106A56, &H1FC91, &H38100F00, &HF5E018B, &HD6030228, &HF8EF0F66, &HF07D290F, &H740AE883, &HE883483C, &H481E7401, &HF01E883, &H9585&, &HA280F00, &HF66D603, &HFF9DC38, &HD6030228, &HDC380F66, &HA280FF8, &HF66D603, &HFF9DC38, &HD6030228, &HDC380F66, &H7D290FF8, &H3A280FF0, &H280FD603, &HFD60332, &HFE07529, &H66F07528, &HF7DC380F, &H32A280F, &H75290FD6, &H7D280FF0
    pvAppendBuffer &H380F66F0, &HFE07DDC, &HD6032228, &HDC380F66, &H380F66FD, &H280FFCDC, &H66D6031A, &HFBDC380F, &H312280F, &H380F66D6, &H280FFADC, &H66D6030A, &HF9DC380F, &HDC380F66, &H380F663A, &H66323CDC, &H7CDD380F, &H438B1032, &H110F5E10, &H5DE58B38, &HC25BE38B, &H8B55000C, &HCEC81EC, &H8B000001, &HC9330C45, &H45C75756, &H40201F4, &H8DF18B08, &H45C70678, &H402010F8, &HBD148D80, &H4&
    pvAppendBuffer &HFC45C766, &H7D89361B, &HE85589E4, &H840FD285, &HE7&, &H8D0C7D8B, &H5589F455, &HF4958DF0, &HC1FFFFFE, &HD02B02E0, &H8908458B, &H958DEC55, &HFFFFFEF4, &H4589C22B, &H73F73B08, &HB0048D1B, &HFEF48D8D, &HC103FFFF, &H4D52E850, &H558B0000, &HE9C933EC, &H86&, &HF0B5848B, &H33FFFFFE, &HC4589D2, &HF7F7C68B, &H7508FF83, &H24C68B0A, &HB0043C07, &H8A027401, &H75D285C1, &H6E0F6632
    pvAppendBuffer &HC0330C4D, &H66D2570F, &H8BC06E0F, &HF66F045, &HF66C862, &HF66D162, &HC2DF3A, &H6600B60F, &HC2163A0F, &HFFD03301, &H5589F045, &H8427EB0C, &H662374C0, &HC4D6E0F, &H570FC033, &H6E0F66D2, &H620F66C0, &H620F66C8, &H3A0F66D1, &H6600C2DF, &H45163A0F, &H558B000C, &H33028BEC, &H84890C45, &HFFFEF4B5, &H4C283FF, &H4608458B, &H3BEC5589, &H820FE875, &HFFFFFF41, &H8BE47D8B, &H858D1055
    pvAppendBuffer &HFFFFFEF8, &HD62BF08B, &H8901778D, &H758B0875, &H86D8310, &H6E0F6601, &HF660858, &H6604406E, &H66086E0F, &HFC506E0F, &HD0620F66, &HCB620F66, &HD1620F66, &H214290F, &H7510408D, &HF068D4, &H8D510000, &HFFFEF485, &H61E850FF, &H8B000037, &HCC483CF, &H304E1C1, &HD233144D, &H8506280F, &H8B0B74D2, &H74C22BC7, &H380F6605, &H290FC0DB, &HC6834201, &H10E98310, &HE076D73B, &HE58B5E5F
    pvAppendBuffer &H10C25D, &H83EC8B55, &H565310EC, &H8DC03357, &H3340F07D, &HA20F53C9, &H895BF38B, &H4778907, &H89084F89, &H45F70C57, &HF8&, &HF75F7402, &HF845&, &H56740008, &H8B107D8B, &H8758BC7, &HF799CE8B, &H3E283D1, &H302EFC1, &H2F8C1C2, &H8906C083, &H4468D06, &HC140D0F7, &HE08302E8, &HC1414003, &HE18302E9, &H40C18303, &H8986148D, &H1FC96, &H8E0C8D00, &HFF575251, &H8E890C75
    pvAppendBuffer &H200&, &HFFFDCDE8, &H40C033FF, &HC03302EB, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H228EC81, &H75FF0000, &HD8858D1C, &HFFFFFFFD, &H50501875, &H501C458D, &H50F4458D, &HFFFADBE8, &H1475FFFF, &HFFF4458D, &H8D501C75, &HE850DC45, &H431&, &HC110458B, &HFF5004E8, &H458D0875, &HC75FFDC, &H33BE850, &HE58B0000, &H18C25D, &H81EC8B55, &H228EC, &H1C75FF00, &HFDD8858D, &H75FFFFFF
    pvAppendBuffer &H8D505018, &H8D501C45, &HE850F445, &HFFFFFA88, &H8D1475FF, &H75FFF445, &H458D501C, &HDEE850DC, &H8B000003, &HE8C11045, &H75FF5004, &HDC458D08, &H500C75FF, &H359E8, &H5DE58B00, &H550018C2, &HEC81EC8B, &H210&, &H8D2875FF, &HFFFDF085, &H2475FFFF, &H458D5050, &H458D5028, &H35E850F4, &HFFFFFFFA, &H458D0875, &HFF106AF4, &HC6A1475, &HFF2075FF, &H75FF1C75, &H1075FF18, &HFF0C75FF
    pvAppendBuffer &HE8502875, &HD3C&, &HC25DE58B, &H8B550024, &H10EC81EC, &HFF000002, &H858D2875, &HFFFFFDF0, &H502475FF, &H28458D50, &HF4458D50, &HF9E6E850, &H106AFFFF, &H8D0C75FF, &H75FFF445, &HFF0C6A08, &H75FF2075, &H1875FF1C, &HFF1475FF, &H75FF1075, &H9BE85028, &H8B00000E, &H24C25DE5, &HEC8B5500, &H7D8B5751, &H89C03318, &HFF85FC45, &H8B536374, &H8B560C5D, &H10758B0B, &H4D89F12B, &HFFE3B18
    pvAppendBuffer &HC085F742, &HB60F1A75, &H50561445, &H308458B, &H19E850C1, &H8B000035, &HC483184D, &HFC458B0C, &HC75C985, &H3B42D233, &H440F1075, &HFC4589C2, &H3B0E048D, &HE751045, &HFF0875FF, &H55FF2075, &H23831C, &H330102EB, &H2BFC458B, &H5EA475FE, &HE58B5F5B, &H1CC25D, &H56EC8B55, &H8B20758B, &HE883C6, &H86840F, &H75FF0000, &H2475FF28, &HE883016A, &H83617401, &H458D01E8, &H75FF5014
    pvAppendBuffer &HC75FF10, &H740875FF, &H67E848, &H75FF0000, &H184D8B28, &H382475FF, &H20741C4D, &H50FE468D, &H1075FF51, &HFF0C75FF, &H2EE80875, &HFFFFFFFF, &H458D2875, &H2475FF1C, &H25EB016A, &H50FF468D, &H1075FF51, &HFF0C75FF, &HEE80875, &HEBFFFFFF, &H1FE81F, &HD7EB0000, &H301C458A, &H458D1445, &H75FF5014, &HC75FF10, &HE80875FF, &H5&, &H24C25D5E, &HEC8B5500, &HFF2075FF, &H75FF1C75
    pvAppendBuffer &H1875FF1C, &HFF1475FF, &H75FF1075, &H875FF0C, &H4E8&, &H1CC25D00, &HEC8B5500, &H530C4D8B, &H56145D8B, &H57003983, &H74187D8B, &H74FF854A, &H10458B46, &H12BF78B, &H420FC73B, &H3018BF0, &H53560845, &H33C1E850, &H4D8B0000, &H2BDE030C, &HCC483FE, &H758B3101, &H75313910, &H875FF33, &H852475FF, &HFF0575FF, &H3EB2055, &H8B1C55FF, &H20830C45, &H8B19EB00, &H14EB1075, &H2475FF53
    pvAppendBuffer &H575FE3B, &HEB2055FF, &H1C55FF03, &HFE2BDE03, &HE873FE3B, &H2E74FF85, &H8B0C458B, &H2BC68B08, &H3BF78BC1, &HF0420FC7, &H5608458B, &H5053C103, &H3356E8, &HC458B00, &HC483DE03, &H2B30010C, &H10758BFE, &H5E5FD575, &H20C25D5B, &HEC8B5500, &H1C7D8B57, &H4474FF85, &HC5D8B53, &H3B8356, &H75FF1075, &H2475FF08, &H8B2055FF, &H3891045, &H458B03EB, &H8B032B10, &HF3B39F7, &H45033342
    pvAppendBuffer &HFF505608, &H75FF1475, &H63EEE818, &H33290000, &H1187501, &HFE2B1475, &H5B5EC375, &H20C25D5F, &HEC8B5500, &H5310EC83, &H56145D8B, &H8B08758B, &H89008B06, &HDB850845, &H8B575074, &H7D290C7D, &H8D068B10, &H5751F04D, &HFF0476FF, &H75FF0850, &H8468D08, &HF0458D50, &H9DE85050, &HFF000063, &H468D0875, &HE8505708, &H329F&, &H8D0875FF, &H8B50F045, &HC7031045, &H328DE850, &H7D030000
    pvAppendBuffer &H18C48308, &H7501EB83, &H5B5E5FB8, &HC25DE58B, &H8B550010, &H10EC83EC, &H8758B56, &H147D8B57, &H88B068B, &H85084D89, &H8B4E74FF, &H568D0C45, &H5D8B5308, &H89C32B10, &H52510C45, &H8D50C303, &HE850F045, &H6334&, &H4E8D068B, &H4D8D5108, &H76FF51F0, &H450FF04, &H8D0875FF, &H53500846, &H3226E8, &H84D8B00, &H8B08568D, &HC4830C45, &H83D9030C, &HC27501EF, &H8B5E5F5B, &H10C25DE5
    pvAppendBuffer &HEC8B5500, &H8B08558B, &H458B0C4D, &H4428910, &H8908428D, &HFF31FF0A, &HE8501475, &H31EB&, &H5D0CC483, &H550010C2, &H6AE8EC8B, &HB9FFFFDD, &HCA6877, &H4000E981, &HC10300CA, &H51084D8B, &H1475FF50, &HFF74418D, &H75FF1075, &H50406A0C, &H5034418D, &HFFFE74E8, &H10C25DFF, &HEC8B5500, &H5370EC83, &H14758B56, &HB1E85657, &H8D000046, &HF88B044E, &HD07D8951, &H46A3E8, &H84E8D00
    pvAppendBuffer &H51F84589, &HE8CC4589, &H4694&, &H890C4E8D, &H8951F445, &H85E8C845, &H8B000046, &H89560875, &H4589F045, &H4676E8C4, &H4E8D0000, &HD4458904, &HC0458951, &H4667E8, &H84E8D00, &H51E84589, &HE8BC4589, &H4658&, &H890C4E8D, &H8951E445, &H49E8B845, &H8B000046, &H89560C75, &H4589E045, &H463AE8B4, &H4E8D0000, &HFC458904, &HB0458951, &H462BE8, &H84E8D00, &H51084589, &HE8AC4589
    pvAppendBuffer &H461C&, &H890C4E8D, &H89511445, &HDE8A845, &H8B000046, &H89561075, &H45890C45, &H45FEE8A4, &H4E8D0000, &H51D88B04, &HE8A05D89, &H45F0&, &H89084E8D, &H8951DC45, &HE1E89C45, &H8D000045, &H45890C4E, &H458951D8, &H45D2E898, &H4D8B0000, &H8BF08BDC, &H7589D855, &H9045C794, &HA&, &H758B06EB, &H107D8BEC, &H8BD47D03, &HDF33FC45, &H310C3C1, &HFC4589C3, &HC1D44533, &HF8030CC0
    pvAppendBuffer &H7D89DF33, &HFC7D8B10, &H308C3C1, &HFC7D89FB, &H458BF833, &HE84503F8, &H4589C833, &H8458BF8, &H310C1C1, &H7C7C1C1, &H33084589, &HC0C1E845, &HF845010C, &HC1F84D33, &H4D0108C1, &HDC4D8908, &H33084D8B, &HF4458BC8, &H33E44503, &HF44589D0, &HC114458B, &HC20310C2, &H8907C1C1, &H45331445, &HCC0C1E4, &H33F44501, &HC2C1F455, &H14550108, &H89104D01, &H558BD855, &H8BD03314, &H4503F045
    pvAppendBuffer &H89F033E0, &H458BF045, &H10C6C10C, &HC2C1C603, &HC458907, &HC1E04533, &H45010CC0, &HF07533F0, &H108C6C1, &H75890C75, &HC758BEC, &H458BF033, &H104533EC, &H110C0C1, &H45891445, &H14458BEC, &HC6C1C133, &HCC0C107, &H8B104501, &H4D33EC4D, &H8C1C110, &H89144D01, &H4D8BEC4D, &H8BC83314, &HC1C10C45, &HE84D8907, &H3F84D8B, &HC1D933CA, &HC30310C3, &H330C4589, &HF4558BC2, &H30CC0C1
    pvAppendBuffer &H33C803D6, &HF84D89D9, &HC10C4D8B, &HCB0308C3, &H330C4D89, &HFC458BC8, &H8907C1C1, &H4D8BE44D, &HC1CA33DC, &HC10310C1, &H33FC4589, &HF0758BC6, &H30CC0C1, &H33D003F7, &HF45589CA, &HC1FC558B, &HD10308C1, &H33FC5589, &H8458BD0, &H8907C2C1, &H558BE055, &HC1D633D8, &HC20310C2, &H33084589, &H87D8BC7, &H30CC0C1, &H89D633F0, &HC2C1F075, &H89FA0308, &HF833087D, &H8307C7C1, &H8901906D
    pvAppendBuffer &H850FD47D, &HFFFFFE5A, &H3D0458B, &H45891045, &HCC458BD0, &H89F84503, &H458BCC45, &HF44503C8, &H8BC84589, &H4503BC45, &HBC4589E8, &H3B8458B, &H4589E445, &HB4458BB8, &H89E04503, &H458BB445, &H84503AC, &H89A05D01, &H458BAC45, &H144503A8, &H8B985501, &H5D8B1855, &HA84589D0, &H3A4458B, &H45890C45, &H94458BA4, &H88EC4503, &H9445891A, &HE8C1C38B, &H1428808, &HE8C1C38B, &H2428810
    pvAppendBuffer &HC1C47501, &H5A8818EB, &HCC5D8B03, &H5A88C38B, &H8E8C104, &H8B054288, &HC07D01C3, &H8810E8C1, &H7D8B0642, &HB07D01FC, &HC19C4D01, &H5A8818EB, &HC85D8B07, &H5A88C38B, &H8E8C108, &H8B094288, &H10E8C1C3, &HC10A4288, &H5A8818EB, &HC45D8B0B, &H5A88C38B, &H8E8C10C, &H8B0D4288, &H10E8C1C3, &HC10E4288, &H5A8818EB, &HC05D8B0F, &H5A88C38B, &H8E8C110, &H8B114288, &H10E8C1C3, &HC1124288
    pvAppendBuffer &H5A8818EB, &HBC5D8B13, &H5A88C38B, &H8E8C114, &H8B154288, &H10E8C1C3, &HC1164288, &H5A8818EB, &HB85D8B17, &H5A88C38B, &H8E8C118, &H8B194288, &H10E8C1C3, &HC11A4288, &H5A8818EB, &HB45D8B1B, &H5A88C38B, &H8E8C11C, &H8B1D4288, &H10E8C1C3, &HC11E4288, &H5A8818EB, &HB05D8B1F, &H5A88C38B, &H8E8C120, &H8B214288, &H10E8C1C3, &HC1224288, &H5A8818EB, &HAC5D8B23, &H5A88C38B, &H8E8C124
    pvAppendBuffer &H8B254288, &H10E8C1C3, &HC1264288, &H5A8818EB, &HA85D8B27, &H5A88C38B, &H8E8C128, &H8B294288, &H10E8C1C3, &HC12A4288, &H5A8818EB, &HA45D8B2B, &H5A88C38B, &H8E8C12C, &H8B2D4288, &H10E8C1C3, &HC12E4288, &H5A8818EB, &HA05D8B2F, &H5A88C38B, &H8E8C130, &H8D314288, &HC38B3C4A, &HC118EBC1, &H428810E8, &H335A8832, &H8B9C5D8B, &H345A88C3, &H8808E8C1, &HC38B3542, &H8810E8C1, &HEBC13642
    pvAppendBuffer &H375A8818, &H8B985D8B, &H385A88C3, &H8808E8C1, &HC38B3942, &H8810E8C1, &HEBC13A42, &H3B5A8818, &H8B94558B, &H8E8C1C2, &H41881188, &H5FC28B01, &HC110E8C1, &H885E18EA, &H51880241, &HE58B5B03, &H14C25D, &H56EC8B55, &H8B1075FF, &H75FF0875, &H62E8560C, &H6A000045, &H1475FF10, &H5020468D, &H2CA6E8, &H18458B00, &H830CC483, &H89007466, &H5D5E7846, &H550014C2, &H8B56EC8B, &HFF570875
    pvAppendBuffer &H76FF0C75, &H207E8D30, &H10468D57, &HCAE85650, &H8BFFFFFA, &HC9337856, &H75010780, &HCA3B410B, &H4800674, &HF5740139, &HC25D5E5F, &H8B550008, &H10EC83EC, &H6AF0458D, &H2075FF10, &H2C49E850, &HC4830000, &HF0458D0C, &HFF006A50, &H75FF2475, &H1875FF1C, &HFF1475FF, &H75FF1075, &H875FF0C, &H3F0BE8, &H5DE58B00, &H550020C2, &H75FFEC8B, &HFF016A24, &H75FF2075, &H1875FF1C, &HFF1475FF
    pvAppendBuffer &H75FF1075, &H875FF0C, &H3EE3E8, &H20C25D00, &HEC8B5500, &HFFD780E8, &H7841B9FF, &HE98100CA, &HCA4000, &H4D8BC103, &HFF505108, &H18B1475, &HFF1075FF, &H30FF0C75, &H5028418D, &H5018418D, &HFFF888E8, &H10C25DFF, &HEC8B5500, &H8B084D8B, &H41890C45, &H10458B2C, &H5D304189, &H55000CC2, &H8B56EC8B, &H346A0875, &HE856006A, &H2BC0&, &H830C4D8B, &H8B002C66, &H28668301, &H30468900
    pvAppendBuffer &H8910458B, &H468D0446, &HFF0E8908, &H1475FF31, &H2B75E850, &HC4830000, &HC25D5E18, &H8B550010, &H20EC81EC, &H53000004, &HDB335756, &HFD6085C7, &HDB41FFFF, &H706A0000, &HFD70858D, &H9D89FFFF, &HFFFFFD64, &H85C75053, &HFFFFFD68, &H1&, &HFD6C9D89, &H55E8FFFF, &H8B00002B, &H858D0C75, &HFFFFFF60, &H50561F6A, &H2B1EE8, &H1F468A00, &H8018C483, &HFFFF60A5, &H3F24F8FF, &H8588400C
    pvAppendBuffer &HFFFFFF7F, &HFBE0858D, &H75FFFFFF, &HB0E85010, &H6A00004A, &H570F591E, &H60B58DC0, &H6AFFFFFE, &H130F661E, &HFFFE6085, &H68BD8DFF, &HF3FFFFFE, &HF6659A5, &H8D804513, &H7D8D8075, &H130F6688, &HFFFEE085, &H6AA5F3FF, &HB58D591E, &HFFFFFEE0, &HFE649D89, &HBD8DFFFF, &HFFFFFEE8, &HF3845D89, &H59206AA5, &HFBE0B58D, &HFEBBFFFF, &H8D000000, &HFFFDE0BD, &H33A5F3FF, &HBD8947FF, &HFFFFFE60
    pvAppendBuffer &H8B807D89, &HCBB60FC3, &H8303F8C1, &HB60F07E1, &HFF6005B4, &H858DFFFF, &HFFFFFDE0, &HF723EED3, &H458D5056, &H8AE85080, &H56000042, &HFE60858D, &H8D50FFFF, &HFFFEE085, &H76E850FF, &H8D000042, &HFFFEE085, &H458D50FF, &H858D5080, &HFFFFFCE0, &HE379E850, &H858DFFFF, &HFFFFFEE0, &H80458D50, &HFFE85050, &H8D000048, &HFFFE6085, &H858D50FF, &HFFFFFDE0, &HE0858D50, &H50FFFFFE, &HFFE34EE8
    pvAppendBuffer &H60858DFF, &H50FFFFFE, &HFDE0858D, &H5050FFFF, &H48D1E8, &HE0858D00, &H50FFFFFC, &HFE60858D, &HE850FFFF, &H48A9&, &H5080458D, &HFC60858D, &HE850FFFF, &H4899&, &H5080458D, &HFEE0858D, &H8D50FFFF, &HE8508045, &H3697&, &HFCE0858D, &H8D50FFFF, &HFFFDE085, &H858D50FF, &HFFFFFEE0, &H367DE850, &H858D0000, &HFFFFFEE0, &H80458D50, &HE0858D50, &H50FFFFFC, &HFFE2D2E8, &HE0858DFF
    pvAppendBuffer &H50FFFFFE, &H5080458D, &H4858E850, &H458D0000, &H858D5080, &HFFFFFDE0, &H4833E850, &H858D0000, &HFFFFFC60, &H60858D50, &H50FFFFFE, &HFEE0858D, &HE850FFFF, &H482E&, &HFD60858D, &H8D50FFFF, &HFFFEE085, &H458D50FF, &H14E85080, &H8D000036, &HFFFE6085, &H458D50FF, &HE8505080, &HFFFFE26F, &H5080458D, &HFEE0858D, &H5050FFFF, &H35F2E8, &H60858D00, &H50FFFFFC, &HFE60858D, &H8D50FFFF
    pvAppendBuffer &HE8508045, &H35DB&, &HFBE0858D, &H8D50FFFF, &HFFFDE085, &H858D50FF, &HFFFFFE60, &H35C1E850, &H858D0000, &HFFFFFCE0, &HE0858D50, &H50FFFFFD, &H479CE8, &H858D5600, &HFFFFFDE0, &H80458D50, &H40EFE850, &H8D560000, &HFFFE6085, &H858D50FF, &HFFFFFEE0, &H40DBE850, &HEB830000, &H20890F01, &H8DFFFFFE, &HFFFEE085, &HE85050FF, &H26B4&, &HFEE0858D, &H8D50FFFF, &H50508045, &H3562E8
    pvAppendBuffer &H80458D00, &H875FF50, &H37D0E8, &H5B5E5F00, &HC25DE58B, &H8B55000C, &H20EC83EC, &H59076A57, &H45C6C033, &H7D8D09E0, &H66ABF3E1, &H458DAAAB, &H75FF50E0, &H875FF0C, &HFFFCE1E8, &HE58B5FFF, &H8C25D, &H81EC8B55, &H114EC, &H57565300, &H758BC033, &H8DDB3308, &H5D88E17D, &HABABABE0, &H8DAAAB66, &H5050E045, &HFF0C75FF, &H7D830456, &H1E750C24, &H75FF0C6A, &HF0458D20, &H2809E850
    pvAppendBuffer &HC4830000, &H5D89660C, &HFC5D88FD, &H1FF45C6, &H458D32EB, &H858D50E0, &HFFFFFEEC, &H1E23E850, &H75FF0000, &HEC858D24, &HFFFFFFFE, &HE8502075, &H1CFA&, &H50F0458D, &HFEEC858D, &HE850FFFF, &H1D87&, &H50E0458D, &HFF3C858D, &HE850FFFF, &H1DF1&, &H8D1C75FF, &HFFFF3C85, &H1875FFFF, &H1CAAE850, &HC0330000, &H8DD05D88, &HABABD17D, &HAAAB66AB, &H50F0458D, &H8D0C75FF, &H50568C45
    pvAppendBuffer &HFFFBDAE8, &H6A046AFF, &H8C458D0C, &HFBB7E850, &H106AFFFF, &H50D0458D, &H8C458D50, &HFB6FE850, &H75FFFFFF, &H3C858D14, &HFFFFFFFF, &HE8501075, &H1C7A&, &H50C0458D, &HFF3C858D, &HE850FFFF, &H1D07&, &H8D2C758B, &H5056D045, &H50C0458D, &H5826E850, &H4D8D0000, &H74F685C0, &H28558B18, &HD02BC18B, &H320A048A, &H41D80A01, &H7501EE83, &H75DB84F3, &H1475FF16, &HFF8C458D, &H75FF3075
    pvAppendBuffer &HAE85010, &H33FFFFFB, &H3303EBF6, &HC03346F6, &HABE07D8D, &H6A506A, &H8AABABAB, &H7D8DE045, &HABC033F0, &H8AABABAB, &H7D8DF045, &HABC033D0, &H8AABABAB, &H7D8DD045, &HABC033C0, &H8AABABAB, &H858DC045, &HFFFFFF3C, &H26E2E850, &H8D8A0000, &HFFFFFF3C, &H6A8C458D, &H50006A34, &H26CFE8, &H8C4D8A00, &H8B18C483, &H5B5E5FC6, &HC25DE58B, &H8B55002C, &H14EC81EC, &H53000001, &HC0335756
    pvAppendBuffer &H3308758B, &HE17D8DDB, &HABE05D88, &HAB66ABAB, &HE0458DAA, &H75FF5050, &H456FF0C, &HC247D83, &HC6A1E75, &H8D2075FF, &HE850F045, &H265B&, &H660CC483, &H88FD5D89, &H45C6FC5D, &H32EB01FF, &H50E0458D, &HFEEC858D, &HE850FFFF, &H1C75&, &H8D2475FF, &HFFFEEC85, &H2075FFFF, &H1B4CE850, &H458D0000, &H858D50F0, &HFFFFFEEC, &H1BD9E850, &H458D0000, &H858D50E0, &HFFFFFF3C, &H1C43E850
    pvAppendBuffer &H75FF0000, &H3C858D1C, &HFFFFFFFF, &HE8501875, &H1AFC&, &H5D88C033, &HD17D8DD0, &H66ABABAB, &H458DAAAB, &H75FF50F0, &H8C458D0C, &H2CE85056, &H6AFFFFFA, &H8D0C6A04, &HE8508C45, &HFFFFFA09, &H458D106A, &H8D5050D0, &HE8508C45, &HFFFFF9C1, &H8D1475FF, &H75FF8C45, &H1075FF28, &HF9AFE850, &H75FFFFFF, &H3C858D14, &HFFFFFFFF, &HE8502875, &H1ABA&, &H5D88C033, &HC17D8DC0, &H66ABABAB
    pvAppendBuffer &H458DAAAB, &H858D50C0, &HFFFFFF3C, &H1B39E850, &H75FF0000, &HD0458D30, &HC0458D50, &H2C75FF50, &H5657E8, &H8DC03300, &HABABE07D, &H458AABAB, &HF07D8DE0, &HABABC033, &H458AABAB, &HD07D8DF0, &H6AABC033, &HABAB5350, &HD0458AAB, &H33C07D8D, &HABABABC0, &HC0458AAB, &HFF3C858D, &HE850FFFF, &H254C&, &HFF3C858A, &H346AFFFF, &H538C458D, &H253AE850, &H458A0000, &H18C4838C, &H8B5B5E5F
    pvAppendBuffer &H2CC25DE5, &HEC8B5500, &H8B0C558B, &H8B56104D, &H68B0875, &H1890233, &H3304468B, &H41890442, &H8468B04, &H89084233, &H468B0841, &HC42330C, &H5E0C4189, &HCC25D, &H51EC8B55, &HC5D8B53, &H7D8B5756, &H45C76608, &H8BE100FC, &HD1C18B0F, &H1E183E8, &H578B0389, &HD1C28B04, &H1E283E8, &HB1FE1C1, &H1FE2C1C8, &H8B044B89, &HC68B0877, &HE683E8D1, &HC1D00B01, &H53891FE6, &HC4F8B08
    pvAppendBuffer &HE8D1C18B, &HB01E183, &H73895FF0, &H44B60F0C, &HE0C1FC0D, &H5E033118, &H5DE58B5B, &H550008C2, &H8B56EC8B, &H56570875, &H394BE8, &HC7D8B00, &H468D0789, &H3DE85004, &H89000039, &H468D0447, &H31E85008, &H89000039, &H468D0847, &H25E8500C, &H89000039, &H5E5F0C47, &H8C25D, &H83EC8B55, &H575620EC, &HF633106A, &H56E0458D, &H243EE850, &H106A0000, &H8D0C75FF, &HE850F045, &H240B&
    pvAppendBuffer &H83087D8B, &H100F18C4, &HC68BE04D, &H6A1FE083, &HC82B591F, &HF8C1C68B, &H87048B05, &H1A8E8D3, &H100F0C74, &HF66F045, &H110FC8EF, &H458DE04D, &HE85050F0, &HFFFFFF00, &H80FE8146, &H7C000000, &H8D106AC9, &HFF50E045, &HBCE81075, &H83000023, &H5E5F0CC4, &HC25DE58B, &H8B55000C, &H8458BEC, &H8B00100F, &HF660C45, &HF1BE870, &H280FD528, &H100FDD, &H6610458B, &H1BC8700F, &H66C5280F
    pvAppendBuffer &HC1443A0F, &H3A0F6601, &H6610D144, &H66D0EF0F, &HD9443A0F, &H3A0F6600, &HF11E944, &HF66C228, &H6608DA73, &H8F8730F, &HEAEF0F66, &HD8EF0F66, &HFE5280F, &HF66C328, &H661FD472, &H1FD0720F, &HF3720F66, &HC8280F01, &HFC730F66, &H730F6604, &HF660CD8, &H6604F973, &H66CBEB0F, &H1F5720F, &H66D9280F, &H66E5EB0F, &H1FF3720F, &HE0EB0F66, &H66C1280F, &H1EF0720F, &HD8EF0F66, &H66C1280F
    pvAppendBuffer &H19F0720F, &HD8EF0F66, &H66D3280F, &H4DB730F, &HFA730F66, &HEF0F660C, &HCA280FD1, &H66C2280F, &H2D1720F, &HD0720F66, &HEF0F6601, &HC2280FC8, &HD0720F66, &HEF0F6607, &HEF0F66C8, &HEF0F66CB, &HEF0F66CA, &H700F66CC, &H110F1BC1, &HCC25D00, &HEC8B5500, &HC758B56, &H87D8B57, &HE837FF56, &H52C1&, &H5004468D, &HE80477FF, &H52B5&, &H5008468D, &HE80877FF, &H52A9&, &H500C468D
    pvAppendBuffer &HE80C77FF, &H529D&, &HC25D5E5F, &H8B550008, &H44EC83EC, &H8758B56, &HA8BE83, &H74000000, &HDEE85606, &H33000033, &H84B60FC9, &H880E&, &H8D448900, &HF98341BC, &H83EE7210, &H5600FC65, &H334BE8, &HBC458D00, &HF3E85650, &H8B000032, &HC9330C55, &H888E048A, &H83411104, &HF47210F9, &HAC68&, &H56006A00, &H2243E8, &H83068A00, &H8B5E0CC4, &H8C25DE5, &HEC8B5500, &H8758B56
    pvAppendBuffer &HAC68&, &H56006A00, &H2223E8, &HC4D8B00, &HFCBA&, &HFF106A00, &HB60F1075, &H44468901, &H141B60F, &HF484689, &H890241B6, &HB60F4C46, &HE0830341, &H5046890F, &H441B60F, &H4689C223, &H41B60F54, &H58468905, &H641B60F, &HF5C4689, &H830741B6, &H46890FE0, &H41B60F60, &H89C22308, &HB60F6446, &H46890941, &H41B60F68, &H6C46890A, &HB41B60F, &H890FE083, &HB60F7046, &HC2230C41
    pvAppendBuffer &HF744689, &H890D41B6, &HB60F7846, &H46890E41, &H41B60F7C, &H84A6830F, &H0&, &H890FE083, &H8086&, &H88868D00, &H50000000, &H215AE8, &H18C48300, &HCC25D5E, &HEC8B5500, &HFFCCD8E8, &HA7BEB9FF, &HE98100CA, &HCA4000, &H4D8BC103, &HFF505108, &H818D1075, &HA8&, &H6A0C75FF, &H818D5010, &H98&, &HED07E850, &HC25DFFFF, &H8B55000C, &HCEC83EC, &HE8575653, &HFFFFCC99
    pvAppendBuffer &HBE084D8B, &HCAAE2C, &H815A406A, &HCA4000EE, &H8DF00300, &H418B6479, &H6AE2F760, &H8B070300, &H56515ADA, &HDA13F08B, &H8D58406A, &HE183084E, &H50C12B3F, &H80685252, &H6A000000, &H7D8B5740, &H20478D08, &HEC0EE850, &H458DFFFF, &HA40F50F4, &HC15303F3, &HE85603E6, &H50F0&, &H458D086A, &HE85750F4, &HC5&, &H560C758B, &HB2E837FF, &H8D000050, &HFF500446, &HA6E80477, &H8D000050
    pvAppendBuffer &HFF500846, &H9AE80877, &H8D000050, &HFF500C46, &H8EE80C77, &H8D000050, &HFF501046, &H82E81077, &H8D000050, &HFF501446, &H76E81477, &H8D000050, &HFF501846, &H6AE81877, &H8D000050, &HFF501C46, &H5EE81C77, &H6A000050, &H57006A68, &H205FE8, &HCC48300, &H8B5B5E5F, &H8C25DE5, &HEC8B5500, &H8758B56, &H6A686A, &H2042E856, &HC4830000, &H6706C70C, &HC76A09E6, &HAE850446, &H46C7BB67
    pvAppendBuffer &H6EF37208, &HC46C73C, &HA54FF53A, &H7F1046C7, &HC7510E52, &H688C1446, &H46C79B05, &H83D9AB18, &H1C46C71F, &H5BE0CD19, &H4C25D5E, &HEC8B5500, &HFFCB64E8, &HAE2CB9FF, &HE98100CA, &HCA4000, &H4D8BC103, &HFF505108, &H418D1075, &HC75FF64, &H8D50406A, &HE8502041, &HFFFFEB99, &HCC25D, &H83EC8B55, &H458D40EC, &H75FF50C0, &HA7E808, &H306A0000, &H50C0458D, &HE80C75FF, &H1F8B&
    pvAppendBuffer &H8B0CC483, &H8C25DE5, &HEC8B5500, &H8758B56, &HC868&, &H56006A00, &H1F93E8, &HCC48300, &H9ED806C7, &H46C7C105, &HBB9D5D04, &H846C7CB, &H367CD507, &H2A0C46C7, &HC7629A29, &HDD171046, &H46C73070, &H59015A14, &H1846C791, &HF70E5939, &HD81C46C7, &HC7152FEC, &HB312046, &H46C7FFC0, &H33266724, &H2846C767, &H68581511, &H872C46C7, &HC78EB44A, &H8FA73046, &H46C764F9, &HC2E0D34
    pvAppendBuffer &H3846C7DB, &HBEFA4FA4, &H1D3C46C7, &H5E47B548, &H4C25D, &H1B8E9, &HEC8B5500, &H458B5151, &H80B908, &H56530000, &HC4B08D57, &H8B000000, &HC080&, &H8BE1F700, &H3FA8BD8, &HD7831E, &HFFCA54E8, &H8558BFF, &HCAAFD2B9, &HE9815200, &HCA4000, &H4B8DC103, &HE1835010, &H80B87F, &HC12B0000, &H6A006A50, &H80B800, &H50500000, &H40428D56, &HE9DEE850, &H458DFFFF, &H6A50F8
    pvAppendBuffer &HC5E8006A, &H6A00004E, &HF8458D08, &H875FF50, &H13CE8, &HF8458D00, &HDFA40F50, &HE3C15703, &HA5E85303, &H8B00004E, &H458D085D, &H50086AF8, &H11BE853, &H758B0000, &H73FF560C, &HE833FF04, &H4E88&, &H5008468D, &HFF0C73FF, &H79E80873, &H8D00004E, &HFF501046, &H73FF1473, &H4E6AE810, &H468D0000, &H73FF5018, &H1873FF1C, &H4E5BE8, &H20468D00, &H2473FF50, &HE82073FF, &H4E4C&
    pvAppendBuffer &H5028468D, &HFF2C73FF, &H3DE82873, &H8D00004E, &HFF503046, &H73FF3473, &H4E2EE830, &H468D0000, &H73FF5038, &H3873FF3C, &H4E1FE8, &HC86800, &H6A0000, &H1DF6E853, &HC4830000, &H5B5E5F0C, &HC25DE58B, &H8B550008, &H758B56EC, &HC86808, &H6A0000, &H1DD6E856, &HC4830000, &H806C70C, &HC7F3BCC9, &HE6670446, &H46C76A09, &HCAA73B08, &HC46C784, &HBB67AE85, &H2B1046C7, &HC7FE94F8
    pvAppendBuffer &HF3721446, &H46C73C6E, &H1D36F118, &H1C46C75F, &HA54FF53A, &HD12046C7, &HC7ADE682, &H527F2446, &H46C7510E, &H3E6C1F28, &H2C46C72B, &H9B05688C, &H6B3046C7, &HC7FB41BD, &HD9AB3446, &H46C71F83, &H7E217938, &H3C46C713, &H5BE0CD19, &H4C25D5E, &HEC8B5500, &HFFC8C0E8, &HAFD2B9FF, &HE98100CA, &HCA4000, &H4D8BC103, &HFF505108, &H818D1075, &HC4&, &H680C75FF, &H80&, &H40418D50
    pvAppendBuffer &HE8EFE850, &HC25DFFFF, &H8B55000C, &HE85756EC, &HFFFFC872, &H8B087D8B, &H8D0C8D0F, &H4&, &H8B10FF51, &H8DF08B0F, &H48D0C, &H57510000, &H1CD5E856, &HC4830000, &H5FC68B0C, &H4C25D5E, &HEC8B5500, &H8758B56, &H8B0C75FF, &H8468D0E, &H476FF50, &H8B0451FF, &H4E8B2C56, &H5ED60330, &HC98504EB, &H80490874, &H1080A44, &HC25DF474, &H8B550008, &H84D8BEC, &H3940C033, &H830F7E01
    pvAppendBuffer &H7500813C, &H810C8309, &H13B40FF, &HCFFF17C, &H4C25D81, &HEC8B5500, &H530C458B, &HDB335756, &H831A788D, &H458918C0, &H47B60F0C, &HC88B99FF, &H458BF28B, &HC6D830C, &HB60F08, &HC2A40F99, &HC1F20B08, &HC80B08E0, &HF07B60F, &H9908CEA4, &HC1F87F8D, &HF20B08E1, &HB60FC80B, &HA40F0947, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0A47, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0B47
    pvAppendBuffer &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0C47, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0D47, &HC19908CE, &HF20B08E1, &H458BC80B, &HD80C8908, &H4D87489, &H4FB8343, &HFF6B820F, &H5E5FFFFF, &H8C25D5B, &HEC8B5500, &H530C458B, &HDB335756, &H832A788D, &H458928C0, &H47B60F0C, &HC88B99FF, &H458BF28B, &HC6D830C, &HB60F08, &HC2A40F99, &HC1F20B08, &HC80B08E0, &HF07B60F, &H9908CEA4
    pvAppendBuffer &HC1F87F8D, &HF20B08E1, &HB60FC80B, &HA40F0947, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0A47, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0B47, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0C47, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0D47, &HC19908CE, &HF20B08E1, &H458BC80B, &HD80C8908, &H4D87489, &H6FB8343, &HFF6B820F, &H5E5FFFFF, &H8C25D5B, &HEC8B5500, &H8D60EC83, &H75FFE045
    pvAppendBuffer &H8EE8500C, &H6AFFFFFE, &HE0458D04, &H3C0EE850, &HC0850000, &HC0330474, &H6A567FEB, &HE0458D04, &HC660E850, &H90BEFFFF, &H3000000, &H9EE850C6, &H8300003B, &H147401F8, &H47E8046A, &H3FFFFC6, &H458D50C6, &HE85050E0, &H4A13&, &H458D006A, &H2FE850E0, &H83FFFFC6, &H8D5050C0, &HE850A045, &HFFFFC9C6, &H50A0458D, &HFFC961E8, &H74C085FF, &HEBC03304, &H8758B23, &H51A04D8D, &HC6014E8D
    pvAppendBuffer &HE8510406, &HD2&, &H51C04D8D, &H51214E8D, &HC5E8&, &H40C03300, &H5DE58B5E, &H550008C2, &HEC81EC8B, &H90&, &HFFD0458D, &HE8500C75, &HFFFFFE91, &H458D066A, &H61E850D0, &H8500003B, &H330774C0, &H8AE9C0, &H6A560000, &HD0458D06, &HC5B0E850, &H70BEFFFF, &H3000001, &HEEE850C6, &H8300003A, &H147401F8, &H97E8066A, &H3FFFFC5, &H458D50C6, &HE85050D0, &H4963&, &H458D006A
    pvAppendBuffer &H7FE850D0, &H5FFFFC5, &H110&, &H70858D50, &H50FFFFFF, &HFFCABDE8, &H70858DFF, &H50FFFFFF, &HFFC8D7E8, &H74C085FF, &HEBC03304, &H8758B26, &HFF708D8D, &H8D51FFFF, &H6C6014E, &HAEE85104, &H8D000000, &H8D51A04D, &HE851314E, &HA1&, &H5E40C033, &HC25DE58B, &H8B550008, &H8458BEC, &H8B575653, &H488D0C7D, &H89F63318, &H588D084D, &HF7448A1A, &HB1018807, &HF7448A28, &HFF438806
    pvAppendBuffer &H8BF7048B, &HE804F754, &HFFFFD2A4, &H20B10388, &H8BF7048B, &HE804F754, &HFFFFD294, &H8D014388, &HC8BF85B, &HF7448BF7, &HC1AC0F04, &H18E8C118, &H8B0A4B88, &H448BF70C, &HAC0F04F7, &HE8C110C1, &HB4B8810, &H8BF70C8B, &HF04F744, &HC108C1AC, &H4B8808E8, &HF7048A0C, &H84D8B46, &H8808E983, &H4D890D43, &H4FE8308, &H5E5F8772, &H8C25D5B, &HEC8B5500, &H5308458B, &H7D8B5756, &H28488D0C
    pvAppendBuffer &H4D89F633, &H2A588D08, &H7F7448A, &H28B10188, &H6F7448A, &H8BFF4388, &H548BF704, &HDE804F7, &H88FFFFD2, &H8B20B103, &H548BF704, &HFDE804F7, &H88FFFFD1, &H5B8D0143, &HF70C8BF8, &H4F7448B, &H18C1AC0F, &H8818E8C1, &HC8B0A4B, &HF7448BF7, &HC1AC0F04, &H10E8C110, &H8B0B4B88, &H448BF70C, &HAC0F04F7, &HE8C108C1, &HC4B8808, &H46F7048A, &H83084D8B, &H438808E9, &H84D890D, &H7206FE83
    pvAppendBuffer &H5B5E5F87, &H8C25D, &H83EC8B55, &H8B5320EC, &HC033085D, &HC758B56, &H59066A57, &H8DE44589, &H45C7E87D, &H3E0&, &H8DABF300, &H53500146, &HFFFBE0E8, &H43E80FF, &H468D0F75, &H438D5021, &HCEE85020, &HEBFFFFFB, &H53046A76, &H57207B8D, &H4346E8, &HE8046A00, &HFFFFC3AA, &H5010C083, &H50E0458D, &H5AE85757, &H6A000043, &H57575304, &H42F9E8, &HE8046A00, &HFFFFC38A, &H5010C083
    pvAppendBuffer &HFFC381E8, &H30C083FF, &HE8575750, &H3D7E&, &H18A5E857, &H68A0000, &H1240F8B, &H83C0B60F, &H3B9901E1, &H330675C8, &H74C23BC0, &H57046A12, &HFFC351E8, &H10C083FF, &H20E85750, &H5F000047, &HE58B5B5E, &H8C25D, &H83EC8B55, &H8B5330EC, &HC033085D, &HC758B56, &H590A6A57, &H8DD44589, &H45C7D87D, &H3D0&, &H8DABF300, &H53500146, &HFFFBD0E8, &H43E80FF, &H468D0F75, &H438D5031
    pvAppendBuffer &HBEE85030, &HEBFFFFFB, &H53066A7D, &H57307B8D, &H4286E8, &HE8066A00, &HFFFFC2EA, &HB005&, &H458D5000, &H575750D0, &H4298E8, &H53066A00, &H37E85757, &H6A000042, &HC2C8E806, &HB0BBFFFF, &H3000000, &HBBE850C3, &H5FFFFC2, &HE0&, &HE8575750, &H3CB6&, &H1871E857, &H68A0000, &H1240F8B, &H83C0B60F, &H3B9901E1, &H330675C8, &H74C23BC0, &H57066A11, &HFFC289E8, &H50C303FF
    pvAppendBuffer &H4659E857, &H5E5F0000, &H5DE58B5B, &H550008C2, &HEC81EC8B, &HA0&, &HFF60858D, &H75FFFFFF, &H61E85008, &HFFFFFFFE, &H458D0C75, &H62E850E0, &H6AFFFFFA, &HE0458D00, &H60858D50, &H50FFFFFF, &H50A0458D, &HFFC5E1E8, &HA0458DFF, &H1075FF50, &HFFFD05E8, &HA0458DFF, &HC570E850, &HD8F7FFFF, &H8B40C01B, &HCC25DE5, &HEC8B5500, &HF0EC81, &H858D0000, &HFFFFFF10, &H500875FF, &HFFFEC7E8
    pvAppendBuffer &HC75FFFF, &H50D0458D, &HFFFAB8E8, &H8D006AFF, &H8D50D045, &HFFFF1085, &H858D50FF, &HFFFFFF70, &HC730E850, &H858DFFFF, &HFFFFFF70, &H1075FF50, &HFFFD3CE8, &H70858DFF, &H50FFFFFF, &HFFC53BE8, &H1BD8F7FF, &HE58B40C0, &HCC25D, &H83EC8B55, &H458D40EC, &H75FF56C0, &HA9E85008, &H8BFFFFFD, &H458D0C75, &H468D50C0, &H406C601, &HFC68E850, &H458DFFFF, &H468D50E0, &H5BE85021, &H33FFFFFC
    pvAppendBuffer &H8B5E40C0, &H8C25DE5, &HEC8B5500, &H8D60EC83, &HFF56A045, &HE8500875, &HFFFFFE2C, &H8D0C758B, &H8D50A045, &H6C60146, &HC2E85004, &H8DFFFFFC, &H8D50D045, &HE8503146, &HFFFFFCB5, &H5E40C033, &HC25DE58B, &H8B550008, &H80EC81EC, &H53000000, &H107D8B57, &H535B046A, &H36C2E857, &HC0850000, &HC0330774, &H115E9, &H57535600, &HFFC115E8, &H90BEFF, &HC6030000, &H3653E850, &HF8830000
    pvAppendBuffer &H53107401, &HFFC0FDE8, &H50C603FF, &HCCE85757, &H6A000044, &HEBE85700, &H83FFFFC0, &H8D5050C0, &HE8508045, &HFFFFC482, &H80458D53, &HC0D4E850, &HC603FFFF, &H3617E850, &HF8830000, &H53137401, &HFFC0C1E8, &H50C603FF, &H5080458D, &H448DE850, &H8D530000, &HE8508045, &H3644&, &H774C085, &H96E9C033, &H8B000000, &H458D1475, &HE8565080, &HFFFFFB66, &H8D0875FF, &HE850C045, &HFFFFF895
    pvAppendBuffer &HFFC081E8, &H9005FF, &H8D500000, &H8D50C045, &H8D508045, &HE850E045, &H3CE2&, &H8D0C75FF, &HE850C045, &HFFFFF86D, &HC058E853, &H9005FFFF, &H50000000, &H50E0458D, &H50C0458D, &H50E0458D, &H3A49E8, &H3BE85300, &HBBFFFFC0, &H90&, &H5750C303, &H3A73E857, &H27E80000, &H3FFFFC0, &HE04D8DC3, &H51515750, &H3C91E8, &HE04D8D00, &H204E8D51, &HFAE0E851, &HC033FFFF, &H5B5F5E40
    pvAppendBuffer &HC25DE58B, &H8B550010, &HC0EC81EC, &H53000000, &H107D8B57, &H535B066A, &H3582E857, &HC0850000, &HC0330774, &H129E9, &H57535600, &HFFBFD5E8, &H170BEFF, &HC6030000, &H3513E850, &HF8830000, &H53107401, &HFFBFBDE8, &H50C603FF, &H8CE85757, &H6A000043, &HABE85700, &H5FFFFBF, &H110&, &H40858D50, &H50FFFFFF, &HFFC4E9E8, &H858D53FF, &HFFFFFF40, &HBF8CE850, &HC603FFFF, &H34CFE850
    pvAppendBuffer &HF8830000, &H53167401, &HFFBF79E8, &H50C603FF, &HFF40858D, &H5050FFFF, &H4342E8, &H858D5300, &HFFFFFF40, &H34F6E850, &HC0850000, &HC0330774, &H9CE9&, &H14758B00, &HFF40858D, &H5650FFFF, &HFFFAACE8, &H875FFFF, &H50A0458D, &HFFF7F4E8, &HBF30E8FF, &H7005FFFF, &H50000001, &H50A0458D, &HFF40858D, &H8D50FFFF, &HE850D045, &H3CF5&, &H8D0C75FF, &HE850A045, &HFFFFF7C9, &HBF04E853
    pvAppendBuffer &H7005FFFF, &H50000001, &H50D0458D, &H50A0458D, &H50D0458D, &H38F5E8, &HE7E85300, &HBBFFFFBE, &H170&, &H5750C303, &H391FE857, &HD3E80000, &H3FFFFBE, &HD04D8DC3, &H51515750, &H3CA4E8, &HD04D8D00, &H304E8D51, &HFA23E851, &HC033FFFF, &H5B5F5E40, &HC25DE58B, &H8B550010, &HB0EC81EC, &H8D000001, &HFFFE5085, &HFF5653FF, &HE8500875, &HFFFFFA94, &H8D10758B, &HFFFF3085, &HE85056FF
    pvAppendBuffer &HFFFFF691, &H5020468D, &H5090458D, &HFFF684E8, &H5E046AFF, &H30858D56, &H50FFFFFF, &H33FFE8, &HFC08500, &H37485, &H458D5600, &HEDE85090, &H85000033, &H62850FC0, &H56000003, &HFF30858D, &HE850FFFF, &HFFFFBE3E, &H90BB&, &H50C30300, &H337CE8, &H1F88300, &H33F850F, &H8D560000, &HE8509045, &HFFFFBE1E, &HE850C303, &H3361&, &HF01F883, &H32485, &HE8565700, &HFFFFBE06
    pvAppendBuffer &H8D50C303, &H8D509045, &HE850E045, &H383D&, &H8D0C75FF, &HFFFF5085, &HF6E850FF, &HE8FFFFF5, &HFFFFBDE2, &H8D50C303, &H8D50E045, &HFFFF5085, &HE85050FF, &H3A46&, &HFFBDC9E8, &H50C303FF, &H50E0458D, &HFF30858D, &H8D50FFFF, &HFFFF1085, &H27E850FF, &H5600003A, &HFE50858D, &H8D50FFFF, &HFFFEB085, &H96E850FF, &H5600003F, &HFE70858D, &H8D50FFFF, &HFFFED085, &H82E850FF, &H5600003F
    pvAppendBuffer &HFFBD81E8, &H50C083FF, &H70858D50, &H50FFFFFF, &H3F6CE8, &H6BE85600, &H83FFFFBD, &H8D5070C0, &HE8509045, &H3F59&, &HBD58E856, &HC083FFFF, &H858D5010, &HFFFFFF70, &HB0858D50, &H50FFFFFE, &H50E0458D, &H3CFCE8, &HD0858D00, &H50FFFFFE, &HFEB0858D, &H8D50FFFF, &H8D509045, &HFFFF7085, &H2CE850FF, &H56FFFFC4, &HFFBD19E8, &H10C083FF, &HE0458D50, &H52E85050, &H8D000037, &H8D50E045
    pvAppendBuffer &HFFFED085, &H858D50FF, &HFFFFFEB0, &HCC7AE850, &H6583FFFF, &HEBE800D0, &H83FFFFBC, &H458950C0, &H50858DD4, &H89FFFFFE, &H858DD845, &HFFFFFEB0, &H8DDC4589, &HFFFF1085, &HE85056FF, &H3E0A&, &H858DD88B, &HFFFFFF50, &HFBE85056, &H3B00003D, &HD8470FC3, &HFF50858D, &H738DFFFF, &HE85056FF, &H410A&, &H574C20B, &HEB47FF33, &H56FF3302, &HFF10858D, &HE850FFFF, &H40F2&, &H574C20B
    pvAppendBuffer &HEB5E026A, &HBF63302, &HB0458DF7, &H748B046A, &H5056D0B5, &H3E68E8, &H8D046A00, &H8D502046, &HFFFEF085, &H56E850FF, &H6A00003E, &HE0458D04, &H316EE850, &H65830000, &H738D00E4, &HE045C7FE, &H1&, &H880FF685, &HDD&, &H50E0458D, &HFEF0858D, &H8D50FFFF, &HE850B045, &HFFFFBC4A, &H50858D56, &H50FFFFFF, &H4081E8, &H74C20B00, &H47FF3305, &HFF3302EB, &H10858D56, &H50FFFFFF
    pvAppendBuffer &H4069E8, &H74C20B00, &H58026A05, &HC03302EB, &H7C8BF80B, &HFF85D0BD, &H82840F, &H46A0000, &H70858D57, &H50FFFFFF, &H3DD4E8, &H8D046A00, &H8D502047, &HE8509045, &H3DC5&, &H50E0458D, &H5090458D, &HFF70858D, &HE850FFFF, &HFFFFCB3C, &HAFE8046A, &H83FFFFBB, &H8D5010C0, &HFFFF7085, &H458D50FF, &H858D50B0, &HFFFFFE90, &H3B53E850, &H858D0000, &HFFFFFEF0, &HB0458D50, &H90458D50
    pvAppendBuffer &H70858D50, &H50FFFFFF, &HFFC286E8, &H8D046AFF, &HFFFE9085, &H458D50FF, &HE85050E0, &H3ACE&, &HF01EE83, &HFFFF2389, &HE8046AFF, &HFFFFBB56, &H5010C083, &H50E0458D, &H358FE850, &H458D0000, &H858D50E0, &HFFFFFEF0, &HB0458D50, &HCABAE850, &H46AFFFF, &H50B0458D, &HFFBB29E8, &H90BEFF, &HC6030000, &H3067E850, &H835F0000, &H147401F8, &HFE8046A, &H3FFFFBB, &H458D50C6, &HE85050B0
    pvAppendBuffer &H3EDB&, &H858D046A, &HFFFFFF30, &HB0458D50, &H303BE850, &HD8F70000, &HEB40C01B, &H5EC03302, &H5DE58B5B, &H55000CC2, &HEC81EC8B, &H280&, &HFD80858D, &H5653FFFF, &H875FF57, &HF786E850, &H758BFFFF, &HD0858D10, &H56FFFFFE, &HF373E850, &H468DFFFF, &H858D5030, &HFFFFFF60, &HF363E850, &H66AFFFF, &H858D575F, &HFFFFFED0, &H302EE850, &HC0850000, &H392850F, &H8D570000, &HFFFF6085
    pvAppendBuffer &H19E850FF, &H85000030, &H7D850FC0, &H57000003, &HFED0858D, &HE850FFFF, &HFFFFBA6A, &H170BB, &H50C30300, &H2FA8E8, &H1F88300, &H35A850F, &H8D570000, &HFFFF6085, &H47E850FF, &H3FFFFBA, &H8AE850C3, &H8300002F, &H850F01F8, &H33C&, &HBA30E857, &HC303FFFF, &H60858D50, &H50FFFFFF, &H50C0458D, &H3464E8, &HC75FF00, &HFF00858D, &HE850FFFF, &HFFFFF2CD, &HFFBA09E8, &H50C303FF
    pvAppendBuffer &H50C0458D, &HFF00858D, &H5050FFFF, &H37D4E8, &HB9F0E800, &HC303FFFF, &HC0458D50, &HD0858D50, &H50FFFFFE, &HFEA0858D, &HE850FFFF, &H37B5&, &H80858D57, &H50FFFFFD, &HFE10858D, &HE850FFFF, &H3BBD&, &HB0858D57, &H50FFFFFD, &HFE40858D, &HE850FFFF, &H3BA9&, &HB9A8E857, &H738DFFFF, &H50C603A0, &HFF30858D, &HE850FFFF, &H3B91&, &HB990E857, &H4005FFFF, &H50000001, &HFF60858D
    pvAppendBuffer &HE850FFFF, &H3B79&, &HB978E857, &H5E8DFFFF, &H50C303A0, &HFF30858D, &H8D50FFFF, &HFFFE1085, &H458D50FF, &H1AE850C0, &H8D000039, &HFFFE4085, &H858D50FF, &HFFFFFE10, &H60858D50, &H50FFFFFF, &HFF30858D, &HE850FFFF, &HFFFFC157, &HB934E857, &HC303FFFF, &HC0458D50, &H6EE85050, &H8D000033, &H8D50C045, &HFFFE4085, &H858D50FF, &HFFFFFE10, &HC8E1E850, &H6583FFFF, &H7E800F0, &H3FFFFB9
    pvAppendBuffer &HF44589C6, &HFD80858D, &H4589FFFF, &H10858DF8, &H89FFFFFE, &H858DFC45, &HFFFFFEA0, &H27E85057, &H8B00003A, &H858DD8, &H57FFFFFF, &H3A18E850, &HC33B0000, &H8DD8470F, &HFFFF0085, &HFF738DFF, &H27E85056, &HB00003D, &H330574C2, &H2EB47FF, &H8D56FF33, &HFFFEA085, &HFE850FF, &HB00003D, &H6A0574C2, &H2EB5E02, &HF70BF633, &H6A90458D, &HB5748B06, &HE85056F0, &H3A85&, &H468D066A
    pvAppendBuffer &H858D5030, &HFFFFFE70, &H3A73E850, &H66A0000, &H50C0458D, &H2D8BE8, &HC4658300, &HFE738D00, &H1C045C7, &H85000000, &HE8880FF6, &H8D000000, &H8D50C045, &HFFFE7085, &H458D50FF, &HEFE85090, &H56FFFFB9, &HFF00858D, &HE850FFFF, &H3C9E&, &H574C20B, &HEB47FF33, &H56FF3302, &HFEA0858D, &HE850FFFF, &H3C86&, &H574C20B, &HEB58026A, &HBC03302, &HBD7C8BF8, &HFFF85F0, &H8D84&
    pvAppendBuffer &H57066A00, &HFF30858D, &HE850FFFF, &H39F1&, &H478D066A, &H858D5030, &HFFFFFF60, &H39DFE850, &H458D0000, &H858D50C0, &HFFFFFF60, &H30858D50, &H50FFFFFF, &HFFC79EE8, &HE8066AFF, &HFFFFB7C6, &HB005&, &H858D5000, &HFFFFFF30, &H90458D50, &HE0858D50, &H50FFFFFD, &H3768E8, &H70858D00, &H50FFFFFE, &H5090458D, &HFF60858D, &H8D50FFFF, &HFFFF3085, &HA8E850FF, &H6AFFFFBF, &HE0858D06
    pvAppendBuffer &H50FFFFFD, &H50C0458D, &H36E0E850, &HEE830000, &H18890F01, &H6AFFFFFF, &HB768E806, &HB005FFFF, &H50000000, &H50C0458D, &H319FE850, &H458D0000, &H858D50C0, &HFFFFFE70, &H90458D50, &HC715E850, &H66AFFFF, &H5090458D, &HFFB739E8, &H170BEFF, &HC6030000, &H2C77E850, &HF8830000, &H6A147401, &HB720E806, &HC603FFFF, &H90458D50, &HECE85050, &H6A00003A, &HD0858D06, &H50FFFFFE, &H5090458D
    pvAppendBuffer &H2C4CE8, &H1BD8F700, &H2EB40C0, &H5E5FC033, &H5DE58B5B, &H55000CC2, &H8B56EC8B, &H68B0875, &H485048D, &H50000000, &H2A05E856, &HD3E80000, &H56FFFFB6, &H5E0850FF, &H4C25D, &H8BEC8B55, &HC18B084D, &H8107E8C1, &H7F7F7FE1, &H10125FF, &HC9030101, &H331BC06B, &H4C25DC1, &HEC8B5500, &HFFB6B4E8, &H8A98B9FF, &HE98100CA, &HCA4000, &H4D8BC103, &HFF505108, &H418D1075, &HC75FF30
    pvAppendBuffer &H8D50106A, &HE8502041, &HFFFFD6E9, &HCC25D, &H8BEC8B55, &H458B084D, &H75FF5010, &H3841010C, &H3C518351, &HFFB3E800, &HC25DFFFF, &H8B55000C, &H758B56EC, &H487E8308, &H560D7501, &H20E8&, &H4846C700, &H2&, &H110458B, &HFF504046, &H56830C75, &HE8560044, &HFFFFFF81, &HCC25D5E, &HEC8B5500, &H8758B56, &H85304E8B, &H6A2474C9, &HC12B5810, &H20468D50, &H6AC103, &HAAEE850
    pvAppendBuffer &HC4830000, &H20468D0C, &H9E85650, &H83000000, &H5E003066, &H4C25D, &H83EC8B55, &H458D10EC, &H505756F0, &HE80C75FF, &HFFFFE5F3, &H8D087D8B, &H778DF045, &H50565610, &HFFE548E8, &H565756FF, &H5F4C57FF, &H5DE58B5E, &H550008C2, &H5151EC8B, &H8758B56, &H1487E83, &H7E830674, &HA750248, &HFF7BE856, &H6683FFFF, &H4E8B0048, &HF8458D38, &H3C468B50, &H3C8A40F, &H5003E1C1, &H3A4AE851
    pvAppendBuffer &H86A0000, &H50F8458D, &HFECFE856, &H4E8BFFFF, &HF8458D40, &H44468B50, &H3C8A40F, &H5003E1C1, &H3A26E851, &H86A0000, &H50F8458D, &HFEABE856, &H75FFFFFF, &H10468D0C, &HE713E850, &H8B5EFFFF, &H8C25DE5, &HEC8B5500, &H5310EC83, &H7D8B5756, &H6A506A08, &HD5E85700, &H83000009, &HFF570CC4, &H38E80C75, &H33FFFFE5, &H40C933C0, &H48478953, &HF38BA20F, &HF05D8D5B, &H73890389, &H84B8904
    pvAppendBuffer &HE80C5389, &HFFFFB511, &H2F845F6, &HCA715AB9, &HB9057500, &HCA70DC, &H4000E981, &HC10300CA, &H5F4C4789, &HE58B5B5E, &H8C25D, &H8BEC8B55, &H33561455, &H1EA83F6, &H458B3078, &H5D8B530C, &HD82B5708, &H29903C8D, &HC8B1045, &H3C0333B, &H3C013CE, &HD0830F, &H8B01EA83, &H10458BF0, &H8D380C89, &HE279FC7F, &HC68B5B5F, &H10C25D5E, &HEC8B5500, &H758B5653, &H99C68B10, &H571FE283
    pvAppendBuffer &HC1023C8D, &HE68105FF, &H8000001F, &H834E0579, &H8B46E0CE, &HDE8B0C55, &H2B59206A, &HD3C28BCE, &H8BDBF7E8, &HD3DB1BCE, &H33D823E2, &H105D89C9, &H8B085D8B, &H54011075, &HC91304BB, &H475F685, &H2C74C985, &H8BB448B, &HC103D233, &HC603D213, &H8BB4489, &H8300D283, &HD28503C7, &H348D1274, &H1C033BB, &H4768D16, &HD08BC013, &HF175D285, &H5D5B5E5F, &H55000CC2, &HEC83EC8B, &H14558B38
    pvAppendBuffer &HC5D8B53, &H5756C38B, &HC22BFF33, &HFEC4589, &H21688, &HEC558B00, &H458BC88B, &H5E1C108, &H8BFC4D89, &HF685B834, &H83470C75, &H4D8920E9, &H1E7E9FC, &HE8560000, &HFFFFCA2D, &HCA8BD08B, &HD3D05589, &H7ED285E6, &H1478D17, &H107DC33B, &H6A08458B, &HCA2B5920, &H4B8548B, &HF20BEAD3, &H8B205D8B, &H1C65F7C6, &H83FC458B, &H452BE1C3, &H3F28BD0, &HF07589D8, &H79F45D89, &HE0FB831F
    pvAppendBuffer &H19F8E0F, &HDBF70000, &HDB33CB8B, &H7589EED3, &HF45D89F0, &H840FF685, &H189&, &H1F25C38B, &H79800000, &HC8834805, &H4D8B40E0, &HD845890C, &H8399C38B, &HC2031FE2, &HC114558B, &HC82B05F8, &H2BD8458B, &HFC085CA, &H8985&, &HD8452100, &H89CA034A, &H4D89E855, &HC0570FE4, &H45130F66, &HF045C7C8, &H1&, &H8C0FCF3B, &H111&, &H8B08458B, &H48D105D, &HC84D8B88, &H8BF84589
    pvAppendBuffer &H4589CC45, &H79D285E0, &HEBC03304, &H93048B03, &HC103E6F7, &H5513D0F7, &H89C933E0, &H558BD055, &H130203F8, &HF04503C9, &H4D130289, &HE8558BD8, &H4F86D83, &HE065834A, &HD8658300, &HF04D8900, &H49E44D8B, &H89E85589, &HCF3BE44D, &H7DD04D8B, &HA8E9B4, &H65830000, &H34A00E8, &HD45589CA, &H3BD04D89, &H968C0FCF, &H6A000000, &HD82B5B20, &H1E445C7, &H8B000000, &H570F0845, &H130F66C0
    pvAppendBuffer &H5D89DC45, &HD85D8BCC, &H8B88048D, &H4589E04D, &HDC458BF8, &H85E04589, &H330479D2, &H8B06EBC0, &H48B1045, &H8BE6F790, &HE07503F0, &HCB8BD113, &H33E05589, &HD3D68BC0, &HE8550BE2, &HF7F84D8B, &H131103D2, &HE45503C0, &H8300D083, &H8900D865, &HC18BE445, &H89CC4D8B, &H8BD68B10, &HE883F075, &H8BEAD304, &H5589D04D, &HD4558BE8, &HF845894A, &HD4558949, &H3BD04D89, &HD84D8BCF, &H5D8B997D
    pvAppendBuffer &H187D83F4, &H530A7400, &H1875FF56, &HFFFD60E8, &HEC558BFF, &H8BFC4D8B, &H5D8B0845, &HFFA3B0C, &HFFFDFE8E, &HC5D8BFF, &H8B14558B, &HC933087D, &H247EDB85, &H2B10458B, &H90348DD3, &H850A048D, &H330479C0, &H8B02EBC0, &H8F043906, &H8776772, &H4C68341, &HE47CCB3B, &HC6583, &HC7FF538D, &H10845, &HD2850000, &H458B3978, &H2BF28B10, &H147503F3, &H85B01C8D, &H330479F6, &H8B02EBC9
    pvAppendBuffer &HF7C0330B, &H970C03D1, &H4D03C013, &H970C8908, &H4E0C4513, &HC6583, &H8304EB83, &H458901EA, &H83D47908, &H7400187D, &H6A006A0C, &H1875FF01, &HFFFCBCE8, &H5B5E5FFF, &HC25DE58B, &H8B55001C, &H28EC83EC, &H145D8B53, &HFB835756, &HB88E0F32, &H8B000001, &HC38B1875, &H8BC22B99, &H56C38BF8, &HFF57FFD1, &HC72B1075, &HFFF07D89, &H45890C75, &H875FFFC, &HFFFFC5E8, &H10458BFF, &HFC758B56
    pvAppendBuffer &HF8148D56, &H520C458B, &H8DD85589, &H458BB80C, &H4D895108, &HB8048DDC, &HE4458950, &HFFFF9DE8, &H184D8BFF, &H3301468D, &H8D046AF6, &H7A8D8114, &H89328904, &H518DF455, &H89378904, &H3289EC55, &H89F0558B, &H3189E07D, &H7ED2855F, &H2BC00346, &H81048DC2, &H8BE84589, &H2B40FC45, &H81048DC2, &H890C4D8B, &H458BF845, &H89C12B08, &H558B0C55, &H84589E8, &H8BF85D8B, &H7D010804, &H8B0389F8
    pvAppendBuffer &H8BCF0301, &H289145D, &H6D83D703, &H458B010C, &HFFE17508, &H458BFC75, &H75FF50EC, &H89E850E4, &H8BFFFFFB, &H75FF184D, &H8B0189FC, &HFF50E045, &HE850DC75, &HFFFFFB74, &H8BF44D8B, &H8942FC55, &H184D8B01, &H89D1048D, &HC28B0C45, &H304E0C1, &HFF5250C1, &H75FF0C75, &HDFE851F4, &H8BFFFFFE, &H558B184D, &H8418DF0, &H890C7189, &H3089DC45, &H8912048D, &H31890471, &H85084589, &H8B1E7EC0
    pvAppendBuffer &H5D8BFC45, &HC22B4010, &H8D08558B, &H48BC10C, &H18946B3, &HF23BCF03, &H5D8BF47C, &HFC758B14, &H5036048D, &H50DC458B, &H50D875FF, &HFFFAFBE8, &HEC4D8BFF, &H468D0189, &H184D8B01, &H8B00348D, &H50560C45, &H3E85051, &H8B000002, &H8D56FC4D, &HD82B0141, &H310458B, &H8DD92BDB, &HFF509804, &HE8500C75, &HFFFFFAC4, &HF685F08B, &H87840F, &H558B0000, &HFCC28310, &H339A148D, &H133201C0
    pvAppendBuffer &H8BD72BC0, &H75F685F0, &H8B6EEBF2, &HC8D1055, &H7EC9851B, &H8BC03306, &H8BABF3FA, &HC8D0845, &H76C83B98, &HC758B54, &H83DA148D, &H46AFCC2, &H8D105589, &H75899E34, &H65835F14, &HCF2B0018, &H753BDA8B, &H8B21760C, &HF7F72B01, &H18450326, &H100D283, &HD28303, &H5589DF2B, &HC753B18, &H458BE577, &H10558B08, &H2B18758B, &H8B3389D7, &H55891475, &H77C83B10, &H5B5E5FC1, &HC25DE58B
    pvAppendBuffer &H8B550014, &H10EC83EC, &H145D8B53, &HFB835756, &HA48E0F32, &H8B000000, &HC38B0C7D, &H9908758B, &HD38BC22B, &H458BC88B, &H2BF9D118, &HFC4D89D1, &H8DF85589, &H5253D81C, &H50C8048D, &H508F048D, &H508E048D, &HFFFD75E8, &HFC4D8BFF, &H5318458B, &H88048D51, &HF8458B50, &H5087048D, &HFFA0E856, &HFF53FFFF, &H5D8BFC75, &H8B575318, &H48DF87D, &H8BE850BE, &H8BFFFFFF, &HFF85FC75, &H458B1B7E
    pvAppendBuffer &HB00C8D10, &H8D77048D, &H28B8314, &H8904528D, &H4498D01, &H7501EF83, &H8D5356F1, &H5350B304, &HFFF987E8, &H14458BFF, &H1075FF56, &H5083048D, &HF976E853, &H8AE9FFFF, &H85000000, &H8B097EDB, &HCB8B107D, &HABF3C033, &H8B08458B, &HC8D107D, &H9F348D98, &H3B147589, &H8B6A76C8, &H48D0C45, &H85D8B98, &H89FCC083, &HE983FC45, &HC758904, &HFF84589, &HF66C057, &H3BF04513, &H8B3C76F7
    pvAppendBuffer &HDE8BF445, &H89F8758B, &H458B1845, &H104589F0, &HEB83068B, &H8D21F704, &H303FC76, &H300D283, &H3891045, &H83185513, &H89001865, &HDF3B1055, &H758BDE77, &H85D8B14, &H83FC458B, &H758904EE, &H77CB3B14, &H5B5E5FA5, &HC25DE58B, &H8B550014, &H558B51EC, &HFC658314, &H1EA8300, &H11445C7, &H78000000, &H8458B39, &H758B5653, &H7D8B570C, &H901C8D10, &HF82BF02B, &H331E0C8B, &H3D1F7C0
    pvAppendBuffer &H3C0130B, &HC89144D, &HFC5B8D1F, &H83FC4513, &H8300FC65, &H458901EA, &H5FDD7914, &HE58B5B5E, &H10C25D, &H81EC8B55, &H80EC&, &H758B5600, &H206A570C, &H807D8D59, &HFDBEA5F3, &H8D000000, &H50508045, &H2088E8, &H2FE8300, &HFE831274, &HFF0D7404, &H458D0C75, &HE8505080, &HE83&, &H7901EE83, &H87D8BDA, &H6A80758D, &HA5F35920, &HE58B5E5F, &H8C25D, &H53EC8B55, &H75FF5756
    pvAppendBuffer &HF636E808, &HD88BFFFF, &HF62EE853, &HD08BFFFF, &HF626E852, &HF88BFFFF, &H8B087D33, &H33C78BF7, &H8CFC1C3, &HC0C1F233, &HC1CE8B08, &HC13310C9, &HC633C733, &H33C3335F, &H5B5E0845, &H4C25D, &H56EC8B55, &HFF08758B, &HFFAAE836, &H76FFFFFF, &HE8068904, &HFFFFFFA0, &H890876FF, &H95E80446, &HFFFFFFFF, &H46890C76, &HFF8AE808, &H4689FFFF, &HC25D5E0C, &H8B550004, &H5D8B53EC, &HF575608
    pvAppendBuffer &HF077BB6, &HF0243B6, &HF0B73B6, &HC10F53B6, &HF80B08E7, &H34BB60F, &HD43B60F, &HB08E7C1, &H8E6C1F8, &H843B60F, &HB08E7C1, &H8E2C1F8, &H643B60F, &HE1C1F00B, &H43B60F08, &H8E6C101, &HB60FF00B, &HE6C10C43, &HFF00B08, &HB0A43B6, &H43B60FD0, &H8E2C105, &HB60FD00B, &H8E2C103, &HB60FD00B, &HC80B0E43, &HF0C5389, &HC10943B6, &HC80B08E1, &HF087389, &H890443B6, &HE1C1047B
    pvAppendBuffer &HC80B5F08, &H5B0B895E, &H4C25D, &H56EC8B55, &HABF0E857, &H758BFFFF, &H693BF08, &HC7030000, &HE836FF50, &H1FD3&, &HD7E80689, &H3FFFFAB, &H76FF50C7, &H1FC1E804, &H46890000, &HABC4E804, &HC703FFFF, &H876FF50, &H1FAEE8, &H8468900, &HFFABB1E8, &H50C703FF, &HE80C76FF, &H1F9B&, &HC46895F, &H4C25D5E, &HEC8B5500, &HC7D83, &HFF561574, &H758B0C75, &H56006A08, &H2FE8&
    pvAppendBuffer &H83068A00, &H5D5E0CC4, &H550008C2, &H558BEC8B, &H8458B10, &H85F08B56, &H571274D2, &H2B0C7D8B, &H370C8AF8, &H83460E88, &HF57501EA, &HC35D5E5F, &H8BEC8B55, &HC985104D, &HB60F1F74, &H8B560C45, &H1C069F1, &H57010101, &HC1087D8B, &HABF302E9, &HE183CE8B, &H5FAAF303, &H8458B5E, &H8B55C35D, &H758B56EC, &H51E85608, &H8BFFFFF4, &H33CE8BD0, &H10C9C1D6, &HC108C2C1, &HD13308CE, &HC233D633
    pvAppendBuffer &H4C25D5E, &HEC8B5500, &H8758B56, &HCBE836FF, &HFFFFFFFF, &H6890476, &HFFFFC1E8, &H876FFFF, &HE8044689, &HFFFFFFB6, &H890C76FF, &HABE80846, &H89FFFFFF, &H5D5E0C46, &H550004C2, &HEC83EC8B, &HC4658340, &H21C03300, &H5653E445, &H59066A57, &H7D8DDB33, &H43066AC8, &HF3C05D89, &H46A59AB, &H89E87D8D, &HABF3E05D, &H458D575F, &H93E850C0, &H83FFFFAA, &H8D5010C0, &HE850C045, &H1F0D&
    pvAppendBuffer &HC0458D57, &H2BC0E850, &H708D0000, &H8D26EBFF, &HE850E045, &H2A03&, &HC0458D56, &H2ECCE850, &HC20B0000, &HFF570E74, &H458D0875, &HE85050E0, &H29BA&, &HE0458D4E, &HF33B5057, &H75FFD177, &H2C3BE808, &H5E5F0000, &H5DE58B5B, &H550004C2, &HEC83EC8B, &HA4658360, &H21C03300, &H5653D445, &H590A6A57, &H7D8DDB33, &H430A6AA8, &HF3A05D89, &H66A59AB, &H89D87D8D, &HABF3D05D, &H458D575F
    pvAppendBuffer &HFFE850A0, &H5FFFFA9, &HB0&, &HA0458D50, &H1E77E850, &H8D570000, &HE850A045, &H2B2A&, &HEBFF708D, &HD0458D26, &H296DE850, &H8D560000, &HE850A045, &H2E36&, &HE74C20B, &H875FF57, &H50D0458D, &H2924E850, &H8D4E0000, &H5057D045, &HD177F33B, &HE80875FF, &H2BA5&, &H8B5B5E5F, &H4C25DE5, &HEC8B5500, &H5310EC83, &H75FF5756, &HE114E80C, &H75FFFFFF, &HFC458908, &HFFE109E8
    pvAppendBuffer &HE8F08BFF, &HFFFFA97E, &H79405, &HF7E85000, &H89FFFFE0, &H6BE80845, &H5FFFFA9, &H798&, &HE0E4E850, &HFF33FFFF, &H8947D88B, &H53E8F07D, &HE9FFFFA9, &HA6&, &HFFA949E8, &H79405FF, &H56500000, &HFFBDA0E8, &HFC085FF, &H10B84, &HE836FF00, &HC23&, &H8BF84589, &H30FFFC45, &HC16E8, &H45895000, &HF8458BF4, &H75FF5650, &HB9F6E8FC, &H4D8BFFFF, &H763939F8, &H83018B0F
    pvAppendBuffer &H7500813C, &H1894807, &HF177C73B, &H39F44D8B, &H8B0F7639, &H813C8301, &H48077500, &HC73B0189, &H75FFF177, &HF1EDE8FC, &H458BFFFF, &HFC7589F8, &H5608758B, &HF475FF53, &H89F84589, &H2EE8085D, &HF7FFFFBB, &HD88BF05D, &HF1C9E856, &H75FFFFFF, &HF1C1E8F4, &HABE8FFFF, &H8BFFFFA8, &H9805F875, &H50000007, &HBCFFE856, &HC085FFFF, &HFF46850F, &HE856FFFF, &HFFFFF19F, &HE8FC75FF, &HFFFFF197
    pvAppendBuffer &HE80875FF, &HFFFFF18F, &HF07D83, &HB58D0F, &H458B0000, &HE830FF0C, &HB5F&, &H7D89F08B, &H89C933FC, &H3E39F875, &H8F8C0F, &H458B0000, &H4568D0C, &H8BF04589, &HF05D29C7, &H29085D89, &H758B0875, &H8B063B0C, &H307F0875, &H3F0458B, &H10048BC6, &H8BF44589, &H24EBFC45, &HF135E856, &H75FFFFFF, &HF12DE8FC, &H75FFFFFF, &HF125E808, &HE853FFFF, &HFFFFF11F, &H4DEBC033, &HF46583
    pvAppendBuffer &H57F033B, &HEB16048B, &H8BC03302, &HF02BF475, &HF12BD0F7, &HC9853289, &HF03B0774, &HEB41C91B, &H1BC63B06, &H8BD9F7C9, &HF685FC45, &HFF8758B, &H8340F845, &H458904C2, &H7E063BFC, &H3E895385, &HFFF0D2E8, &H8BDE8BFF, &H5B5E5FC3, &HC25DE58B, &H8B550008, &H20EC83EC, &H758B5653, &H1E8B5710, &HE7C1FB8B, &HE47D8902, &HFFA799E8, &H10FF57FF, &HC933D08B, &H85105589, &H8B0F7EDB, &H8BC12B06
    pvAppendBuffer &H4898604, &HCB3B418A, &H7D8BF17C, &HC558B08, &H78BCA8B, &H470F023B, &H8D318BCF, &HC33B3604, &HC38B0A7F, &H8BC22B99, &H46FED1F0, &HE0C1C68B, &HE0458902, &HFFA74DE8, &HC1CE8BFF, &HFF5102E1, &H89D68B10, &H172BFC45, &HC7ED285, &H8BFC7D8B, &HF3C033CA, &H87D8BAB, &HC933078B, &H1A7EC085, &H8DFC7D8B, &H7D8B9714, &H41C12B08, &H8987048B, &H4528D02, &HC83B078B, &H7E8EF7C, &H8BFFFFA7
    pvAppendBuffer &H2E1C1CE, &H8B10FF51, &HCE8B0C55, &H7D89F88B, &H890A2BF8, &HC985084D, &HC033077E, &H4D8BABF3, &H33028B08, &H7EC085FF, &HF8558B1A, &H8B8A0C8D, &HC72B0C55, &H82048B47, &H498D0189, &H3B028B04, &H8BEF7CF8, &H3E7C1FE, &HE8E87D89, &HFFFFA6B6, &H5610FF57, &HE8084589, &H97B&, &HE7C1F88B, &HEC7D8902, &HFFA69DE8, &H10FF57FF, &H89107D8B, &H3F8B0C45, &HBCEFE857, &H4589FFFF, &H74C085F4
    pvAppendBuffer &HD3C88B18, &H1FB83E7, &H206A0F7E, &H8BC82B59, &H408B1045, &HBE8D304, &HC75FFF8, &H875FF56, &HFFF875FF, &H13E8FC75, &HFFFFFFF5, &HF603F475, &HF0758957, &H100FE8, &H6A5000, &H1075FF53, &HFF57FE8B, &H2CE80875, &H3BFFFFF2, &HFC78BDF, &H8950C34C, &H20E8F045, &H8B000009, &HD233F05D, &HDB85F08B, &H458B197E, &H8DFB2B08, &HE8BB83C, &HCA2B078B, &H47F8D42, &H3B8E0489, &H83EF7CD3
    pvAppendBuffer &H1076013E, &H3C83068B, &H8750086, &H83068948, &HF07701F8, &H8BEC75FF, &HE8530C5D, &H1913&, &HFFA5E1E8, &H50FF53FF, &HE875FF08, &H53085D8B, &H18FEE8, &HA5CCE800, &HFF53FFFF, &H75FF0850, &H105D8BE4, &H18E9E853, &HB7E80000, &H53FFFFA5, &H8B0850FF, &H5D8BE07D, &HE85357FC, &H18D3&, &HFFA5A1E8, &H50FF53FF, &HF85D8B08, &HC0E85357, &HE8000018, &HFFFFA58E, &H850FF53, &H5EC68B5F
    pvAppendBuffer &H5DE58B5B, &H55000CC2, &HEC83EC8B, &H5D8B532C, &H4438D10, &H890100F6, &H1175E045, &HC75FF53, &HE80875FF, &H36D&, &H361E9, &H53575600, &HE80875FF, &HFFFFB76B, &HF08B1B8B, &H5D89CB8B, &H5E1C1DC, &HBBCBE851, &HF88BFFFF, &H1075FF57, &HFFFB90E8, &H1075FFFF, &H57084589, &HFD70E856, &H8956FFFF, &H2CE8F045, &HFFFFFFEE, &HE8571075, &HFFFFB72F, &HFC458957, &HFFEE1AE8, &HC1FB8BFF
    pvAppendBuffer &H7D8902E7, &HA4FCE8D8, &HFF57FFFF, &HEC458910, &H197EDB85, &H8DE04D8B, &HD003FC57, &H18BF38B, &H8904498D, &HFC528D02, &H7501EE83, &HA4D4E8F1, &HFF57FFFF, &H33F08B10, &HF47589C9, &H2D7EDB85, &H8308558B, &HC283FCC7, &H8BF70304, &HF3B087D, &H28B047D, &HC03302EB, &H83410689, &HEE8304C2, &H7CCB3B04, &HF4758BE9, &HE7C1FB8B, &H8458B02, &HEDA1E850, &H8BE8FFFF, &H57FFFFA4, &HD08B10FF
    pvAppendBuffer &H85E85589, &H8B107EDB, &HC1C033CF, &HFA8B02E9, &HFB8BABF3, &H5302E7C1, &HE8525656, &HFFFFF6AE, &HDB85C033, &H558B2D7E, &HF04D8BE8, &H83FCC283, &HD70304C1, &H3BF07D8B, &H8B047D07, &H3302EB31, &H403289F6, &H8304C183, &HC33B04EA, &HFB8BE97C, &H8B02E7C1, &HE851F04D, &HFFFFED3B, &HE6C1F38B, &HE0758903, &HFFA41DE8, &H10FF56FF, &H7589F08B, &HA410E808, &HCB8BFFFF, &H5103E1C1, &HC93310FF
    pvAppendBuffer &H85F84589, &H8B2C7EDB, &HC38BFC7D, &H8303E0C1, &HF003FCC0, &H3B04578D, &H8B047D0F, &H3302EB02, &H410689C0, &H8304C283, &HCB3B04EE, &HFB8BE97C, &HFF02E7C1, &HDCE8FC75, &H53FFFFEC, &H696E8, &H3F36B00, &HE6C1F003, &HD4758902, &HFFA3B5E8, &H10FF56FF, &H330C4D8B, &HFC4589D2, &H55891F6A, &H89018BF0, &H855EE445, &H8B317EC0, &H993C8DD8, &HCE8BC033, &H85E0D340, &H83107507, &H77901EE
    pvAppendBuffer &H83421F6A, &H3B5E04EF, &H8BE57CD3, &HFB8BDC5D, &H8BE4458B, &H55890C4D, &H2E7C1F0, &H8D0FD03B, &H97&, &H880FF685, &H80&, &HFF084D8B, &H48DFC75, &H75FF530F, &HE85050F8, &HFFFFF1FE, &HFC75FF53, &HFFF475FF, &H75FFEC75, &H3E6E8F8, &H458B0000, &H8BCE8B0C, &H2BC03310, &HD340F055, &HC4D8BE0, &H74910485, &HFC75FF2C, &H53F8458B, &H30875FF, &HE875FFC7, &HF1C0E850, &HFF53FFFF
    pvAppendBuffer &H75FFFC75, &HEC75FFF4, &HE80875FF, &H3A8&, &HEB084D8B, &H8458B0C, &H89F84D8B, &H4589084D, &H1EE83F8, &H558B8979, &HC4D8BF0, &H891F6A42, &H3B5EF055, &H698C0F11, &H53FFFFFF, &HFFFC75FF, &H75FFF475, &H875FFEC, &H36BE8, &H10458B00, &HA4E830FF, &H33000005, &H85F08BD2, &H8B1E7EDB, &HC7030845, &H8B0C4589, &H8B0E8BF8, &H42CA2B07, &H89047F8D, &HD33B8E04, &H7D8BEF7C, &H13E83D8
    pvAppendBuffer &H68B1076, &H863C83, &H89480875, &H1F88306, &H75FFF077, &HFC5D8BD4, &H1595E853, &H63E80000, &H53FFFFA2, &HFF0850FF, &H5D8BE075, &H80E85308, &HE8000015, &HFFFFA24E, &H850FF53, &H8BE075FF, &HE853F85D, &H156B&, &HFFA239E8, &H50FF53FF, &HF45D8B08, &H58E85357, &HE8000015, &HFFFFA226, &H850FF53, &H57EC5D8B, &H1545E853, &H13E80000, &H53FFFFA2, &H8B0850FF, &H5357E85D, &H1532E8
    pvAppendBuffer &HA200E800, &HFF53FFFF, &H8B5F0850, &H8B5B5EC6, &HCC25DE5, &HEC8B5500, &H532CEC83, &H10758B56, &H75FF5657, &HB3F9E808, &H1E8BFFFF, &HE7C1FB8B, &HE0458902, &H89E45D89, &HC7E8D47D, &H57FFFFA1, &HD08B10FF, &H5589C933, &H7EDB85F8, &H2B068B0F, &H86048BC1, &H418A0489, &HF17CCB3B, &HFFA1A5E8, &H10FF57FF, &H8BE0758B, &H89F88BD3, &H162BEC7D, &H97ED285, &HC033CA8B, &H7D8BABF3, &H33068BEC
    pvAppendBuffer &H7EC085C9, &H97148D14, &H8B41C12B, &H2898604, &H8B04528D, &H7CC83B06, &HC1FB8BEF, &H7D8903E7, &HA160E8DC, &HFF57FFFF, &H89F08B10, &H53E8F475, &H57FFFFA1, &H458910FF, &H1B048D08, &HD7EC085, &HFE8BC88B, &HABF3C033, &HE7C1FB8B, &H44C75303, &H1FC37, &HFCE80000, &H8B000003, &H2E6C1F0, &HE8D87589, &HFFFFA11E, &H8B10FF56, &HF6330C4D, &H6AF04589, &HFC75891F, &H855F018B, &H8D267EC0
    pvAppendBuffer &HD98B8114, &HCF8BC033, &H85E0D340, &H83107502, &H77901EF, &H83461F6A, &H3B5F04EA, &H8BE57C33, &H7589E45D, &HF8758BFC, &HE856368B, &HFFFFB739, &H85E84589, &H8B1874C0, &H83E6D3C8, &HF7E01FB, &H2B59206A, &HF8458BC8, &HD304408B, &H56F00BE8, &HA73E8, &HC4D8B00, &H8BE44589, &H94E9FC45, &H85000000, &H85880FFF, &H8B000000, &H75FFF475, &H9E048DF0, &H875FF53, &H3FE85050, &HFFFFFFEF
    pvAppendBuffer &H48DE875, &HE475FF1B, &HFF53006A, &HFF50F875, &H60E80875, &H8BFFFFEC, &HCF8B0C45, &HC033108B, &H40FC552B, &H4D8BE0D3, &H9104850C, &H75FF2F74, &H8458BF0, &H75FF5653, &H98048DEC, &HEEFCE850, &H75FFFFFF, &H1B048DE8, &H6AE475FF, &H75FF5300, &HE85650F8, &HFFFFEC1F, &HEB0C4D8B, &H8BC68B08, &H45890875, &H1EF8308, &H458B8479, &HF47589FC, &H891F6A40, &H3B5FFC45, &H648C0F01, &H8BFFFFFF
    pvAppendBuffer &H30FF1045, &H2F2E8, &H8BD23300, &H7EDB85F0, &HF4458B17, &H8B983C8D, &H2B078B0E, &H7F8D42CA, &H8E048904, &HEF7CD33B, &H76013E83, &H83068B10, &H7500863C, &H6894808, &H7701F883, &HDC7D8BF0, &H57F45D8B, &H12E9E853, &HB7E80000, &H53FFFF9F, &HFF0850FF, &H5D8BD875, &HD4E853F0, &HE8000012, &HFFFF9FA2, &H850FF53, &H57085D8B, &H12C1E853, &H8FE80000, &H53FFFF9F, &H8B0850FF, &H7D8BD45D
    pvAppendBuffer &HE85753F8, &H12AB&, &HFF9F79E8, &H50FF57FF, &H5D8B5308, &H98E853EC, &HE8000012, &HFFFF9F66, &H850FF53, &H50E0458B, &HFFE86AE8, &HC68B5FFF, &HE58B5B5E, &HCC25D, &H8BEC8B55, &H8B530845, &H5756145D, &H6B187D8B, &H48D0CF7, &H144589B8, &H5756F303, &H1075FF53, &HF024E850, &H8B56FFFF, &H3C8D1875, &HFF5756BB, &HE8530C75, &HFFFFEDCE, &H8D085D8B, &H53503604, &H2DE85357, &H8BFFFFEA
    pvAppendBuffer &HC9331455, &H8B104589, &H7EF685F9, &HBB048B16, &H528D0289, &HBB0C8904, &H7CFE3B47, &H14558BF0, &H8510458B, &H852A75C0, &H8B267EF6, &HFA8B0C5D, &H43B078B, &H4108758B, &H3B04C783, &H3BF17CCE, &H8B0E7DCE, &H48D087D, &H87048B31, &H768B043B, &HFF52560B, &HE8520C75, &HFFFFF0F2, &H5D5B5E5F, &H550014C2, &HEC81EC8B, &H108&, &HF0C458B, &H5756C057, &H8D593C6A, &HFFFEF8B5, &H130F66FF
    pvAppendBuffer &HFFFEF885, &HBD8DFF, &HC7FFFFFF, &H10F845, &HA5F30000, &HFEF8B58D, &HCE8BFFFF, &H2BFC7589, &HC4589C1, &H300C8B53, &H448BFE8B, &HDB330430, &H8910758B, &H4589F44D, &HFF5150F0, &HFF04DE74, &HA7E8DE34, &H1FFFFAB, &HF44D8B07, &H43045711, &H8DF0458B, &HFB83087F, &H8BDE7210, &H458BFC75, &H8C6830C, &H1F86D83, &H75FC7589, &H5BF633B8, &H266A006A, &H7CF5B4FF, &HFFFFFFFF, &HFF78F5B4
    pvAppendBuffer &H67E8FFFF, &H1FFFFAB, &HFEF8F584, &H9411FFFF, &HFFFEFCF5, &HFE8346FF, &H8BD5720F, &HB58D087D, &HFFFFFEF8, &HFF59206A, &HA5F30875, &HFFB48FE8, &H875FFFF, &HFFB487E8, &H8B5E5FFF, &HCC25DE5, &HEC8B5500, &H5310EC83, &HDB335756, &HC75FF53, &H1475FF53, &HFFAB15E8, &H75FF53FF, &HF045890C, &HFF53F28B, &H3E81875, &H53FFFFAB, &H891075FF, &HFA8BF445, &H1875FF53, &HFFAAF1E8, &H75FF53FF
    pvAppendBuffer &HFC458910, &H1475FF53, &HE8F85589, &HFFFFAADE, &H458BD88B, &H6ADE03F4, &HD6135E00, &HD713D803, &HD77D73B, &HD83B0472, &H75010773, &HF85583FC, &H8458B01, &H4D0BC933, &H5FDE0BF0, &H3308895E, &HFC5503C9, &H13045889, &H5089F84D, &HC488908, &H5DE58B5B, &H550014C2, &H4D8BEC8B, &HF6335608, &HC18B0DEB, &HD1C22B99, &H41C82BF8, &H838E348D, &HEE7F32F9, &H5D5EC68B, &H550004C2, &H8B53EC8B
    pvAppendBuffer &H5756085D, &H49D348D, &HE8000000, &HFFFF9CF6, &H5610FF56, &H6AF88B, &HF192E857, &HC483FFFF, &H8B1F890C, &H5B5E5FC7, &H4C25D, &H83EC8B55, &H8B5330EC, &H5756085D, &H75FF066A, &HBEE8530C, &H6A00001E, &HFF206A06, &H458D0C75, &H74E850D0, &H6A000012, &H8DF08B06, &H458D084B, &H50FA8BD0, &HC3835151, &H1123E838, &HC6030000, &H75FF066A, &H1303890C, &H8458BD7, &H8910C083, &H50500453
    pvAppendBuffer &H1108E8, &H84D8B00, &H4189066A, &HD0458D40, &H89515150, &H44E84451, &H8B000020, &HF003084D, &H518BFA13, &H8BD62B30, &HF71B3471, &H7234713B, &H3B05771D, &H16763051, &HEBFFCF83, &H85B8D03, &H7B113B01, &H23038B04, &HC73B0443, &H895FEF74, &H895E3471, &H8B5B3051, &H8C25DE5, &HEC8B5500, &H108EC81, &H56530000, &HC75FF57, &HFF78858D, &HE850FFFF, &H919&, &HFF78858D, &HE850FFFF
    pvAppendBuffer &HFFFFB2BC, &HFF78858D, &HE850FFFF, &HFFFFB2B0, &HFF78858D, &HE850FFFF, &HFFFFB2A4, &HFEF89D8D, &H45C7FFFF, &H20C&, &H8BF63300, &HFFFF788D, &H7C858BFF, &H81FFFFFF, &HFFEDE9, &H1B086A00, &HF88D89C6, &H89FFFFFE, &HFFFEFC85, &H548B5FFF, &H448BF83B, &H8C8BFC3B, &HFFFF783D, &HF85589FF, &H10C2AC0F, &H7C3D848B, &H83FFFFFF, &H748901E2, &HCA2BFC3B, &HE981C61B, &HFFFF&, &HF83D8C89
    pvAppendBuffer &H1BFFFFFE, &H3D8489C6, &HFFFFFEFC, &HF845B70F, &HF83B4489, &H8308C783, &HB27278FF, &HFF688D8B, &H858BFFFF, &HFFFFFF6C, &HFF0558B, &HF10C1AC, &HFF6885B7, &HE183FFFF, &H68858901, &H2BFFFFFF, &H6CB589D1, &H8BFFFFFF, &HCE1BF44D, &H7FFFEA81, &H95890000, &HFFFFFF70, &HC033CE1B, &HFF748D89, &HF40FFFF, &H8310CAAC, &HF9C101E2, &H50C22B10, &HFEF8858D, &H8D50FFFF, &HFFFF7885, &HB2E850FF
    pvAppendBuffer &H83000007, &HF010C6D, &HFFFF1E85, &H8558BFF, &H78F5848A, &H8BFFFFFF, &HFF78F58C, &H488FFFF, &HF5848B72, &HFFFFFF7C, &H8C1AC0F, &H8808F8C1, &H4601724C, &H7210FE83, &H5B5E5FD7, &HC25DE58B, &H8B550008, &H84D8BEC, &H8B56D233, &H6A570C75, &H5FF12B11, &H30E048B, &HFD00301, &HEAC1C2B6, &H8D018908, &HEF830449, &H5FE97501, &H8C25D5E, &HEC8B5500, &HC75FF56, &H5608758B, &HFFFFC1E8
    pvAppendBuffer &H44468DFF, &H15E85650, &H5E000001, &H8C25D, &H83EC8B55, &H8B5344EC, &H5756085D, &H8B59116A, &HBC7D8DF3, &H5BE8A5F3, &H5FFFF9A, &H544&, &HBC458D50, &HFF88E850, &H458BFFFF, &H25D0F7FC, &H80&, &H8D5E116A, &HD0F7FF50, &HC11FEAC1, &HD0231FE8, &HF7BC458D, &H89C32BD2, &H7A8D0845, &H8BD7F701, &H428D180C, &H23032301, &H8BC80BCF, &HB890845, &H83045B8D, &HE77501EE, &H8B5B5E5F
    pvAppendBuffer &H4C25DE5, &HEC8B5500, &H8D44EC83, &H6A56BC45, &H56F63344, &HEE96E850, &H4D8BFFFF, &HCC48308, &HA8918B, &HD2850000, &HB60F1174, &H983184, &H44890000, &H3B46BCB5, &H8DEF72F2, &H44C7BC45, &H1BC95, &H51500000, &HFFFF24E8, &HE58B5EFF, &H4C25D, &H8BEC8B55, &H3356084D, &HD68B57F6, &H1403FE8B, &HC2B60FB9, &H8908EAC1, &H8347B904, &HEE7C10FF, &H8B405103, &H2EAC1C2, &H6B03E083
    pvAppendBuffer &H418905D2, &HB1140340, &HC1C2B60F, &H48908EA, &HFE8346B1, &H1EE7C10, &H5E5F4051, &H4C25D, &H83EC8B55, &H8B534CEC, &H458D0C5D, &H2BD233B4, &H895756D8, &HFF33FC5D, &HD285C933, &H48D1E78, &H85D8B93, &H3B4758D, &H8B048BF0, &H8D06AF0F, &HF803FC76, &H7ECA3B41, &HFC5D8BF0, &H8901728D, &HCE8BF875, &H7D11FE83, &H85D8B2D, &HC62BC28B, &H830C758B, &H348D44C6, &H8B048B86, &H8D06AF0F
    pvAppendBuffer &HC069FC76, &H140&, &H8341F803, &HE97C11F9, &H8BF8758B, &H7C89FC5D, &HD68BB495, &H7C11FA83, &HB4458D95, &HFF2AE850, &H7D8BFFFF, &HB4758D08, &HF359116A, &H5B5E5FA5, &HC25DE58B, &H8B550008, &HC558BEC, &H3344EC83, &H4B60FC9, &H8D448911, &HF98341BC, &H8DF27C10, &H45C7BC45, &H1FC&, &H75FF5000, &HFE07E808, &HE58BFFFF, &H8C25D, &H81EC8B55, &H178EC, &H57565300, &H7D8DC033
    pvAppendBuffer &H88DB33B1, &H6AABB05D, &HC75FF0C, &HAB66ABAB, &HB4458DAA, &HECF1E850, &HC483FFFF, &HD05D880C, &H7D8DC033, &H59076AD1, &H46AABF3, &H8DAAAB66, &H6A50B045, &H875FF20, &HFF34858D, &HE850FFFF, &HFFFFC000, &H458D206A, &H8D5050D0, &HFFFF3485, &HD0E850FF, &H8DFFFFBA, &H8D50E045, &H8D50D045, &HFFFE8885, &H8EE850FF, &H6AFFFFCA, &HC0335908, &HF3D07D8D, &HD0458AAB, &H458D206A, &H8D5050D0
    pvAppendBuffer &HFFFF3485, &H9CE850FF, &H33FFFFBA, &HD07D8DC0, &HFF59086A, &HABF31475, &H8DD0458A, &H75FFF17D, &H88C03310, &HABABF05D, &HAAAB66AB, &HFE88858D, &HE850FFFF, &HFFFFCB01, &HF714458B, &HFE083D8, &HF0458D50, &H88858D50, &H50FFFFFE, &HFFCAE8E8, &H1C758BFF, &H7D8BC033, &H45894018, &H4539560C, &HFF167524, &H858D2075, &HFFFFFF34, &H34E85057, &H56FFFFBA, &HEB2075FF, &H858D5701, &HFFFFFE88
    pvAppendBuffer &HCAB3E850, &HC68BFFFF, &HE083D8F7, &H458D500F, &H858D50F0, &HFFFFFE88, &HCA9BE850, &H458DFFFF, &HFF5350F0, &H7AE81475, &H8D00001C, &H5350F845, &H1C6FE856, &H106A0000, &H458D535B, &H858D50F0, &HFFFFFE88, &HCA6FE850, &H7D83FFFF, &H2C750124, &H8D2875FF, &HFFFE8885, &H2FE850FF, &H6AFFFFC9, &H34858D7C, &H6AFFFFFF, &HC1E85000, &H8AFFFFEB, &HFFFF3485, &HCC483FF, &H73EBC033, &H50C0458D
    pvAppendBuffer &HFE88858D, &HE850FFFF, &HFFFFC902, &H8D287D8B, &HC18BC04D, &HF82BD232, &H320F048A, &H41D00A01, &H7501EB83, &H187D8BF3, &H2075FF56, &H1275D284, &H34858D57, &H50FFFFFF, &HFFB96AE8, &HC5D21FF, &H22E805EB, &H6AFFFFEB, &H34858D7C, &H6AFFFFFF, &H59E85000, &H8AFFFFEB, &HFFFF348D, &HC07D8DFF, &H330CC483, &HABABABC0, &HC04D8AAB, &H5F0C458B, &HE58B5B5E, &H24C25D, &H8BEC8B55, &HB60F0855
    pvAppendBuffer &H4AB60F02, &H8E0C101, &HB60FC10B, &HE0C1024A, &HFC10B08, &HC1034AB6, &HC10B08E0, &H4C25D, &H8BEC8B55, &HB60F0855, &HB60F0342, &HE0C1024A, &HFC10B08, &HC1014AB6, &HC10B08E0, &HC10AB60F, &HC10B08E0, &H4C25D, &H53EC8B55, &H33085D8B, &HC1CB8BD2, &H564110E9, &H8D006A57, &HFFFFFF81, &H8BF1F77F, &H10E6C1F0, &HE3F7C68B, &HCE03C88B, &HD283D1F7, &H83C03300, &HD2F701C1, &HC203C013
    pvAppendBuffer &HE8C1E6F7, &H3FA8B1F, &H8BF80BFF, &H8BE3F7C7, &H3CA8BF0, &HBAF7&, &H13588000, &H72CA3BC8, &H33D3F718, &H13F303C0, &H1C683C0, &H4900D083, &H3B4FC803, &HEBEC73CA, &H3D23319, &HEBD013F3, &HC9334707, &HC813F303, &HFA81D103, &H80000000, &HC78BEF72, &H5D5B5E5F, &H550004C2, &H458BEC8B, &HC558B08, &H406A5756, &H104D2B59, &HFFA320E8, &H104D8BFF, &H458BF08B, &H8BFA8B08, &H2DE80C55
    pvAppendBuffer &HBFFFFA3, &H5FD70BC6, &HCC25D5E, &HEC8B5500, &H5324EC83, &H7D8B5756, &H75FF5708, &HAA40E80C, &HFF57FFFF, &H45891075, &HAA34E808, &HFF57FFFF, &H45891475, &HAA28E80C, &HF78BFFFF, &HD1F04589, &H75FF56EE, &HAA18E818, &HFF56FFFF, &HD88B1C75, &HE8EC5D89, &HFFFFAA0A, &H2075FF56, &HE8104589, &HFFFFA9FE, &HE8458953, &HFFCC9DE8, &H458950FF, &HCCFCE81C, &H758BFFFF, &H8BE85610, &H50FFFFCC
    pvAppendBuffer &HE8204589, &HFFFFCCEA, &HFF1C75FF, &H10E80C75, &HFFFFFFA7, &H45892075, &HC75FFFC, &HFFA702E8, &H75FF53FF, &HF84589FC, &HE80875FF, &HFFFFEF57, &HF875FF56, &H75FFD88B, &H145D8908, &HFFEF46E8, &H56F08BFF, &H18758953, &HFFA91CE8, &H79C085FF, &HEC75FF19, &HE853F38B, &HFFFFA4C6, &H89D88B56, &HB0E81445, &H8BFFFFDD, &H53561875, &HFFAA44E8, &H1075FFFF, &HFFF44589, &HDEE8E875, &H56FFFFA6
    pvAppendBuffer &H89F475FF, &HE850E045, &HFFFFA6E5, &H89F075FF, &HE850DC45, &HFFFFA68B, &H85E44589, &H8B1A74FF, &HD88B2475, &HE853574F, &HFFFFA879, &H85460688, &H8BF175FF, &H758B145D, &H1C75FF18, &HFFDD56E8, &H2075FFFF, &HFFDD4EE8, &HFC75FFFF, &HFFDD46E8, &HF875FFFF, &HFFDD3EE8, &H38E853FF, &H56FFFFDD, &HFFDD32E8, &HF475FFFF, &HFFDD2AE8, &HE075FFFF, &HFFDD22E8, &HDC75FFFF, &HFFDD1AE8, &H875FFFF
    pvAppendBuffer &HFFDD12E8, &HC75FFFF, &HFFDD0AE8, &HF075FFFF, &HFFDD02E8, &HEC75FFFF, &HFFDCFAE8, &H1075FFFF, &HFFDCF2E8, &HE875FFFF, &HFFDCEAE8, &HE475FFFF, &HFFDCE2E8, &H5B5E5FFF, &HC25DE58B, &H8B550020, &H8B5653EC, &H56570875, &HE80C75FF, &HFFFFA892, &H1075FF56, &H5D89D88B, &HA884E80C, &HFF56FFFF, &HF88B1475, &HE8087D89, &HFFFFA876, &H89535750, &H10E81045, &H8BFFFFEE, &H74F685D8, &H187D8B15
    pvAppendBuffer &HE853564E, &HFFFFA79D, &H85470788, &H8BF175F6, &H75FF087D, &HDC7DE80C, &HE857FFFF, &HFFFFDC77, &HE81075FF, &HFFFFDC6F, &HDC69E853, &H5E5FFFFF, &H14C25D5B, &HEC8B5500, &H458B5151, &H5D8B5310, &HF7564808, &H1045C7D0, &H10&, &H7D8B5799, &H89DF2B0C, &H5589FC45, &H3B348BF8, &H43B548B, &H4F8B078B, &H23C63304, &HCA33FC45, &H33F84D23, &H89D133F0, &H54893B34, &H731043B, &H31087F8D
    pvAppendBuffer &H6D83FC4F, &HD1750110, &H8B5B5E5F, &HCC25DE5, &HEC8B5500, &H8B0C558B, &HD12B084D, &H5E106A56, &H890A048B, &H8498D01, &HFC0A448B, &H83FC4189, &HEC7501EE, &H8C25D5E, &HEC8B5500, &H106A5653, &H105D395B, &H7D833674, &H5A752010, &H570C758B, &H53087D8B, &H30E85756, &H53FFFFE7, &H5010468D, &H5010478D, &HFFE722E8, &H18C483FF, &HFF9295E8, &H53105FF, &H47890000, &H2AEB5F30, &H5308758B
    pvAppendBuffer &H560C75FF, &HFFE702E8, &H75FF53FF, &H10468D0C, &HE6F5E850, &HC483FFFF, &H9268E818, &H2005FFFF, &H89000005, &H5B5E3046, &HCC25D, &H83EC8B55, &H8B536CEC, &H5756085D, &HA0BFF633, &H8B000001, &H104B8B03, &H8BF84589, &H45890443, &H8438BF0, &H8BEC4589, &H45890C43, &H14438BE0, &H8BE84589, &H45891843, &H1C438BE4, &H89945D8D, &HDF2BF44D, &H89DC4589, &H7D89D875, &H10FE83D4, &H758B1773
    pvAppendBuffer &H71E8560C, &H83FFFFFB, &H458904C6, &H3B0489FC, &HEB0C7589, &H17E8D5D, &H8D0FE683, &HE083FD47, &H85548B0F, &H83C78B94, &H4C8B0FE0, &HC18B9485, &H890EC0C1, &HC18BFC45, &H3107C8C1, &HC28BFC45, &H3103E9C1, &HCA8BFC4D, &HC10DC0C1, &HC8330FC1, &H330AEAC1, &HF8478DCA, &H83FC558B, &H7D8B0FE0, &H3D103D4, &H3948554, &H8994B554, &H5489FC55, &H8FE894B5, &H8BFFFF91, &HD68BF475, &HCAC1CE8B
    pvAppendBuffer &H7C1C10B, &HCE8BD133, &HF706C9C1, &HE47523D6, &HC8BD133, &H4C78338, &H3E8458B, &HF44523CA, &H33FC4D03, &H89F103F0, &H4D8BD47D, &H3D18BF8, &HC18BDC75, &HC10AC0C1, &HD0330DCA, &HC8C1C18B, &H8BD03302, &HC88BF045, &H33F84523, &H4D23F84D, &H8BC833EC, &H4589E445, &H8BD103DC, &H4D8BE845, &HE44589E0, &H458BCE03, &HE84589F4, &H89EC458B, &H458BE045, &HEC4589F0, &H89F8458B, &H48DF045
    pvAppendBuffer &HD8758B32, &HF44D8946, &H89F84589, &HFF81D875, &H2A0&, &HFEDF820F, &H5D8BFFFF, &H15E5F08, &HF0458B03, &H8B044301, &H4301EC45, &HE0458B08, &H8B0C4301, &H4301E845, &HE4458B14, &H1184301, &H458B104B, &H1C4301DC, &H5B6043FF, &HC25DE58B, &H8B550008, &HDCEC81EC, &H8B000000, &H56530845, &H4488B57, &H8BF84D89, &H108B0848, &H8BD84D89, &H4D890C48, &H10488BEC, &H8BD44D89, &H4D891448
    pvAppendBuffer &H18488BD0, &H8BB84D89, &H4D891C48, &H20488BB4, &H8BE84D89, &H4D892448, &H28488BFC, &H8BCC4D89, &H4D892C48, &H30488BC8, &H8BC44D89, &H4D893448, &H38488BC0, &H893C408B, &H858DAC45, &HFFFFFF24, &HB9B04D89, &H2A0&, &H33E05589, &H89C12BD2, &H4D89BC55, &HA44589DC, &H7310FA83, &HC7D8B36, &H5004478D, &HFFF98FE8, &HD88B57FF, &H85E8F633, &H8BFFFFF9, &HF00BDC4D, &H83A4458B, &H5D8908C7
    pvAppendBuffer &HF07589F4, &H890C7D89, &H7489081C, &HD6E90408, &H8D000000, &HE083FE42, &H8B3D6A0F, &HFF24C5BC, &HB48BFFFF, &HFFFF28C5, &H1428DFF, &H830FE083, &H57560FE2, &H8BA85589, &HFF24C58C, &H9C8BFFFF, &HFFFF28C5, &HE44D89FF, &HFFFA0EE8, &H56136AFF, &HF4458957, &HE8F05589, &HFFFFF9FF, &H33F44D8B, &HF0458BC8, &H6F7AC0F, &HC233086A, &H5306EEC1, &H33E475FF, &H89C633CF, &H4589F44D, &HF9D9E8F0
    pvAppendBuffer &H16AFFFF, &HE475FF53, &HFA8BF08B, &HFFF9CAE8, &HF44D8BFF, &H458BF033, &H8BFA33E4, &HAC0FF055, &HF03307D8, &H8B07EBC1, &HFB33BC45, &HD713CE03, &H83F9C083, &H8C030FE0, &HFFFF24C5, &HC59413FF, &HFFFFFF28, &H3A8458B, &HFF24C58C, &H4D89FFFF, &HC59413F4, &HFFFFFF28, &H24C58C89, &H89FFFFFF, &H9489F055, &HFFFF28C5, &HE8758BFF, &H75FF296A, &H64E856FC, &H6AFFFFF9, &HFC75FF12, &HFA8BD88B
    pvAppendBuffer &HF955E856, &HE6AFFFF, &H33FC75FF, &H56FA33D8, &HFFF946E8, &H33D833FF, &H8EE8E8FA, &H4D8BFFFF, &H8BD6F7DC, &HD2F7FC55, &H1C03276A, &HF875FF08, &H4087C13, &H23C47523, &H4D8BC055, &HE84D23CC, &H33C8458B, &HFC4523F1, &HDE03D033, &H13E0758B, &HF45D03FA, &HF07D1356, &H89B05D03, &H7D13A85D, &HE47D89AC, &HFFF8F2E8, &HFF226AFF, &HF88BF875, &HE856DA8B, &HFFFFF8E3, &H75FF1C6A, &H33F833F8
    pvAppendBuffer &HD4E856DA, &H8BFFFFF8, &HDA33D84D, &H33EC558B, &HF85533F8, &H7533F18B, &HD05523E0, &H23EC458B, &H7523F845, &H23D033D4, &H458BE04D, &H89F133C4, &HF703B045, &H89C0458B, &HDA13AC45, &H89CC458B, &H458BC445, &HA84D8BC8, &H8BB84D03, &H5513E455, &HA87503B4, &H8BC04589, &H5D13E845, &HCC4589E4, &H89FC458B, &H558BFC55, &HB85589D4, &H89D0558B, &H558BB455, &HD45589D8, &H89EC558B, &H558BD055
    pvAppendBuffer &HD85589E0, &H89F8558B, &H4D8BE84D, &HEC5589DC, &H8B08C183, &H8942BC55, &H7589C845, &HF85D89E0, &H89BC5589, &HF981DC4D, &H520&, &HFDA6820F, &H458BFFFF, &HD8558B08, &H5FD44D8B, &H758B3001, &H45811B4, &H8B085001, &H5011EC55, &H1048010C, &H11D04D8B, &H558B1448, &H185001B8, &H11E84D8B, &H48011C70, &HFC4D8B20, &H8B244811, &H4801CC4D, &HC84D8B28, &H8B2C4811, &H4801C44D, &HC04D8B30
    pvAppendBuffer &H8B344811, &H4801B04D, &HAC4D8B38, &HFF3C4811, &HC080&, &H8B5B5E00, &H8C25DE5, &HEC8B5500, &H85D8B53, &HB60F5756, &HB60F077B, &HB60F0A43, &HB60F0B73, &HE7C10F53, &HFF80B08, &HF034BB6, &HC10D43B6, &HF80B08E7, &HF08E6C1, &HE7C103B6, &HC1F80B08, &HB60F08E2, &HF00B0E43, &HF08E1C1, &HC10143B6, &HF00B08E6, &H443B60F, &HB08E6C1, &H43B60FF0, &HFD00B02, &HC10543B6, &HD00B08E2
    pvAppendBuffer &H843B60F, &HB08E2C1, &H43B60FD0, &H89C80B06, &HB60F047B, &HE1C10943, &H89C80B08, &HB60F0873, &HE1C10C43, &HC80B5F08, &H5E0C5389, &H5D5B0B89, &H550004C2, &H458BEC8B, &H74C08508, &HC4D8B10, &H974C985, &H400000C6, &H7501E983, &H8C25DF7, &HEC8B5500, &HFF0C75FF, &H75FF0C75, &HEE01E808, &HC25DFFFF, &H8B550008, &H10558BEC, &H758B5653, &H7D8B570C, &H6AF22B08, &H5BFA2B10, &H2B160C8B
    pvAppendBuffer &H16448B0A, &H4421B04, &H8D170C89, &H44890852, &HEB83FC17, &H5FE57501, &HC25D5B5E, &H8B55000C, &HE85756EC, &HFFFF8C5E, &HBF08758B, &H588&, &HFF50C703, &H41E836, &H6890000, &HFF8C45E8, &H50C703FF, &HE80476FF, &H2F&, &HE8044689, &HFFFF8C32, &HFF50C703, &H1CE80876, &H89000000, &H1FE80846, &H3FFFF8C, &H76FF50C7, &H9E80C, &H895F0000, &H5D5E0C46, &H550004C2, &H558BEC8B
    pvAppendBuffer &H5D8B530C, &HC1C38B08, &HCB8B18E8, &H8E9C156, &HFC9B60F, &H8B1034B6, &H10E8C1C3, &HFC0B60F, &HC1110CB6, &HB60F08E6, &HC60B1004, &HB08E0C1, &HCBB60FC1, &H5E08E0C1, &HCB60F5B, &H5DC10B11, &H550008C2, &H5653EC8B, &H87D8B57, &H458BDB33, &H44B60F0C, &H8B990158, &H8BF28BC8, &HA40F0C45, &HE1C108CE, &H4B60F08, &HC8039958, &H13DF0C89, &HDF7489F2, &HFB834304, &H81D37210, &H7FFF7867
    pvAppendBuffer &H67830000, &H5E5F007C, &H8C25D5B, &HEC8B5500, &H570F5151, &H7D8B57C0, &H130F6614, &H558BF845, &HF84D8BFC, &H7374FF85, &H530C458B, &H10758B56, &H4529F02B, &H10758908, &H306348B, &HC758930, &H8B10758B, &H748B0C5D, &H70130406, &H89D90304, &HF2130C5D, &H5D8B183B, &H3B057508, &H23740470, &H7704703B, &H8B077210, &HC4D3908, &HC9330773, &HEBD23341, &HC0570F0E, &H45130F66, &HFC558BF8
    pvAppendBuffer &H8BF84D8B, &H74890C7D, &H758B0403, &H33C8910, &H8308C083, &H7501146D, &H8B5B5E9F, &HE58B5FC1, &H10C25D, &H8BEC8B55, &HC9850C4D, &H57561F74, &H8D087D8B, &HFFF8CD0C, &HF78BFFFF, &H8302E9C1, &H67830027, &HC7830004, &H5FA5F308, &H8C25D5E, &HEC8B5500, &H83104D8B, &H575601E9, &H458B3078, &H8758B0C, &H148DF02B, &H167C8BC8, &H16048B04, &H77047A3B, &H3B1F7226, &H3B207702, &H1672047A
    pvAppendBuffer &H23B0477, &HEA831072, &H1E98308, &HC033DB79, &HC25D5E5F, &HC883000C, &H33F5EBFF, &HF0EB40C0, &H33EC8B55, &HC4D39C9, &H558B1276, &HCA048B08, &H4CA440B, &H3B410D75, &HF1720C4D, &H5D40C033, &H330008C2, &H55F8EBC0, &HEC83EC8B, &H147D830C, &HC0570F00, &H45130F66, &H5D8B53F8, &H8B6376FC, &H558B0C4D, &H58406A08, &H2B10452B, &HF44589CA, &H56F8458B, &H890C4D89, &H8B57FC45, &HC78B113C
    pvAppendBuffer &H411748B, &H4D8BD68B, &H9783E810, &H4D8BFFFF, &HBD30B08, &H189FC45, &H5189C78B, &H8BD68B04, &H89E8F44D, &H8BFFFF97, &HDA8B0C4D, &H8308558B, &H458908C2, &H146D83FC, &H8558901, &H5E5FBD75, &H458B03EB, &H5BD38BF8, &HC25DE58B, &H8B550010, &H28EC83EC, &H758B5653, &H7D8B570C, &H56046A08, &HB9BE857, &H468B0000, &HC0570F2C, &HE06583, &H8BE44589, &H45893046, &H34468BE8, &H8BEC4589
    pvAppendBuffer &H45893846, &H3C468BF0, &H4589046A, &HD8458DF4, &H5050016A, &H45130F66, &HFF29E8D8, &H46AFFFF, &H458DD88B, &H575750D8, &HFFFDE0E8, &H384E8BFF, &H468BD803, &H3C568B30, &HE06583, &HF46583, &H33E44589, &H34460BC0, &H4589046A, &HD8458DE8, &H5050016A, &H89EC4D89, &HE8E8F055, &H6AFFFFFE, &H8DD80304, &H5750D845, &HFD9FE857, &H6583FFFF, &HD80300E4, &HF20468B, &H4589C057, &H24468BD8
    pvAppendBuffer &H8BDC4589, &H45892846, &H38468BE0, &H8BF04589, &H46A3C46, &H8DF44589, &H5750D845, &H130F6657, &H66E8E845, &H3FFFFFD, &H244E8BD8, &H460BC033, &H34568B28, &H8BDC4589, &H45893046, &HBC033F8, &H45892C46, &H38468BE0, &H8BE84589, &H45893C46, &HBC033EC, &H46A2046, &H8DF44589, &H8950D845, &HCA8BD84D, &H4D895757, &HF05589E4, &HFFFD1CE8, &H2C4E8BFF, &H568BD803, &HBC03334, &H570F3046
    pvAppendBuffer &HE46583C0, &HDC458900, &H6A20468B, &HF0458904, &H89D8458D, &HC933D84D, &H50284E0B, &H55895757, &H130F66E0, &H4D89E845, &HC31E8F4, &H568B0000, &H8BD82B24, &H570F3046, &HD84589C0, &H468B20B1, &HDC458934, &H8938468B, &H468BE045, &HE445893C, &H6620468B, &HE845130F, &HFF95D3E8, &H2C560BFF, &H4589046A, &HD8458DF0, &H89575750, &HECE8F455, &H8B00000B, &HD82B344E, &H5D89C033, &H38460BFC
    pvAppendBuffer &H89305E8B, &H4589D84D, &H20568BDC, &H460BC033, &H8B20B13C, &HF633247E, &H8BE45589, &H45890C55, &H28428BE0, &HE82C528B, &HFFFF9565, &HF06583, &H46AF80B, &H8DF45D89, &H535BD845, &HE87D8950, &H7D8BF20B, &H89575708, &H90E8EC75, &H8B00000B, &H758B0C4D, &HE06583FC, &H83F02B00, &H8B00F065, &H45893841, &H3C418BD8, &H8BDC4589, &H45892441, &H28418BE4, &H8BE84589, &H45892C41, &H34418BEC
    pvAppendBuffer &HF4458953, &H50D8458D, &H50E85757, &H2B00000B, &H531E79F0, &HFF876DE8, &H10C083FF, &HE8575750, &HFFFFFBE9, &HEB78F003, &H8B5B5E5F, &H8C25DE5, &H75F68500, &HE8575315, &HFFFF874A, &H5010C083, &HFFFC8CE8, &H1F883FF, &HE853DE74, &HFFFF8736, &H5010C083, &H4E85757, &H2B00000B, &H55D2EBF0, &HEC83EC8B, &H758B5674, &H66A570C, &H57307E8D, &HE8EC7D89, &HFFFFFCA8, &H850FC085, &HAB&
    pvAppendBuffer &H8C5D8D53, &H66ADE2B, &HF8C458D, &H6650C057, &HF045130F, &HFFFC0BE8, &H8D066AFF, &HE850BC45, &HFFFFFC00, &H8C458D57, &HE9FEE850, &H66AFFFF, &HFBEEE857, &H4D8BFFFF, &H89C68BF4, &H4D8BF84D, &H8BF98BF0, &H45C7F875, &H9FC&, &H3148B00, &H4C8B1003, &H48130403, &H13D70304, &H75103BCE, &H4483B05, &H483B2074, &H720D7704, &H73103B04, &H47FF3307, &HEEBF633, &H66C0570F, &HF045130F
    pvAppendBuffer &H8BF4758B, &H1089F07D, &H83044889, &H6D8308C0, &HB97501FC, &H6AEC7D8B, &H1E85706, &H8BFFFFFC, &HC0850C75, &HFF5C840F, &H6A5BFFFF, &H8654E806, &HB0BFFFFF, &HEB000000, &H8648E816, &HC703FFFF, &HE8565650, &HA17&, &H37E8066A, &H3FFFF86, &HE85650C7, &HFFFFFB79, &HC085066A, &HFF56DB7F, &H1AE80875, &H5F000008, &H5DE58B5E, &H550008C2, &H7D83EC8B, &H75FF0410, &H875FF0C, &H4BE80775
    pvAppendBuffer &HEBFFFFFC, &HFED5E805, &HC25DFFFF, &H8B55000C, &H1875FFEC, &HFF1075FF, &H75FF0C75, &HFA6BE808, &HC20BFFFF, &H75FF1275, &H1475FF18, &HE80875FF, &HFFFFFB1D, &H1178C085, &HFF1875FF, &H75FF1475, &H875FF08, &H996E8, &H14C25D00, &HEC8B5500, &HC8EC81, &H8B560000, &HFF561475, &H3DE80C75, &H85FFFFFB, &H560E74C0, &HE80875FF, &HFFFFFAB4, &H202E9, &H56575300, &H8D0C75FF, &HFFFF3885
    pvAppendBuffer &H7AE850FF, &H8B000007, &H858D105D, &HFFFFFF68, &HE8505356, &H769&, &HC8458D56, &HFA82E850, &H6583FFFF, &H458D00CC, &H56FF3398, &H7D895047, &HFA6EE8C8, &H8D56FFFF, &HFFFF6885, &H858D50FF, &HFFFFFF38, &HFA87E850, &H8FE9FFFF, &H8B000001, &HFFFF3885, &HC0570FFF, &HF66C723, &H83F84513, &HE7500C8, &HFF38858D, &HE850FFFF, &H6BB&, &H858B74EB, &HFFFFFF68, &HC883C723, &H8D117500
    pvAppendBuffer &HFFFF6885, &HA0E850FF, &HE9000006, &HF2&, &H8E0FC985, &H9C&, &HFF68858D, &H8D50FFFF, &HFFFF3885, &HE85050FF, &H8B3&, &H38858D56, &H50FFFFFF, &H672E8, &H458D5600, &H458D5098, &HAE850C8, &H85FFFFFA, &H560C79C0, &HC8458D53, &H36E85050, &H56FFFFF9, &H5098458D, &H50C8458D, &H879E850, &H458B0000, &H83C723C8, &H137400C8, &H458D5356, &HE85050C8, &HFFFFF911, &H5589F88B
    pvAppendBuffer &H8B09EB0C, &H7D8BFC45, &HC4589F8, &HC8458D56, &H615E850, &H7D0B0000, &HA8840F0C, &H8B000000, &H81C0F544, &HC4F54C, &H89800000, &HE9C0F544, &H93&, &HFF38858D, &H8D50FFFF, &HFFFF6885, &HE85050FF, &H817&, &H68858D56, &H50FFFFFF, &H5D6E8, &H458D5600, &H458D50C8, &H6EE85098, &H85FFFFF9, &H560C79C0, &H98458D53, &H9AE85050, &H56FFFFF8, &H50C8458D, &H5098458D, &H7DDE850
    pvAppendBuffer &H458B0000, &H83C72398, &H137400C8, &H458D5356, &HE8505098, &HFFFFF875, &H5589F88B, &H8B09EB0C, &H7D8BFC45, &HC4589F8, &H98458D56, &H579E850, &H7D0B0000, &H8B10740C, &H8190F544, &H94F54C, &H89800000, &H5690F544, &HFF68858D, &H8D50FFFF, &HFFFF3885, &HF6E850FF, &H6AFFFFF8, &HC88B5F01, &HFC98556, &HFFFE6685, &HC8458DFF, &H875FF50, &H58CE8, &H5E5B5F00, &HC25DE58B, &H8B550010
    pvAppendBuffer &H80EC81EC, &H53000000, &H46A5756, &H75FF535B, &H4B4E814, &HFF530000, &HF08B1075, &HFF80458D, &HE8500C75, &H34A&, &HA0458D53, &H498E850, &HF88B0000, &H874FF85, &H100C781, &HCEB0000, &H80458D53, &H480E850, &HF88B0000, &H73FE3B53, &H80458D0C, &H875FF50, &HFAE9&, &HC0458D00, &HF83EE850, &H8D53FFFF, &HE850E045, &HFFFFF834, &HC62BC78B, &HEEC1F08B, &HE0835306, &H501A743F
    pvAppendBuffer &H8D1475FF, &H48DC045, &HB8E850F0, &H89FFFFF8, &H89E0F544, &HEBE4F554, &H1475FF0F, &H8DC0458D, &HE850F004, &H4D9&, &H85D8B53, &HF7F2E853, &H6383FFFF, &H3C70004, &H1&, &H815E046A, &H100FF, &H56117700, &H8D1475FF, &HE850C045, &HFFFFF7FD, &H7978C085, &HA0458D56, &HE0458D50, &HF7EBE850, &HC085FFFF, &H40751478, &H80458D56, &HC0458D50, &HF7D7E850, &HC085FFFF, &H8D562E7F
    pvAppendBuffer &H8D50C045, &H50508045, &H652E8, &H74C20B00, &H8D53560C, &H5050A045, &H642E8, &H458D5600, &H458D50E0, &HE85050A0, &H633&, &H8DE0758B, &H46AE045, &H1FE6C150, &H3EEE8, &H8D046A00, &HE850C045, &H3E3&, &H4FDC7509, &HFFFF6BE9, &H458D56FF, &HE8535080, &H425&, &H8B5B5E5F, &H10C25DE5, &HEC8B5500, &HC0EC81, &H56530000, &H5B066A57, &H1475FF53, &H34DE8, &H75FF5300
    pvAppendBuffer &H8DF08B10, &HFFFF4085, &HC75FFFF, &H1E0E850, &H8D530000, &HFFFF7085, &H2BE850FF, &H8B000003, &H74FF85F8, &H80C78108, &HEB000001, &H858D530F, &HFFFFFF40, &H310E850, &HF88B0000, &H73FE3B53, &H40858D0F, &H50FFFFFF, &HE90875FF, &H110&, &H50A0458D, &HFFF6CBE8, &H458D53FF, &HC1E850D0, &H8BFFFFF6, &H8BC62BC7, &H6EEC1F0, &H3FE08353, &HFF501A74, &H458D1475, &HF0048DA0, &HF745E850
    pvAppendBuffer &H4489FFFF, &H5489D0F5, &HFEBD4F5, &H8D1475FF, &H48DA045, &H66E850F0, &H53000003, &H53085D8B, &HFFF67FE8, &H46383FF, &H103C700, &H6A000000, &HFF815E06, &H180&, &HFF561577, &H458D1475, &H8AE850A0, &H85FFFFF6, &H88880FC0, &H56000000, &HFF70858D, &H8D50FFFF, &HE850D045, &HFFFFF671, &H1778C085, &H8D564C75, &HFFFF4085, &H458D50FF, &H5AE850A0, &H85FFFFF6, &H56377FC0, &H50A0458D
    pvAppendBuffer &HFF40858D, &H5050FFFF, &H4D2E8, &H74C20B00, &H8D53560F, &HFFFF7085, &HE85050FF, &H4BF&, &HD0458D56, &H70858D50, &H50FFFFFF, &H4ADE850, &H758B0000, &HD0458DD0, &HC150066A, &H68E81FE6, &H6A000002, &HA0458D06, &H25DE850, &H75090000, &H58E94FCC, &H56FFFFFF, &HFF40858D, &H5350FFFF, &H29CE8, &H5B5E5F00, &HC25DE58B, &H8B550010, &H60EC83EC, &HFFA0458D, &H75FF1475, &HC75FF10
    pvAppendBuffer &H6CE850, &H75FF0000, &HA0458D14, &H875FF50, &HFFFA5AE8, &H5DE58BFF, &H550010C2, &HEC83EC8B, &HA0458D60, &HFF1075FF, &HE8500C75, &H27D&, &H8D1075FF, &HFF50A045, &H30E80875, &H8BFFFFFA, &HCC25DE5, &HEC8B5500, &HFF1875FF, &H75FF1075, &H875FF0C, &H406E8, &H74C20B00, &H1875FF11, &HFF1475FF, &H75FF0875, &HF49FE808, &HC25DFFFF, &H8B550014, &H64EC83EC, &HF144D8B, &H5653C057
    pvAppendBuffer &H66FF3357, &HCC45130F, &H8DCC758B, &HFFFF4D04, &H8947FFFF, &H458BE845, &H2BDB33D0, &H130F66F9, &H7D89D445, &H1F148DE4, &H33C0570F, &H130F66FF, &HD93BF845, &H8BD7420F, &H5589FC7D, &HFD33BF0, &HBD87&, &H10758B00, &HC22BC38B, &H8DF47D89, &H458BC634, &HEC7589F8, &H3BFC4589, &H97830FD1, &HFF000000, &H458B0476, &HFF36FF0C, &HFF04D074, &H458DD034, &HC6E850AC, &H8BFFFFE1, &HBC7D8DF0
    pvAppendBuffer &HA510EC83, &H8BA5A5A5, &H83F08BFC, &H458D10EC, &HA5A5A59C, &H8DFC8BA5, &HA550CC75, &HE8A5A5A5, &HFFFF8D63, &H7D8DF08B, &HA5A5A5CC, &HD8458BA5, &H77C8453B, &H8B087211, &H453BD445, &H330773C4, &HC93340C0, &H570F0EEB, &H130F66C0, &H4D8BDC45, &HDC458BE0, &H8BFC4501, &H558BF47D, &H8BF913F0, &H8B42EC75, &HEE83144D, &HF47D8908, &H89F05589, &HD33BEC75, &HFF61860F, &H458BFFFF, &HCC758BD0
    pvAppendBuffer &H558B06EB, &HFC5589F8, &H8B08558B, &H3489FC4D, &HD4758BDA, &H4DA4489, &HD8458B43, &H8BD44D89, &H7D89144D, &HE47D8BD8, &H89CC7589, &H5D3BD045, &HEE820FE8, &H3FFFFFE, &H74895FC9, &H895EF8CA, &H5BFCCA44, &HC25DE58B, &H8B550010, &H75FF56EC, &H8758B0C, &H2EE856, &HC88B0000, &H2374C985, &HFCCE548B, &HCE7C8B57, &HEBF633F8, &HD7AC0F07, &H46EAD101, &HC20BC78B, &HE1C1F375, &H418D5F06
    pvAppendBuffer &H5EC603C0, &H8C25D, &H8BEC8B55, &HE9830C4D, &H8B117801, &H48B0855, &HCA440BCA, &H83057504, &HF27901E9, &H5D01418D, &H550008C2, &H5151EC8B, &HF0C458B, &H5756C057, &H66087D8B, &HF845130F, &H3BC7348D, &H8B3276F7, &H8B53F845, &H4589FC5D, &HFC4E8B08, &H8B08EE83, &HFC28B16, &HB01C8AC, &H65830845, &HE9D10008, &H689CB0B, &H4E89DA8B, &H1FE3C104, &HD977F73B, &H8B5E5F5B, &H8C25DE5
    pvAppendBuffer &HEC8B5500, &H8510558B, &H8B1E74D2, &H8B56084D, &HF12B0C75, &H890E048B, &H8498D01, &HFC0E448B, &H83FC4189, &HEC7501EA, &HCC25D5E, &HEC8B5500, &H8B64EC83, &HD2331045, &HC0570F53, &HC8D5642, &HFFFFFF45, &H130F66FF, &H4D89CC45, &H2BC933E0, &H130F66D0, &H8B57D445, &H4D89D07D, &HDC5589E4, &HF0A1C8D, &HD233C057, &H45130F66, &HEC758BEC, &H420FC83B, &HFD93BDA, &H11587, &HC558B00
    pvAppendBuffer &HC32BC18B, &H8DFC7589, &H558BC23C, &HE87D89F0, &H8BF85589, &H89C32BC1, &HD83BF045, &HEB870F, &H77FF0000, &HC458B04, &H74FF37FF, &H34FF04D8, &HAC458DD8, &HDF87E850, &H7D8DFFFF, &HA5F08BBC, &H3BA5A5A5, &H4373F05D, &H8BC84D8B, &HC0558BC1, &HE8C1F28B, &HFC45011F, &H83C4458B, &H3300F855, &HC1A40FFF, &H1FEEC101, &HF90BC003, &H7D89F00B, &HBC458BF4, &H1C2A40F, &H3F07589, &HC47589C0
    pvAppendBuffer &H89C87D89, &H5589BC45, &H8B0CEBC0, &H4589C845, &HC4458BF4, &H83F04589, &H758D10EC, &H8DFC8BBC, &HEC839C45, &HA5A5A510, &H8DFC8BA5, &HA550CC75, &HE8A5A5A5, &HFFFF8ACF, &H7D8DF08B, &HA5A5A5CC, &HD8458BA5, &H77F4453B, &H8B087211, &H453BD445, &H330773F0, &HC93340C0, &H570F0EEB, &H130F66C0, &H4D8BEC45, &HEC458BF0, &H8BFC758B, &HF003F855, &H13E87D8B, &HFC7589D1, &H43E44D8B, &H8908EF83
    pvAppendBuffer &H7D89F855, &HFD93BE8, &HFFFF0686, &HD07D8BFF, &H558B03EB, &H8458BF0, &H89CC5D8B, &HD88BC81C, &H8910458B, &H4104CB7C, &H8BD87D8B, &H5589D45D, &HDC558BD8, &H89CC5D89, &H7589D07D, &HE44D89D4, &HFE04D3B, &HFFFE9782, &H84D8BFF, &H7C89C003, &H5E5FFCC1, &HF8C15C89, &H5DE58B5B, &H55000CC2, &H5151EC8B, &H57C0570F, &H66147D8B, &HF845130F, &H8BFC558B, &HFF85F84D, &H458B6B74, &H1045290C
    pvAppendBuffer &H53084529, &H105D8B56, &H342B308B, &HC758903, &H1B04708B, &H8B040374, &HD92B0C5D, &H1B0C5D89, &H8B183BF2, &H575085D, &H7404703B, &H4703B23, &H7771072, &H4D39088B, &H3307760C, &HD23341C9, &H570F0EEB, &H130F66C0, &H558BF845, &HF84D8BFC, &H890C7D8B, &H7489033C, &HC0830403, &H146D8308, &H5EA27501, &H5FC18B5B, &HC25DE58B, &H8B550010, &H84D8BEC, &H5756D233, &H330C7D8B, &H83C78BF6
    pvAppendBuffer &HAB0F3FE0, &H20F883C6, &H33D6430F, &H40F883F2, &HC1D6430F, &H342306EF, &HF95423F9, &H5FC68B04, &H8C25D5E, &HEC8B5500, &H8B08558B, &HC4D8BC2, &H8818E8C1, &HC1C28B01, &H418810E8, &HC1C28B01, &H418808E8, &H3518802, &H8C25D, &H8BEC8B55, &HCA8B0855, &HC5D8B53, &HE8C1C38B, &H758B5618, &H8B068810, &H10E8C1C3, &H8B014688, &H8E8C1C3, &H8B024688, &HC1AC0FC3, &H35E8818, &H8818E8C1
    pvAppendBuffer &HC38B044E, &HAC0FCA8B, &HE8C110C1, &H88C28B10, &HAC0F054E, &H468808D8, &H7568806, &H5E08EBC1, &HCC25D5B, &HEC8B5500, &H8B08558B, &H5D8B53CA, &HFC38B0C, &H5608C1AC, &HC110758B, &HC38B08E8, &H4E881688, &HFCA8B01, &HC110C1AC, &HC38B10E8, &HF024E88, &HC118C2AC, &H568818E8, &H88C38B03, &HE8C1045E, &H5468808, &HE8C1C38B, &H18EBC110, &H88064688, &H5B5E075E, &HCC25D, &H8BEC8B55
    pvAppendBuffer &HD2851455, &H4D8B1F74, &H758B5610, &H7D8B570C, &H2BF12B08, &HE048AF9, &H4880132, &HEA83410F, &H5FF27501, &H10C25D5E, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 34347 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub
