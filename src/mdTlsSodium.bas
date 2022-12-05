Attribute VB_Name = "mdTlsSodium"
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
Private Const MODULE_NAME As String = "mdTlsSodium"

#Const ImplLibSodium = (ASYNCSOCKET_NO_LIBSODIUM = 0)
#Const ImplTlsServer = (ASYNCSOCKET_NO_TLSSERVER = 0)
#Const ImplUseShared = (ASYNCSOCKET_USE_SHARED <> 0)
#Const ImplUseDebugLog = (USE_DEBUG_LOG <> 0)
#Const ImplCaptureTraffic = CLng(ASYNCSOCKET_CAPTURE_TRAFFIC) '--- bitmask: 1 - traffic

'=========================================================================
' API
'=========================================================================

'--- for CryptAcquireContext
Private Const PROV_RSA_AES                              As Long = 24
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_ECC_SIGNATURE                        As Long = 47
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
'--- for CryptCreateHash
Private Const CALG_RC2                                  As Long = &H6602&
Private Const CALG_HMAC                                 As Long = &H8009&
Private Const CALG_SHA_256                              As Long = &H800C&
Private Const CALG_SHA_384                              As Long = &H800D&
Private Const CALG_SHA_512                              As Long = &H800E&
'--- for CryptGet/SetHashParam
Private Const HP_HASHVAL                                As Long = 2
Private Const HP_HMAC_INFO                              As Long = 5
'--- for CryptImportKey
Private Const PLAINTEXTKEYBLOB                          As Long = 8
Private Const CUR_BLOB_VERSION                          As Long = 2
Private Const CRYPT_EXPORTABLE                          As Long = &H1
Private Const CRYPT_IPSEC_HMAC_KEY                      As Long = &H100
'--- for BCrypt keys
Private Const BCRYPT_ECDH_PUBLIC_P256_MAGIC             As Long = &H314B4345
Private Const BCRYPT_ECDH_PRIVATE_P256_MAGIC            As Long = &H324B4345
Private Const BCRYPT_ECDH_PUBLIC_P384_MAGIC             As Long = &H334B4345
Private Const BCRYPT_ECDH_PRIVATE_P384_MAGIC            As Long = &H344B4345
Private Const BCRYPT_ECDSA_PRIVATE_P256_MAGIC           As Long = &H32534345
Private Const BCRYPT_ECDSA_PRIVATE_P384_MAGIC           As Long = &H34534345
Private Const BCRYPT_ECDSA_PRIVATE_P521_MAGIC           As Long = &H36534345
Private Const BCRYPT_ECDH_PUBLIC_GENERIC_MAGIC          As Long = &H504B4345
Private Const BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC         As Long = &H564B4345
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_RSA_SSA_PSS                         As String = "1.2.840.113549.1.1.10"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpStr As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
'--- msvbvm60
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
#If ImplTlsServer Or Not ImplLibSodium Then
    Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
    '--- version
    Private Declare Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
    Private Declare Function VerQueryValue Lib "version" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Any, puLen As Long) As Long
#End If
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long, dwBufLen As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
'--- crypt32
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
'--- bcrypt
#If ImplTlsServer Or Not ImplLibSodium Then
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptImportKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptSignHash Lib "bcrypt" (ByVal hKey As Long, pPaddingInfo As Any, pbInput As Any, ByVal cbInput As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptExportKey Lib "bcrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByRef cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSecretAgreement Lib "bcrypt" (ByVal hPrivKey As Long, ByVal hPubKey As Long, ByRef phSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroySecret Lib "bcrypt" (ByVal hSecret As Long) As Long
    Private Declare Function BCryptDeriveKey Lib "bcrypt" (ByVal hSharedSecret As Long, ByVal pwszKDF As Long, ByVal pParameterList As Long, ByVal pbDerivedKey As Long, ByVal cbDerivedKey As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByRef hKey As Long, ByVal dwLength As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinalizeKeyPair Lib "bcrypt" (ByVal hKey As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
#End If
#If Not ImplLibSodium Then
    Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, hKey As Long, ByVal pbKeyObject As Long, ByVal cbKeyObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDecrypt Lib "bcrypt" (ByVal hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
#End If
'--- libsodium
#If ImplLibSodium Then
    Private Declare Function sodium_init Lib "libsodium" () As Long
    Private Declare Function crypto_scalarmult_curve25519 Lib "libsodium" (lpOut As Any, lpConstN As Any, lpConstP As Any) As Long
    Private Declare Function crypto_scalarmult_curve25519_base Lib "libsodium" (lpOut As Any, lpConstN As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_is_available Lib "libsodium" () As Long
    Private Declare Function crypto_aead_aes256gcm_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
#End If

Private Type CRYPT_DATA_BLOB
    cbData              As Long
    pbData              As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId            As Long
    Parameters          As CRYPT_DATA_BLOB
End Type

Private Type CRYPT_ECC_PRIVATE_KEY_INFO
    dwVersion           As Long
    PrivateKey          As CRYPT_DATA_BLOB
    szCurveOid          As Long
    PublicKey           As CRYPT_DATA_BLOB
End Type

Private Type CRYPT_PRIVATE_KEY_INFO
    dwVersion           As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_DATA_BLOB
    pAttributes         As Long
End Type

Private Type BCRYPT_PSS_PADDING_INFO
    pszAlgId            As Long
    cbSalt              As Long
End Type

Private Type BCRYPT_ECCKEY_BLOB
    dwMagic             As Long
    cbKey               As Long
    Buffer(0 To 255)    As Byte
End Type
Private Const sizeof_BCRYPT_ECCKEY_BLOB As Long = 8

Private Type BLOBHEADER
    bType               As Byte
    bVersion            As Byte
    reserved            As Integer
    aiKeyAlg            As Long
    cbKeySize           As Long
    Buffer(0 To 255)    As Byte
End Type
Private Const sizeof_BLOBHEADER As Long = 12

Private Type HMAC_INFO
    HashAlgid           As Long
    pbInnerString       As Long
    cbInnerString       As Long
    pbOuterString       As Long
    cbOuterString       As Long
End Type

Private Type BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    cbSize              As Long
    dwInfoVersion       As Long
    pbNonce             As Long
    cbNonce             As Long
    pbAuthData          As Long
    cbAuthData          As Long
    pbTag               As Long
    cbTag               As Long
    pbMacContext        As Long
    cbMacContext        As Long
    cbAAD               As Long
    lPad                As Long
    cbData(7)           As Byte
    dwFlags             As Long
    lPad2               As Long
End Type

Private Type CERT_ECC_SIGNATURE
    rValue             As CRYPT_DATA_BLOB
    sValue             As CRYPT_DATA_BLOB
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS                             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|21|Decryption failed|22|Record overflow|30|Decompression failure|40|Handshake failure|41|No certificate|42|Bad certificate|43|Unsupported certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown certificate authority|50|Decode error|51|Decrypt error|70|Protocol version|71|Insufficient security|80|Internal error|90|User canceled|100|No renegotiation|109|Missing extension|110|Unsupported extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_VL_STATES                             As String = "0|New|1|Closed|2|HandshakeStart|3|ExpectServerHello|4|ExpectExtensions|5|ExpectServerFinished|6|ExpectClientHello|7|ExpectClientFinished|8|PostHandshake|9|Shutdown"
Private Const STR_VL_MESSAGE_NAMES                      As String = "1|client_hello|2|server_hello|4|new_session_ticket|5|end_of_early_data|8|encrypted_extensions|11|certificate|12|server_key_exchange|13|certificate_request|14|server_hello_done|15|certificate_verify|16|client_key_exchange|20|finished|24|key_update|25|compressed_certificate|254|message_hash"
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
Private Const TLS_HANDSHAKE_KEY_UPDATE                  As Long = 24
'Private Const TLS_HANDSHAKE_COMPRESSED_CERTIFICATE      As Long = 25
Private Const TLS_HANDSHAKE_MESSAGE_HASH                As Long = 254
'--- TLS Extensions from https://www.iana.org/assignments/tls-extensiontype-values/tls-extensiontype-values.xhtml
Private Const TLS_EXTENSION_SERVER_NAME                 As Long = 0
'Private Const TLS_EXTENSION_STATUS_REQUEST              As Long = 5
Private Const TLS_EXTENSION_SUPPORTED_GROUPS            As Long = 10
Private Const TLS_EXTENSION_EC_POINT_FORMAT             As Long = 11
Private Const TLS_EXTENSION_SIGNATURE_ALGORITHMS        As Long = 13
Private Const TLS_EXTENSION_ALPN                        As Long = 16
Private Const TLS_EXTENSION_EXTENDED_MASTER_SECRET      As Long = 23
Private Const TLS_EXTENSION_SUPPORTED_VERSIONS          As Long = 43
Private Const TLS_EXTENSION_COOKIE                      As Long = 44
Private Const TLS_EXTENSION_CERTIFICATE_AUTHORITIES     As Long = 47
Private Const TLS_EXTENSION_POST_HANDSHAKE_AUTH         As Long = 49
Private Const TLS_EXTENSION_KEY_SHARE                   As Long = 51
Private Const TLS_EXTENSION_RENEGOTIATION_INFO          As Long = &HFF01
'--- TLS Cipher Suites from http://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-4
Private Const TLS_CS_AES_256_GCM_SHA384                 As Long = &H1302
Private Const TLS_CS_CHACHA20_POLY1305_SHA256           As Long = &H1303
Private Const TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384 As Long = &HC02C&
Private Const TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384  As Long = &HC030&
Private Const TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA8&
Private Const TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA9&
Private Const TLS_CS_RSA_WITH_AES_256_GCM_SHA384        As Long = &H9D
'--- TLS Supported Groups from https://www.iana.org/assignments/tls-parameters/tls-parameters.xhtml#tls-parameters-8
Private Const TLS_GROUP_SECP256R1                       As Long = 23
Private Const TLS_GROUP_SECP384R1                       As Long = 24
Private Const TLS_GROUP_X25519                          As Long = 29
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
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA256          As Long = &H809
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA384          As Long = &H80A
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA512          As Long = &H80B
Private Const TLS_ALERT_LEVEL_WARNING                   As Long = 1
Private Const TLS_ALERT_LEVEL_FATAL                     As Long = 2
Private Const TLS_COMPRESS_NULL                         As Long = 0
Private Const TLS_SERVER_NAME_TYPE_HOSTNAME             As Long = 0
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = (TLS_MAX_PLAINTEXT_RECORD_SIZE + 1 + 255) '-- 1 byte content type + 255 bytes AEAD padding
Private Const TLS_HELLO_RANDOM_SIZE                     As Long = 32
Private Const TLS_AAD_SIZE                              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const TLS_LEGACY_AAD_SIZE                       As Long = 13    '--- for TLS 1.2
Private Const TLS_VERIFY_DATA_SIZE                      As Long = 12
'--- crypto constants
Private Const LNG_X25519_KEYSZ                          As Long = 32
Private Const LNG_SECP256R1_KEYSZ                       As Long = 32
Private Const LNG_SECP384R1_KEYSZ                       As Long = 48
Private Const LNG_SECP521R1_KEYSZ                       As Long = 66 '--- CEIL(521/8)
Private Const LNG_MD5_HASHSZ                            As Long = 16
Private Const LNG_SHA1_HASHSZ                           As Long = 20
Private Const LNG_SHA224_HASHSZ                         As Long = 28
Private Const LNG_SHA256_HASHSZ                         As Long = 32
Private Const LNG_SHA384_HASHSZ                         As Long = 48
Private Const LNG_SHA512_HASHSZ                         As Long = 64
Private Const LNG_CHACHA20_KEYSZ                        As Long = 32
Private Const LNG_CHACHA20POLY1305_IVSZ                 As Long = 12
Private Const LNG_CHACHA20POLY1305_TAGSZ                As Long = 16
Private Const LNG_AES256_KEYSZ                          As Long = 32
Private Const LNG_AESGCM_IVSZ                           As Long = 12
Private Const LNG_AESGCM_TAGSZ                          As Long = 16
'--- errors
Private Const ERR_CONNECTION_CLOSED                     As String = "Connection closed"
Private Const ERR_GENER_KEYPAIR_FAILED                  As String = "Failed generating key pair (%1)"
Private Const ERR_UNSUPPORTED_EXCH_GROUP                As String = "Unsupported exchange group (%1)"
Private Const ERR_UNSUPPORTED_CIPHER_SUITE              As String = "Unsupported cipher suite (%1)"
Private Const ERR_UNSUPPORTED_SIGNATURE_SCHEME          As String = "Unsupported signature scheme (%1)"
Private Const ERR_UNSUPPORTED_CERTIFICATE               As String = "Unsupported certificate"
Private Const ERR_UNSUPPORTED_PRIVATE_KEY               As String = "Unsupported private key"
Private Const ERR_UNSUPPORTED_CURVE_TYPE                As String = "Unsupported curve type (%1)"
Private Const ERR_UNSUPPORTED_PROTOCOL                  As String = "Invalid protocol version"
Private Const ERR_ENCRYPTION_FAILED                     As String = "Encryption failed"
Private Const ERR_SIGNATURE_FAILED                      As String = "Certificate signature failed (%1)"
Private Const ERR_DECRYPTION_FAILED                     As String = "Decryption failed"
Private Const ERR_SERVER_HANDSHAKE_FAILED               As String = "Handshake verification failed"
Private Const ERR_RECORD_MAC_FAILED                     As String = "MAC verification failed"
Private Const ERR_HELLO_RETRY_FAILED                    As String = "HelloRetryRequest failed"
Private Const ERR_NEGOTIATE_SIGNATURE_FAILED            As String = "Negotiate signature type failed"
Private Const ERR_CALL_FAILED                           As String = "Call failed (%1)"
Private Const ERR_RECORD_TOO_BIG                        As String = "Record size too big"
Private Const ERR_FATAL_ALERT                           As String = "Received fatal alert"
Private Const ERR_UNEXPECTED_RECORD_TYPE                As String = "Unexpected record type (%1)"
Private Const ERR_UNEXPECTED_MSG_TYPE                   As String = "Unexpected message type for %1 state (%2)"
Private Const ERR_UNEXPECTED_EXTENSION                  As String = "Unexpected extension (%1)"
Private Const ERR_INVALID_STATE_HANDSHAKE               As String = "Invalid state for handshake content (%1)"
Private Const ERR_INVALID_REMOTE_KEY                    As String = "Invalid remote key size"
Private Const ERR_INVALID_SIZE_EXTENSION                As String = "Invalid data size for %1"
Private Const ERR_NO_HANDSHAKE_MESSAGES                 As String = "Missing handshake messages"
Private Const ERR_NO_PREVIOUS_SECRET                    As String = "Missing previous %1"
Private Const ERR_NO_REMOTE_RANDOM                      As String = "Missing remote random"
Private Const ERR_NO_SERVER_CERTIFICATE                 As String = "Missing server certificate"
Private Const ERR_NO_SUPPORTED_CIPHER_SUITE             As String = "Missing supported ciphersuite"
Private Const ERR_NO_PRIVATE_KEY                        As String = "Missing server private key"
Private Const ERR_NO_SERVER_COMPILED                    As String = "Server-side TLS not compiled (ASYNCSOCKET_NO_TLSSERVER = 1)"
'--- numeric
Private Const LNG_OUT_OF_MEMORY                         As Long = 8
Private Const MAX_RSA_KEY                               As Long = 8192 '--- in bits

Private m_uData                     As UcsCryptoData
Private m_baHelloRetryRandom()      As Byte
Public g_oRequestSocket             As Object

Private Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsIgnoreServerCertificateErrors = 2 ^ 2
    ucsTlsSupportAll = ucsTlsSupportTls12 Or ucsTlsSupportTls13
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
    ucsTlsStateExpectClientFinished = 7
#End If
    ucsTlsStatePostHandshake = 8
    ucsTlsStateShutdown = 9
End Enum

Private Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoExchX25519 = 1
    ucsTlsAlgoExchSecp256r1
    ucsTlsAlgoExchSecp384r1
    ucsTlsAlgoExchSecp521r1
    ucsTlsAlgoExchCertificate
    '--- ciphers
    ucsTlsAlgoBulkChacha20Poly1305 = 11
    ucsTlsAlgoBulkAesGcm256
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
    '--- handshake
    LocalSessionID()    As Byte
    LocalExchRandom()   As Byte
    LocalExchPrivate()  As Byte
    LocalExchPublic()   As Byte
    LocalExchRsaEncrPriv() As Byte
    LocalCertificates   As Collection
    LocalPrivateKey     As Collection
    LocalSignatureScheme As Long
    LocalLegacyVerifyData() As Byte
    RemoteSessionID()   As Byte
    RemoteExchRandom()  As Byte
    RemoteExchPublic()  As Byte
    RemoteCertificates  As Collection
    RemoteExtensions    As Collection
    RemoteTickets       As Collection
    RemoteSupportedGroups As Collection
    RemoteCertStatuses  As Collection
    '--- crypto settings
    ProtocolVersion     As Long
    ExchGroup           As Long
    ExchAlgo            As UcsTlsCryptoAlgorithmsEnum
    CipherSuite         As Long
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
End Type

Private Type UcsCryptoData
    hProv               As Long
    hResult             As Long
    ApiSource           As String
End Type

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
        pvTlsGetRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
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
    Err.Raise vbObjectError, , ERR_NO_SERVER_COMPILED
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
        pvTlsGetRandom .LocalExchRandom, TLS_HELLO_RANDOM_SIZE
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
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lSize <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Input" & vbCrLf & TlsDesignDumpArray(baInput, Size:=lSize)
            End If
        #End If
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
            If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
                pvTlsSetLastError uCtx, vbObjectError, MODULE_NAME & "." & FUNC_NAME, .LastError, .LastAlertCode
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
        #If (ImplCaptureTraffic And 1) <> 0 Then
            If lSize <> 0 Then
                .TrafficDump.Add FUNC_NAME & ".Input (undecrypted)" & vbCrLf & TlsDesignDumpArray(baInput, Size:=lSize)
            End If
        #End If
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
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
    If lState <= UBound(vTexts) Then
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
    If lMessageType <= UBound(vTexts) Then
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
    If lExtType <= UBound(vTexts) Then
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
#If ImplTlsServer Then
            ElseIf pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                pvTlsSetupExchGroup uCtx, TLS_GROUP_SECP256R1
            ElseIf pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                pvTlsSetupExchGroup uCtx, TLS_GROUP_SECP384R1
#End If
            End If
        End If
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Handshake Header
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_CLIENT_HELLO
            pvBufferWriteBlockStart uOutput, Size:=3
                pvBufferWriteLong uOutput, TLS_LOCAL_LEGACY_VERSION, Size:=2
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
#If ImplTlsServer Then
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP256R1 Then
                                    pvBufferWriteLong uOutput, TLS_GROUP_SECP256R1, Size:=2
                                End If
                            End If
                            If pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP384R1 Then
                                    pvBufferWriteLong uOutput, TLS_GROUP_SECP384R1, Size:=2
                                End If
                            End If
#End If
                        pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockEnd uOutput
                    '--- Extension - EC Point Formats
                    pvArrayByte baTemp, 0, TLS_EXTENSION_EC_POINT_FORMAT, 0, 2, 1, 0
                    pvBufferWriteArray uOutput, baTemp          '--- uncompressed only
                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- Extension - Extended Master Secret
                        pvArrayByte baTemp, 0, TLS_EXTENSION_EXTENDED_MASTER_SECRET, 0, 0
                        pvBufferWriteArray uOutput, baTemp      '--- supported
                        '--- Extension - Renegotiation Info
                        pvBufferWriteLong uOutput, TLS_EXTENSION_RENEGOTIATION_INFO, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            pvBufferWriteBlockStart uOutput
                                pvBufferWriteArray uOutput, .LocalLegacyVerifyData
                            pvBufferWriteBlockEnd uOutput
                        pvBufferWriteBlockEnd uOutput
                    End If
                    '--- Extension - Signature Algorithms
                    pvBufferWriteLong uOutput, TLS_EXTENSION_SIGNATURE_ALGORITHMS, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvBufferWriteBlockStart uOutput, Size:=2
                            For Each vElem In Array(TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512, _
                                                    TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
                                                    TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512, _
                                                    TLS_SIGNATURE_RSA_PKCS1_SHA224, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, _
                                                    TLS_SIGNATURE_RSA_PKCS1_SHA512, TLS_SIGNATURE_RSA_PKCS1_SHA1, TLS_SIGNATURE_ECDSA_SHA1)
                                pvBufferWriteLong uOutput, vElem, Size:=2
                            Next
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
                        If .HelloRetryRequest Then
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
#If ImplTlsServer Then
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
#End If
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
            pvTlsGetHandshakeHash uCtx, baHandshakeHash
#If ImplTlsServer Then
            If .CertRequestSignatureScheme > 0 Then
                '--- Client Certificate Verify
                lMessagePos = uOutput.Size
                pvBufferWriteLong uOutput, TLS_HANDSHAKE_CERTIFICATE_VERIFY
                pvBufferWriteBlockStart uOutput, Size:=3
                    pvBufferWriteLong uOutput, .CertRequestSignatureScheme, Size:=2
                    pvBufferWriteBlockStart uOutput, Size:=2
                        pvArrayWriteEOF .HandshakeMessages.Data, .HandshakeMessages.Size
                        pvTlsSignatureSign baSignature, .LocalPrivateKey, .CertRequestSignatureScheme, .HandshakeMessages.Data
                        pvBufferWriteArray uOutput, baSignature
                    pvBufferWriteBlockEnd uOutput
                pvBufferWriteBlockEnd uOutput
                pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
            End If
#End If
        pvBufferWriteRecordEnd uOutput, uCtx
        '--- Legacy Change Cipher Spec
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, uCtx
            pvBufferWriteLong uOutput, 1
        pvBufferWriteRecordEnd uOutput, uCtx
        pvTlsDeriveLegacySecrets uCtx, baHandshakeHash
        '--- Record Header
        pvBufferWriteRecordStart uOutput, TLS_CONTENT_TYPE_HANDSHAKE, uCtx
            '--- Client Handshake Finished
            lMessagePos = uOutput.Size
            pvBufferWriteLong uOutput, TLS_HANDSHAKE_FINISHED
            pvBufferWriteBlockStart uOutput, Size:=3
                pvTlsGetHandshakeHash uCtx, baHandshakeHash
                pvTlsKdfLegacyPrf baVerifyData, .DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                pvBufferWriteArray uOutput, baVerifyData
                .LocalLegacyVerifyData = baVerifyData
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
#If ImplTlsServer Then
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
                            pvTlsSignatureSign baSignature, .LocalPrivateKey, .CertRequestSignatureScheme, uVerify.Data
                            pvBufferWriteArray uOutput, baSignature
                        pvBufferWriteBlockEnd uOutput
                    pvBufferWriteBlockEnd uOutput
                    pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
                End If
                '--- Record Type
                pvBufferWriteLong uOutput, TLS_CONTENT_TYPE_HANDSHAKE
            pvBufferWriteRecordEnd uOutput, uCtx
        End If
#End If
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

#If ImplTlsServer Then
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
                    pvBufferWriteArray uOutput, .RemoteSessionID
                pvBufferWriteBlockEnd uOutput
                '--- Cipher Suite
                pvBufferWriteLong uOutput, IIf(.HelloRetryRequest, .HelloRetryCipherSuite, .CipherSuite), Size:=2
                '--- Legacy Compression Method
                pvBufferWriteLong uOutput, TLS_COMPRESS_NULL
                '--- Extensions
                pvBufferWriteBlockStart uOutput, Size:=2
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
                pvBufferWriteBlockEnd uOutput
            pvBufferWriteBlockEnd uOutput
            pvTlsAppendHandshakeHash uCtx, uOutput.Data, lMessagePos, uOutput.Size - lMessagePos
        pvBufferWriteRecordEnd uOutput, uCtx
        If .HelloRetryRequest Or .HelloRetryCipherSuite = 0 Then
            '--- Legacy Change Cipher Spec
            pvArrayByte baTemp, TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1
            pvBufferWriteArray uOutput, baTemp
        End If
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
                pvBufferWriteLong uOutput, .LocalSignatureScheme, Size:=2
                pvBufferWriteBlockStart uOutput, Size:=2
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvBufferWriteString uVerify, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0)
                    pvBufferWriteArray uVerify, baHandshakeHash
                    pvBufferWriteEOF uVerify
                    pvTlsSignatureSign baSignature, .LocalPrivateKey, .LocalSignatureScheme, uVerify.Data
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
#End If

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
    
    On Error GoTo EH
    With uCtx
    Do While uInput.Pos + 5 <= uInput.Size
        lRecordPos = uInput.Pos
        pvBufferReadLong uInput, lRecordType
        pvBufferReadLong uInput, lRecordProtocol, Size:=2
        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lRecordSize
            If lRecordSize > IIf(lRecordType = TLS_CONTENT_TYPE_APPDATA, TLS_MAX_ENCRYPTED_RECORD_SIZE, TLS_MAX_PLAINTEXT_RECORD_SIZE) Then
                GoTo RecordTooBig
            End If
            If uInput.Pos + lRecordSize > uInput.Size Then
                '--- back off and bail out early
                uInput.Stack.Remove 1
                uInput.Pos = lRecordPos
                Exit Do
            End If
            '--- try to decrypt record
            If pvArraySize(.RemoteTrafficKey) > 0 And lRecordSize > .TagSize Then
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
                    pvBufferWriteLong uAad, lEnd - uInput.Pos, Size:=2
                    Debug.Assert uAad.Size = TLS_LEGACY_AAD_SIZE
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
                    '--- trim zero padding at the end of decrypted record
                    Do While lEnd > uInput.Pos
                        lEnd = lEnd - 1
                        If uInput.Data(lEnd) <> 0 Then
                            Exit Do
                        End If
                    Loop
                    lRecordType = uInput.Data(lEnd)
                End If
            Else
                lEnd = uInput.Pos + lRecordSize
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
                    .RemoteTrafficKey = .RemoteLegacyNextTrafficKey
                    .RemoteTrafficIV = .RemoteLegacyNextTrafficIV
                    .RemoteTrafficSeqNo = 0
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
                    If .LastAlertCode = uscTlsAlertCloseNotify Then
                        pvTlsSetLastError uCtx, AlertCode:=uscTlsAlertCloseNotify
                    End If
                End Select
            Case TLS_CONTENT_TYPE_HANDSHAKE
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
                If .State < ucsTlsStatePostHandshake Then
                    GoTo UnexpectedRecordType
                End If
                pvBufferWriteBlob .DecrBuffer, VarPtr(uInput.Data(uInput.Pos)), lEnd - uInput.Pos
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
RecordTooBig:
    sError = ERR_RECORD_TOO_BIG
    eAlertCode = uscTlsAlertDecodeError
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
    sError = ERR_RECORD_TOO_BIG
    eAlertCode = uscTlsAlertUnexpectedMessage
    GoTo QH
RecordMacFailed:
    sError = ERR_RECORD_MAC_FAILED
    eAlertCode = uscTlsAlertBadRecordMac
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshake(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshake"
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
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
    
    On Error GoTo EH
    With uCtx
    Do While uInput.Pos < lInputEnd
        lMessagePos = uInput.Pos
        pvBufferReadLong uInput, lMessageType
        pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lMessageSize
            If uInput.Pos + lMessageSize > lInputEnd Then
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
                    Do While uInput.Pos < lBlockEnd
                        pvBufferReadLong uInput, lExtType, Size:=2
                        #If ImplUseDebugLog Then
'                            DebugLog MODULE_NAME, FUNC_NAME, "EncryptedExtensions " & pvTlsGetExtensionName(lExtType)
                        #End If
                        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                            lExtEnd = uInput.Pos + lExtSize
                            Select Case lExtType
                            Case TLS_EXTENSION_ALPN
                                pvBufferReadBlockStart uInput, Size:=2
                                    pvBufferReadBlockStart uInput, BlockSize:=lStringSize
                                        pvBufferReadString uInput, .AlpnNegotiated, lStringSize
                                    pvBufferReadBlockEnd uInput
                                pvBufferReadBlockEnd uInput
                            Case TLS_EXTENSION_SUPPORTED_GROUPS
                                If lExtSize < 2 Then
                                    GoTo InvalidSize
                                End If
                                Set .RemoteSupportedGroups = New Collection
                                Do While uInput.Pos < lExtEnd
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
                            uInput.Pos = uInput.Pos + lCertSize '--- skip RemoteCertReqContext
                        pvBufferReadBlockEnd uInput
                    End If
                    Set .RemoteCertificates = New Collection
                    pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lCertSize
                        lCertEnd = uInput.Pos + lCertSize
                        Do While uInput.Pos < lCertEnd
                            pvBufferReadBlockStart uInput, Size:=3, BlockSize:=lCertSize
                                pvBufferReadArray uInput, baCert, lCertSize
                                .RemoteCertificates.Add baCert
                            pvBufferReadBlockEnd uInput
                            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lCertSize
                                    '--- certificate extensions -> skip
                                    uInput.Pos = uInput.Pos + lCertSize
                                pvBufferReadBlockEnd uInput
                            End If
                        Loop
                    pvBufferReadBlockEnd uInput
                Case TLS_HANDSHAKE_CERTIFICATE_VERIFY
                    pvBufferReadLong uInput, lSignatureScheme, Size:=2
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lCertSize
                        pvBufferReadArray uInput, baSignature, lCertSize
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
                    If InStrB(uVerify.Data, baMessage) = 0 Then
                        GoTo ServerHandshakeFailed
                    End If
                    .State = ucsTlsStatePostHandshake
                Case TLS_HANDSHAKE_SERVER_KEY_EXCHANGE
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
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
                            pvBufferReadArray uInput, .RemoteExchPublic, lSignatureSize
                        pvBufferReadBlockEnd uInput
                        lSignSize = uInput.Pos - lSignPos
                        '--- signature
                        pvBufferReadLong uInput, lSignatureScheme, Size:=2
                        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSignatureSize
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
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case TLS_HANDSHAKE_SERVER_HELLO_DONE
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        .State = ucsTlsStateExpectServerFinished
                        uInput.Pos = uInput.Pos + lMessageSize
                    Else
                        GoTo UnexpectedMessageType
                    End If
#If ImplTlsServer Then
                Case TLS_HANDSHAKE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, uInput, sError, eAlertCode) Then
                        GoTo QH
                    End If
#End If
                Case Else
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
                Case TLS_HANDSHAKE_FINISHED
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        pvBufferReadArray uInput, baMessage, lMessageSize
                        pvTlsGetHandshakeHash uCtx, baHandshakeHash
                        pvTlsKdfLegacyPrf uVerify.Data, .DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, TLS_VERIFY_DATA_SIZE
                        If InStrB(uVerify.Data, baMessage) = 0 Then
                            GoTo ServerHandshakeFailed
                        End If
                        .State = ucsTlsStatePostHandshake
                        pvTlsResetHandshakeHash uCtx
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case Else
                    GoTo UnexpectedMessageType
                End Select
#If ImplTlsServer Then
            Case ucsTlsStateExpectClientHello
                Select Case lMessageType
                Case TLS_HANDSHAKE_CLIENT_HELLO
                    If Not pvTlsParseHandshakeClientHello(uCtx, uInput, uInput.Pos + lMessageSize, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If .ExchGroup = 0 Or .CipherSuite = 0 Then
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
            Case ucsTlsStateExpectClientFinished
                Select Case lMessageType
                Case TLS_HANDSHAKE_FINISHED
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvTlsHkdfExpandLabel baTemp, .DigestAlgo, .RemoteTrafficSecret, "finished", baEmpty, .DigestSize
                    pvTlsHkdfExtract uVerify.Data, .DigestAlgo, baTemp, baHandshakeHash
                    If InStrB(uVerify.Data, baMessage) = 0 Then
                        GoTo ServerHandshakeFailed
                    End If
                    .State = ucsTlsStatePostHandshake
                Case Else
                    GoTo UnexpectedMessageType
                End Select
                '--- post-process ucsTlsStateExpectClientFinished
                If .State = ucsTlsStatePostHandshake Then
                    pvTlsGetHandshakeHash uCtx, baHandshakeHash
                    pvTlsDeriveApplicationSecrets uCtx, baHandshakeHash
                    pvTlsResetHandshakeHash uCtx
                    Set .RemoteTickets = New Collection
                End If
#End If
            Case ucsTlsStatePostHandshake
                Select Case lMessageType
                Case TLS_HANDSHAKE_HELLO_REQUEST '--- Hello Request
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        Debug.Assert lMessageSize = 0
                        #If ImplUseDebugLog Then
                            DebugLog MODULE_NAME, FUNC_NAME, "Received Hello Request. Will renegotiate"
                        #End If
                        .State = ucsTlsStateExpectServerHello
                        '--- renegotiate ephemeral keys too
                        .ExchGroup = 0
                        .CipherSuite = 0
                        pvTlsBuildClientHello uCtx, .SendBuffer
                    Else
                        GoTo UnexpectedMessageType
                    End If
                Case TLS_HANDSHAKE_NEW_SESSION_TICKET
                    pvBufferReadArray uInput, baMessage, lMessageSize
                    If Not .RemoteTickets Is Nothing Then
                        .RemoteTickets.Add baMessage
                    End If
                Case TLS_HANDSHAKE_KEY_UPDATE
                    #If ImplUseDebugLog Then
                        DebugLog MODULE_NAME, FUNC_NAME, "Received TLS_HANDSHAKE_KEY_UPDATE"
                    #End If
                    If lMessageSize = 1 Then
                        lRequestUpdate = uInput.Data(uInput.Pos)
                    Else
                        lRequestUpdate = -1
                    End If
                    pvTlsDeriveKeyUpdate uCtx, lRequestUpdate <> 0
                    If lRequestUpdate <> 0 Then
                        '--- ack by TLS_HANDSHAKE_KEY_UPDATE w/ update_not_requested(0)
                        pvArrayByte baTemp, TLS_HANDSHAKE_KEY_UPDATE, 0, 0, 1, 0
                        pvTlsBuildApplicationData uCtx, .SendBuffer, baTemp, 0, UBound(baTemp) + 1, TLS_CONTENT_TYPE_APPDATA
                    End If
                    uInput.Pos = uInput.Pos + lMessageSize
#If ImplTlsServer Then
                Case TLS_HANDSHAKE_CERTIFICATE_REQUEST
                    If Not pvTlsParseHandshakeCertificateRequest(uCtx, uInput, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    pvTlsBuildClientHandshakeFinished uCtx, .SendBuffer
                    pvTlsResetHandshakeHash uCtx
#End If
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
    sError = Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
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
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeServerHello(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Const FUNC_NAME     As String = "pvTlsParseHandshakeServerHello"
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lLegacyVersion  As Long
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
    If pvArraySize(m_baHelloRetryRandom) = 0 Then
        pvTlsGetHelloRetryRandom m_baHelloRetryRandom
    End If
    With uCtx
        .ProtocolVersion = lRecordProtocol
        pvBufferReadLong uInput, lLegacyVersion, Size:=2
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
                Do While uInput.Pos < lBlockEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "ServerHello " & pvTlsGetExtensionName(lExtType)
                    #End If
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        lExtEnd = uInput.Pos + lExtSize
                        Select Case lExtType
                        Case TLS_EXTENSION_KEY_SHARE
                            .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadLong uInput, lExchGroup, Size:=2
                            pvTlsSetupExchGroup uCtx, lExchGroup
                            #If ImplUseDebugLog Then
                                DebugLog MODULE_NAME, FUNC_NAME, "With exchange group " & pvTlsGetExchGroupName(.ExchGroup)
                            #End If
                            If .HelloRetryRequest Then
                                .HelloRetryExchGroup = lExchGroup
                            Else
                                If lExtSize <= 4 Then
                                    GoTo InvalidSize
                                End If
                                pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lPublicSize
                                    pvBufferReadArray uInput, .RemoteExchPublic, lPublicSize
                                pvBufferReadBlockEnd uInput
                            End If
                        Case TLS_EXTENSION_SUPPORTED_VERSIONS
                            If lExtSize <> 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadLong uInput, .ProtocolVersion, Size:=2
                        Case TLS_EXTENSION_COOKIE
                            If Not .HelloRetryRequest Then
                                GoTo UnexpectedExtension
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lCookieSize
                                pvBufferReadArray uInput, .HelloRetryCookie, lCookieSize
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_ALPN
                            pvBufferReadBlockStart uInput, Size:=2
                                pvBufferReadBlockStart uInput, BlockSize:=lNameSize
                                    pvBufferReadString uInput, .AlpnNegotiated, lNameSize
                                pvBufferReadBlockEnd uInput
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
    sError = Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
    eAlertCode = uscTlsAlertDecodeError
    GoTo QH
UnexpectedExtension:
    sError = Replace(ERR_UNEXPECTED_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsParseHandshakeClientHello(uCtx As UcsTlsContext, uInput As UcsBuffer, ByVal lInputEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
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
    Dim eExchAlgo       As UcsTlsCryptoAlgorithmsEnum
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lProtocolVersion As Long
    Dim lSignatureScheme As Long
    Dim cCipherPrefs    As Collection
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baPrivKey()     As Byte
    Dim uKeyInfo        As UcsKeyInfo
    Dim lNameType       As Long
    Dim lNameSize       As Long
    Dim sName           As String
    Dim cAlpnPrefs      As Collection
    Dim lAlpnPref       As Long
    Dim lKeySize        As Long
    
    On Error GoTo EH
    Set cCipherPrefs = New Collection
    For Each vElem In pvTlsGetSortedCipherSuites(ucsTlsSupportTls13)
        cCipherPrefs.Add cCipherPrefs.Count, "#" & vElem
    Next
    lCipherPref = 1000
    With uCtx
        If SearchCollection(.LocalPrivateKey, 1, RetVal:=baPrivKey) Then
            If Not pvAsn1DecodePrivateKey(baPrivKey, uKeyInfo) Then
                GoTo UnsupportedCertificate
            End If
        End If
        .ProtocolVersion = lRecordProtocol
        pvBufferReadLong uInput, lLegacyVersion, Size:=2
        If lLegacyVersion < TLS_PROTOCOL_VERSION_TLS12 Then
            GoTo UnsupportedProtocol
        End If
        pvBufferReadArray uInput, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE
        pvBufferReadBlockStart uInput, BlockSize:=lSize
            pvBufferReadArray uInput, .RemoteSessionID, lSize
        pvBufferReadBlockEnd uInput
        pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
            lEnd = uInput.Pos + lSize
            Do While uInput.Pos < lEnd
                pvBufferReadLong uInput, lIdx, Size:=2
                If .HelloRetryRequest And lIdx <> .HelloRetryCipherSuite Then
                    lIdx = 0
                End If
                If SearchCollection(cCipherPrefs, "#" & lIdx, RetVal:=vElem) Then
                    If vElem < lCipherPref Then
                        lCipherSuite = lIdx
                        lCipherPref = vElem
                    End If
                End If
            Loop
        pvBufferReadBlockEnd uInput
        If lCipherSuite = 0 Then
            GoTo NoCipherSuite
        End If
        pvTlsSetupCipherSuite uCtx, lCipherSuite
        pvBufferReadBlockStart uInput
            pvBufferReadLong uInput, lLegacyCompress
        pvBufferReadBlockEnd uInput
        Debug.Assert lLegacyCompress = 0
        '--- extensions
        If Not .HelloRetryRequest Then
            Set .RemoteExtensions = New Collection
        End If
        If uInput.Pos < lInputEnd Then
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                Do While uInput.Pos < lEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    #If ImplUseDebugLog Then
'                        DebugLog MODULE_NAME, FUNC_NAME, "ClientHello " & pvTlsGetExtensionName(lExtType)
                    #End If
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        lExtEnd = uInput.Pos + lExtSize
                        Select Case lExtType
                        Case TLS_EXTENSION_SERVER_NAME
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                Do While uInput.Pos < lBlockEnd
                                    pvBufferReadLong uInput, lNameType
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lNameSize
                                        If lNameType = TLS_SERVER_NAME_TYPE_HOSTNAME Then
                                            pvBufferReadString uInput, .SniRequested, lNameSize
                                        Else
                                            uInput.Pos = uInput.Pos + lNameSize '--- skip
                                        End If
                                    pvBufferReadBlockEnd uInput
                                Loop
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_ALPN
                            Set cAlpnPrefs = New Collection
                            For Each vElem In Split(.AlpnProtocols, "|")
                                cAlpnPrefs.Add cAlpnPrefs.Count, "#" & vElem
                            Next
                            lAlpnPref = 1000
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                Do While uInput.Pos < lBlockEnd
                                    pvBufferReadBlockStart uInput, BlockSize:=lNameSize
                                        pvBufferReadString uInput, sName, lNameSize
                                        If SearchCollection(cAlpnPrefs, "#" & sName, RetVal:=vElem) Then
                                            If vElem < lAlpnPref Then
                                                .AlpnNegotiated = sName
                                                lAlpnPref = vElem
                                            End If
                                        End If
                                    pvBufferReadBlockEnd uInput
                                Loop
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_KEY_SHARE
                            .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                Do While uInput.Pos < lBlockEnd
                                    pvBufferReadLong uInput, lExchGroup, Size:=2
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                        If .HelloRetryRequest And lExchGroup <> .HelloRetryExchGroup Then
                                            lExchGroup = 0
                                        End If
                                        Select Case lExchGroup
                                        Case TLS_GROUP_X25519
                                            lKeySize = LNG_X25519_KEYSZ
                                            eExchAlgo = ucsTlsAlgoExchX25519
#If ImplTlsServer Then
                                        Case TLS_GROUP_SECP256R1
                                            lKeySize = LNG_SECP256R1_KEYSZ
                                            eExchAlgo = ucsTlsAlgoExchSecp256r1
                                        Case TLS_GROUP_SECP384R1
                                            lKeySize = LNG_SECP384R1_KEYSZ
                                            eExchAlgo = ucsTlsAlgoExchSecp384r1
#End If
                                        Case Else
                                            eExchAlgo = 0
                                        End Select
                                        Select Case True
                                        Case eExchAlgo = 0, Not pvCryptoIsSupported(eExchAlgo)
                                            lExchGroup = 0
                                            uInput.Pos = uInput.Pos + lBlockSize
                                        End Select
                                        If lExchGroup <> 0 Then
                                            If lBlockSize <> lKeySize Then
                                                GoTo InvalidRemoteKey
                                            End If
                                            pvBufferReadArray uInput, .RemoteExchPublic, lBlockSize
                                            pvTlsSetupExchGroup uCtx, lExchGroup
                                            #If ImplUseDebugLog Then
                                                DebugLog MODULE_NAME, FUNC_NAME, "With exchange group " & pvTlsGetExchGroupName(.ExchGroup)
                                            #End If
                                        End If
                                    pvBufferReadBlockEnd uInput
                                    If lExchGroup <> 0 Then
                                        uInput.Pos = lBlockEnd
                                    End If
                                Loop
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_SIGNATURE_ALGORITHMS
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                Do While uInput.Pos < lExtEnd
                                    pvBufferReadLong uInput, lSignatureScheme, Size:=2
                                    If pvTlsMatchSignatureScheme(uCtx, lSignatureScheme, uKeyInfo) Then
                                        .LocalSignatureScheme = lSignatureScheme
                                        uInput.Pos = lExtEnd
                                    End If
                                Loop
                                If .LocalSignatureScheme = 0 Then
                                    GoTo NegotiateSignatureFailed
                                End If
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_SUPPORTED_GROUPS
                            If lExtSize < 2 Then
                                GoTo InvalidSize
                            End If
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                If uInput.Pos + lBlockSize <> lExtEnd Or lBlockSize = 0 Then
                                    GoTo InvalidSize
                                End If
                                Set .RemoteSupportedGroups = New Collection
                                Do While uInput.Pos < lExtEnd
                                    pvBufferReadLong uInput, lExchGroup, Size:=2
                                    .RemoteSupportedGroups.Add lExchGroup, "#" & lExchGroup
                                Loop
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_SUPPORTED_VERSIONS
                            pvBufferReadBlockStart uInput, BlockSize:=lBlockSize
                                Do While uInput.Pos < lExtEnd
                                    pvBufferReadLong uInput, lProtocolVersion, Size:=2
                                    If lProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                        uInput.Pos = lExtEnd
                                    End If
                                Loop
                            pvBufferReadBlockEnd uInput
                            If lProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 Then
                                GoTo UnsupportedProtocol
                            End If
                            .ProtocolVersion = lProtocolVersion
                        Case Else
                            If .HelloRetryRequest Then
                                If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
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
            pvBufferReadBlockEnd uInput
        End If
        #If ImplUseDebugLog Then
            DebugLog MODULE_NAME, FUNC_NAME, "Using " & pvTlsGetCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName
        #End If
    End With
    '--- success
    pvTlsParseHandshakeClientHello = True
QH:
    Exit Function
UnsupportedCertificate:
    sError = ERR_UNSUPPORTED_CERTIFICATE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
UnsupportedProtocol:
    sError = ERR_UNSUPPORTED_PROTOCOL
    eAlertCode = uscTlsAlertProtocolVersion
    GoTo QH
NoCipherSuite:
    sError = ERR_NO_SUPPORTED_CIPHER_SUITE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
InvalidSize:
    sError = Replace(ERR_INVALID_SIZE_EXTENSION, "%1", pvTlsGetExtensionName(lExtType))
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
    eAlertCode = uscTlsAlertIllegalParameter
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function

#If ImplTlsServer Then
Private Function pvTlsParseHandshakeCertificateRequest(uCtx As UcsTlsContext, uInput As UcsBuffer, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lSignatureScheme As Long
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim uKeyInfo        As UcsKeyInfo
    Dim baDName()       As Byte
    Dim lDnSize         As Long
    Dim baPrivKey()     As Byte
    Dim baSignatureSchemes() As Byte
    Dim lSigPos         As Long
    Dim oCallback       As Object
    Dim bConfirmed      As Boolean
    
    On Error GoTo EH
    With uCtx
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            pvBufferReadBlockStart uInput, Size:=1, BlockSize:=lSize
                pvBufferReadArray uInput, .CertRequestContext, lSize
            pvBufferReadBlockEnd uInput
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                Do While uInput.Pos < lEnd
                    pvBufferReadLong uInput, lExtType, Size:=2
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lExtSize
                        Select Case lExtType
                        Case TLS_EXTENSION_SIGNATURE_ALGORITHMS
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                pvBufferReadArray uInput, baSignatureSchemes, lBlockSize
                            pvBufferReadBlockEnd uInput
                        Case TLS_EXTENSION_CERTIFICATE_AUTHORITIES
                            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lBlockSize
                                lBlockEnd = uInput.Pos + lBlockSize
                                Set .CertRequestCaDn = New Collection
                                Do While uInput.Pos < lBlockEnd
                                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lDnSize
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
            pvBufferReadBlockStart uInput, Size:=1, BlockSize:=lSize
                uInput.Pos = uInput.Pos + lSize '--- skip certificate_types
            pvBufferReadBlockEnd uInput
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                pvBufferReadArray uInput, baSignatureSchemes, lSize
            pvBufferReadBlockEnd uInput
            pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lSize
                lEnd = uInput.Pos + lSize
                Set .CertRequestCaDn = New Collection
                Do While uInput.Pos < lEnd
                    pvBufferReadBlockStart uInput, Size:=2, BlockSize:=lDnSize
                        pvBufferReadArray uInput, baDName, lDnSize
                        .CertRequestCaDn.Add baDName
                    pvBufferReadBlockEnd uInput
                Loop
            pvBufferReadBlockEnd uInput
        End If
        Do
            If SearchCollection(.LocalPrivateKey, 1, RetVal:=baPrivKey) Then
                If Not pvAsn1DecodePrivateKey(baPrivKey, uKeyInfo) Then
                    GoTo UnsupportedPrivateKey
                End If
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
    Exit Function
UnsupportedPrivateKey:
    sError = ERR_UNSUPPORTED_PRIVATE_KEY
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
EH:
    sError = Err.Description & " [" & Err.Source & "]"
    eAlertCode = uscTlsAlertInternalError
End Function
#End If

Private Function pvTlsMatchSignatureScheme(uCtx As UcsTlsContext, ByVal lSignatureScheme As Long, uKeyInfo As UcsKeyInfo) As Boolean
    Dim bHasEnoughBits  As Boolean
    
    '--- PSS w/ SHA512 fails on short key lengths (min PSS size is 2 + lHashSize + lSaltSize where lSaltSize=lHashSize)
    bHasEnoughBits = (uKeyInfo.BitLen + 7) \ 8 > 2 + 2 * pvTlsSignatureHashSize(lSignatureScheme)
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        If bHasEnoughBits And uKeyInfo.AlgoObjId = szOID_RSA_RSA Then
            pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPss)
        End If
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If bHasEnoughBits And uKeyInfo.AlgoObjId = szOID_RSA_SSA_PSS Then
            pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoPaddingPss)
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        If uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P256 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP256R1_SHA256 Then
            pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp256r1)
        ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P384 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP384R1_SHA384 Then
            pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp384r1)
        ElseIf uKeyInfo.AlgoObjId = szOID_ECC_CURVE_P521 And lSignatureScheme = TLS_SIGNATURE_ECDSA_SECP521R1_SHA512 Then
            pvTlsMatchSignatureScheme = pvCryptoIsSupported(ucsTlsAlgoExchSecp521r1)
        End If
    End Select
End Function

Private Sub pvTlsSetupExchGroup(uCtx As UcsTlsContext, ByVal lExchGroup As Long)
    Const FUNC_NAME     As String = "pvTlsSetupExchGroup"
    
    With uCtx
        If .ExchGroup <> lExchGroup Then
            .ExchGroup = lExchGroup
            Select Case lExchGroup
            Case TLS_GROUP_X25519
                .ExchAlgo = ucsTlsAlgoExchX25519
#If ImplLibSodium Then
                If Not pvCryptoEcdhCurve25519MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "Curve25519")
                End If
#Else
                If Not pvBCryptEcdhMakeKey(0, .LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "Curve25519")
                End If
#End If
#If ImplTlsServer Then
            Case TLS_GROUP_SECP256R1, TLS_GROUP_SECP384R1
                .ExchAlgo = IIf(lExchGroup = TLS_GROUP_SECP256R1, ucsTlsAlgoExchSecp256r1, ucsTlsAlgoExchSecp384r1)
                If Not pvBCryptEcdhMakeKey(IIf(.ExchAlgo = ucsTlsAlgoExchSecp256r1, 256, 384), .LocalExchPrivate, .LocalExchPublic) Then
                    Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_GENER_KEYPAIR_FAILED, "%1", IIf(.ExchAlgo = ucsTlsAlgoExchSecp256r1, "secp256r1", "secp384r1"))
                End If
#End If
            Case Else
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_EXCH_GROUP, "%1", "0x" & Hex$(.ExchGroup))
            End Select
        End If
    End With
End Sub

Private Sub pvTlsSetupExchRsaCertificate(uCtx As UcsTlsContext, baCert() As Byte)
    Const FUNC_NAME     As String = "pvTlsSetupExchRsaCertificate"
    Const MAX_RSA_BYTES As Long = MAX_RSA_KEY / 8
    Dim pCertContext    As Long
    Dim lPtr            As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim lAlignedSize    As Long
    Dim lResult         As Long
    
    pvCryptoClearApiError
    With uCtx
        .ExchAlgo = ucsTlsAlgoExchCertificate
        pvTlsGetRandom .LocalExchPrivate, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2 '--- always 48
        Call CopyMemory(.LocalExchPrivate(0), TLS_LOCAL_LEGACY_VERSION, 2)
        pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
        If pvCryptoSetApiError(pCertContext, "CertCreateCertificateContext") Then
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
        lPtr = UnsignedAdd(lPtr, 56)                                        '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
        lResult = CryptImportPublicKeyInfo(m_uData.hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lPtr, hKey)
        If pvCryptoSetApiError(lResult, "CryptImportPublicKeyInfo") Then
            GoTo QH
        End If
        lSize = pvArraySize(.LocalExchPrivate)
        lAlignedSize = (lSize + MAX_RSA_BYTES - 1 And -MAX_RSA_BYTES) + MAX_RSA_BYTES
        pvArrayAllocate .LocalExchRsaEncrPriv, lAlignedSize, FUNC_NAME & ".LocalExchRsaEncrPriv"
        Call CopyMemory(.LocalExchRsaEncrPriv(0), .LocalExchPrivate(0), lSize)
        lResult = CryptEncrypt(hKey, 0, 1, 0, .LocalExchRsaEncrPriv(0), lSize, lAlignedSize)
        If pvCryptoSetApiError(lResult, "CryptEncrypt") Then
            GoTo QH
        End If
        If UBound(.LocalExchRsaEncrPriv) <> lSize - 1 Then
            pvArrayReallocate .LocalExchRsaEncrPriv, lSize, FUNC_NAME & ".LocalExchRsaEncrPriv"
        End If
        pvArrayReverse .LocalExchRsaEncrPriv
    End With
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Sub

Private Sub pvTlsSetupCipherSuite(uCtx As UcsTlsContext, ByVal lCipherSuite As Long)
    Const FUNC_NAME     As String = "pvTlsSetupCipherSuite"
    
    With uCtx
        If .CipherSuite <> lCipherSuite Then
            .CipherSuite = lCipherSuite
            .BulkAlgo = 0
            .KeySize = 0
            .IvSize = 0
            .IvExplicitSize = 0
            .TagSize = 0
            .DigestAlgo = 0
            .DigestSize = 0
            .UseRsaKeyTransport = False
            Select Case lCipherSuite
            Case TLS_CS_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = LNG_SHA256_HASHSZ
            Case TLS_CS_AES_256_GCM_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .DigestAlgo = ucsTlsAlgoDigestSha384
                .DigestSize = LNG_SHA384_HASHSZ
            End Select
            Select Case lCipherSuite
            Case TLS_CS_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                .BulkAlgo = ucsTlsAlgoBulkChacha20Poly1305
                .KeySize = LNG_CHACHA20_KEYSZ
                .IvSize = LNG_CHACHA20POLY1305_IVSZ
                .TagSize = LNG_CHACHA20POLY1305_TAGSZ
            Case TLS_CS_AES_256_GCM_SHA384, TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .BulkAlgo = ucsTlsAlgoBulkAesGcm256
                .KeySize = LNG_AES256_KEYSZ
                .IvSize = LNG_AESGCM_IVSZ
                If lCipherSuite <> TLS_CS_AES_256_GCM_SHA384 Then
                    .IvExplicitSize = 8
                End If
                .TagSize = LNG_AESGCM_TAGSZ
            End Select
            Select Case lCipherSuite
            Case TLS_CS_RSA_WITH_AES_256_GCM_SHA384
                .UseRsaKeyTransport = True
            End Select
            If .BulkAlgo = 0 Or .DigestAlgo = 0 Then
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_CIPHER_SUITE, "%1", "0x" & Hex$(.CipherSuite))
            End If
        End If
    End With
End Sub

Private Function pvTlsGetSortedCipherSuites(ByVal eFilter As UcsTlsLocalFeaturesEnum) As Collection
    Dim oRetVal     As Collection
    
    Set oRetVal = New Collection
    If (eFilter And ucsTlsSupportTls13) <> 0 Then
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            oRetVal.Add TLS_CS_AES_256_GCM_SHA384
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkChacha20Poly1305) Then
            oRetVal.Add TLS_CS_CHACHA20_POLY1305_SHA256
        End If
    End If
    If (eFilter And ucsTlsSupportTls12) <> 0 Then
        If pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
            oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
            oRetVal.Add TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        End If
        If pvCryptoIsSupported(ucsTlsAlgoBulkChacha20Poly1305) Then
            oRetVal.Add TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
            oRetVal.Add TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        End If
        '--- no perfect forward secrecy -> least preferred
        If pvCryptoIsSupported(ucsTlsAlgoExchCertificate) Then
            If pvCryptoIsSupported(ucsTlsAlgoBulkAesGcm256) Then
                oRetVal.Add TLS_CS_RSA_WITH_AES_256_GCM_SHA384
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
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
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
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_HANDSHAKE_MESSAGES
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
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "RemoteTrafficSecret")
        End If
        pvTlsHkdfExpandLabel .RemoteTrafficSecret, .DigestAlgo, .RemoteTrafficSecret, "traffic upd", baEmpty, .DigestSize
        pvTlsHkdfExpandLabel .RemoteTrafficKey, .DigestAlgo, .RemoteTrafficSecret, "key", baEmpty, .KeySize
        pvTlsHkdfExpandLabel .RemoteTrafficIV, .DigestAlgo, .RemoteTrafficSecret, "iv", baEmpty, .IvSize
        .RemoteTrafficSeqNo = 0
        pvTlsLogSecret uCtx, IIf(.IsServer, "CLIENT", "SERVER") & "_TRAFFIC_SECRET_0", .RemoteTrafficSecret
        If bLocalUpdate Then
            If pvArraySize(.LocalTrafficSecret) = 0 Then
                Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_NO_PREVIOUS_SECRET, "%1", "LocalTrafficSecret")
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

Private Sub pvTlsDeriveLegacySecrets(uCtx As UcsTlsContext, baHandshakeHash() As Byte)
    Const FUNC_NAME     As String = "pvTlsDeriveLegacySecrets"
    Dim baPreMasterSecret() As Byte
    Dim uRandom         As UcsBuffer
    Dim uExpanded       As UcsBuffer
    
    With uCtx
        If pvArraySize(.RemoteExchRandom) = 0 Then
            Err.Raise vbObjectError, FUNC_NAME, ERR_NO_REMOTE_RANDOM
        End If
        Debug.Assert pvArraySize(.LocalExchRandom) = TLS_HELLO_RANDOM_SIZE
        Debug.Assert pvArraySize(.RemoteExchRandom) = TLS_HELLO_RANDOM_SIZE
        pvTlsGetSharedSecret baPreMasterSecret, .ExchAlgo, .LocalExchPrivate, .RemoteExchPublic
        If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_EXTENDED_MASTER_SECRET) Then
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "extended master secret", baHandshakeHash, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2    '--- always 48
        Else
            pvBufferWriteArray uRandom, .LocalExchRandom
            pvBufferWriteArray uRandom, .RemoteExchRandom
            pvBufferWriteEOF uRandom
            pvTlsKdfLegacyPrf .MasterSecret, .DigestAlgo, baPreMasterSecret, "master secret", uRandom.Data, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2    '--- always 48
        End If
        pvTlsLogSecret uCtx, "CLIENT_RANDOM", .MasterSecret
        uRandom.Size = 0
        pvBufferWriteArray uRandom, .RemoteExchRandom
        pvBufferWriteArray uRandom, .LocalExchRandom
        pvBufferWriteEOF uRandom
        pvTlsKdfLegacyPrf uExpanded.Data, .DigestAlgo, .MasterSecret, "key expansion", uRandom.Data, 2 * (.KeySize + .IvSize)
        pvBufferReadArray uExpanded, .LocalTrafficKey, .KeySize
        pvBufferReadArray uExpanded, .RemoteLegacyNextTrafficKey, .KeySize
        pvTlsGetRandom .LocalTrafficIV, .IvSize
        pvBufferReadBlob uExpanded, VarPtr(.LocalTrafficIV(0)), .IvSize - .IvExplicitSize
        pvTlsGetRandom .RemoteLegacyNextTrafficIV, .IvSize
        pvBufferReadBlob uExpanded, VarPtr(.RemoteLegacyNextTrafficIV(0)), .IvSize - .IvExplicitSize
        .LocalTrafficSeqNo = 0
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
    With uCtx
        pvTlsGetHash baRetVal, .DigestAlgo, .HandshakeMessages.Data, Size:=.HandshakeMessages.Size
    End With
End Sub

Private Sub pvTlsAppendHandshakeHash(uCtx As UcsTlsContext, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    pvBufferWriteBlob uCtx.HandshakeMessages, VarPtr(baInput(lPos)), lSize
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
    Dim lHashAlgId      As Long
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        lHashAlgId = CALG_SHA_256
    Case ucsTlsAlgoDigestSha384
        lHashAlgId = CALG_SHA_384
    Case ucsTlsAlgoDigestSha512
        lHashAlgId = CALG_SHA_512
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported hash type " & eHash
    End Select
    If Not pvCryptoHash(baRetVal, lHashAlgId, baInput, Pos, Size) Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHash")
    End If
End Sub

Private Sub pvTlsGetHmac(baRetVal() As Byte, ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Const FUNC_NAME     As String = "pvTlsGetHmac"
    Dim lHashAlgId      As Long
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        lHashAlgId = CALG_SHA_256
    Case ucsTlsAlgoDigestSha384
        lHashAlgId = CALG_SHA_384
    Case ucsTlsAlgoDigestSha512
        lHashAlgId = CALG_SHA_512
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported HMAC type " & eHash
    End Select
    If Not pvCryptoHmac(baRetVal, lHashAlgId, baKey, baInput, Pos, Size) Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHmac")
    End If
End Sub

Private Sub pvTlsGetHelloRetryRandom(baRetVal() As Byte)
    pvArrayByte baRetVal, &HCF, &H21, &HAD, &H74, &HE5, &H9A, &H61, &H11, &HBE, &H1D, &H8C, &H2, &H1E, &H65, &HB8, &H91, &HC2, &HA2, &H11, &H16, &H7A, &HBB, &H8C, &H5E, &H7, &H9E, &H9, &HE2, &HC8, &HA8, &H33, &H9C
End Sub

Private Function pvTlsBulkDecrypt(ByVal eBulk As UcsTlsCryptoAlgorithmsEnum, baRemoteIV() As Byte, baRemoteKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Const FUNC_NAME     As String = "pvTlsBulkDecrypt"
    
    Select Case eBulk
#If ImplLibSodium Then
    Case ucsTlsAlgoBulkChacha20Poly1305
        If Not pvCryptoAeadChacha20Poly1305Decrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
#End If
    Case ucsTlsAlgoBulkAesGcm256
        If Not pvCryptoAeadAesGcmDecrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported bulk type " & eBulk
    End Select
    '--- success
    pvTlsBulkDecrypt = True
QH:
End Function

Private Sub pvTlsBulkEncrypt(ByVal eBulk As UcsTlsCryptoAlgorithmsEnum, baLocalIV() As Byte, baLocalKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long)
    Const FUNC_NAME     As String = "pvTlsBulkEncrypt"
    
    Select Case eBulk
#If ImplLibSodium Then
    Case ucsTlsAlgoBulkChacha20Poly1305
        If Not pvCryptoAeadChacha20Poly1305Encrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoAeadChacha20Poly1305Encrypt")
        End If
#End If
    Case ucsTlsAlgoBulkAesGcm256
        If Not pvCryptoAeadAesGcmEncrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_ENCRYPTION_FAILED, "%1", "CryptoAeadAesGcmEncrypt")
        End If
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported bulk type " & eBulk
    End Select
End Sub

Private Sub pvTlsGetSharedSecret(baRetVal() As Byte, ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte)
    Const FUNC_NAME     As String = "pvTlsGetSharedSecret"
    
    Select Case eKeyX
    Case ucsTlsAlgoExchX25519
#If ImplLibSodium Then
        If Not pvCryptoEcdhCurve25519SharedSecret(baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoEcdhCurve25519SharedSecret")
        End If
#Else
        If Not pvBCryptEcdhSharedSecret(0, baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "BCryptEcdhSharedSecret")
        End If
#End If
#If ImplTlsServer Then
    Case ucsTlsAlgoExchSecp256r1, ucsTlsAlgoExchSecp384r1
        If Not pvBCryptEcdhSharedSecret(IIf(eKeyX = ucsTlsAlgoExchSecp256r1, 256, 384), baRetVal, baPriv, baPub) Then
            Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "BCryptEcdhSharedSecret")
        End If
#End If
    Case ucsTlsAlgoExchCertificate
        baRetVal = baPriv
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, "Unsupported exchange " & eKeyX
    End Select
End Sub

Private Function pvTlsGetExchGroupName(ByVal lExchGroup As Long) As String
    Select Case lExchGroup
    Case TLS_GROUP_X25519
        pvTlsGetExchGroupName = "X25519"
#If ImplTlsServer Then
    Case TLS_GROUP_SECP256R1
        pvTlsGetExchGroupName = "secp256r1"
    Case TLS_GROUP_SECP384R1
        pvTlsGetExchGroupName = "secp384r1"
#End If
    Case Else
        pvTlsGetExchGroupName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lExchGroup))
    End Select
End Function

Private Function pvTlsGetCipherSuiteName(ByVal lCipherSuite As Long) As String
    Select Case lCipherSuite
    Case TLS_CS_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "TLS_AES_256_GCM_SHA384"
    Case TLS_CS_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "TLS_CHACHA20_POLY1305_SHA256"
    Case TLS_CS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "ECDHE-RSA-AES256-GCM-SHA384"
    Case TLS_CS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-RSA-CHACHA20-POLY1305"
    Case TLS_CS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsGetCipherSuiteName = "ECDHE-ECDSA-CHACHA20-POLY1305"
    Case TLS_CS_RSA_WITH_AES_256_GCM_SHA384
        pvTlsGetCipherSuiteName = "AES256-GCM-SHA384"
    Case Else
        pvTlsGetCipherSuiteName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lCipherSuite))
    End Select
End Function

#If ImplTlsServer Then
Private Function pvTlsSignatureName(ByVal lSignatureScheme As Long) As String
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        pvTlsSignatureName = "RSA_PKCS1_SHA1"
    Case TLS_SIGNATURE_ECDSA_SHA1
        pvTlsSignatureName = "ECDSA_SHA1"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA224
        pvTlsSignatureName = "RSA_PKCS1_SHA224"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256
        pvTlsSignatureName = "RSA_PKCS1_SHA256"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA384
        pvTlsSignatureName = "RSA_PKCS1_SHA384"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA512
        pvTlsSignatureName = "RSA_PKCS1_SHA512"
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        pvTlsSignatureName = "ECDSA_SECP256R1_SHA256"
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        pvTlsSignatureName = "ECDSA_SECP384R1_SHA384"
    Case TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        pvTlsSignatureName = "ECDSA_SECP521R1_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256
        pvTlsSignatureName = "RSA_PSS_RSAE_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384
        pvTlsSignatureName = "RSA_PSS_RSAE_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        pvTlsSignatureName = "RSA_PSS_RSAE_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        pvTlsSignatureName = "RSA_PSS_PSS_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        pvTlsSignatureName = "RSA_PSS_PSS_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        pvTlsSignatureName = "RSA_PSS_PSS_SHA512"
    Case Else
        pvTlsSignatureName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lSignatureScheme))
    End Select
End Function

Private Sub pvTlsSignatureSign(baRetVal() As Byte, cPrivKey As Collection, ByVal lSignatureScheme As Long, baVerifyData() As Byte)
    Const FUNC_NAME     As String = "pvTlsSignatureSign"
    Dim baPrivKey()     As Byte
    Dim uKeyInfo        As UcsKeyInfo
    
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, FUNC_NAME, "Signing with " & pvTlsSignatureName(lSignatureScheme) & " signature"
    #End If
    If Not SearchCollection(cPrivKey, 1, RetVal:=baPrivKey) Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_NO_PRIVATE_KEY
    End If
    If Not pvAsn1DecodePrivateKey(baPrivKey, uKeyInfo) Then
        Err.Raise vbObjectError, FUNC_NAME, ERR_UNSUPPORTED_PRIVATE_KEY
    End If
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        pvBCryptRsaPssSign baRetVal, uKeyInfo.KeyBlob, lSignatureScheme, baVerifyData
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        pvBCryptEcdsaSign baRetVal, uKeyInfo.KeyBlob, lSignatureScheme, baVerifyData
    Case Else
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_UNSUPPORTED_SIGNATURE_SCHEME, "%1", "0x" & Hex$(lSignatureScheme))
    End Select
    If pvArraySize(baRetVal) = 0 Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_SIGNATURE_FAILED, "%1", pvTlsSignatureName(lSignatureScheme))
    End If
End Sub
#End If

Private Function pvTlsSignatureVerify(baCert() As Byte, ByVal lSignatureScheme As Long, baVerifyData() As Byte, baSignature() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    #If baCert And lSignatureScheme And baVerifyData And baSignature And sError And eAlertCode Then '--- touch args
    #End If
    pvTlsSignatureVerify = True
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
                End If
            pvBufferWriteBlockEnd uOutput
            #If (ImplCaptureTraffic And 1) <> 0 Then
                If lMessageSize <> 0 Then
                    .TrafficDump.Add FUNC_NAME & ".Output (unencrypted)" & vbCrLf & TlsDesignDumpArray(uOutput.Data, lRecordPos, uOutput.Size - lRecordPos - .TagSize)
                End If
            #End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                pvTlsBulkEncrypt .BulkAlgo, baLocalIV, .LocalTrafficKey, uOutput.Data, lRecordPos, TLS_AAD_SIZE, uOutput.Data, lMessagePos, lMessageSize
            ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                pvTlsBulkEncrypt .BulkAlgo, baLocalIV, .LocalTrafficKey, uAad.Data, 0, uAad.Size, uOutput.Data, lMessagePos, lMessageSize
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
            Call CopyMemory(baTemp(0), lValue, 4)
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
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), 4)
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
        Else
            lValue = 0
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

Private Sub pvArrayReverse(baData() As Byte)
    Dim lIdx            As Long
    Dim bTemp           As Byte
    
    For lIdx = 0 To UBound(baData) \ 2
        bTemp = baData(lIdx)
        baData(lIdx) = baData(UBound(baData) - lIdx)
        baData(UBound(baData) - lIdx) = bTemp
    Next
End Sub

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

'=========================================================================
' Crypto
'=========================================================================

Private Function pvCryptoIsSupported(ByVal eAlgo As UcsTlsCryptoAlgorithmsEnum) As Boolean
    Select Case eAlgo
    Case ucsTlsAlgoExchSecp256r1, ucsTlsAlgoExchSecp384r1
        #If ImplTlsServer Then
            pvCryptoIsSupported = (RealOsVersion >= ucsOsvWin10)
        #End If
    Case ucsTlsAlgoBulkChacha20Poly1305
        #If ImplLibSodium Then
            pvCryptoIsSupported = True
        #End If
    Case ucsTlsAlgoBulkAesGcm256
        #If ImplLibSodium Then
            pvCryptoIsSupported = (crypto_aead_aes256gcm_is_available() <> 0)
        #Else
            pvCryptoIsSupported = (RealOsVersion >= ucsOsvWin7)
        #End If
    Case Else
        pvCryptoIsSupported = True
    End Select
End Function

#If ImplTlsServer Then
Private Sub pvBCryptRsaPssSign(baRetVal() As Byte, baKeyBlob() As Byte, ByVal lSignatureScheme As Long, baMessage() As Byte)
    Const FUNC_NAME     As String = "pvBCryptRsaPssSign"
    Const BCRYPT_PAD_PSS As Long = 8
    Dim lHashAlgId      As Long
    Dim hAlg            As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim baHash()        As Byte
    
    pvCryptoClearApiError
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("RSA"), 0, 0)
    If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
        GoTo QH
    End If
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("CAPIPRIVATEBLOB"), hKey, baKeyBlob(0), UBound(baKeyBlob) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptImportKeyPair") Then
        GoTo QH
    End If
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        lHashAlgId = CALG_SHA_256
        uPadInfo.pszAlgId = StrPtr("SHA256")
        uPadInfo.cbSalt = LNG_SHA256_HASHSZ
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        lHashAlgId = CALG_SHA_384
        uPadInfo.pszAlgId = StrPtr("SHA384")
        uPadInfo.cbSalt = LNG_SHA384_HASHSZ
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        lHashAlgId = CALG_SHA_512
        uPadInfo.pszAlgId = StrPtr("SHA512")
        uPadInfo.cbSalt = LNG_SHA512_HASHSZ
    End Select
    If Not pvCryptoHash(baHash, lHashAlgId, baMessage) Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHash")
    End If
    pvArrayAllocate baRetVal, 4096, FUNC_NAME & ".baRetVal"
    hResult = BCryptSignHash(hKey, uPadInfo, baHash(0), UBound(baHash) + 1, baRetVal(0), UBound(baRetVal) + 1, lSize, BCRYPT_PAD_PSS)
    If pvCryptoSetApiResult(hResult, "BCryptSignHash") Then
        GoTo QH
    End If
    pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Sub

Private Sub pvBCryptEcdsaSign(baRetVal() As Byte, baKeyBlob() As Byte, ByVal lSignatureScheme As Long, baMessage() As Byte)
    Const FUNC_NAME     As String = "pvBCryptEcdsaSign"
    Dim lHashAlgId      As Long
    Dim sHashAlgId    As String
    Dim hAlg            As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim hResult         As Long
    Dim baHash()        As Byte
    Dim uEccKey         As BCRYPT_ECCKEY_BLOB
    Dim baTemp()        As Byte
    Dim uEccSig         As CERT_ECC_SIGNATURE
    Dim lResult         As Long
    
    pvCryptoClearApiError
    Select Case lSignatureScheme
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        lHashAlgId = CALG_SHA_256
        sHashAlgId = "ECDSA_P256"
        uEccKey.dwMagic = BCRYPT_ECDSA_PRIVATE_P256_MAGIC
        uEccKey.cbKey = LNG_SECP256R1_KEYSZ
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        lHashAlgId = CALG_SHA_384
        sHashAlgId = "ECDSA_P384"
        uEccKey.dwMagic = BCRYPT_ECDSA_PRIVATE_P384_MAGIC
        uEccKey.cbKey = LNG_SECP384R1_KEYSZ
    Case TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        lHashAlgId = CALG_SHA_512
        sHashAlgId = "ECDSA_P521"
        uEccKey.dwMagic = BCRYPT_ECDSA_PRIVATE_P521_MAGIC
        uEccKey.cbKey = LNG_SECP521R1_KEYSZ
    End Select
    If Not pvCryptoHash(baHash, lHashAlgId, baMessage) Then
        Err.Raise vbObjectError, FUNC_NAME, Replace(ERR_CALL_FAILED, "%1", "CryptoHash")
    End If
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr(sHashAlgId), 0, 0)
    If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
        GoTo QH
    End If
    Debug.Assert UBound(uEccKey.Buffer) >= UBound(baKeyBlob)
    Call CopyMemory(uEccKey.Buffer(0), baKeyBlob(0), UBound(baKeyBlob) + 1)
    hResult = BCryptImportKeyPair(hAlg, 0, StrPtr("ECCPRIVATEBLOB"), hKey, uEccKey, sizeof_BCRYPT_ECCKEY_BLOB + UBound(baKeyBlob) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptImportKeyPair") Then
        GoTo QH
    End If
    pvArrayAllocate baTemp, 1024, FUNC_NAME & ".baTemp"
    hResult = BCryptSignHash(hKey, ByVal 0, baHash(0), UBound(baHash) + 1, baTemp(0), UBound(baTemp) + 1, lSize, 0)
    If pvCryptoSetApiResult(hResult, "BCryptSignHash") Then
        GoTo QH
    End If
    pvArrayReallocate baTemp, lSize, FUNC_NAME & ".baTemp"
    '--- ASN.1 encode signature
    pvArrayReverse baTemp
    uEccSig.rValue.pbData = VarPtr(baTemp(uEccKey.cbKey))
    uEccSig.rValue.cbData = uEccKey.cbKey
    uEccSig.sValue.pbData = VarPtr(baTemp(0))
    uEccSig.sValue.cbData = uEccKey.cbKey
    lSize = 1024
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    lResult = CryptEncodeObjectEx(X509_ASN_ENCODING, X509_ECC_SIGNATURE, uEccSig, 0, 0, baRetVal(0), lSize)
    If pvCryptoSetApiError(lResult, "CryptEncodeObjectEx") Then
        GoTo QH
    End If
    pvArrayReallocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Sub
#End If

Private Function pvAsn1DecodePrivateKey(baPrivKey() As Byte, uRetVal As UcsKeyInfo) As Boolean
    Const FUNC_NAME     As String = "Asn1DecodePrivateKey"
    Dim lPkiPtr         As Long
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim lSize           As Long
    Dim uEccKeyInfo     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim lResult         As Long

    pvCryptoClearApiError
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) <> 0 Then
        Call CopyMemory(uPrivKey, ByVal lPkiPtr, Len(uPrivKey))
        lResult = CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uPrivKey.PrivateKey.pbData, uPrivKey.PrivateKey.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize)
        If pvCryptoSetApiError(lResult, "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)") Then
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
        Call CopyMemory(uRetVal.BitLen, uRetVal.KeyBlob(12), 4)
    Else
        lResult = CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize)
        If pvCryptoSetApiError(lResult, "CryptDecodeObjectEx(X509_ECC_PRIVATE_KEY)") Then
            GoTo QH
        End If
        Call CopyMemory(uEccKeyInfo, ByVal lKeyPtr, Len(uEccKeyInfo))
        uRetVal.AlgoObjId = pvToStringA(uEccKeyInfo.szCurveOid)
        If uEccKeyInfo.PublicKey.cbData > 0 Then
            pvArrayAllocate uRetVal.KeyBlob, uEccKeyInfo.PublicKey.cbData - 1 + uEccKeyInfo.PrivateKey.cbData, FUNC_NAME & ".uRetVal.KeyBlob"
            Debug.Assert UBound(uRetVal.KeyBlob) + 1 >= uEccKeyInfo.PublicKey.cbData - 1
            Call CopyMemory(uRetVal.KeyBlob(0), ByVal uEccKeyInfo.PublicKey.pbData + 1, uEccKeyInfo.PublicKey.cbData - 1)
            lSize = uEccKeyInfo.PublicKey.cbData - 1
        Else
            pvArrayAllocate uRetVal.KeyBlob, 3 * uEccKeyInfo.PrivateKey.cbData, FUNC_NAME & ".uRetVal.KeyBlob"
            lSize = 2 * uEccKeyInfo.PrivateKey.cbData
        End If
        Debug.Assert UBound(uRetVal.KeyBlob) + 1 - lSize >= uEccKeyInfo.PrivateKey.cbData
        Call CopyMemory(uRetVal.KeyBlob(lSize), ByVal uEccKeyInfo.PrivateKey.pbData, uEccKeyInfo.PrivateKey.cbData)
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
    pvCryptoCheckApiError FUNC_NAME
End Function

Private Function pvCryptoInit() As Boolean
    Const FUNC_NAME     As String = "CryptoInit"
    Dim lResult         As Long
    
    pvCryptoClearApiError
    With m_uData
#If ImplLibSodium Then
        If GetModuleHandle(StrPtr("libsodium.dll")) = 0 Then
            If LoadLibrary(StrPtr(App.Path & "\libsodium.dll")) = 0 Then
                Call LoadLibrary(StrPtr(App.Path & "\..\..\lib\libsodium.dll"))
            End If
            If sodium_init() < 0 Then
                pvCryptoSetApiResult LNG_OUT_OF_MEMORY, "sodium_init"
                GoTo QH
            End If
        End If
#End If
        If .hProv = 0 Then
            lResult = CryptAcquireContext(.hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
            If pvCryptoSetApiError(lResult, "CryptAcquireContext") Then
                GoTo QH
            End If
        End If
    End With
    '--- success
    pvCryptoInit = True
QH:
    pvCryptoCheckApiError FUNC_NAME
End Function

#If ImplLibSodium Then
Private Function pvCryptoEcdhCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccCurve25519MakeKey"
    
    pvArrayAllocate baPrivate, LNG_X25519_KEYSZ, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baPublic, LNG_X25519_KEYSZ, FUNC_NAME & ".baPublic"
    pvCryptoRandomBytes VarPtr(baPrivate(0)), LNG_X25519_KEYSZ
    '--- fix privkey randomness
    baPrivate(0) = baPrivate(0) And &HF8
    baPrivate(LNG_X25519_KEYSZ - 1) = (baPrivate(LNG_X25519_KEYSZ - 1) And &H7F) Or &H40
    Call crypto_scalarmult_curve25519_base(baPublic(0), baPrivate(0))
    '--- success
    pvCryptoEcdhCurve25519MakeKey = True
End Function

Private Function pvCryptoEcdhCurve25519SharedSecret(baRetVal() As Byte, baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoEccCurve25519SharedSecret"
    
    Debug.Assert UBound(baPrivate) >= LNG_X25519_KEYSZ - 1
    Debug.Assert UBound(baPublic) >= LNG_X25519_KEYSZ - 1
    pvArrayAllocate baRetVal, LNG_X25519_KEYSZ, FUNC_NAME & ".baRetVal"
    Call crypto_scalarmult_curve25519(baRetVal(0), baPrivate(0), baPublic(0))
    '--- success
    pvCryptoEcdhCurve25519SharedSecret = True
End Function
#End If

#If ImplTlsServer Or Not ImplLibSodium Then
Private Function pvBCryptEcdhMakeKey(ByVal lKeySize As Long, baPriv() As Byte, baPub() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvBCryptEcdhMakeKey"
    Dim hAlgEcdh        As Long
    Dim hResult         As Long
    Dim hKeyPair        As Long
    Dim baBlob()        As Byte
    Dim cbResult        As Long
    
    pvCryptoClearApiError
    If lKeySize = 0 Then
        hResult = BCryptOpenAlgorithmProvider(hAlgEcdh, StrPtr("ECDH"), 0, 0)
        If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
            GoTo QH
        End If
        hResult = BCryptSetProperty(hAlgEcdh, StrPtr("ECCCurveName"), StrPtr("Curve25519"), 22, 0)
        If pvCryptoSetApiResult(hResult, "BCryptSetProperty(ECCCurveName)") Then
            GoTo QH
        End If
    Else
        hResult = BCryptOpenAlgorithmProvider(hAlgEcdh, StrPtr("ECDH_P" & lKeySize), 0, 0)
        If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
            GoTo QH
        End If
    End If
    hResult = BCryptGenerateKeyPair(hAlgEcdh, hKeyPair, lKeySize, 0)
    If pvCryptoSetApiResult(hResult, "BCryptGenerateKeyPair") Then
        GoTo QH
    End If
    hResult = BCryptFinalizeKeyPair(hKeyPair, 0)
    If pvCryptoSetApiResult(hResult, "BCryptFinalizeKeyPair") Then
        GoTo QH
    End If
    pvArrayAllocate baBlob, 1024, FUNC_NAME & ".baBlob"
    hResult = BCryptExportKey(hKeyPair, 0, StrPtr("ECCPRIVATEBLOB"), VarPtr(baBlob(0)), UBound(baBlob) + 1, cbResult, 0)
    If pvCryptoSetApiResult(hResult, "BCryptExportKey(ECCPRIVATEBLOB)") Then
        GoTo QH
    End If
    If Not pvBCryptEcdhFromKeyBlob(baPriv, baBlob, cbResult) Then
        GoTo QH
    End If
    hResult = BCryptExportKey(hKeyPair, 0, StrPtr("ECCPUBLICBLOB"), VarPtr(baBlob(0)), UBound(baBlob) + 1, cbResult, 0)
    If pvCryptoSetApiResult(hResult, "BCryptExportKey(ECCPUBLICBLOB)") Then
        GoTo QH
    End If
    If Not pvBCryptEcdhFromKeyBlob(baPub, baBlob, cbResult) Then
        GoTo QH
    End If
    '--- success
    pvBCryptEcdhMakeKey = True
QH:
    If hKeyPair <> 0 Then
        Call BCryptDestroyKey(hKeyPair)
    End If
    If hAlgEcdh <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlgEcdh, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Function

Private Function pvBCryptEcdhSharedSecret(ByVal lKeySize As Long, baRetVal() As Byte, baPriv() As Byte, baPub() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvBCryptEcdhSharedSecret"
    Dim hAlgEcdh        As Long
    Dim hPrivKey        As Long
    Dim hPubKey         As Long
    Dim hAgreedSecret   As Long
    Dim cbAgreedSecret  As Long
    Dim hResult         As Long
    Dim baBlob()        As Byte
    
    pvCryptoClearApiError
    If lKeySize = 0 Then
        hResult = BCryptOpenAlgorithmProvider(hAlgEcdh, StrPtr("ECDH"), 0, 0)
        If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
            GoTo QH
        End If
        hResult = BCryptSetProperty(hAlgEcdh, StrPtr("ECCCurveName"), StrPtr("Curve25519"), 22, 0)
        If pvCryptoSetApiResult(hResult, "BCryptSetProperty(ECCCurveName)") Then
            GoTo QH
        End If
    Else
        hResult = BCryptOpenAlgorithmProvider(hAlgEcdh, StrPtr("ECDH_P" & lKeySize), 0, 0)
        If pvCryptoSetApiResult(hResult, "BCryptOpenAlgorithmProvider") Then
            GoTo QH
        End If
    End If
    If Not pvBCryptEcdhToKeyBlob(baBlob, baPriv) Then
        GoTo QH
    End If
    hResult = BCryptImportKeyPair(hAlgEcdh, 0, StrPtr("ECCPRIVATEBLOB"), hPrivKey, baBlob(0), UBound(baBlob) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptImportKeyPair(ECCPRIVATEBLOB)") Then
        GoTo QH
    End If
    If Not pvBCryptEcdhToKeyBlob(baBlob, baPub) Then
        GoTo QH
    End If
    hResult = BCryptImportKeyPair(hAlgEcdh, 0, StrPtr("ECCPUBLICBLOB"), hPubKey, baBlob(0), UBound(baBlob) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptImportKeyPair(ECCPUBLICBLOB)") Then
        GoTo QH
    End If
    hResult = BCryptSecretAgreement(hPrivKey, hPubKey, hAgreedSecret, 0)
    If pvCryptoSetApiResult(hResult, "BCryptSecretAgreement") Then
        GoTo QH
    End If
    pvArrayAllocate baRetVal, 1024, FUNC_NAME & ".baRetVal"
    hResult = BCryptDeriveKey(hAgreedSecret, StrPtr("TRUNCATE"), 0, VarPtr(baRetVal(0)), UBound(baRetVal) + 1, cbAgreedSecret, 0)
    If pvCryptoSetApiResult(hResult, "BCryptDeriveKey") Then
        GoTo QH
    End If
    pvArrayReallocate baRetVal, cbAgreedSecret, FUNC_NAME & ".baRetVal"
    pvArrayReverse baRetVal
    '--- success
    pvBCryptEcdhSharedSecret = True
QH:
    If hAgreedSecret <> 0 Then
        Call BCryptDestroySecret(hAgreedSecret)
    End If
    If hPrivKey <> 0 Then
        Call BCryptDestroyKey(hPrivKey)
    End If
    If hPubKey <> 0 Then
        Call BCryptDestroyKey(hPubKey)
    End If
    If hAlgEcdh <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlgEcdh, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Function

Private Function pvBCryptEcdhToKeyBlob(baRetVal() As Byte, baKey() As Byte, Optional ByVal lSize As Long = -1) As Boolean
    Const FUNC_NAME     As String = "pvBCryptEcdhToKeyBlob"
    Const TAG_UNCOMPRESSED As Long = 4
    Dim lMagic          As Long
    Dim lPartSize       As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim lSum            As Long
    
    If lSize < 0 Then
        lSize = pvArraySize(baKey)
    End If
    Select Case lSize
    Case LNG_X25519_KEYSZ
        lMagic = BCRYPT_ECDH_PUBLIC_GENERIC_MAGIC
        lPartSize = LNG_X25519_KEYSZ
        lSize = 2 * LNG_X25519_KEYSZ
    Case 1 + 2 * LNG_SECP256R1_KEYSZ
        Debug.Assert baKey(0) = TAG_UNCOMPRESSED
        lMagic = BCRYPT_ECDH_PUBLIC_P256_MAGIC
        lPartSize = LNG_SECP256R1_KEYSZ
        lPos = 1
    Case 1 + 2 * LNG_SECP384R1_KEYSZ
        Debug.Assert baKey(0) = TAG_UNCOMPRESSED
        lMagic = BCRYPT_ECDH_PUBLIC_P384_MAGIC
        lPartSize = LNG_SECP384R1_KEYSZ
        lPos = 1
    Case 3 * LNG_SECP256R1_KEYSZ
        For lIdx = LNG_SECP256R1_KEYSZ To 2 * LNG_SECP256R1_KEYSZ - 1
            lSum = lSum Or baKey(lIdx)
        Next
        If lSum = 0 Then
            lMagic = BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC
            lPartSize = LNG_X25519_KEYSZ
        Else
            lMagic = BCRYPT_ECDH_PRIVATE_P256_MAGIC
            lPartSize = LNG_SECP256R1_KEYSZ
        End If
    Case 3 * LNG_SECP384R1_KEYSZ
        lMagic = BCRYPT_ECDH_PRIVATE_P384_MAGIC
        lPartSize = LNG_SECP384R1_KEYSZ
    Case Else
        Err.Raise vbObjectError, , "Unrecognized key size"
    End Select
    pvArrayAllocate baRetVal, 8 + lSize - lPos, FUNC_NAME & ".baRetVal"
    Call CopyMemory(baRetVal(0), lMagic, 4)
    Call CopyMemory(baRetVal(4), lPartSize, 4)
    Call CopyMemory(baRetVal(8), baKey(lPos), UBound(baKey) + 1 - lPos)
    '--- success
    pvBCryptEcdhToKeyBlob = True
End Function

Private Function pvBCryptEcdhFromKeyBlob(baRetVal() As Byte, baBlob() As Byte, Optional ByVal lSize As Long = -1) As Boolean
    Const FUNC_NAME     As String = "pvBCryptEcdhFromKeyBlob"
    Const TAG_UNCOMPRESSED As Long = 4
    Dim lMagic          As Long
    Dim lPartSize       As Long
    
    If lSize < 0 Then
        lSize = pvArraySize(baBlob)
    End If
    Call CopyMemory(lMagic, baBlob(0), 4)
    Select Case lMagic
    Case BCRYPT_ECDH_PUBLIC_P256_MAGIC, BCRYPT_ECDH_PUBLIC_P384_MAGIC
        Call CopyMemory(lPartSize, baBlob(4), 4)
        Debug.Assert lPartSize = LNG_SECP256R1_KEYSZ Or lPartSize = LNG_SECP384R1_KEYSZ
        pvArrayAllocate baRetVal, 1 + 2 * lPartSize, FUNC_NAME & ".baRetVal"
        Debug.Assert lSize >= 8 + 2 * lPartSize
        baRetVal(0) = TAG_UNCOMPRESSED
        Call CopyMemory(baRetVal(1), baBlob(8), 2 * lPartSize)
    Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_GENERIC_MAGIC
        Call CopyMemory(lPartSize, baBlob(4), 4)
        Debug.Assert lPartSize = LNG_SECP256R1_KEYSZ Or lPartSize = LNG_SECP384R1_KEYSZ
        pvArrayAllocate baRetVal, 3 * lPartSize, FUNC_NAME & ".baRetVal"
        Debug.Assert lSize >= 8 + 3 * lPartSize
        Call CopyMemory(baRetVal(0), baBlob(8), 3 * lPartSize)
    Case BCRYPT_ECDH_PUBLIC_GENERIC_MAGIC
        Call CopyMemory(lPartSize, baBlob(4), 4)
        Debug.Assert lPartSize = LNG_X25519_KEYSZ
        pvArrayAllocate baRetVal, lPartSize, FUNC_NAME & ".baRetVal"
        Debug.Assert lSize >= 8 + lPartSize
        Call CopyMemory(baRetVal(0), baBlob(8), lPartSize)
    Case Else
        Err.Raise vbObjectError, , "Unknown BCrypt magic (&H" & Hex$(lMagic) & ")"
    End Select
    '--- success
    pvBCryptEcdhFromKeyBlob = True
End Function
#End If

Private Function pvCryptoHash(baRetVal() As Byte, ByVal lHashAlgId As Long, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHash"
    Dim hHash           As Long
    Dim lHashSize       As Long
    Dim lResult         As Long
    
    pvCryptoClearApiError
    lResult = CryptCreateHash(m_uData.hProv, lHashAlgId, 0, 0, hHash)
    If pvCryptoSetApiError(lResult, "CryptCreateHash") Then
        GoTo QH
    End If
    If Size < 0 Then
        Size = pvArraySize(baInput)
    End If
    If Size > 0 Then
        lResult = CryptHashData(hHash, baInput(Pos), Size, 0)
        If pvCryptoSetApiError(lResult, "CryptHashData") Then
            GoTo QH
        End If
    End If
    lHashSize = 1024
    pvArrayAllocate baRetVal, lHashSize, FUNC_NAME & ".baRetVal"
    lResult = CryptGetHashParam(hHash, HP_HASHVAL, baRetVal(0), lHashSize, 0)
    If pvCryptoSetApiError(lResult, "CryptGetHashParam") Then
        GoTo QH
    End If
    pvArrayReallocate baRetVal, lHashSize, FUNC_NAME & ".baRetVal"
    '--- success
    pvCryptoHash = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Function

Private Function pvCryptoHmac(baRetVal() As Byte, ByVal lHashAlgId As Long, baKey() As Byte, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Boolean
    Const FUNC_NAME     As String = "CryptoHmac"
    Dim uBlob           As BLOBHEADER
    Dim hKey            As Long
    Dim hHash           As Long
    Dim uInfo           As HMAC_INFO
    Dim lHashSize       As Long
    Dim lResult         As Long
    
    pvCryptoClearApiError
    uBlob.bType = PLAINTEXTKEYBLOB
    uBlob.bVersion = CUR_BLOB_VERSION
    uBlob.aiKeyAlg = CALG_RC2
    Debug.Assert UBound(uBlob.Buffer) >= UBound(baKey)
    uBlob.cbKeySize = UBound(baKey) + 1
    Call CopyMemory(uBlob.Buffer(0), baKey(0), uBlob.cbKeySize)
    lResult = CryptImportKey(m_uData.hProv, uBlob, sizeof_BLOBHEADER + uBlob.cbKeySize, 0, CRYPT_EXPORTABLE Or CRYPT_IPSEC_HMAC_KEY, hKey)
    If pvCryptoSetApiError(lResult, "CryptImportKey") Then
        GoTo QH
    End If
    lResult = CryptCreateHash(m_uData.hProv, CALG_HMAC, hKey, 0, hHash)
    If pvCryptoSetApiError(lResult, "CryptCreateHash") Then
        GoTo QH
    End If
    uInfo.HashAlgid = lHashAlgId
    lResult = CryptSetHashParam(hHash, HP_HMAC_INFO, uInfo, 0)
    If pvCryptoSetApiError(lResult, "CryptSetHashParam") Then
        GoTo QH
    End If
    If Size < 0 Then
        Size = pvArraySize(baInput)
    End If
    If Size > 0 Then
        lResult = CryptHashData(hHash, baInput(Pos), Size, 0)
        If pvCryptoSetApiError(lResult, "CryptHashData") Then
            GoTo QH
        End If
    End If
    lHashSize = 1024
    pvArrayAllocate baRetVal, lHashSize, FUNC_NAME & ".baRetVal"
    lResult = CryptGetHashParam(hHash, HP_HASHVAL, baRetVal(0), lHashSize, 0)
    If pvCryptoSetApiError(lResult, "CryptGetHashParam") Then
        GoTo QH
    End If
    pvArrayReallocate baRetVal, lHashSize, FUNC_NAME & ".baRetVal"
    '--- success
    pvCryptoHmac = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    pvCryptoCheckApiError FUNC_NAME
End Function

#If ImplLibSodium Then
Private Function pvCryptoAeadChacha20Poly1305Encrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_CHACHA20POLY1305_TAGSZ
    If lAdSize > 0 Then
        lAdPtr = VarPtr(baAad(lAadPos))
    End If
    Call crypto_aead_chacha20poly1305_ietf_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
    '--- success
    pvCryptoAeadChacha20Poly1305Encrypt = True
End Function

Private Function pvCryptoAeadChacha20Poly1305Decrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    If crypto_aead_chacha20poly1305_ietf_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
        '--- success
        pvCryptoAeadChacha20Poly1305Decrypt = True
    End If
End Function
#End If

Private Function pvCryptoAeadAesGcmEncrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_AESGCM_TAGSZ
    If lAdSize > 0 Then
        lAdPtr = VarPtr(baAad(lAadPos))
    End If
#If ImplLibSodium Then
    Call crypto_aead_aes256gcm_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
#Else
    Const FUNC_NAME     As String = "pvCryptoAeadAesGcmEncrypt"
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim hAlg            As Long
    Dim hKey            As Long
    Dim uInfo           As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    Dim lResult         As Long

    pvCryptoClearApiError
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("AES"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ChainingMode"), StrPtr("ChainingModeGCM"), 32, 0)
    If pvCryptoSetApiResult(hResult, "BCryptSetProperty(ChainingMode)") Then
        GoTo QH
    End If
    hResult = BCryptGenerateSymmetricKey(hAlg, hKey, 0, 0, VarPtr(baKey(0)), UBound(baKey) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptGenerateSymmetricKey") Then
        GoTo QH
    End If
    With uInfo
        .cbSize = Len(uInfo)
        .dwInfoVersion = 1
        .pbNonce = VarPtr(baNonce(0))
        .cbNonce = UBound(baNonce) + 1
        .pbAuthData = lAdPtr
        .cbAuthData = lAdSize
        .pbTag = VarPtr(baBuffer(lPos + lSize))
        .cbTag = LNG_AESGCM_TAGSZ
    End With
    hResult = BCryptEncrypt(hKey, VarPtr(baBuffer(lPos)), lSize, VarPtr(uInfo), 0, 0, VarPtr(baBuffer(lPos)), lSize, lResult, 0)
    Debug.Assert lResult = lSize
    If pvCryptoSetApiResult(hResult, "BCryptEncrypt") Then
        GoTo QH
    End If
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
#End If
    '--- success
    pvCryptoAeadAesGcmEncrypt = True
End Function

Private Function pvCryptoAeadAesGcmDecrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
#If ImplLibSodium Then
    If crypto_aead_aes256gcm_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
        '--- success
        pvCryptoAeadAesGcmDecrypt = True
    End If
#Else
    Const FUNC_NAME     As String = "pvCryptoAeadAesGcmDecrypt"
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim hAlg            As Long
    Dim hKey            As Long
    Dim uInfo           As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    Dim lResult         As Long

    pvCryptoClearApiError
    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("AES"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ChainingMode"), StrPtr("ChainingModeGCM"), 32, 0)
    If pvCryptoSetApiResult(hResult, "BCryptSetProperty(ChainingMode)") Then
        GoTo QH
    End If
    hResult = BCryptGenerateSymmetricKey(hAlg, hKey, 0, 0, VarPtr(baKey(0)), UBound(baKey) + 1, 0)
    If pvCryptoSetApiResult(hResult, "BCryptGenerateSymmetricKey") Then
        GoTo QH
    End If
    With uInfo
        .cbSize = Len(uInfo)
        .dwInfoVersion = 1
        .pbNonce = VarPtr(baNonce(0))
        .cbNonce = UBound(baNonce) + 1
        .pbAuthData = VarPtr(baAad(lAadPos))
        .cbAuthData = lAdSize
        .pbTag = VarPtr(baBuffer(lPos + lSize - LNG_AESGCM_TAGSZ))
        .cbTag = LNG_AESGCM_TAGSZ
    End With
    hResult = BCryptDecrypt(hKey, VarPtr(baBuffer(lPos)), lSize - LNG_AESGCM_TAGSZ, VarPtr(uInfo), 0, 0, VarPtr(baBuffer(lPos)), lSize - LNG_AESGCM_TAGSZ, lResult, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    Debug.Assert lResult = lSize - LNG_AESGCM_TAGSZ
    '--- success
    pvCryptoAeadAesGcmDecrypt = True
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    pvCryptoCheckApiError FUNC_NAME
#End If
End Function

Private Sub pvCryptoRandomBytes(ByVal lPtr As Long, ByVal lSize As Long)
    Call CryptGenRandom(m_uData.hProv, lSize, lPtr)
End Sub

Private Sub pvCryptoClearApiError()
    m_uData.hResult = 0
    m_uData.ApiSource = vbNullString
End Sub

Private Function pvCryptoSetApiResult(ByVal hResult As Long, sApiSource As String) As Boolean
    If hResult <> 0 Then
        m_uData.hResult = hResult
        m_uData.ApiSource = sApiSource
        '--- success
        pvCryptoSetApiResult = True
    End If
End Function

Private Function pvCryptoSetApiError(ByVal lResult As Long, sApiSource As String) As Boolean
    If lResult = 0 Then
        m_uData.hResult = Err.LastDllError
        m_uData.ApiSource = sApiSource
        '--- success
        pvCryptoSetApiError = True
    End If
End Function

Private Sub pvCryptoCheckApiError(sSource As String)
    Const LNG_FACILITY_WIN32 As Long = &H80070000

    If LenB(m_uData.ApiSource) <> 0 Then
        Err.Raise IIf(m_uData.hResult < 0, m_uData.hResult, m_uData.hResult Or LNG_FACILITY_WIN32), sSource & "." & m_uData.ApiSource
    End If
End Sub

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
    Static Map(0 To &H3FF) As Long
    Dim baRetVal()      As Byte
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lSize           As Long
    Dim lPtr            As Long
    
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
    lSize = ((UBound(baInput) + 1) \ 4) * 3
    pvArrayAllocate baRetVal, lSize, FUNC_NAME & ".baRetVal"
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(baRetVal(lPtr), lChar, 3)
        lPtr = lPtr + 3
    Next
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
    On Error Resume Next
    RetVal = oCol.Item(Index)
    pvCallCollectionItem = -Abs(Err.Number)
End Function

Private Function pvCallCollectionRemove(ByVal oCol As Collection, Index As Variant) As Long
    On Error Resume Next
    oCol.Remove Index
    pvCallCollectionRemove = -Abs(Err.Number)
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

#If ImplTlsServer Or Not ImplLibSodium Then
Private Property Get RealOsVersion() As UcsOsVersionEnum
    Static lVersion     As Long
    Dim baBuffer()      As Byte
    Dim lPtr            As Long
    Dim lSize           As Long
    Dim aVer(0 To 9)    As Integer
    
    If lVersion = 0 Then
        ReDim baBuffer(0 To 8192) As Byte
        Call GetFileVersionInfo(StrPtr("kernel32.dll"), 0, UBound(baBuffer), baBuffer(0))
        Call VerQueryValue(baBuffer(0), StrPtr("\"), lPtr, lSize)
        Call CopyMemory(aVer(0), ByVal lPtr, 20)
        lVersion = aVer(9) * 100 + aVer(8)
    End If
    RealOsVersion = lVersion
End Property
#End If

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

Private Function TlsDesignDumpArray(baData() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As String
    If Size < 0 Then
        Size = UBound(baData) + 1 - Pos
    End If
    If Size > 0 Then
        TlsDesignDumpArray = TlsDesignDumpMemory(VarPtr(baData(Pos)), Size)
    End If
End Function

Private Function TlsDesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
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
#End If
