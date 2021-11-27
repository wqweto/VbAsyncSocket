Attribute VB_Name = "Module1"
Option Explicit

'=========================================================================
' API
'=========================================================================

'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_NEWKEYSET                           As Long = &H8
Private Const CRYPT_DELETEKEYSET                        As Long = &H10
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
'Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
'--- for CryptSignHash
Private Const AT_SIGNATURE                              As Long = 2
Private Const RSA1024BIT_KEY                            As Long = &H4000000
'--- for CertGetCertificateContextProperty
Private Const CERT_KEY_PROV_INFO_PROP_ID                As Long = 2
'--- for PFXImportCertStore
Private Const CRYPT_EXPORTABLE                          As Long = &H1
'--- for CryptExportKey
Private Const PRIVATEKEYBLOB                            As Long = 7
'--- for CryptAcquireCertificatePrivateKey
Private Const CRYPT_ACQUIRE_CACHE_FLAG                  As Long = &H1
Private Const CRYPT_ACQUIRE_SILENT_FLAG                 As Long = &H40
Private Const CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG       As Long = &H10000
'--- for CertStrToName
Private Const CERT_OID_NAME_STR                         As Long = 2
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
'--- BLOBs magic
Private Const BCRYPT_RSAPRIVATE_MAGIC                   As Long = &H32415352
Private Const BCRYPT_ECDH_PRIVATE_P256_MAGIC            As Long = &H324B4345
Private Const BCRYPT_ECDH_PRIVATE_P384_MAGIC            As Long = &H344B4345
Private Const BCRYPT_ECDH_PRIVATE_P521_MAGIC            As Long = &H364B4345

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
'--- Crypt32
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvParameters As Long, phCryptProvOrNCryptKey As Long, pdwKeySpec As Long, pfCallerFreeProvOrNCryptKey As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertStrToName Lib "crypt32" Alias "CertStrToNameW" (ByVal dwCertEncodingType As Long, ByVal pszX500 As Long, ByVal dwStrType As Long, ByVal pvReserved As Long, pbEncoded As Any, pcbEncoded As Long, ByVal ppszError As Long) As Long
Private Declare Function CertCreateSelfSignCertificate Lib "crypt32" (ByVal hCryptProvOrNCryptKey As Long, pSubjectIssuerBlob As Any, ByVal dwFlags As Long, pKeyProvInfo As Any, ByVal pSignatureAlgorithm As Long, pStartTime As Any, pEndTime As Any, ByVal pExtensions As Long) As Long

Private Type CRYPT_DATA_BLOB
    cbData              As Long
    pbData              As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId            As Long
    Parameters          As CRYPT_DATA_BLOB
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
    PrivateKey          As CRYPT_DATA_BLOB
    pAttributes         As Long
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

Private Const LNG_FACILITY_WIN32                        As Long = &H80070000

Public Function PkiGenerSelfSignedCertificate(cCerts As Collection, cPrivKey As Collection, Optional ByVal Subject As String) As Boolean
    Const FUNC_NAME     As String = "PkiGenerSelfSignedCertificate"
    Dim hProv           As Long
    Dim hKey            As Long
    Dim sName           As String
    Dim baName()        As Byte
    Dim lSize           As Long
    Dim uName           As CRYPT_DATA_BLOB
    Dim uExpire         As SYSTEMTIME
    Dim uInfo           As CRYPT_KEY_PROV_INFO
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, 0) = 0 Then
        If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptAcquireContext"
            GoTo QH
        End If
    End If
    If CryptGenKey(hProv, AT_SIGNATURE, RSA1024BIT_KEY Or CRYPT_EXPORTABLE, hKey) = 0 Then
        GoTo QH
    End If
    If Left$(Subject, 3) <> "CN=" Then
        If LenB(Subject) = 0 Then
            Subject = LCase$(Environ$("COMPUTERNAME") & IIf(LenB(Environ$("USERDNSDOMAIN")) <> 0, "." & Environ$("USERDNSDOMAIN"), vbNullString))
        End If
        sName = "CN=""" & Replace(Subject, """", """""") & """" & ",OU=""" & Replace(Environ$("USERDOMAIN") & "\" & Environ$("USERNAME"), """", """""") & """,O=""Self-Signed Certificate"""
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
    If pvPkiAppendCertContext(pCertContext, cCerts, cPrivKey) Then
        '--- success
        PkiGenerSelfSignedCertificate = True
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

Private Function pvPkiAppendCertContext(ByVal pCertContext As Long, cCerts As Collection, cPrivKey As Collection) As Boolean
    Const FUNC_NAME     As String = "pvPkiAppendCertContext"
    Dim uCertContext    As CERT_CONTEXT
    Dim baBuffer()      As Byte
    Dim baPrivKey()     As Byte
    
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
        Set cPrivKey = New Collection
        cPrivKey.Add baPrivKey
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
        Err.Raise vbObjectError, , "NCrypt not supported"
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
            Err.Raise vbObjectError, , "Export ECC private key not implemented"
        Case Else
            #If ImplUseDebugLog Then
                DebugLog MODULE_NAME, FUNC_NAME, Replace(ERR_UNKNOWN_CAPI_MAGIC, "%1", "&H" & Hex$(lMagic)), vbLogEventTypeWarning
            #End If
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

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
End Sub
