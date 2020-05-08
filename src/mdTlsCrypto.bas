Attribute VB_Name = "mdTlsCrypto"
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

'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
''--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
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

Private m_baBuffer()                As Byte
Private m_lBuffIdx                  As Long
Private m_uData                     As UcsCryptoThunkData
Public g_oRequestSocket             As cTlsSocket

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
' Functions
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
    Dim baCompressed()  As Byte
    Dim lIdx            As Long
    
    pvArrayAllocate baPrivate, m_uData.Ecc256KeySize, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baCompressed, m_uData.Ecc256KeySize + 1, FUNC_NAME & ".baCompressed"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baCompressed(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    If Not CryptoEccSecp256r1UncompressKey(baPublic, baCompressed) Then
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
    Dim baCompressed()  As Byte
    Dim lIdx            As Long
        
    pvArrayAllocate baPrivate, m_uData.Ecc384KeySize, FUNC_NAME & ".baPrivate"
    pvArrayAllocate baCompressed, 1 + m_uData.Ecc384KeySize, FUNC_NAME & ".baCompressed"
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc384KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp384r1MakeKey), baCompressed(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx > MAX_RETRIES Then
        GoTo QH
    End If
    If Not CryptoEccSecp384r1UncompressKey(baPublic, baCompressed) Then
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

Private Sub pvArrayAllocate(baRetVal() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
    Else
        baRetVal = vbNullString
    End If
    Debug.Assert RedimStats(sFuncName, lSize)
End Sub

Private Sub pvArrayReallocate(baArray() As Byte, ByVal lSize As Long, sFuncName As String)
    If lSize > 0 Then
        ReDim Preserve baArray(0 To lSize - 1) As Byte
    Else
        baArray = vbNullString
    End If
    Debug.Assert RedimStats(sFuncName, lSize)
End Sub

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Property

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
    
    pvPatchMethodTrampoline AddressOf mdTlsCrypto.pvCallCollectionItem, IDX_COLLECTION_ITEM
    pvCallCollectionItem = pvCallCollectionItem(oCol, Index, RetVal)
End Function

Private Function pvCallCollectionRemove(ByVal oCol As Collection, Index As Variant) As Long
    Const IDX_COLLECTION_REMOVE As Long = 10
    
    pvPatchMethodTrampoline AddressOf mdTlsCrypto.pvCallCollectionRemove, IDX_COLLECTION_REMOVE
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
    pvAppendBuffer &H76397160, &H76401BC0, &H763973C0, &H0&, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &H0&, &H0&, &H0&, &H1&, &HFFFFFFFF, &H27D2604B, &H3BCE3C3E, &HCC53B0F6, &H651D06B0, &H769886BC, &HB3EBBD55, &HAA3A93E7, &H5AC635D8, &HD898C296, &HF4A13945, &H2DEB33A0, &H77037D81, &H63A440F2, &HF8BCE6E5, &HE12C4247, &H6B17D1F2, &H37BF51F5, &HCBB64068, &H6B315ECE, &H2BCE3357
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
    ReDim m_baBuffer(0 To 44160 - 1) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer &HBA2238, &H2780&, &H2AA0&, &H3FE0&, &H46A0&, &H4740&, &H48D0&, &H4E10&, &H3F30&, &H4630&, &H4700&, &H4780&, &H4A10&, &H3420&, &H3470&, &H3250&, &H34E0&, &H3570&, &H34B0&, &H3900&, &H3990&, &H3580&, &H26A0&, &H2660&, &H1BB0&, &H1B30&, &H7DB0&, &HCCCCCCCC, &HE8&, &H752D5800, &H500B940, &HB94000
    pvAppendBuffer &HCCC3008B, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &HE8&, &H952D5800, &H500B940, &HB94000, &HCCCCCCC3, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H8B5368EC, &HE853105D, &H8D60&, &H850FC085, &H15B&, &HC758B56, &H57C8458D, &HA9E85056, &H8B00009D, &H458D087D, &H8D5750C8, &HE8509845, &H9D38&, &H50C8458D, &H9D8EE850, &H56530000, &H9D26E856, &H53530000, &H9D7FE8
    pvAppendBuffer &HFF6AE800, &HB005FFFF, &H50000000, &HE8575753, &H950C&, &HFFFF57E8, &HB005FF, &H53500000, &HF9E85353, &HE8000094, &HFFFFFF44, &HB005&, &H57535000, &H9DA6E853, &H57530000, &H9CDEE857, &H29E80000, &H5FFFFFF, &HB0&, &H53575750, &H94CBE8, &HFF16E800, &HB005FFFF, &H50000000, &HE8575753, &H94B8&, &HE857006A, &HA9A0&, &H2574C20B, &HFFFEF7E8, &HB005FF, &H57500000
    pvAppendBuffer &H878AE857, &H8B570000, &HA172E8F0, &HE6C10000, &H2C77091F, &HEB0C758B, &H61E85706, &H570000A1, &H9CDAE853, &HC5E80000, &H5FFFFFE, &HB0&, &H98458D50, &HE8535350, &H9D24&, &HFFFEAFE8, &HB005FF, &H8D500000, &H53509845, &H9D0EE853, &H99E80000, &H5FFFFFE, &HB0&, &H458D5350, &HE8505098, &H9CF8&, &H5098458D, &H2DE85757, &HE800009C, &HFFFFFE78, &HB005&, &H458D5000
    pvAppendBuffer &H505750C8, &H9CD7E8, &HE8575300, &HA1B0&, &HA9E85356, &H8D0000A1, &H5650C845, &HA19FE8, &H5B5E5F00, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &H83EC8B55, &H8B5348EC, &HE853105D, &H8C10&, &H850FC085, &H147&, &HC758B56, &H57D8458D, &H59E85056, &H8B00009C, &H458D087D, &H8D5750D8, &HE850B845, &H9BE8&, &H50D8458D, &H9C3EE850, &H56530000, &H9BD6E856, &H53530000, &H9C2FE8
    pvAppendBuffer &HFDEAE800, &HC083FFFF, &H57535010, &H93CEE857, &HD9E80000, &H83FFFFFD, &H535010C0, &HBDE85353, &HE8000093, &HFFFFFDC8, &H5010C083, &HE8535753, &H9C5C&, &HE8575753, &H9B94&, &HFFFDAFE8, &H10C083FF, &H53575750, &H9393E8, &HFD9EE800, &HC083FFFF, &H57535010, &H9382E857, &H6A0000, &HA82AE857, &HC20B0000, &H81E82374, &H83FFFFFD, &H575010C0, &H8856E857, &H8B570000, &HA05EE8F0
    pvAppendBuffer &HE6C10000, &H1C77091F, &HEB0C758B, &H4DE85706, &H570000A0, &H9B96E853, &H51E80000, &H83FFFFFD, &H8D5010C0, &H5350B845, &H9BE2E853, &H3DE80000, &H83FFFFFD, &H8D5010C0, &H5350B845, &H9BCEE853, &H29E80000, &H83FFFFFD, &H535010C0, &H50B8458D, &H9BBAE850, &H458D0000, &H575750B8, &H9AEFE8, &HFD0AE800, &HC083FFFF, &H458D5010, &H505750D8, &H9B9BE8, &HE8575300, &HA0A4&, &H9DE85356
    pvAppendBuffer &H8D0000A0, &H5650D845, &HA093E8, &H5B5E5F00, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H5608758B, &H8A73E8, &H74C08500, &H30468D17, &H8A66E850, &HC0850000, &H1B80A74, &H5E000000, &H4C25D, &H5D5EC033, &HCC0004C2, &H56EC8B55, &H5608758B, &H8A73E8, &H74C08500, &H20468D17, &H8A66E850, &HC0850000, &H1B80A74, &H5E000000, &H4C25D, &H5D5EC033, &HCC0004C2
    pvAppendBuffer &H81EC8B55, &HF8EC&, &H5D8B5300, &H98458D0C, &H50535756, &H9FA7E8, &H30438D00, &HF8458950, &HFF38858D, &HE850FFFF, &H9F94&, &H8D1475FF, &HFFFF0885, &H858D50FF, &HFFFFFF68, &H38858D50, &H50FFFFFF, &H5098458D, &H803E8, &H105D8B00, &H9DAAE853, &H708D0000, &H7EF685FE, &H1F0F60, &HA9E85356, &HB0000A6, &HB80775C2, &H1&, &HC03302EB, &HC140048D, &H8D8D04E0, &HFFFFFF08
    pvAppendBuffer &H958DC803, &HFFFFFF68, &H4D89D003, &HD8F75114, &H8DFC5589, &HFFFF38BD, &HF80352FF, &H3985D8D, &HE85357D8, &H4A8&, &H75FF5357, &HFC75FF14, &H29BE8, &H105D8B00, &H7FF6854E, &H53006AA3, &HA64BE8, &H75C20B00, &H1B807, &H2EB0000, &H48DC033, &H4E0C140, &HFF089D8D, &HD803FFFF, &HFF688D8D, &H353FFFF, &H38BD8DC8, &H51FFFFFF, &H4D89F82B, &H98758D10, &H5657F02B, &H44FE8
    pvAppendBuffer &HFB6AE800, &HB005FFFF, &H50000000, &HFF68858D, &H8D50FFFF, &H8D509845, &HE850C845, &H99C0&, &HC8458D57, &HF5E85050, &HFF000098, &H458D0C75, &HE85050C8, &H98E8&, &HFFFB33E8, &HB005FF, &H8D500000, &H5050C845, &H9153E8, &HF875FF00, &H50C8458D, &H98C6E850, &H8D560000, &H5050C845, &H98BBE8, &H53565700, &HE81075FF, &H1E0&, &H50C8458D, &HFF08858D, &H8D50FFFF, &HFFFF6885
    pvAppendBuffer &H89E850FF, &H8B00000A, &H858D0875, &HFFFFFF68, &H29E85650, &H8D00009E, &HFFFF0885, &H468D50FF, &H19E85030, &H5F00009E, &HE58B5B5E, &H10C25D, &H81EC8B55, &HA8EC&, &H5D8B5300, &HB8458D0C, &H50535756, &H9E57E8, &H20438D00, &HF8458950, &HFF78858D, &HE850FFFF, &H9E44&, &H8D1475FF, &HFFFF5885, &H458D50FF, &H858D5098, &HFFFFFF78, &HB8458D50, &H6D6E850, &H5D8B0000, &H4DE85310
    pvAppendBuffer &H8300009C, &H458902E8, &H7EC08514, &H1F0F5B, &HF9E85350, &HB0000A4, &HB80775C2, &H1&, &HC03302EB, &H8D05E0C1, &HFFFF589D, &H8DD803FF, &HC803984D, &HFF78B58D, &HF753FFFF, &HFC4D89D8, &H8DF00351, &HF803B87D, &H91E85756, &H56000004, &H75FF5357, &H1F6E8FC, &H458B0000, &H105D8B14, &H14458948, &HA87FC085, &HE853006A, &HA4A0&, &H575C20B, &HEB01488D, &HC1C93302, &H9D8D05E1
    pvAppendBuffer &HFFFFFF58, &H4D89D903, &H458D5310, &H8DC10398, &HFFFF78BD, &HF92B50FF, &H2BB8758D, &HE85657F1, &H43C&, &HFFF9C7E8, &H10C083FF, &H98458D50, &HB8458D50, &HD8458D50, &H9852E850, &H8D570000, &H5050D845, &H9787E8, &HC75FF00, &H50D8458D, &H977AE850, &H95E80000, &H83FFFFF9, &H8D5010C0, &H5050D845, &H9217E8, &HF875FF00, &H50D8458D, &H975AE850, &H8D560000, &H5050D845, &H974FE8
    pvAppendBuffer &H8D565700, &H45039845, &HE8505310, &H140&, &H50D8458D, &HFF58858D, &H8D50FFFF, &HE8509845, &H93C&, &H8D08758B, &H56509845, &H9CEFE8, &H58858D00, &H50FFFFFF, &H5020468D, &H9CDFE8, &H5B5E5F00, &HC25DE58B, &HCCCC0010, &HCCCCCCCC, &H83EC8B55, &H565330EC, &HF912E857, &H5D8BFFFF, &HB00508, &H758B0000, &H56535010, &H50D0458D, &H976BE8, &HD0458D00, &H1E85050, &H8D000097
    pvAppendBuffer &H5350D045, &H9696E853, &H458D0000, &H565650D0, &H968BE8, &HF8D6E800, &H758BFFFF, &HB0050C, &H7D8B0000, &H57565014, &H9732E857, &H8D570000, &HE850D045, &H96C8&, &HFFF8B3E8, &HB005FF, &H53500000, &H50D0458D, &H9712E850, &H9DE80000, &H5FFFFF8, &HB0&, &H10458B50, &HD0458D50, &HF9E85050, &HE8000096, &HFFFFF884, &HB005&, &H458B5000, &H50505310, &H96E3E8, &H10458B00
    pvAppendBuffer &HE8565650, &H9618&, &HFFF863E8, &HB005FF, &H8D500000, &H5350D045, &H53105D8B, &H96BFE8, &H57575300, &H95F7E8, &HF842E800, &HB005FFFF, &H50000000, &HE8575756, &H96A4&, &H50D0458D, &H9B7AE853, &H5E5F0000, &H5DE58B5B, &HCC0010C2, &H83EC8B55, &H565320EC, &HF812E857, &H5D8BFFFF, &H10C08308, &H5010758B, &H458D5653, &H9DE850E0, &H8D000096, &H5050E045, &H9633E8, &HE0458D00
    pvAppendBuffer &HE8535350, &H95C8&, &H50E0458D, &HBDE85656, &HE8000095, &HFFFFF7D8, &H830C758B, &H7D8B10C0, &H57565014, &H9666E857, &H8D570000, &HE850E045, &H95FC&, &HFFF7B7E8, &H10C083FF, &H458D5350, &HE85050E0, &H9648&, &HFFF7A3E8, &H10C083FF, &H10458B50, &HE0458D50, &H31E85050, &HE8000096, &HFFFFF78C, &H5010C083, &H5310458B, &H1DE85050, &H8B000096, &H56501045, &H9552E856, &H6DE80000
    pvAppendBuffer &H83FFFFF7, &H8D5010C0, &H5350E045, &H53105D8B, &H95FBE8, &H57575300, &H9533E8, &HF74EE800, &HC083FFFF, &H57565010, &H95E2E857, &H458D0000, &HE85350E0, &H9AE8&, &H8B5B5E5F, &H10C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &H90EC&, &H57565300, &HFFF70FE8, &H85D8BFF, &HB005&, &H107D8B00, &H8D575350, &HE850A045, &H9568&, &H50A0458D, &H94FEE850
    pvAppendBuffer &H458D0000, &H535350A0, &H9493E8, &HA0458D00, &HE8575750, &H9488&, &HFFF6D3E8, &HC5D8BFF, &HB005&, &H14758B00, &H8D565350, &HE850A045, &H8C6C&, &HFFF6B7E8, &HB005FF, &H53500000, &H19E85656, &HE8000095, &HFFFFF6A4, &HB005&, &H75FF5000, &HD0458D08, &H1E85057, &H8D000095, &H5350D045, &H9436E853, &H81E80000, &H5FFFFF6, &HB0&, &H75FF5750, &HD0458D08, &H8C1EE850
    pvAppendBuffer &H57560000, &H9477E8, &HF662E800, &HB005FFFF, &H50000000, &H50D0458D, &HC1E85757, &HE8000094, &HFFFFF64C, &HB005&, &H8B575000, &H858D087D, &HFFFFFF70, &HA5E85057, &H8D000094, &HFFFF7085, &H565650FF, &H93D7E8, &HF622E800, &HB005FFFF, &H50000000, &HE8565653, &H9484&, &H50A0458D, &HFF70858D, &HE850FFFF, &H9414&, &HFFF5FFE8, &HB005FF, &H8D500000, &H8D50D045, &HFFFF7085
    pvAppendBuffer &HE85050FF, &H9458&, &HFFF5E3E8, &HB005FF, &H57500000, &HFF70858D, &H8D50FFFF, &HE850D045, &H943C&, &H50A0458D, &H50D0458D, &H936EE850, &HB9E80000, &H5FFFFF5, &HB0&, &H458D5350, &HE85350D0, &H9418&, &HFF70858D, &H5750FFFF, &H98EBE8, &H5B5E5F00, &HC25DE58B, &HCCCC0010, &H83EC8B55, &H565360EC, &HF582E857, &H5D8BFFFF, &H10C08308, &H50107D8B, &H458D5753, &HDE850C0
    pvAppendBuffer &H8D000094, &H5050C045, &H93A3E8, &HC0458D00, &HE8535350, &H9338&, &H50C0458D, &H2DE85757, &HE8000093, &HFFFFF548, &H830C5D8B, &H758B10C0, &H56535014, &H50C0458D, &H8B23E8, &HF52EE800, &HC083FFFF, &H56535010, &H93C2E856, &H1DE80000, &H83FFFFF5, &HFF5010C0, &H458D0875, &HE85057E0, &H93AC&, &H50E0458D, &HE1E85353, &HE8000092, &HFFFFF4FC, &H5010C083, &H875FF57, &H50E0458D
    pvAppendBuffer &H8ADBE8, &HE8575600, &H9324&, &HFFF4DFE8, &H10C083FF, &HE0458D50, &HE8575750, &H9370&, &HFFF4CBE8, &H10C083FF, &H7D8B5750, &HA0458D08, &H59E85057, &H8D000093, &H5650A045, &H928EE856, &HA9E80000, &H83FFFFF4, &H535010C0, &H3DE85656, &H8D000093, &H8D50C045, &HE850A045, &H92D0&, &HFFF48BE8, &H10C083FF, &HE0458D50, &HA0458D50, &H19E85050, &HE8000093, &HFFFFF474, &H5010C083
    pvAppendBuffer &HA0458D57, &HE0458D50, &H9302E850, &H458D0000, &H458D50C0, &HE85050E0, &H9234&, &HFFF44FE8, &H10C083FF, &H458D5350, &HE85350E0, &H92E0&, &H50A0458D, &H97E6E857, &H5E5F0000, &H5DE58B5B, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H8B5630EC, &H56570875, &HE81075FF, &H975C&, &H570C7D8B, &HE81475FF, &H9750&, &H50D0458D, &H8017E8, &H18458B00, &H1D045C7
    pvAppendBuffer &HC7000000, &HD445&, &HC0850000, &H8D500A74, &HE850D045, &H9728&, &H50D0458D, &H6DE85657, &H8D000003, &H5750D045, &HF402E856, &H458DFFFF, &H75FF50D0, &H1075FF14, &H353E8, &H8B5E5F00, &H14C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H8B5620EC, &H56570875, &HE81075FF, &H973C&, &H570C7D8B, &HE81475FF, &H9730&, &H50E0458D, &H7FF7E8, &H18458B00, &H1E045C7
    pvAppendBuffer &HC7000000, &HE445&, &HC0850000, &H8D500A74, &HE850E045, &H9708&, &H50E0458D, &H3DE85657, &H8D000003, &H5750E045, &HF502E856, &H458DFFFF, &H75FF50E0, &H1075FF14, &H323E8, &H8B5E5F00, &H14C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H24448B53, &H244C8B0C, &H8BE1F710, &H24448BD8, &H2464F708, &H8BD80314, &HF7082444, &H5BD303E1, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H7340F980, &H20F98015, &HA50F0673, &HC3E0D3C2, &HC033D08B, &HD31FE180, &HC033C3E2, &HCCC3D233, &H7340F980, &H20F98015, &HAD0F0673, &HC3EAD3D0, &HD233C28B, &HD31FE180, &HC033C3E8, &HCCC3D233, &H8BEC8B55, &H56531045, &H8D08758B, &H8B577848, &H568D0C7D, &H77F13B78, &H73D03B04, &H784F8D0B, &H3077F13B, &H2C72D73B, &H10BBF82B, &H2B000000, &H38148BF0, &H4C8B1003, &H48130438, &H8408D04
    pvAppendBuffer &HF8305489, &HFC304C89, &H7501EB83, &H5B5E5FE4, &HCC25D, &H488DD78B, &H2BDE8B10, &H2BD82BD0, &H4B8FE, &H768D0000, &H20498D20, &HD041100F, &H374C100F, &HD40F66E0, &H4E110FC8, &H4C100FE0, &H100FE00A, &HF66E041, &H110FC8D4, &H83E00B4C, &HD27501E8, &H5D5B5E5F, &HCC000CC2, &HCCCCCCCC, &H8BEC8B55, &HEC831C55, &H20458B08, &H8758B56, &HC7D8B57, &H4513D703, &H89168910, &H453B0446
    pvAppendBuffer &H720F7710, &H73D73B04, &H1B809, &HC9330000, &H570F0EEB, &H130F66C0, &H4D8BF845, &HF8458BFC, &H5F244503, &H3284D13, &H46891445, &H13C68B08, &H4E89184D, &HE58B5E0C, &H24C25D, &HCCCCCCCC, &H8BEC8B55, &H4D8B0C55, &H31028B08, &H4428B01, &H8B044131, &H41310842, &HC428B08, &H5D0C4131, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H4D8B08EC, &H10558B08, &H18B5653
    pvAppendBuffer &HC104598D, &HF63302EA, &H89105589, &H48DF85D, &H485&, &HFC458900, &H74D28557, &HC558B42, &H83107D8B, &H666602C2, &H841F0F, &H0&, &HFE4AB60F, &HF04528D, &HC1FB42B6, &HC80B08E1, &HFC42B60F, &HB08E1C1, &H42B60FC8, &H8E1C1FD, &HC89C80B, &HF73B46B3, &H458BD672, &HB9D78BFC, &H1&, &H4D89FF33, &HFF03B0C, &H8D83&, &H2BC68B00, &H83048DC2, &HF084589, &H441F&
    pvAppendBuffer &HFCB35C8B, &H875FA3B, &H89FF3341, &H4EB0C4D, &H2D75FF85, &HFFF0D7E8, &H58805FF, &HC3C10000, &HE8535008, &H7848&, &HC1E8D88B, &H8BFFFFF0, &HB60F0C4D, &H6880884, &HE0C10000, &HEBD83318, &H6FA831D, &HFF831E76, &HE8197504, &HFFFFF0A0, &H58805, &HE8535000, &H7814&, &H458BD88B, &H10558B08, &H3347088B, &H4C083CB, &H89F85D8B, &HC890845, &H4D8B46B3, &HFC753B0C, &H5E5F8272
    pvAppendBuffer &H5DE58B5B, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D30EC, &H1075FFD0, &H8E5EE850, &H458D0000, &H458B50D0, &HE8505008, &H8DF0&, &H8D1075FF, &H5050D045, &H8DE3E8, &HD0458D00, &HC458B50, &HD5E85050, &H8B00008D, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D20EC, &H1075FFE0, &H8E3EE850, &H458D0000, &H458B50E0, &HE8505008, &H8DD0&
    pvAppendBuffer &H8D1075FF, &H5050E045, &H8DC3E8, &HE0458D00, &HC458B50, &HB5E85050, &H8B00008D, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H56530CEC, &H570C758B, &H7D893E8B, &HBD1C8DF8, &H0&, &HE8F45D89, &HFFFFEFA0, &HFF008B53, &H89D233D0, &HFF85FC45, &H9066117E, &HCA2B0E8B, &H898E0C8B, &H3B42900C, &H8BF17CD7, &H38B085D, &H3B084589, &H8D067FC7, &H45890147
    pvAppendBuffer &H85348D08, &H0&, &HFFEF63E8, &H8B56FF, &H558BD0FF, &HC458908, &H127ED285, &H8B0C7D8B, &H2E9C1CE, &HABF3C033, &H8BF87D8B, &H1B90C45, &H39000000, &H8D177C0B, &HD603FC50, &H401F0F, &H8D8B048B, &H4289FC52, &HB3B4104, &H5D8BF27E, &H531B8BFC, &H507E8, &H85D08B00, &H8B1A74D2, &H83E3D3CA, &H117E01FF, &HB9FC458B, &H20&, &H408BCA2B, &HBE8D304, &HE85352D8, &H6B80&
    pvAppendBuffer &H1475FF50, &HFC7D8B57, &H875FF57, &HE80C75FF, &H43AC&, &H85105D8B, &HB82E74DB, &H1&, &H257C0339, &H8B0C7D8B, &HC7830855, &HFE034AFC, &H478D285, &H2EB0F8B, &HC89C933, &H83404A83, &H33B04EF, &H7D8BEA7E, &HF44D8BFC, &HD74C985, &HC6C78B, &H1408D00, &H7501E983, &HEE8EE8F5, &H8B57FFFF, &HD0FF0840, &H850C5D8B, &H851174DB, &H8B0D74F6, &HC6C3&, &H8301408D, &HF57501EE
    pvAppendBuffer &HFFEE6BE8, &H408B53FF, &H5FD0FF08, &HE58B5B5E, &H10C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H53EC8B55, &H560C5D8B, &H8D1B8B57, &H49D3C, &H39E80000, &H57FFFFEE, &HD0FF008B, &H6AF08B57, &HE9E85600, &H8300004B, &H1E890CC4, &HFF56006A, &H75FF0C75, &HFE56E808, &H3E83FFFF, &H90117601, &H3C83068B, &H8750086, &H83068948, &HF07701F8, &H5EC68B5F, &H8C25D5B, &HCCCCCC00, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H458B1CEC, &H57565308, &H458B088B, &HF44D890C, &HC83B008B, &H4589D88B, &HD94F0FEC, &HF05D8953, &H9D3C8D, &HE8000000, &H6090&, &H348DC703, &H85&, &HE4758900, &HFFEDAFE8, &H8B56FF, &HD08BD0FF, &H5589F633, &H7EDB85F8, &HC458B55, &HCB8BFA03, &H8B98148D, &H45890845, &HC458BFC, &H8BFC4529, &H83B0845, &H458B087F, &H10048BFC, &HC03302EB, &H89F85D8B, &H458BB304
    pvAppendBuffer &H7F083B0C, &HEB028B04, &H8BC03302, &H8946F05D, &H4EA8307, &H4904C783, &HCB7CF33B, &H8DF8558B, &H9D3C&, &HC38B0000, &H304E0C1, &H458950C2, &HDA048D08, &H48D5053, &HE8525017, &H44C0&, &H8BF44D8B, &H3411045, &H4D89EC4D, &H74C0850C, &H3B008B0C, &H8D067FC8, &H4D890148, &H8D348D0C, &H4&, &HFFED07E8, &H8B56FF, &H8B56D0FF, &H57006AF8, &H4AB7E8, &HC458B00, &HC483C933
    pvAppendBuffer &H8D07890C, &HC23B0151, &H758B257C, &H83DB0308, &HD33BFCC6, &H68B047F, &HC03302EB, &H489C085, &HCA450F97, &H4EE8342, &HE67E173B, &H8B0C458B, &HF891075, &H7B74F685, &H1BA&, &HC0570F00, &H45130F66, &H7CC23BE8, &H8DC68B6A, &H758B045F, &H89C72BEC, &H758BF475, &HFC7589E8, &H8910758B, &H173BF045, &H38B0A7F, &H8B084589, &H7EBF045, &H845C7, &H3B000000, &H8B057F16, &H2EB1834
    pvAppendBuffer &HC033F633, &H13087503, &HFC7503C0, &H45133389, &HFC4589F4, &HF445C7, &H85000000, &H3B0574F6, &HCA4F0FD1, &H4210758B, &H83F0458B, &H553B04C3, &H8BAF7E0C, &HF89F85D, &H1474DB85, &H85E44D8B, &H8B0D74C9, &HC6C3&, &H8301408D, &HF57501E9, &HFFEC1BE8, &H408B53FF, &H8BD0FF08, &H5B5E5FC7, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H85560C4D, &H8B3578C9, &H68B0875
    pvAppendBuffer &H3B02E0C1, &H8B297DC8, &HE28399C1, &HC1C20303, &HE18102F8, &H80000003, &H83490579, &H8B41FCC9, &HC1048644, &HE8D303E1, &H5EC0B60F, &H8C25D, &H5D5EC033, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H56530855, &H570C758B, &H1E8B3A8B, &H7501FF83, &H39C03308, &H440F0442, &H1FB83F8, &HC0330875, &HF044639, &HFB3BD844, &H4F0FC38B, &H74C085C7, &H8DD62B33, &H5589860C, &H1F0F08
    pvAppendBuffer &H47EC73B, &H3EBD233, &H3B0A148B, &H33047EC3, &H8B06EBF6, &H72D63B31, &H77D63B24, &H8558B14, &H8304E983, &HD87501E8, &HC0335E5F, &H8C25D5B, &HB85E5F00, &H1&, &H8C25D5B, &H835E5F00, &H5D5BFFC8, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HC0830C45, &HE2839903, &H57565303, &HC1021C8D, &H3C8D02FB, &H49D&, &HEAFEE800, &H8B57FFFF, &H57D1FF08, &H6AF08B, &H48AEE856
    pvAppendBuffer &HC4830000, &H831E890C, &H117C01FB, &H9D0C8D, &HC1000000, &H7E8D02E9, &HF3C03304, &HC5D8BAB, &H4774DB85, &HDD0C8D, &HF000000, &H441F&, &H8308458B, &H894B08E9, &H108A0C4D, &H8458940, &HC085C38B, &HC0830379, &H2F8C103, &H81863C8D, &H1FE1&, &H49057980, &H41E0C983, &HD3C2B60F, &H44709E0, &H850C4D8B, &H83C575DB, &H1076013E, &H3C83068B, &H8750086, &H83068948, &HF07701F8
    pvAppendBuffer &H5EC68B5F, &H8C25D5B, &HCCCCCC00, &HCCCCCCCC, &H8BEC8B55, &HC9850C4D, &H8B564278, &H68B0875, &H3B05E0C1, &H8B347DC8, &HE28399C1, &HC1C2031F, &HE18105F8, &H8000001F, &H83490579, &HBA41E0C9, &H1&, &H7D83E2D3, &H9740010, &H4865409, &HCC25D5E, &H21D2F700, &H5E048654, &HCC25D, &H56EC8B55, &HBA08758B, &H10&, &H90FF3357, &H20B9&, &H2BC68B00, &H85E8D3CA, &H8B0675C0
    pvAppendBuffer &H3E6D3CA, &H75FAD1FA, &H5FC78BE7, &H4C25D5E, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H53990845, &H561FE283, &H21C8D57, &H4305FBC1, &H49D348D, &HE8000000, &HFFFFE9B0, &HFF088B56, &HF88B56D1, &HE857006A, &H4760&, &H890CC483, &HFF016A1F, &HE8570875, &HFFFFFF30, &H5E5FC78B, &H4C25D5B, &HCCCCCC00, &HCCCCCCCC, &H83EC8B55, &H565320EC, &H3308758B, &H4D8957C9
    pvAppendBuffer &HCE0481EC, &H10000, &H83CE048B, &H4CE54, &H4CE5C8B, &H10D8AC0F, &H8910FBC1, &HF983E845, &HC715750F, &H1FC45, &HD08B0000, &HF045C7, &H89000000, &H22EBF85D, &H66C0570F, &HF445130F, &H89F8458B, &H458BF045, &H130F66F4, &H558BE045, &HFC4589E0, &H89E4458B, &HF983F845, &H1798D0F, &HC01B006A, &HAF0FD8F7, &HFC552BC7, &H348D256A, &HF8458BC6, &H50F0451B, &HF5D2E852, &H4D8BFFFF
    pvAppendBuffer &H13C103E8, &H1E883D3, &H100DA83, &HEC458B06, &H8B045611, &HA40F0875, &HE1C110CB, &HC60C2910, &H4D89CF8B, &HC65C19EC, &H10F98304, &HFF4F820F, &H5E5FFFFF, &H5DE58B5B, &HCC0004C2, &HCCCCCCCC, &H83EC8B55, &H558B10EC, &HF57560C, &HB60F0AB6, &HE1C10142, &HFC80B08, &HC10242B6, &HC80B08E1, &H342B60F, &HB08E1C1, &H42B60FC8, &HF04D8905, &H44AB60F, &HB08E1C1, &H42B60FC8, &H8E1C106
    pvAppendBuffer &HB60FC80B, &HE1C10742, &HFC80B08, &H890942B6, &HB60FF44D, &HE1C1084A, &HFC80B08, &HC10A42B6, &HC80B08E1, &HB42B60F, &HB08E1C1, &H42B60FC8, &HF84D890C, &HD4AB60F, &HB08E0C1, &H42B60FC8, &H8E1C10E, &HB60FC80B, &HE1C10F42, &H89C80B08, &H4D8BFC4D, &H8D398B08, &HC78B0471, &H304E0C1, &HF0458DF0, &H55E85056, &H83FFFFF6, &HC78310EE, &H8D2D74FF, &HE850F045, &H4474&, &H50F0458D
    pvAppendBuffer &H450BE8, &H458D5600, &H31E850F0, &H8DFFFFF6, &HE850F045, &H4418&, &H8310EE83, &HD37501EF, &H50F0458D, &H4447E8, &HF0458D00, &H44DEE850, &H8D560000, &HE850F045, &HFFFFF604, &H8B10758B, &HC28BF055, &HC1F44D8B, &H68818E8, &HE8C1C28B, &H1468810, &HE8C1C28B, &H2468808, &HE8C1C18B, &H3568818, &H8B044688, &H10E8C1C1, &H8B054688, &H8E8C1C1, &H88064688, &H4D8B074E, &HC1C18BF8
    pvAppendBuffer &H468818E8, &HC1C18B08, &H468810E8, &HC1C18B09, &H468808E8, &HB4E880A, &H8BFC4D8B, &H18E8C1C1, &H8B0C4688, &H10E8C1C1, &H8B0D4688, &H8E8C1C1, &H5F0E4688, &H5E0F4E88, &HC25DE58B, &HCCCC000C, &H83EC8B55, &H565310EC, &HC558B57, &HF085D8B, &HB60F0AB6, &HE1C10142, &H4738D08, &HB60FC80B, &HE1C10242, &HFC80B08, &HC10342B6, &HC80B08E1, &H542B60F, &HFF04D89, &HC1044AB6, &HC80B08E1
    pvAppendBuffer &H642B60F, &HB08E1C1, &H42B60FC8, &H8E1C107, &HB60FC80B, &H4D890942, &H4AB60FF4, &H8E1C108, &HB60FC80B, &HE1C10A42, &HFC80B08, &HC10B42B6, &HC80B08E1, &HC42B60F, &HFF84D89, &HC10D4AB6, &HC80B08E0, &HE42B60F, &HB08E1C1, &H42B60FC8, &H8E1C10F, &H458DC80B, &H895056F0, &HDDE8FC4D, &HBFFFFFF4, &H1&, &H3910C683, &H902E763B, &H50F0458D, &H6D87E8, &HF0458D00, &H6C1EE850
    pvAppendBuffer &H458D0000, &H85E850F0, &H56000044, &H50F0458D, &HFFF4ABE8, &HC68347FF, &H723B3B10, &HF0458DD3, &H6D5AE850, &H458D0000, &HF1E850F0, &H5600006B, &H50F0458D, &HFFF487E8, &H10758BFF, &H8BF0558B, &HF44D8BC2, &H8818E8C1, &HC1C28B06, &H468810E8, &HC1C28B01, &H468808E8, &HC1C18B02, &H568818E8, &H4468803, &HE8C1C18B, &H5468810, &HE8C1C18B, &H6468808, &H8B074E88, &HC18BF84D, &H8818E8C1
    pvAppendBuffer &HC18B0846, &H8810E8C1, &HC18B0946, &H8808E8C1, &H4E880A46, &HFC4D8B0B, &HE8C1C18B, &HC468818, &HE8C1C18B, &HD468810, &HE8C1C18B, &HE468808, &HF4E885F, &HE58B5B5E, &HCC25D, &HCCCCCCCC, &H56EC8B55, &H6808758B, &HF4&, &HE856006A, &H435C&, &H8310458B, &HF8830CC4, &H83357410, &H1A7418F8, &H7520F883, &H75FF503C, &HE06C70C, &H56000000, &HFFF3F7E8, &HC25D5EFF, &H186A000C
    pvAppendBuffer &HC70C75FF, &HC06&, &HE1E85600, &H5EFFFFF3, &HCC25D, &H75FF106A, &HA06C70C, &H56000000, &HFFF3CBE8, &HC25D5EFF, &HCCCC000C, &HCCCCCCCC, &H81EC8B55, &H100EC, &H51E85600, &HBEFFFFE5, &HB95940, &H4000EE81, &HF00300B9, &HFFE53FE8, &H2875FFFF, &HB957C0B9, &HF445C700, &H10&, &H812475FF, &HB94000E9, &HF8758900, &H4589C103, &H858DFC, &H50FFFFFF, &HFFFF43E8, &H875FFFF
    pvAppendBuffer &HFF00858D, &H106AFFFF, &H6A1475FF, &H2075FF0C, &HFF1C75FF, &H75FF1875, &HC75FF10, &HF4458D50, &HF3AE850, &H8B5E0000, &H24C25DE5, &HCCCCCC00, &H81EC8B55, &H100EC, &HD1E85600, &HBEFFFFE4, &HB95940, &H4000EE81, &HF00300B9, &HFFE4BFE8, &H2875FFFF, &HB957C0B9, &HF445C700, &H10&, &H812475FF, &HB94000E9, &HF8758900, &H4589C103, &H858DFC, &H50FFFFFF, &HFFFEC3E8, &HFF106AFF
    pvAppendBuffer &H858D0C75, &HFFFFFF00, &H6A0875FF, &H2075FF0C, &HFF1C75FF, &H75FF1875, &H1075FF14, &HF4458D50, &H107AE850, &H8B5E0000, &H24C25DE5, &HCCCCCC00, &H51EC8B55, &H185D8B53, &H4589C033, &H74DB85FC, &H10558B71, &H560C4D8B, &H11845C7, &H57000000, &HF28B398B, &HDE3BF72B, &H85F3420F, &HF1D75C0, &H561445B6, &H8458B50, &HE850C703, &H41C0&, &H830C4D8B, &H458B0CC4, &H10558BFC, &H975FF85
    pvAppendBuffer &H440FF23B, &H45891845, &H3E048DFC, &H1775C23B, &HFF0875FF, &H55FF2075, &HC4D8B1C, &HC710558B, &H1&, &H102EB00, &HFC458B31, &HA075DE2B, &H8B5B5E5F, &H1CC25DE5, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &H8B20758B, &HE883C6, &HE8836074, &HAC840F01, &H53000000, &H5701E883, &H7414458D, &H287D8B6D, &H57245D8B, &H50016A53, &HFF1075FF, &H75FF0C75, &HB6E808, &H4D8B0000, &H38535718
    pvAppendBuffer &H2F741C4D, &H8BFE468D, &H51501075, &HC75FF56, &HE80875FF, &HFFFFFF18, &H16A5357, &H501C458D, &HC75FF56, &HE80875FF, &H84&, &H5D5E5B5F, &H8D0024C2, &H758BFF46, &H56515010, &HFF0C75FF, &HE9E80875, &H5FFFFFFE, &HC25D5E5B, &H75FF0024, &H105D8B28, &H8B2475FF, &H758B0C7D, &H50016A08, &HE8565753, &H48&, &H8D2875FF, &H75FF1C45, &H50016A24, &HE8565753, &H34&, &H5D5E5B5F
    pvAppendBuffer &HFF0024C2, &H458A2875, &H2475FF1C, &H8D144530, &H16A1445, &H1075FF50, &HFF0C75FF, &HDE80875, &H5E000000, &H24C25D, &HCCCCCCCC, &HCCCCCCCC, &HFFEC8B55, &H458B2075, &HFF50501C, &H75FF1875, &H1075FF14, &HFF0C75FF, &H11E80875, &H5D000000, &HCC001CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H458B0C4D, &H5D8B5324, &H56118B14, &H5718758B, &H5974D285, &H5574F685, &H8B10458B
    pvAppendBuffer &H3BC22BFE, &HF8420FC6, &H4503C28B, &H50535708, &H3FEBE8, &HC458B00, &HF72BDF03, &H10CC483, &H107D8B38, &H458B3839, &HFF297524, &H85500875, &HFF0D75F6, &H4D8B2055, &H24458B0C, &H14EB3189, &H8B1C55FF, &H458B0C4D, &H1C724, &HEB000000, &H107D8B03, &H1972F73B, &HF73B5053, &H55FF0575, &HFF03EB20, &H458B1C55, &H3F72B24, &H73F73BDF, &H74F685E7, &HC458B2E, &HC78B088B, &HFE8BC12B
    pvAppendBuffer &H420FC63B, &H8458BF8, &H53C10357, &H3F6EE850, &H458B0000, &H83DF030C, &H38010CC4, &H7D8BF72B, &H5FD57510, &HC25D5B5E, &HCCCC0020, &HCCCCCCCC, &H8BEC8B55, &HEC831C4D, &H7D8B5708, &H74C98518, &H5D8B5376, &H3B83560C, &HFF117500, &H75FF0875, &H2055FF24, &H8B10458B, &H3891C4D, &HF18B038B, &H2B10558B, &H89C13BD0, &H420F1855, &H89C033F0, &HF685FC75, &H5D8B2F74, &H89DF2B14, &H9066F85D
    pvAppendBuffer &H8DFC758B, &HC8A3814, &H18558B13, &H8B085503, &HC32F85D, &H38148D02, &H3B0A8840, &H8BE172C6, &H4D8B0C5D, &H2B33291C, &H147501CE, &H4D89FE03, &H75C9851C, &H5F5B5E91, &HC25DE58B, &HCCCC0020, &HE8EC8B55, &HFFFFE148, &HB96620B9, &HE98100, &H300B940, &H84D8BC1, &H75FF5051, &H74418D14, &HFF1075FF, &H406A0C75, &H34418D50, &HFF3EE850, &HC25DFFFF, &HCCCC0010, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H4D8B6CEC, &H57565314, &H359B60F, &H241B60F, &H751B60F, &HC108E2C1, &HD80B08E3, &H141B60F, &HB08E3C1, &H1B60FD8, &HB08E3C1, &H41B60FD8, &H89D00B06, &HE2C1D85D, &H41B60F08, &HFD00B05, &HC10441B6, &HD00B08E2, &HA41B60F, &H89F45589, &HB60FD455, &HE2C10B51, &HFD00B08, &HC10941B6, &HD00B08E2, &H841B60F, &HB08E2C1, &H41B60FD0, &HF055890E, &HFD05589, &HC10F51B6
    pvAppendBuffer &HD00B08E2, &HD41B60F, &HB08E2C1, &H41B60FD0, &H84D8B0C, &HB08E2C1, &HF85589D0, &H241B60F, &HFCC5589, &HC10351B6, &HD00B08E2, &H141B60F, &HB08E2C1, &H1B60FD0, &HB08E2C1, &H41B60FD0, &HEC558906, &HFC85589, &HC10751B6, &HD00B08E2, &H541B60F, &HB08E2C1, &H41B60FD0, &H8E2C104, &HB60FD00B, &H55890A41, &HC45589E8, &HB51B60F, &HB08E2C1, &H8E2C1D0, &H941B60F, &HB60FD00B
    pvAppendBuffer &HE2C10841, &HFD00B08, &H890E41B6, &H5589E455, &H51B60FC0, &H8E2C10F, &HB60FD00B, &HE2C10D41, &HFD00B08, &H8B0C41B6, &HE2C10C4D, &H89D00B08, &HB60FE055, &H55890241, &H51B60FBC, &H8E2C103, &HB60FD00B, &HE2C10141, &HFD00B08, &HE2C101B6, &HFD00B08, &H890641B6, &H55890855, &H51B60FB8, &H8E2C107, &HB60FD00B, &HE2C10541, &HFD00B08, &HC10441B6, &HD00B08E2, &HA41B60F, &H89145589
    pvAppendBuffer &HB60FB455, &HE2C10B51, &HFD00B08, &HC10941B6, &HD00B08E2, &H841B60F, &HB08E2C1, &H41B60FD0, &HC55890E, &HFB05589, &HC10F51B6, &HD00B08E2, &HD41B60F, &HB08E2C1, &H41B60FD0, &H8E2C10C, &H5589D00B, &HAC5589FC, &HF10558B, &HF034AB6, &HC10242B6, &HC80B08E1, &H142B60F, &HB08E1C1, &H2B60FC8, &HB08E1C1, &HDC4D89C8, &HFA84D89, &HF0772B6, &HF0642B6, &HF0B7AB6, &HC10E4AB6
    pvAppendBuffer &HF00B08E6, &HF08E7C1, &HC10542B6, &HF00B08E6, &HA9845C7, &HF000000, &HC10442B6, &HF00B08E6, &HA42B60F, &H7589F80B, &H42B60FA4, &H8E7C109, &HB60FF80B, &HE7C10842, &HFF80B08, &HC10F42B6, &HC10B08E0, &HFA07D89, &HC10D4AB6, &HC10B08E0, &HC4AB60F, &HC1DC558B, &HC10B08E0, &H89EC4D8B, &H3EB9C45, &H3105D8B, &H84D8BD9, &H5D89D333, &H10C2C110, &H4D89CA03, &HEC4D3308, &H30CC1C1
    pvAppendBuffer &H89D333D9, &H5D8B105D, &H8C2C108, &H5589DA03, &HF4558BDC, &H33E85503, &H85D89F2, &HC6C1D933, &H144D8B10, &HC3C1CE03, &H144D8907, &HC1E84D33, &HD1030CC1, &H5589F233, &H14558BF4, &H308C6C1, &HEC7589D6, &H3F0758B, &HFE33E475, &H33145589, &H10C7C1D1, &H30C4D8B, &H7C2C1CF, &H330C4D89, &HC1C1E44D, &H33F1030C, &HF07589FE, &HC10C758B, &HF70308C7, &H8B947D89, &H7D03F87D, &H89C733E0
    pvAppendBuffer &HF1330C75, &H8B10C0C1, &HC803FC4D, &H8907C6C1, &H4D33FC4D, &HCC1C1E0, &HC733F903, &H8BF87D89, &HC0C1FC7D, &H89F80308, &HF933FC7D, &H3104D8B, &H7C7C1CA, &H4D89C133, &HC4D8B10, &H310C0C1, &HC4D89C8, &H558BCA33, &HCC1C110, &HC233D103, &H8B105589, &HC0C10C55, &H89D00308, &HD1330C55, &H3F44D8B, &H7C2C1CE, &H89F44D89, &H558BE855, &H8BD133DC, &HC2C1FC4D, &H89CA0310, &HCE33FC4D
    pvAppendBuffer &HC1F4758B, &HF1030CC1, &H7589D633, &HFC758BF4, &H308C2C1, &HFC7589F2, &H4D8BF133, &HC1CF03F0, &H4D8907C6, &HE47589F0, &H33EC758B, &H84D8BF1, &H310C6C1, &H84D89CE, &H7D8BCF33, &HCC1C1F0, &HF733F903, &H8BF07D89, &HC6C1087D, &H89FE0308, &HF933087D, &H3F84D8B, &H7C7C1CB, &H8BE07D89, &HF933947D, &H8BF84D89, &HC7C1144D, &H89CF0310, &HCB33144D, &HC1F85D8B, &HD9030CC1, &H5D89FB33
    pvAppendBuffer &H8C7C1F8, &H8B147D01, &HD933145D, &H5D89CB8B, &H7C1C1EC, &H1986D83, &H89F85D8B, &H850FEC4D, &HFFFFFE40, &H19C4501, &H4D8BCC5D, &H104D03D8, &H8BA85501, &H4D891855, &HD85D8BD8, &H4D8BC38B, &HF44D03D4, &H4D891A88, &HD04D8BD4, &HC1F04D03, &H428808E8, &H89C38B01, &H4D8BD04D, &HC84D01EC, &H3C44D8B, &HE8C1E84D, &H2428810, &H8818EBC1, &H5D8B035A, &H88C38BD4, &HE8C1045A, &H5428808
    pvAppendBuffer &H4D89C38B, &HC04D8BC4, &HC1E44D03, &H428810E8, &HC04D8906, &H3BC4D8B, &H7501E04D, &HA07D01A4, &H8818EBC1, &H5D8B075A, &H88C38BD0, &H4D89085A, &HB84D8BBC, &HC1084D03, &H428808E8, &H89C38B09, &H4D8BB84D, &H144D03B4, &H8810E8C1, &HEBC10A42, &HB5A8818, &H8BCC5D8B, &HB44D89C3, &H3B04D8B, &H5A880C4D, &H8E8C10C, &H8B0D4288, &HB04D89C3, &H3AC4D8B, &HE8C1FC4D, &HE428810, &H8818EBC1
    pvAppendBuffer &H5D8B0F5A, &H89C38BC8, &H5A88AC4D, &H8E8C110, &H8B114288, &H10E8C1C3, &HC1124288, &H5A8818EB, &HC45D8B13, &H5A88C38B, &H8E8C114, &H8B154288, &H10E8C1C3, &HC1164288, &H5A8818EB, &HC05D8B17, &H5A88C38B, &H8E8C118, &H8B194288, &H10E8C1C3, &HC11A4288, &H5A8818EB, &HBC5D8B1B, &H5A88C38B, &H8E8C11C, &H8B1D4288, &H10E8C1C3, &HC11E4288, &H5A8818EB, &HB85D8B1F, &H5A88C38B, &H8E8C120
    pvAppendBuffer &H8B214288, &H10E8C1C3, &HC1224288, &H5A8818EB, &HB45D8B23, &H5A88C38B, &H8E8C124, &H8B254288, &H10E8C1C3, &HC1264288, &H5A8818EB, &HB05D8B27, &H5A88C38B, &H8E8C128, &H8B294288, &H10E8C1C3, &HC12A4288, &H5A8818EB, &H88D98B2B, &HC38B2C5A, &H8808E8C1, &HC38B2D42, &H8810E8C1, &HEBC12E42, &H2F5A8818, &H8BA85D8B, &H305A88C3, &H8808E8C1, &H4A8D3142, &HC1C38B3C, &HE8C118EB, &H32428810
    pvAppendBuffer &H8B335A88, &HC38BA45D, &HC1345A88, &H428808E8, &HC1C38B35, &H428810E8, &H18EBC136, &H8B375A88, &HC38BA05D, &HC1385A88, &H428808E8, &HC1C38B39, &H428810E8, &H18EBC13A, &H8B3B5A88, &HC28B9C55, &H8808E8C1, &H1418811, &HC15FC28B, &HEAC110E8, &H41885E18, &H3518802, &H5DE58B5B, &HCC0014C2, &H56EC8B55, &H8B1075FF, &H75FF0875, &HEDE8560C, &H6A000059, &H1475FF10, &H5020468D, &H37FFE8
    pvAppendBuffer &H18458B00, &HC70CC483, &H7446&, &H46890000, &HC25D5E78, &HCCCC0014, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H5708758B, &HFF0C75FF, &H7E8D3076, &H468D5720, &HE8565010, &HFFFFF944, &H3378568B, &H10780C0, &H3B400B75, &H800674C2, &H74013804, &H5D5E5FF5, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D10EC, &HFF106AF0, &HE8502075, &H378C&, &H8D0CC483, &H6A50F045, &H2475FF00
    pvAppendBuffer &HFF1C75FF, &H75FF1875, &H1075FF14, &HFF0C75FF, &HD9E80875, &H8B000052, &H20C25DE5, &HCCCCCC00, &HFFEC8B55, &H16A2475, &HFF2075FF, &H75FF1C75, &H1475FF18, &HFF1075FF, &H75FF0C75, &H52AEE808, &HC25D0000, &HCCCC0020, &HCCCCCCCC, &HCCCCCCCC, &HE8EC8B55, &HFFFFD9B8, &HB979D0B9, &HE98100, &H300B940, &H84D8BC1, &H75FF5051, &HFF018B14, &H75FF1075, &H8D30FF0C, &H8D502841, &HE8501841
    pvAppendBuffer &HFFFFF7AC, &H10C25D, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H458B084D, &H2C41890C, &H8910458B, &HC25D3041, &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H6A08758B, &H56006A34, &H36EFE8, &HC4D8B00, &H2C46C7, &H8B000000, &H30468901, &H8910458B, &H468D0446, &HC70E8908, &H2846&, &H31FF0000, &H501475FF, &H3693E8, &H18C48300, &H10C25D5E, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H81EC8B55, &H420EC, &H57565300, &H858D706A, &HFFFFFD70, &HFD6085C7, &HDB41FFFF, &H6A0000, &H6485C750, &HFFFFFD, &HC7000000, &HFFFD6885, &H1FF&, &H6C85C700, &HFFFFFD, &HE8000000, &H366C&, &H8D0C758B, &HFFFF6085, &H561F6AFF, &H362AE850, &H468A0000, &H18C4831F, &HFF60A580, &H24F8FFFF, &H88400C3F, &HFFFF7F85, &HE0858DFF, &HFFFFFFFB, &HE8501075, &H6044&, &H8DC0570F
    pvAppendBuffer &HFFFE60B5, &H130F66FF, &HFFFE6085, &H68BD8DFF, &HB9FFFFFE, &H1E&, &H45130F66, &HB9A5F380, &H1E&, &H85130F66, &HFFFFFEE0, &HC780758D, &HFFFE6085, &H1FF&, &H887D8D00, &HFE6485C7, &HFFFF&, &HA5F30000, &H1EB9&, &H8045C700, &H1&, &HFEE0B58D, &H45C7FFFF, &H84&, &HE8BD8D00, &HBBFFFFFE, &HFE&, &H20B9A5F3, &H8D000000, &HFFFBE0B5, &HE0BD8DFF, &HF3FFFFFD
    pvAppendBuffer &HFC38BA5, &HF8C1CBB6, &H7E18303, &H5B4B60F, &HFFFFFF60, &HFDE0858D, &HEED3FFFF, &H5601E683, &H80458D50, &H5616E850, &H8D560000, &HFFFE6085, &H858D50FF, &HFFFFFEE0, &H5602E850, &H858D0000, &HFFFFFEE0, &H80458D50, &HE0858D50, &H50FFFFFC, &HFFE4EBE8, &HE0858DFF, &H50FFFFFE, &H5080458D, &H5DFAE850, &H858D0000, &HFFFFFE60, &HE0858D50, &H50FFFFFD, &HFEE0858D, &HE850FFFF, &HFFFFE4C0
    pvAppendBuffer &HFE60858D, &H8D50FFFF, &HFFFDE085, &HE85050FF, &H5DCC&, &HFCE0858D, &H8D50FFFF, &HFFFE6085, &H99E850FF, &H8D00005D, &H8D508045, &HFFFC6085, &H89E850FF, &H8D00005D, &H8D508045, &HFFFEE085, &H458D50FF, &H95E85080, &H8D000047, &HFFFCE085, &H858D50FF, &HFFFFFDE0, &HE0858D50, &H50FFFFFE, &H477BE8, &HE0858D00, &H50FFFFFE, &H5080458D, &HFCE0858D, &HE850FFFF, &HFFFFE444, &HFEE0858D
    pvAppendBuffer &H8D50FFFF, &H50508045, &H5D53E8, &H80458D00, &HE0858D50, &H50FFFFFD, &H5D23E8, &H60858D00, &H50FFFFFC, &HFE60858D, &H8D50FFFF, &HFFFEE085, &H29E850FF, &H8D00005D, &HFFFD6085, &H858D50FF, &HFFFFFEE0, &H80458D50, &H4712E850, &H858D0000, &HFFFFFE60, &H80458D50, &HE1E85050, &H8DFFFFE3, &H8D508045, &HFFFEE085, &HE85050FF, &H46F0&, &HFC60858D, &H8D50FFFF, &HFFFE6085, &H458D50FF
    pvAppendBuffer &HD9E85080, &H8D000046, &HFFFBE085, &H858D50FF, &HFFFFFDE0, &H60858D50, &H50FFFFFE, &H46BFE8, &HE0858D00, &H50FFFFFC, &HFDE0858D, &HE850FFFF, &H5C8C&, &HE0858D56, &H50FFFFFD, &H5080458D, &H547BE8, &H858D5600, &HFFFFFE60, &HE0858D50, &H50FFFFFE, &H5467E8, &H1EB8300, &HFE1F890F, &H858DFFFF, &HFFFFFEE0, &H91E85050, &H8D000031, &HFFFEE085, &H458D50FF, &HE8505080, &H4660&
    pvAppendBuffer &H5080458D, &HE80875FF, &H49A4&, &H8B5B5E5F, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D20EC, &HE045C6E0, &H75FF5009, &HC0570F0C, &HF945C7, &HFF000000, &H110F0875, &HC766E145, &HFD45&, &H45D60F66, &HFF45C6F1, &HFCAAE800, &HE58BFFFF, &H8C25D, &HCCCCCCCC, &H81EC8B55, &H114EC, &H5D8B5300, &HF0458D08, &H7D8B5756, &HC0570F0C, &H438B5050, &H45C65704
    pvAppendBuffer &HF6600F0, &HC7F145D6, &HF945&, &HC7660000, &HFD45&, &HFF45C6, &H758BD0FF, &HCFE8324, &HFF562075, &H458D2075, &HD1E850D0, &H83000032, &HC7660CC4, &HDD45&, &HDC45C6, &H1DF45C6, &H458D30EB, &H858D50F0, &HFFFFFEEC, &H28EEE850, &HFF560000, &H858D2075, &HFFFFFEEC, &H270EE850, &H458D0000, &H858D50D0, &HFFFFFEEC, &H27BEE850, &H458D0000, &H858D50F0, &HFFFFFF3C, &H28BEE850
    pvAppendBuffer &H75FF0000, &H3C858D1C, &HFFFFFFFF, &HE8501875, &H26BC&, &HC6D0458D, &H5000E045, &H458D5357, &HE945C78C, &H0&, &H66C0570F, &HED45C7, &HF665000, &HC6E145D6, &HE800EF45, &HFFFFFB70, &HC6A046A, &H508C458D, &HFFFB43E8, &H8D106AFF, &H5050E045, &H508C458D, &HFFFAF3E8, &H1475FFFF, &HFF3C858D, &H75FFFFFF, &H81E85010, &H8D000026, &H8D50C045, &HFFFF3C85, &H31E850FF, &H8B000027
    pvAppendBuffer &H458D2C75, &H8D5056E0, &H5050C045, &H7F3FE8, &H8DD23200, &H1BBC045, &H85000000, &H8B1A74F6, &HC88B287D, &HC8AF92B, &H1408D07, &HAFF4832, &H75F32BD1, &H75D284F1, &H1475FF14, &HFF8C458D, &H75FF3075, &H85E85010, &H33FFFFFA, &HC0570FDB, &HF045110F, &HFF0458A, &H8AD04511, &H110FD045, &H458AE045, &H45110FE0, &HC0458AC0, &H858D506A, &HFFFFFF3C, &HE850006A, &H31B4&, &HFF3C8D8A
    pvAppendBuffer &H458DFFFF, &H6A346A8C, &HA1E85000, &H8A000031, &HC4838C4D, &H5FC38B18, &HE58B5B5E, &H2CC25D, &H81EC8B55, &H114EC, &H5D8B5300, &HF0458D08, &H7D8B5756, &HC0570F0C, &H438B5050, &H45C65704, &HF6600F0, &HC7F145D6, &HF945&, &HC7660000, &HFD45&, &HFF45C6, &H758BD0FF, &HCFE8324, &HFF562075, &H458D2075, &H11E850D0, &H83000031, &HC7660CC4, &HDD45&, &HDC45C6, &H1DF45C6
    pvAppendBuffer &H458D30EB, &H858D50F0, &HFFFFFEEC, &H272EE850, &HFF560000, &H858D2075, &HFFFFFEEC, &H254EE850, &H458D0000, &H858D50D0, &HFFFFFEEC, &H25FEE850, &H458D0000, &H858D50F0, &HFFFFFF3C, &H26FEE850, &H75FF0000, &H3C858D1C, &HFFFFFFFF, &HE8501875, &H24FC&, &HC6D0458D, &H5000E045, &H458D5357, &HE945C78C, &H0&, &H66C0570F, &HED45C7, &HF665000, &HC6E145D6, &HE800EF45, &HFFFFF9B0
    pvAppendBuffer &HC6A046A, &H508C458D, &HFFF983E8, &H8D106AFF, &H5050E045, &H508C458D, &HFFF933E8, &H147D8BFF, &H8B8C458D, &H56572875, &H501075FF, &HFFF91FE8, &H8D5657FF, &HFFFF3C85, &HB1E850FF, &H8D000024, &H45C6C045, &H8D5000C0, &HFFFF3C85, &HC945C7FF, &H0&, &H66C0570F, &HCD45C7, &HF665000, &HC6C145D6, &HE800CF45, &H2544&, &H8D3075FF, &H8D50E045, &HFF50C045, &H51E82C75, &HF00007D
    pvAppendBuffer &H110FC057, &H458AF045, &H45110FF0, &HD0458AD0, &HE045110F, &HFE0458A, &H8AC04511, &H506AC045, &HFF3C858D, &H6AFFFF, &H3002E850, &H858A0000, &HFFFFFF3C, &H458D346A, &H50006A8C, &H2FEFE8, &H8C458A00, &H5F18C483, &HE58B5B5E, &H2CC25D, &H8BEC8B55, &H4D8B0C55, &H758B5610, &H33068B08, &H8B018902, &H42330446, &H4418904, &H3308468B, &H41890842, &HC468B08, &H890C4233, &H5D5E0C41
    pvAppendBuffer &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H51EC8B55, &HC5D8B53, &H7D8B5756, &H45C76608, &H8BE100FC, &HD1C18B0F, &H1E183E8, &H578B0389, &HD1C28B04, &H1E283E8, &HB1FE1C1, &H1FE2C1C8, &H8B044B89, &HC68B0877, &HE683E8D1, &HC1D00B01, &H53891FE6, &HC4F8B08, &HE8D1C18B, &HB01E183, &H73895FF0, &H44B60F0C, &HE0C1FC0D, &H5E033118, &H5DE58B5B, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H8BEC8B55, &H8B560C55, &HB60F0875, &H46B60F0E, &H8E1C101, &HB60FC80B, &HE1C10246, &HFC80B08, &HC10346B6, &HC80B08E1, &HB60F0A89, &HB60F044E, &HE1C10546, &HFC80B08, &HC10646B6, &HC80B08E1, &H746B60F, &HB08E1C1, &H44A89C8, &H84EB60F, &H946B60F, &HB08E1C1, &H46B60FC8, &H8E1C10A, &HB60FC80B, &HE1C10B46, &H89C80B08, &HB60F084A, &HB60F0C4E, &HE1C10D46, &HFC80B08, &HC10E46B6
    pvAppendBuffer &HC80B08E1, &HF46B60F, &HB08E1C1, &HC4A89C8, &H8C25D5E, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H575620EC, &H458D106A, &H50006AE0, &H2E7BE8, &HFF106A00, &H458D0C75, &H3DE850F0, &H8B00002E, &HC483087D, &H4D100F18, &H90F633E0, &H1FB9C68B, &H83000000, &HC82B1FE0, &HF8C1C68B, &H87048B05, &H1A8E8D3, &H100F0C74, &HF66F045, &H110FC8EF, &H458DE04D, &HE85050F0, &HFFFFFE90
    pvAppendBuffer &H80FE8146, &H7C000000, &H8D106AC7, &HFF50E045, &HE9E81075, &H8300002D, &H5E5F0CC4, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H570C758B, &H8B087D8B, &HC1C28B17, &H68818E8, &HE8C1C28B, &H1468810, &HE8C1C28B, &H2468808, &H8B035688, &HC18B044F, &H8818E8C1, &HC18B0446, &H8810E8C1, &HC18B0546, &H8808E8C1, &H4E880646, &H84F8B07, &HE8C1C18B, &H8468818
    pvAppendBuffer &HE8C1C18B, &H9468810, &HE8C1C18B, &HA468808, &H8B0B4E88, &HC18B0C4F, &H8818E8C1, &HC18B0C46, &H8810E8C1, &HC18B0D46, &H8808E8C1, &H885F0E46, &H5D5E0F4E, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H8B5644EC, &HBE830875, &HA8&, &H56067400, &H4647E8, &HFC93300, &H441F&, &HE84B60F, &H88&, &HBC8D4489, &H10F98341, &HC756EE72, &HFC45&, &H21E80000, &H8D000045
    pvAppendBuffer &H5650BC45, &H44B7E8, &HC558B00, &H9066C933, &H888E048A, &H83411104, &HF47210F9, &HAC68&, &H56006A00, &H2D07E8, &H83068A00, &H8B5E0CC4, &H8C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H6808758B, &HAC&, &HE856006A, &H2CDC&, &H6A0C4D8B, &H1075FF10, &H8901B60F, &HB60F4446, &H46890141, &H41B60F48, &H4C468902, &H341B60F, &H890FE083, &HB60F5046, &HFC250441
    pvAppendBuffer &H89000000, &HB60F5446, &H46890541, &H41B60F58, &H5C468906, &H741B60F, &H890FE083, &HB60F6046, &HFC250841, &H89000000, &HB60F6446, &H46890941, &H41B60F68, &H6C46890A, &HB41B60F, &H890FE083, &HB60F7046, &HFC250C41, &H89000000, &HB60F7446, &H46890D41, &H41B60F78, &H7C46890E, &HF41B60F, &HC70FE083, &H8486&, &H0&, &H80868900, &H8D000000, &H8886&, &H1E85000, &H8300002C
    pvAppendBuffer &H5D5E18C4, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HE8EC8B55, &HFFFFCE78, &HB9B930B9, &HE98100, &H300B940, &H84D8BC1, &H75FF5051, &HA8818D10, &HFF000000, &H106A0C75, &H98818D50, &H50000000, &HFFEB6BE8, &HCC25DFF, &HCCCCCC00, &HCCCCCCCC, &H83EC8B55, &H565318EC, &HCE32E857, &H75FFFFFF, &HC060BE08, &H40B900B9, &H81000000, &HB94000EE, &H8BF00300, &H8D560845, &H408B6478, &H3E1F760
    pvAppendBuffer &H83D88B07, &HC08300D2, &H3FE08308, &H6A51C82B, &H68006A00, &H80&, &H8B57406A, &HA40F087D, &H558903DA, &H20478DFC, &H5003E3C1, &HE8F85589, &HFFFFEA0C, &H8BFC558B, &H88C28BCB, &HE8C1EF5D, &HE8458818, &HE8C1C28B, &HE9458810, &HE8C1C28B, &HEA458808, &H88F8458A, &HC28BEB45, &H18C1AC0F, &HE8C1086A, &HEC4D8818, &HCB8BC28B, &H10C1AC0F, &H8B10E8C1, &HED4D88C3, &H8D0AC0F, &H8DEE4588
    pvAppendBuffer &HC150E845, &HE85708EA, &H164&, &HC28B178B, &HC10C758B, &H68818E8, &HE8C1C28B, &H1468810, &HE8C1C28B, &H2468808, &H8B035688, &HC18B044F, &H8818E8C1, &HC18B0446, &H8810E8C1, &HC18B0546, &H8808E8C1, &H4E880646, &H84F8B07, &HE8C1C18B, &H8468818, &HE8C1C18B, &H9468810, &HE8C1C18B, &HA468808, &H8B0B4E88, &HC18B0C4F, &H8818E8C1, &HC18B0C46, &H8810E8C1, &HC18B0D46, &H8808E8C1
    pvAppendBuffer &H4E880E46, &H104F8B0F, &HE8C1C18B, &H10468818, &HE8C1C18B, &H11468810, &HE8C1C18B, &H12468808, &H8B134E88, &HC18B144F, &H8818E8C1, &HC18B1446, &H8810E8C1, &HC18B1546, &H8808E8C1, &H4E881646, &H184F8B17, &HE8C1C18B, &H18468818, &HE8C1C18B, &H19468810, &HE8C1C18B, &H1A468808, &H8B1B4E88, &HC18B1C4F, &H8818E8C1, &HC18B1C46, &H8810E8C1, &HC18B1D46, &HE8C1686A, &H88006A08, &H88571E46
    pvAppendBuffer &H29E81F4E, &H8300002A, &H5E5F0CC4, &H5DE58B5B, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H6A08758B, &H56006A68, &H29FFE8, &HCC48300, &HE66706C7, &H46C76A09, &H67AE8504, &H846C7BB, &H3C6EF372, &H3A0C46C7, &HC7A54FF5, &H527F1046, &H46C7510E, &H5688C14, &H1846C79B, &H1F83D9AB, &H191C46C7, &H5E5BE0CD, &H4C25D, &HE8EC8B55, &HFFFFCC18, &HB9C060B9, &HE98100
    pvAppendBuffer &H300B940, &H84D8BC1, &H75FF5051, &H64418D10, &H6A0C75FF, &H418D5040, &H11E85020, &H5DFFFFE9, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D40EC, &H75FF50C0, &HBEE808, &H306A0000, &H50C0458D, &HE80C75FF, &H2930&, &H8B0CC483, &H8C25DE5, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &H6808758B, &HC8&, &HE856006A, &H293C&, &HC70CC483, &H59ED806, &H446C7C1
    pvAppendBuffer &HCBBB9D5D, &H70846C7, &HC7367CD5, &H292A0C46, &H46C7629A, &H70DD1710, &H1446C730, &H9159015A, &H391846C7, &HC7F70E59, &HECD81C46, &H46C7152F, &HC00B3120, &H2446C7FF, &H67332667, &H112846C7, &HC7685815, &H4A872C46, &H46C78EB4, &HF98FA730, &H3446C764, &HDB0C2E0D, &HA43846C7, &HC7BEFA4F, &H481D3C46, &H5D5E47B5, &HCC0004C2, &HCCCCCCCC, &H41BE9, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H458B1CEC, &H988D5308, &HC4&, &HC0808B56, &H57000000, &H80BF&, &H8BE7F700, &H8B3303F0, &HD283C6, &H3C2A40F, &H8903E0C1, &H4589FC55, &HF45589F8, &HFFCAD3E8, &H875FFFF, &HB9C230B9, &HE98100, &H300B940, &H468D50C1, &H8758B10, &H2B7FE083, &H6A57F8, &H8068006A, &H68000000, &H80&, &H40468D53, &HE6CEE850, &H86AFFFF, &HC7E4458D, &HE445&, &H56500000
    pvAppendBuffer &HE845C7, &HE8000000, &H384&, &H8BFC5D8B, &HF8558BC3, &HE8C1CA8B, &HE4458818, &HE8C1C38B, &HE5458810, &HE8C1C38B, &HE6458808, &H88F4458A, &HC38BE745, &H18C1AC0F, &HE8C1086A, &HE84D8818, &HCA8BC38B, &HFEB5588, &HC110C1AC, &HC28B10E8, &HFE94D88, &H8808D8AC, &H458DEA45, &HC15650E4, &H29E808EB, &H8B000003, &HC38B045E, &H4D890E8B, &H18E8C1FC, &H880C7D8B, &HC1C38B07, &H478810E8
    pvAppendBuffer &HC1C38B01, &H478808E8, &HFC38B02, &H8818C1AC, &HE8C1035F, &H44F8818, &H4D8BC38B, &HC1AC0FFC, &H10E8C110, &H8B054F88, &HC18BFC4D, &H8D8AC0F, &H8B064788, &H74F88C6, &H8B08EBC1, &HCB8B0858, &H8B0C508B, &H18E8C1C2, &H8B084788, &H10E8C1C2, &H8B094788, &H8E8C1C2, &H8B0A4788, &HC1AC0FC2, &HB578818, &H8818E8C1, &HC28B0C4F, &HAC0FCB8B, &HE8C110C1, &HD4F8810, &HAC0FC38B, &H478808D0
    pvAppendBuffer &H88C68B0E, &HEAC10F5F, &H10588B08, &H508BCB8B, &HC1C28B14, &H478818E8, &HC1C28B10, &H478810E8, &HC1C28B11, &H478808E8, &HFC28B12, &H8818C1AC, &HE8C11357, &H144F8818, &HCB8BC28B, &H10C1AC0F, &H8B10E8C1, &H154F88C3, &H8D0AC0F, &H8B164788, &H8EAC1C6, &H8B175F88, &HCB8B1858, &H8B1C508B, &H18E8C1C2, &H8B184788, &H10E8C1C2, &H8B194788, &H8E8C1C2, &H8B1A4788, &HC1AC0FC2, &H1B578818
    pvAppendBuffer &H8818E8C1, &HC28B1C4F, &HAC0FCB8B, &HE8C110C1, &H1D4F8810, &HAC0FC38B, &H478808D0, &H88C68B1E, &HEAC11F5F, &H20588B08, &H508BCB8B, &HC1C28B24, &H478818E8, &HC1C28B20, &H478810E8, &HC1C28B21, &H478808E8, &HFC28B22, &H8818C1AC, &HE8C12357, &H244F8818, &HCB8BC28B, &H10C1AC0F, &H8810E8C1, &HC38B254F, &H8D0AC0F, &H8B264788, &H275F88C6, &H8B08EAC1, &HCB8B2858, &H8B2C508B, &H18E8C1C2
    pvAppendBuffer &H8B284788, &H10E8C1C2, &H8B294788, &H8E8C1C2, &H8B2A4788, &HC1AC0FC2, &H2B578818, &H8818E8C1, &HC28B2C4F, &HAC0FCB8B, &HE8C110C1, &H88C38B10, &HAC0F2D4F, &HEAC108D0, &H2E478808, &H5F88C68B, &H38778D2F, &HC868&, &H8B006A00, &HCB8B3058, &H8B34508B, &H18E8C1C2, &H8B304788, &H10E8C1C2, &H8B314788, &H8E8C1C2, &H8B324788, &HC1AC0FC2, &H33578818, &H8818E8C1, &HC28B344F, &HAC0FCB8B
    pvAppendBuffer &HE8C110C1, &H354F8810, &HAC0FC38B, &H478808D0, &H375F8836, &HC1087D8B, &H8B5708EA, &HC28B3C57, &H8B385F8B, &H18E8C1CB, &HC28B0688, &H8810E8C1, &HC28B0146, &H8808E8C1, &HC28B0246, &H18C1AC0F, &HC1035688, &H4E8818E8, &H8BC28B04, &HC1AC0FCB, &H10E8C110, &H4E88C38B, &HD0AC0F05, &H6468808, &H8808EAC1, &H45E8075E, &H83000025, &H5E5F0CC4, &H5DE58B5B, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H56EC8B55, &H6808758B, &HC8&, &HE856006A, &H251C&, &HC70CC483, &HBCC90806, &H446C7F3, &H6A09E667, &H3B0846C7, &HC784CAA7, &HAE850C46, &H46C7BB67, &H94F82B10, &H1446C7FE, &H3C6EF372, &HF11846C7, &HC75F1D36, &HF53A1C46, &H46C7A54F, &HE682D120, &H2446C7AD, &H510E527F, &H1F2846C7, &HC72B3E6C, &H688C2C46, &H46C79B05, &H41BD6B30, &H3446C7FB, &H1F83D9AB, &H793846C7, &HC7137E21
    pvAppendBuffer &HCD193C46, &H5D5E5BE0, &HCC0004C2, &HCCCCCCCC, &HE8EC8B55, &HFFFFC6F8, &HB9C230B9, &HE98100, &H300B940, &H84D8BC1, &H75FF5051, &HC4818D10, &HFF000000, &H80680C75, &H50000000, &H5040418D, &HFFE3EBE8, &HCC25DFF, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &HFF08758B, &HE8B0C75, &H5008468D, &H8B0476FF, &HD0FF0441, &H8B2C568B, &HD6033046, &H44805E48, &H75010802, &H801F0F13, &H0&
    pvAppendBuffer &H874C085, &H2448048, &HF4740108, &H8C25D, &H53EC8B55, &H560C5D8B, &H87D8B57, &H2843B60F, &H8BC88B99, &HCEA40FF2, &H43B60F08, &H8E1C129, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C12A, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C12B, &HBC80B99, &H43B60FF2, &HCEA40F2C, &HE1C19908, &HBF20B08, &H43B60FC8, &HCEA40F2D, &HE1C19908, &HBF20B08, &H43B60FC8, &HCEA40F2E, &HE1C19908
    pvAppendBuffer &HBF20B08, &H43B60FC8, &HCEA40F2F, &HE1C19908, &HBF20B08, &H47789C8, &HB60F0F89, &H8B992043, &HFF28BC8, &HF2143B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF2243B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF2343B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF2443B6, &H9908CEA4, &HB08E1C1, &HFF20BC8, &HF08CEA4, &HC12543B6, &HB9908E1, &HFF20BC8, &HF08CEA4, &HC12643B6, &HB9908E1
    pvAppendBuffer &HFF20BC8, &HF08CEA4, &HC12743B6, &HB9908E1, &H89F20BC8, &H7789084F, &H43B60F0C, &HC88B9918, &HA40FF28B, &HB60F08CE, &HE1C11943, &HC80B9908, &HB60FF20B, &HA40F1A43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F1B43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F1C43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F1D43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F1E43, &HC19908CE, &HF20B08E1
    pvAppendBuffer &HB60FC80B, &HA40F1F43, &HC19908CE, &HF20B08E1, &H7789C80B, &H104F8914, &H1043B60F, &H8BC88B99, &H43B60FF2, &HCEA40F11, &HE1C19908, &HBF20B08, &H43B60FC8, &HCEA40F12, &H8E1C108, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C113, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C114, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C115, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C116, &HBC80B99
    pvAppendBuffer &H43B60FF2, &HCEA40F17, &HE1C19908, &HBF20B08, &H1C7789C8, &HF184F89, &H990843B6, &HF28BC88B, &H943B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &HA43B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &HB43B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &HC43B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &HD43B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &HE43B60F, &H8CEA40F, &H8E1C199, &HC80BF20B
    pvAppendBuffer &HF43B60F, &H8CEA40F, &H8E1C199, &HF20BC80B, &H89204F89, &HB60F2477, &HC88B9903, &HB60FF28B, &HA40F0143, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0243, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0343, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0443, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0543, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0643, &HC19908CE, &HF20B08E1, &HB60FC80B
    pvAppendBuffer &HA40F0743, &HC19908CE, &HC80B08E1, &H7789F20B, &H284F892C, &H5D5B5E5F, &HCC0008C2, &HCCCCCCCC, &H53EC8B55, &H560C5D8B, &H87D8B57, &H1843B60F, &H8BC88B99, &HCEA40FF2, &H43B60F08, &H8E1C119, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C11A, &HBC80B99, &HCEA40FF2, &H43B60F08, &H8E1C11B, &HBC80B99, &H43B60FF2, &HCEA40F1C, &HE1C19908, &HBF20B08, &H43B60FC8, &HCEA40F1D, &HE1C19908
    pvAppendBuffer &HBF20B08, &H43B60FC8, &HCEA40F1E, &HE1C19908, &HBF20B08, &H43B60FC8, &HCEA40F1F, &HE1C19908, &HBF20B08, &H47789C8, &HB60F0F89, &H8B991043, &HFF28BC8, &HF1143B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF1243B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF1343B6, &H9908CEA4, &HB08E1C1, &HFC80BF2, &HF1443B6, &H9908CEA4, &HB08E1C1, &HFF20BC8, &HF08CEA4, &HC11543B6, &HB9908E1
    pvAppendBuffer &HFF20BC8, &HF08CEA4, &HC11643B6, &HB9908E1, &HFF20BC8, &HF08CEA4, &HC11743B6, &HB9908E1, &H89F20BC8, &H7789084F, &H43B60F0C, &HC88B9908, &HA40FF28B, &HB60F08CE, &HE1C10943, &HC80B9908, &HB60FF20B, &HA40F0A43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0B43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0C43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0D43, &HC19908CE, &HF20B08E1
    pvAppendBuffer &HB60FC80B, &HA40F0E43, &HC19908CE, &HF20B08E1, &HB60FC80B, &HA40F0F43, &HC19908CE, &HF20B08E1, &H7789C80B, &H104F8914, &H9903B60F, &HF28BC88B, &H143B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H243B60F, &H8CEA40F, &H9908E1C1, &HF20BC80B, &H343B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H443B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H543B60F, &H8CEA40F, &H8E1C199, &HC80BF20B
    pvAppendBuffer &H643B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H743B60F, &H8CEA40F, &H8E1C199, &HF20BC80B, &H891C7789, &H5E5F184F, &H8C25D5B, &HCCCCCC00, &H81EC8B55, &H90EC&, &HD0458D00, &H500C75FF, &HFFFACBE8, &HD0458DFF, &H4ED2E850, &HC0850000, &HC0330874, &HC25DE58B, &H458D0008, &HDE850D0, &H5FFFFC1, &H170&, &H4DD2E850, &HF8830000, &HE8157401, &HFFFFC0F8, &H17005, &H458D5000
    pvAppendBuffer &HE85050D0, &H6818&, &H458D006A, &HDDE850D0, &H5FFFFC0, &H110&, &H70858D50, &H50FFFFFF, &HFFC45BE8, &H70858DFF, &H50FFFFFF, &HFFC3EFE8, &H75C085FF, &HA0458A9D, &H24084D8B, &H88020401, &H70858D01, &H50FFFFFF, &H5001418D, &HAFE8&, &H1B800, &HE58B0000, &H8C25D, &HCCCCCCCC, &H83EC8B55, &H458D60EC, &HC75FFE0, &HFD2EE850, &H458DFFFF, &H55E850E0, &H8500004E, &H330874C0
    pvAppendBuffer &H5DE58BC0, &H8D0008C2, &HE850E045, &HFFFFC060, &H9005&, &H95E85000, &H8300004D, &H157401F8, &HFFC04BE8, &H9005FF, &H8D500000, &H5050E045, &H697BE8, &H8D006A00, &HE850E045, &HFFFFC030, &H5050C083, &H50A0458D, &HFFC563E8, &HA0458DFF, &HC37AE850, &HC085FFFF, &H458AA575, &H84D8BC0, &H2040124, &H458D0188, &H418D50A0, &H7DE85001, &HB8000002, &H1&, &HC25DE58B, &HCCCC0008
    pvAppendBuffer &H56EC8B55, &HB108758B, &H7D8B5728, &H47B60F0C, &H28468807, &H647B60F, &H8B294688, &H4578B07, &HFFCCFBE8, &H2A4688FF, &H78B20B1, &HE804578B, &HFFFFCCEC, &H8B2B4688, &H4478B0F, &H18C1AC0F, &H8B2C4E88, &H18E8C10F, &HF04478B, &H8810C1AC, &HF8B2D4E, &H8B10E8C1, &HAC0F0447, &H4E8808C1, &HC128B12E, &HB60F08E8, &H2F468807, &HF47B60F, &HF204688, &H880E47B6, &H478B2146, &HC578B08
    pvAppendBuffer &HFFCC9BE8, &H224688FF, &H478B20B1, &HC578B08, &HFFCC8BE8, &H234688FF, &H8B084F8B, &HAC0F0C47, &H4E8818C1, &H84F8B24, &H8B18E8C1, &HAC0F0C47, &H4E8810C1, &H84F8B25, &H8B10E8C1, &HAC0F0C47, &H4E8808C1, &HC128B126, &HB60F08E8, &H46880847, &H47B60F27, &H18468817, &H1647B60F, &H8B194688, &H578B1047, &HCC36E814, &H4688FFFF, &H8B20B11A, &H578B1047, &HCC26E814, &H4688FFFF, &H104F8B1B
    pvAppendBuffer &HF14478B, &H8818C1AC, &H4F8B1C4E, &H18E8C110, &HF14478B, &H8810C1AC, &H4F8B1D4E, &H10E8C110, &HF14478B, &H8808C1AC, &H28B11E4E, &HF08E8C1, &H881047B6, &HB60F1F46, &H46881F47, &H47B60F10, &H1146881E, &H8B18478B, &HD1E81C57, &H88FFFFCB, &H20B11246, &H8B18478B, &HC1E81C57, &H88FFFFCB, &H4F8B1346, &H1C478B18, &H18C1AC0F, &H8B144E88, &HE8C1184F, &H1C478B18, &H10C1AC0F, &H8B154E88
    pvAppendBuffer &HE8C1184F, &H1C478B10, &H8C1AC0F, &HB1164E88, &H8E8C128, &H1847B60F, &HF174688, &H882747B6, &HB60F0846, &H46882647, &H20478B09, &HE824578B, &HFFFFCB6C, &HB10A4688, &H20478B20, &HE824578B, &HFFFFCB5C, &H8B0B4688, &H478B204F, &HC1AC0F24, &H18E8C118, &H8B0C4E88, &H478B204F, &HC1AC0F24, &H10E8C110, &H8B0D4E88, &H478B204F, &HC1AC0F24, &H8E8C108, &HF0E4E88, &H882047B6, &HB60F0F46
    pvAppendBuffer &H6882F47, &H2E47B60F, &HB1014688, &H28478B28, &HE82C578B, &HFFFFCB08, &HB1024688, &H28478B20, &HE82C578B, &HFFFFCAF8, &H8B034688, &H478B284F, &HC1AC0F2C, &H18E8C118, &H8B044E88, &H478B284F, &HC1AC0F2C, &H10E8C110, &H8B054E88, &H478B284F, &HC1AC0F2C, &H8E8C108, &HF064E88, &H5F2847B6, &H5E074688, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &HB108758B, &H7D8B5728, &H47B60F0C
    pvAppendBuffer &H18468807, &H647B60F, &H8B194688, &H4578B07, &HFFCA8BE8, &H1A4688FF, &H78B20B1, &HE804578B, &HFFFFCA7C, &H8B1B4688, &H4478B0F, &H18C1AC0F, &H8B1C4E88, &H18E8C10F, &HF04478B, &H8810C1AC, &HF8B1D4E, &H8B10E8C1, &HAC0F0447, &H4E8808C1, &HC128B11E, &HB60F08E8, &H1F468807, &HF47B60F, &HF104688, &H880E47B6, &H478B1146, &HC578B08, &HFFCA2BE8, &H124688FF, &H478B20B1, &HC578B08
    pvAppendBuffer &HFFCA1BE8, &H134688FF, &H8B084F8B, &HAC0F0C47, &H4E8818C1, &H84F8B14, &H8B18E8C1, &HAC0F0C47, &H4E8810C1, &H84F8B15, &H8B10E8C1, &HAC0F0C47, &H4E8808C1, &HC128B116, &HB60F08E8, &H46880847, &H47B60F17, &H8468817, &H1647B60F, &H8B094688, &H578B1047, &HC9C6E814, &H4688FFFF, &H8B20B10A, &H578B1047, &HC9B6E814, &H4688FFFF, &H104F8B0B, &HF14478B, &H8818C1AC, &H4F8B0C4E, &H18E8C110
    pvAppendBuffer &HF14478B, &H8810C1AC, &H4F8B0D4E, &H10E8C110, &HF14478B, &H8808C1AC, &H28B10E4E, &HF08E8C1, &H881047B6, &HB60F0F46, &H6881F47, &H1E47B60F, &H8B014688, &H578B1847, &HC962E81C, &H4688FFFF, &H8B20B102, &H578B1847, &HC952E81C, &H4688FFFF, &H184F8B03, &HF1C478B, &HC118C1AC, &H4E8818E8, &H184F8B04, &HF1C478B, &HC110C1AC, &H4E8810E8, &H184F8B05, &HF1C478B, &HC108C1AC, &H4E8808E8
    pvAppendBuffer &H47B60F06, &H46885F18, &HC25D5E07, &HCCCC0008, &H83EC8B55, &H8B5330EC, &H570F085D, &H758B56C0, &HD045C70C, &H3&, &HD445C7, &HF000000, &H8DD84511, &HF660146, &H50F845D6, &H45110F53, &HF54AE8E8, &H3E80FFFF, &H8D157504, &H8D503146, &HE8503043, &HFFFFF538, &HE58B5B5E, &H8C25D, &H7B8D5357, &H95E85730, &HE8000059, &HFFFFBB80, &HB005&, &H458D5000, &H575750D0, &H59DFE8
    pvAppendBuffer &H57575300, &H5917E8, &HBB62E800, &HB005FFFF, &H50000000, &HFFBB57E8, &HE005FF, &H57500000, &H50FAE857, &HE8570000, &H19A4&, &HF633068A, &H1240F8B, &H83C0B60F, &H3B9901E1, &H3B0475C8, &H571274F2, &HFFBB27E8, &HB005FF, &H57500000, &H624BE8, &H5B5E5F00, &HC25DE58B, &HCCCC0008, &H83EC8B55, &H8B5320EC, &H570F085D, &H758B56C0, &HE045C70C, &H3&, &HE445C7, &HF000000
    pvAppendBuffer &H8DE84511, &HF660146, &H50F845D6, &HF78EE853, &H3E80FFFF, &H8D157504, &H8D502146, &HE8502043, &HFFFFF77C, &HE58B5B5E, &H8C25D, &H7B8D5357, &HF9E85720, &HE8000058, &HFFFFBAB4, &H5010C083, &H50E0458D, &H45E85757, &H53000059, &H7DE85757, &HE8000058, &HFFFFBA98, &H5010C083, &HFFBA8FE8, &H30C083FF, &HE8575750, &H5074&, &H198EE857, &H68A0000, &HF8BF633, &HB60F0124, &H1E183C0
    pvAppendBuffer &H75C83B99, &H74F23B04, &H61E85710, &H83FFFFBA, &H575010C0, &H6397E8, &H5B5E5F00, &HC25DE58B, &HCCCC0008, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &HF0EC&, &H10858D00, &HFFFFFFFF, &HE8500875, &HFFFFFE48, &H8D0C75FF, &HE850D045, &HFFFFF3BC, &H458D006A, &H858D50D0, &HFFFFFF10, &H70858D50, &H50FFFFFF, &HFFBD93E8, &H70858DFF, &H50FFFFFF, &HE81075FF, &HFFFFFA04, &HFF70858D
    pvAppendBuffer &HE850FFFF, &HFFFFBD18, &HC01BD8F7, &H5DE58B40, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &HA0EC&, &H60858D00, &HFFFFFFFF, &HE8500875, &HFFFFFEA8, &H8D0C75FF, &HE850E045, &HFFFFF65C, &H458D006A, &H858D50E0, &HFFFFFF60, &HA0458D50, &HBED6E850, &H458DFFFF, &H75FF50A0, &HFC0AE810, &H458DFFFF, &HE1E850A0, &HF7FFFFBC, &H40C01BD8, &HC25DE58B, &HCCCC000C, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H458D60EC, &H75FF56A0, &H7DE85008, &H8BFFFFFD, &H458D0C75, &H468D50A0, &H406C601, &HF95AE850, &H458DFFFF, &H468D50D0, &H4DE85031, &HB8FFFFF9, &H1&, &H5DE58B5E, &HCC0008C2, &H83EC8B55, &H458D40EC, &H75FF56C0, &HDE85008, &H8BFFFFFE, &H458D0C75, &H468D50C0, &H406C601, &HFB8AE850, &H458DFFFF, &H468D50E0, &H7DE85021, &HB8FFFFFB, &H1&, &H5DE58B5E, &HCC0008C2
    pvAppendBuffer &H81EC8B55, &HC0EC&, &H7D8B5700, &H8DE85710, &H85000046, &H330974C0, &HE58B5FC0, &H10C25D, &HB8CAE857, &H7005FFFF, &H50000001, &H458FE8, &H1F88300, &HB5E81274, &H5FFFFB8, &H170&, &HE8575750, &H5FD8&, &HE857006A, &HFFFFB8A0, &H11005, &H858D5000, &HFFFFFF40, &HBC1EE850, &H858DFFFF, &HFFFFFF40, &HB882E850, &H7005FFFF, &H50000001, &H4547E8, &H1F88300, &H6DE81874
    pvAppendBuffer &H5FFFFB8, &H170&, &H40858D50, &H50FFFFFF, &H5F8AE850, &H858D0000, &HFFFFFF40, &H45FEE850, &HC0850000, &HFF6D850F, &H8B56FFFF, &H858D1475, &HFFFFFF40, &H45E85650, &HFFFFFFF8, &H458D0875, &HC9E850A0, &HE8FFFFF1, &HFFFFB824, &H17005, &H458D5000, &H858D50A0, &HFFFFFF40, &HD0458D50, &H52DAE850, &H75FF0000, &HA0458D0C, &HF19EE850, &HF9E8FFFF, &H5FFFFB7, &H170&, &HD0458D50
    pvAppendBuffer &HA0458D50, &HD0458D50, &H4D92E850, &HDDE80000, &H5FFFFB7, &H170&, &HE8575750, &H4E00&, &HFFB7CBE8, &H17005FF, &H57500000, &H50D0458D, &H528AE850, &H458D0000, &H468D50D0, &HBDE85030, &H5EFFFFF7, &H1B8&, &HE58B5F00, &H10C25D, &H81EC8B55, &H80EC&, &H7D8B5700, &H6DE85710, &H85000045, &H330974C0, &HE58B5FC0, &H10C25D, &HB77AE857, &H9005FFFF, &H50000000, &H44AFE8
    pvAppendBuffer &H1F88300, &H65E81274, &H5FFFFB7, &H90&, &HE8575750, &H6098&, &HE857006A, &HFFFFB750, &H5050C083, &H5080458D, &HFFBC83E8, &H80458DFF, &HB73AE850, &H9005FFFF, &H50000000, &H446FE8, &H1F88300, &H25E81574, &H5FFFFB7, &H90&, &H80458D50, &H55E85050, &H8D000060, &HE8508045, &H44EC&, &H850FC085, &HFFFFFF7B, &H14758B56, &H5080458D, &HF976E856, &H75FFFFFF, &HC0458D08
    pvAppendBuffer &HF39AE850, &HE5E8FFFF, &H5FFFFB6, &H90&, &HC0458D50, &H80458D50, &HE0458D50, &H531EE850, &H75FF0000, &HC0458D0C, &HF372E850, &HBDE8FFFF, &H5FFFFB6, &H90&, &HE0458D50, &HC0458D50, &HE0458D50, &H4C96E850, &HA1E80000, &H5FFFFB6, &H90&, &HE8575750, &H4F24&, &HFFB68FE8, &H9005FF, &H57500000, &H50E0458D, &H52CEE850, &H458D0000, &H468D50E0, &HF1E85020, &H5EFFFFF8
    pvAppendBuffer &H1B8&, &HE58B5F00, &H10C25D, &HCCCCCCCC, &H81EC8B55, &H280EC, &H80858D00, &H56FFFFFD, &H500875FF, &HFFFA67E8, &H10758BFF, &HFED0858D, &H5056FFFF, &HFFEFD7E8, &H30468DFF, &H60858D50, &H50FFFFFF, &HFFEFC7E8, &HD0858DFF, &H50FFFFFE, &H43CBE8, &HFC08500, &H39C85, &H60858D00, &H50FFFFFF, &H43B7E8, &HFC08500, &H38885, &HD0858D00, &H50FFFFFE, &HFFB5F3E8, &H17005FF
    pvAppendBuffer &HE8500000, &H42B8&, &HF01F883, &H36885, &H60858D00, &H50FFFFFF, &HFFB5D3E8, &H17005FF, &HE8500000, &H4298&, &HF01F883, &H34885, &HE8575300, &HFFFFB5B8, &H17005, &H858D5000, &HFFFFFF60, &HC0458D50, &H4BD2E850, &H75FF0000, &H858D0C, &H50FFFFFF, &HFFEF33E8, &HB58EE8FF, &H7005FFFF, &H50000001, &H50C0458D, &HFF00858D, &H5050FFFF, &H5047E8, &HB572E800, &H7005FFFF
    pvAppendBuffer &H50000001, &H50C0458D, &HFED0858D, &H8D50FFFF, &HFFFEA085, &H25E850FF, &H8D000050, &HFFFD8085, &H858D50FF, &HFFFFFE10, &H5892E850, &H858D0000, &HFFFFFDB0, &H40858D50, &H50FFFFFE, &H587FE8, &HB52AE800, &H1005FFFF, &H50000001, &HFF30858D, &HE850FFFF, &H5868&, &HFFB513E8, &H14005FF, &H8D500000, &HFFFF6085, &H51E850FF, &HE8000058, &HFFFFB4FC, &HB005&, &H858D5000, &HFFFFFF30
    pvAppendBuffer &H10858D50, &H50FFFFFE, &H50C0458D, &H534FE8, &H40858D00, &H50FFFFFE, &HFE10858D, &H8D50FFFF, &HFFFF6085, &H858D50FF, &HFFFFFF30, &HBB9EE850, &HB9E8FFFF, &H5FFFFB4, &HB0&, &HC0458D50, &HD9E85050, &H8D00004A, &H8D50C045, &HFFFE4085, &H858D50FF, &HFFFFFE10, &HC432E850, &H45C7FFFF, &HF0&, &HB486E800, &H1005FFFF, &H89000001, &H858DF445, &HFFFFFD80, &H8DF84589, &HFFFE1085
    pvAppendBuffer &HFC4589FF, &HFEA0858D, &HE850FFFF, &H55F0&, &H858DD88B, &HFFFFFF00, &H55E2E850, &HC33B0000, &H8DD8470F, &HFFFF0085, &HFF738DFF, &HDDE85056, &HB00005E, &HBF0774C2, &H1&, &HFF3302EB, &HA0858D56, &H50FFFFFE, &H5EC3E8, &H74C20B00, &H2BE07, &H2EB0000, &HF70BF633, &H8B90458D, &H56F0B574, &H5756E850, &H468D0000, &H858D5030, &HFFFFFE70, &H5746E850, &H458D0000, &HDE850C0
    pvAppendBuffer &H8D000040, &H45C7FE73, &H1C0&, &HC445C700, &H0&, &H880FF685, &HE8&, &H401F0F, &H50C0458D, &HFE70858D, &H8D50FFFF, &HE8509045, &HFFFFB3FC, &H858D56, &H50FFFFFF, &H5E4FE8, &H74C20B00, &H1BF07, &H2EB0000, &H8D56FF33, &HFFFEA085, &H35E850FF, &HB00005E, &HB80774C2, &H2&, &HC03302EB, &H7C8BC70B, &HFF85F085, &H85840F, &H8D570000, &HFFFF3085, &HBDE850FF
    pvAppendBuffer &H8D000056, &H8D503047, &HFFFF6085, &HADE850FF, &H8D000056, &H8D50C045, &HFFFF6085, &H858D50FF, &HFFFFFF30, &HC2E6E850, &H41E8FFFF, &H5FFFFB3, &HB0&, &H30858D50, &H50FFFFFF, &H5090458D, &HFDE0858D, &HE850FFFF, &H5194&, &HFE70858D, &H8D50FFFF, &H8D509045, &HFFFF6085, &H858D50FF, &HFFFFFF30, &HB9E6E850, &H858DFFFF, &HFFFFFDE0, &HC0458D50, &HA5E85050, &H83000050, &H890F01EE
    pvAppendBuffer &HFFFFFF1C, &HFFB2E7E8, &HB005FF, &H8D500000, &H5050C045, &H4907E8, &HC0458D00, &H70858D50, &H50FFFFFE, &H5090458D, &HFFC263E8, &H90458DFF, &HB2BAE850, &H7005FFFF, &H50000001, &H3F7FE8, &H835B5F00, &H157401F8, &HFFB2A3E8, &H17005FF, &H8D500000, &H50509045, &H59C3E8, &HD0858D00, &H50FFFFFE, &H5090458D, &H3F53E8, &H5ED8F700, &H8B40C01B, &HCC25DE5, &H5EC03300, &HC25DE58B
    pvAppendBuffer &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &H1B0EC, &H50858D00, &H56FFFFFE, &H500875FF, &HFFF737E8, &H10758BFF, &HFF30858D, &H5056FFFF, &HFFEEE7E8, &H20468DFF, &H90458D50, &HEEDAE850, &H858DFFFF, &HFFFFFF30, &H3FFEE850, &HC0850000, &H36E850F, &H458D0000, &HEDE85090, &H8500003F, &H5D850FC0, &H8D000003, &HFFFF3085, &HF9E850FF, &H5FFFFB1, &H90&, &H3F2EE850
    pvAppendBuffer &HF8830000, &H3D850F01, &H8D000003, &HE8509045, &HFFFFB1DC, &H9005&, &H11E85000, &H8300003F, &H850F01F8, &H320&, &HC1E85753, &H5FFFFB1, &H90&, &H90458D50, &HE0458D50, &H4A3EE850, &H75FF0000, &H50858D0C, &H50FFFFFF, &HFFEE4FE8, &HB19AE8FF, &H9005FFFF, &H50000000, &H50E0458D, &HFF50858D, &H5050FFFF, &H4DD3E8, &HB17EE800, &H9005FFFF, &H50000000, &H50E0458D, &HFF30858D
    pvAppendBuffer &H8D50FFFF, &HFFFF1085, &HB1E850FF, &H8D00004D, &HFFFE5085, &H858D50FF, &HFFFFFEB0, &H54FEE850, &H858D0000, &HFFFFFE70, &HD0858D50, &H50FFFFFE, &H54EBE8, &HB136E800, &HC083FFFF, &H858D5050, &HFFFFFF70, &H54D6E850, &H21E80000, &H83FFFFB1, &H8D5070C0, &HE8509045, &H54C4&, &HFFB10FE8, &H10C083FF, &H70858D50, &H50FFFFFF, &HFEB0858D, &H8D50FFFF, &HE850E045, &H4F94&, &HFED0858D
    pvAppendBuffer &H8D50FFFF, &HFFFEB085, &H458D50FF, &H858D5090, &HFFFFFF70, &HB8B6E850, &HD1E8FFFF, &H83FFFFB0, &H8D5010C0, &H5050E045, &H4953E8, &HE0458D00, &HD0858D50, &H50FFFFFE, &HFEB0858D, &HE850FFFF, &HFFFFC09C, &HD045C7, &HE8000000, &HFFFFB0A0, &H8950C083, &H858DD445, &HFFFFFE50, &H8DD84589, &HFFFEB085, &HDC4589FF, &HFF10858D, &HE850FFFF, &H525C&, &H858DD88B, &HFFFFFF50, &H524EE850
    pvAppendBuffer &HC33B0000, &H8DD8470F, &HFFFF5085, &HFF738DFF, &HF9E85056, &HB00005A, &HBF0774C2, &H1&, &HFF3302EB, &H10858D56, &H50FFFFFF, &H5ADFE8, &H74C20B00, &H2BE07, &H2EB0000, &HF70BF633, &H8BB0458D, &H56D0B574, &H53D2E850, &H468D0000, &H858D5020, &HFFFFFEF0, &H53C2E850, &H458D0000, &H89E850E0, &H8D00003C, &H45C7FE73, &H1E0&, &HE445C700, &H0&, &H880FF685, &HD5&
    pvAppendBuffer &H50E0458D, &HFEF0858D, &H8D50FFFF, &HE850B045, &HFFFFB19C, &H50858D56, &H50FFFFFF, &H5A6FE8, &H74C20B00, &H1BF07, &H2EB0000, &H8D56FF33, &HFFFF1085, &H55E850FF, &HB00005A, &HB80774C2, &H2&, &HC03302EB, &H7C8BC70B, &HFF85D085, &H8D577A74, &HFFFF7085, &H41E850FF, &H8D000053, &H8D502047, &HE8509045, &H5334&, &H50E0458D, &H5090458D, &HFF70858D, &HE850FFFF, &HFFFFBF60
    pvAppendBuffer &HFFAF6BE8, &H10C083FF, &H70858D50, &H50FFFFFF, &H50B0458D, &HFE90858D, &HE850FFFF, &H4DF0&, &HFEF0858D, &H8D50FFFF, &H8D50B045, &H8D509045, &HFFFF7085, &H15E850FF, &H8DFFFFB7, &HFFFE9085, &H458D50FF, &HE85050E0, &H4D04&, &HF01EE83, &HFFFF2B89, &HAF16E8FF, &HC083FFFF, &H458D5010, &HE85050E0, &H4798&, &H50E0458D, &HFEF0858D, &H8D50FFFF, &HE850B045, &HFFFFBEE4, &H50B0458D
    pvAppendBuffer &HFFAEEBE8, &H9005FF, &HE8500000, &H3C20&, &HF8835B5F, &HE8157401, &HFFFFAED4, &H9005&, &H458D5000, &HE85050B0, &H5804&, &HFF30858D, &H8D50FFFF, &HE850B045, &H3BF4&, &H1B5ED8F7, &HE58B40C0, &HCC25D, &H8B5EC033, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HC18B084D, &H8107E8C1, &H7F7F7FE1, &H10125FF, &HC9030101, &H331BC06B, &H4C25DC1
    pvAppendBuffer &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &HE8EC8B55, &HFFFFAE78, &HB992F0B9, &HE98100, &H300B940, &H84D8BC1, &H75FF5051, &H30418D10, &H6A0C75FF, &H418D5010, &H71E85020, &H5DFFFFCB, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H458B084D, &H38410110, &H3C5183, &H89104589, &HE95D084D, &HFFFFFFA4, &HCCCCCCCC, &H56EC8B55, &H8308758B, &H7501487E, &H2DE8560D
    pvAppendBuffer &HC7000000, &H24846, &H458B0000, &H40460110, &HC75FF50, &H445683, &HFF72E856, &H5D5EFFFF, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H8B08758B, &HC985304E, &H10B82974, &H2B000000, &H468D50C1, &H6AC10320, &H5DE85000, &H8300000B, &H468D0CC4, &HE8565020, &H10&, &H3046C7, &H5E000000, &H4C25D, &HCCCCCCCC, &H83EC8B55, &H458D10EC, &H505756F0, &HE80C75FF
    pvAppendBuffer &HFFFFDBFC, &H8D08758B, &H7E8DF045, &H50575710, &HFFDB3BE8, &H575657FF, &HFFDC83E8, &H8B5E5FFF, &H8C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H565314EC, &H8B08758B, &HF8834846, &H83057401, &HD7502F8, &HFF62E856, &H46C7FFFF, &H48&, &H385E8B00, &HF3C568B, &H6A03DAA4, &HC1C28B08, &HE8C103E3, &H88CB8B18, &HC28BEC45, &H8810E8C1, &HC28BED45, &H8808E8C1, &HB60FEE45
    pvAppendBuffer &HEF4588C2, &HAC0FC28B, &H558918C1, &H18E8C1FC, &H8BF04D88, &H88CB8BC2, &HAC0FF35D, &HE8C110C1, &H88C38B10, &HAC0FF14D, &H458808D0, &HEC458DF2, &H8EAC150, &HFE56E856, &H5E8BFFFF, &H44568B40, &H3DAA40F, &HC28B086A, &HC103E3C1, &HCB8B18E8, &H8BEC4588, &H10E8C1C2, &H8BED4588, &H8E8C1C2, &HFEE4588, &H4588C2B6, &HFC28BEF, &H8918C1AC, &HE8C1FC55, &HF04D8818, &HCB8BC28B, &HFF35D88
    pvAppendBuffer &HC110C1AC, &HC38B10E8, &HFF14D88, &H8808D0AC, &H458DF245, &HEAC150EC, &HF1E85608, &HFFFFFFFD, &H468D0C75, &H5E85010, &H5EFFFFDC, &H5DE58B5B, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H6A08758B, &H56006A50, &H9DFE8, &HCC48300, &HC75FF56, &HFFDAA3E8, &H4846C7FF, &H1&, &H8C25D5E, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &H5714758B, &HEE83FF33, &H8B327801
    pvAppendBuffer &H8B530C45, &HD82B085D, &H8D104529, &H9066B014, &H8D130C8B, &HC033FC52, &HC013CF03, &H83044A03, &HEE8300D0, &H8BF88B01, &H4C891045, &HE0790410, &H5FC78B5B, &H10C25D5E, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &H107D8B57, &H8399C78B, &HD0031FE2, &H8105FAC1, &H1FE7&, &H4F057980, &H47E0CF83, &H8B0C758B, &HD3C68BCF, &H75FF85E0, &HEBF63304, &H20B909, &HCF2B0000, &H7D8BEED3, &H1C93308
    pvAppendBuffer &H13049744, &H75F685C9, &H74C98504, &H3C03327, &H1C013CE, &H8308974C, &HC28300D0, &H74C08503, &H97148D13, &H528DC88B, &H1C03304, &HC013FC4A, &HF075C085, &HC25D5E5F, &HCCCC000C, &HCCCCCCCC, &H83EC8B55, &H458B38EC, &H14558B0C, &HF08B5653, &HF22BDB33, &HEC758957, &H1FD880F, &H458B0000, &HC1CE8B08, &HD68B05E1, &HFF44D89, &H441F&, &H8598348B, &H430C75F6, &H8920E983, &HCAE9F44D
    pvAppendBuffer &H56000001, &HFFC0D7E8, &H8BF88BFF, &H85E6D3CF, &H8D1A7EFF, &H453B0143, &H8B127D0C, &H20B90845, &H2B000000, &H98548BCF, &HBEAD304, &H204D8BF2, &H65F7C68B, &HF4458B1C, &H2BE1C183, &H3F28BC7, &HFC7589C8, &H79F84D89, &HE0F9831D, &H1838E0F, &HD9F70000, &HC933EED3, &H89FC7589, &HF685F84D, &H16F840F, &HF98B0000, &H1FE781, &H5798000, &HE0CF834F, &H89C18B47, &H8399D07D, &HC2031FE2
    pvAppendBuffer &HC10C558B, &HD02B05F8, &H2B14458B, &H75FF85D0, &HFF488D71, &H1E845C7, &H8D000000, &H570F1134, &H130F66C0, &HF33BC845, &H1018C0F, &H458B0000, &HB03C8D08, &H89CC458B, &H458BE445, &HF04589C8, &H479C985, &H6EBC033, &H8B10458B, &H65F78804, &HF04503FC, &H5513D0F7, &HF05589E4, &H703D233, &HE445C7, &H13000000, &HE84503D2, &HD2830789, &H89494E00, &HEF83E855, &H7DF33B04, &HADE9C3
    pvAppendBuffer &H788D0000, &HDC45C7FF, &H0&, &H8917048D, &HC33BD845, &H9B8C0F, &H4D8B0000, &H20BAD0, &HD12B0000, &H1E845C7, &H89000000, &H570FCC55, &H8558BC0, &H45130F66, &H82048DE0, &H8BF04589, &H4589E445, &HE0458BD4, &HFE44589, &H441F&, &H479FF85, &H6EBC033, &H8B10458B, &HE6F7B804, &H7503F08B, &HD45513E4, &H5589C033, &HD3D68BE4, &HDC550BE2, &HF7F04D8B, &HD445C7D2, &H0&
    pvAppendBuffer &HC0131103, &H89E85503, &HCC4D8B11, &H8300D083, &H4F04F06D, &H8BE84589, &HFC758BC6, &H4D8BE8D3, &HDC4589D0, &H48D8458B, &H3BD84589, &H8BA17DC3, &H758BF84D, &H187D8BFC, &H874FF85, &HE8575651, &HFFFFFD7C, &H8BEC558B, &H458BF44D, &HFDA3B08, &HFFFE1B8E, &HC458BFF, &H8B14558B, &HC933105D, &H2E7EC085, &HF82BFA8B, &HFBB148D, &H441F&, &H850F048D, &H330479C0, &H8B02EBC0, &H8758B02
    pvAppendBuffer &H3B8E348B, &H776272F0, &HC2834109, &HC4D3B04, &H4D8BDE7C, &HC45C70C, &H1&, &H85FF518D, &H8B3578D2, &H3F12BF2, &H3C8D1475, &H85D8BB3, &H479F685, &H2EBC933, &HC0330F8B, &HC03D1F7, &H3C01393, &HC890C4D, &HD08393, &H4EF834E, &H830C4589, &HD87901EA, &H8518458B, &H6A0A74C0, &H50016A00, &HFFFCD7E8, &H5B5E5FFF, &HC25DE58B, &HCCCC001C, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H8B531CEC, &H5756145D, &HF32FB83, &H1B48E, &H187D8B00, &H758BC38B, &H99CB8B0C, &HD1C22B57, &H75FF50F8, &H89C82B10, &HFF56F845, &H4D890875, &HFFC6E8FC, &H4D8BFFFF, &H10458BF8, &HFC75FF57, &H8D8E148D, &H5589C804, &HF44589E4, &HFF08458B, &H8D52F475, &H89508804, &H9DE8EC45, &H8BFFFFFF, &H8D40FC45, &H518D870C, &H1C704, &HC7000000, &H2&, &HC4D8900, &H89044F8D
    pvAppendBuffer &H558BE855, &HF04D89F8, &H1C7&, &H7C70000, &H0&, &H3B7ED285, &H3F85D8B, &H8DC22BC0, &H458B870C, &HC22B40FC, &H8B87148D, &HFE2B087D, &H8D37048B, &H42890452, &H4768DFC, &H8DFC468B, &H41890449, &H1EB83FC, &H5D8BE675, &H187D8B14, &H8BF04D8B, &H5156FC75, &H51EC75FF, &HFFFB8FE8, &H78956FF, &H50E8458B, &H50E475FF, &HFFFB7FE8, &HC4D8BFF, &H4E8D0189, &H8DC18B01, &HE0C1CF14
    pvAppendBuffer &H89C70304, &H51500855, &HC75FF52, &HFEEEE857, &H558BFFFF, &HC7C933F8, &HC47&, &H47C70000, &H8&, &H447C700, &H0&, &HC712348D, &H7&, &H7EF68500, &HFC458B1D, &H40105D8B, &H148DC22B, &H8B048BC7, &H4104528D, &H3BFC4289, &H8BF27CCE, &H758B145D, &H36048DFC, &H8478D50, &HF475FF50, &HFB06E850, &H4789FFFF, &H1468D04, &H8B00348D, &H50560845, &H31E85057, &H8B000002
    pvAppendBuffer &H7D8BFC4D, &H418D5610, &H3D82B01, &H8DD92BDB, &HFF509F04, &HE8500875, &HFFFFFAD4, &HF685F08B, &H9C840F, &H578D0000, &H9A148DFC, &H401F0F, &H528DC033, &H47201FC, &HF08BC013, &HF075F685, &H8B5B5E5F, &H14C25DE5, &H10558B00, &H851B0C8D, &H33067EC9, &HF3FA8BC0, &H8458BAB, &H9D0C8D, &H8D000000, &H1C8DDA14, &H76D83B01, &HC458B55, &H8D01348D, &H7589FC4A, &H8558B14, &H90104D89
    pvAppendBuffer &H3304EB83, &H76F03BFF, &H801F0F27, &H0&, &HEE83038B, &H326F704, &HD283C7, &H458B0101, &HD2830C, &H8B04E983, &H77F03BFA, &H8558BE3, &H8914758B, &H104D8B39, &H8904E983, &HDA3B104D, &H5E5FBE77, &H5DE58B5B, &HCC0014C2, &HCCCCCCCC, &H83EC8B55, &H565310EC, &H5714758B, &HF32FE83, &HB88E&, &H184D8B00, &H7D8BC68B, &HC22B990C, &H1C8DD08B, &H8BFAD1F1, &H89C22BC6, &H5053FC55
    pvAppendBuffer &H8DF84589, &H9534&, &H48D0000, &HF47589D1, &H3E048D50, &H5008758B, &H5096048D, &HFFFD5FE8, &HF4458BFF, &H53184503, &H89FC75FF, &H8B500845, &H48DF845, &HE8565087, &HFFFFFF94, &HFC75FF53, &HFFF85D8B, &H8D571875, &HE8509E04, &HFFFFFF80, &H8B107D8B, &HDB85FC75, &H558B1D7E, &H73048D18, &H3F44D8B, &H82148DCF, &H528D028B, &H8D018904, &HEB830449, &H8BF17501, &H5356185D, &H530875FF
    pvAppendBuffer &HFFF96BE8, &H14458BFF, &H48D5756, &HE8535083, &HFFFFF95C, &H8B5B5E5F, &H14C25DE5, &H105D8B00, &H87EF685, &HC033CE8B, &HABF3FB8B, &H8D08458B, &H5589B314, &HB03C8D14, &H6A76F83B, &H830C458B, &H348D04E8, &H8458BB0, &HF0C7589, &H441F&, &HF04EF83, &HF66C057, &H8BF04513, &H76D33BCA, &HF0458B38, &H89F45D8B, &HF661845, &H441F&, &H768D068B, &H8327F7FC, &H10304E9, &H300D283
    pvAppendBuffer &H1891845, &HDB33D313, &H3B185589, &HE077104D, &H8B14558B, &H458B105D, &HC758B08, &H8904EA83, &HF83B1455, &H5E5FAA77, &H5DE58B5B, &HCC0014C2, &H56EC8B55, &H8314758B, &H45C701EE, &H114&, &H8B357800, &H8B530845, &H8B57105D, &H148D0C7D, &H2BF82BB0, &H170C8BD8, &H33FC528D, &H3D1F7C0, &HC013044A, &H89144D03, &H8304134C, &HEE8300D0, &H14458901, &H5B5FDF79, &H10C25D5E, &HCCCCCC00
    pvAppendBuffer &H81EC8B55, &H80EC&, &H20B900, &H8B530000, &H57560C5D, &H7D8DF38B, &HBEA5F380, &HFD&, &H5080458D, &H2A96E850, &HFE830000, &H83107402, &HB7404FE, &H80458D53, &HA1E85050, &H83000014, &HDC7901EE, &H8D087D8B, &H20B98075, &HF3000000, &H5B5E5FA5, &HC25DE58B, &HCCCC0008, &HCCCCCCCC, &H53EC8B55, &H8758B56, &H71E85657, &H8BFFFFF5, &H69E853D8, &H8BFFFFF5, &H61E852D0, &H8BFFFFF5
    pvAppendBuffer &H8BFE33F8, &H33C78BF7, &H8CFC1C3, &HC0C1F233, &HC1CE8B08, &HC13310C9, &HC633C733, &H33C3335F, &H5B5E0845, &H4C25D, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &HFF08758B, &HFFA2E836, &H76FFFFFF, &HE8068904, &HFFFFFF98, &H890876FF, &H8DE80446, &HFFFFFFFF, &H46890C76, &HFF82E808, &H4689FFFF, &HC25D5E0C, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &H53EC8B55, &H56085D8B, &H7BB60F57, &H43B60F07
    pvAppendBuffer &H73B60F02, &H53B60F0B, &H8E7C10F, &HB60FF80B, &HB60F034B, &HE7C10D43, &HC1F80B08, &HB60F08E6, &HE7C10843, &HC1F80B08, &HB60F08E2, &HF00B0643, &HF08E1C1, &HC10143B6, &HF00B08E6, &HC43B60F, &HB08E6C1, &H43B60FF0, &HFD00B0A, &HC10543B6, &HD00B08E2, &HC103B60F, &HD00B08E2, &HE43B60F, &H5389C80B, &H43B60F0C, &H8E1C109, &H7389C80B, &H43B60F08, &H47B8904, &H5F08E1C1, &H895EC80B
    pvAppendBuffer &HC25D5B0B, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &HFFA2D7E8, &H8758BFF, &H69305, &H36FF5000, &H2A47E8, &HE8068900, &HFFFFA2C0, &H69305, &H76FF5000, &H2A32E804, &H46890000, &HA2AAE804, &H9305FFFF, &H50000006, &HE80876FF, &H2A1C&, &HE8084689, &HFFFFA294, &H69305, &H76FF5000, &H2A06E80C, &H46890000, &HC25D5E0C, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H8BEC8B55, &HD08B0845, &H10758B56, &H1574F685, &HC7D8B57, &HC8AF82B, &H1528D17, &H83FF4A88, &HF27501EE, &HC35D5E5F, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HC985104D, &HB60F1F74, &H8B560C45, &H1C069F1, &H57010101, &HC1087D8B, &HABF302E9, &HE183CE8B, &H5FAAF303, &H8458B5E, &HCCCCC35D, &H56EC8B55, &H5608758B, &HFFF373E8, &H8BD08BFF, &HC1D633CE, &HC2C110C9, &H8CEC108, &HD633D133
    pvAppendBuffer &H5D5EC233, &HCC0004C2, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &HFF08758B, &HFFC2E836, &H76FFFFFF, &HE8068904, &HFFFFFFB8, &H890876FF, &HADE80446, &HFFFFFFFF, &H46890C76, &HFFA2E808, &H4689FFFF, &HC25D5E0C, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H570F60EC, &HA045C7C0, &H1&, &H458D5756, &HA445C7A0, &H0&, &H45110F50, &HD045C7A8, &H1&, &HB845110F, &HD445C7
    pvAppendBuffer &H66000000, &HC845D60F, &HD845110F, &HE845110F, &H45D60F66, &HA156E8F8, &HB005FFFF, &H50000000, &H50A0458D, &H29E7E8, &HA0458D00, &H42CEE850, &H7D8B0000, &HFF708D08, &H7601FE83, &H1F0F2C, &H50D0458D, &H3F36E850, &H8D560000, &HE850A045, &H4BBC&, &HB74C20B, &HD0458D57, &HBDE85050, &H4E00003E, &H7701FE83, &HD0458DD7, &H4DE85750, &H5F000044, &H5DE58B5E, &HCC0004C2, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &HF5640EC, &H45C7C057, &H1C0&, &H458D5700, &HC445C7C0, &H0&, &H45110F50, &HE045C7C8, &H1&, &H45D60F66, &HE445C7D8, &H0&, &HE845110F, &H45D60F66, &HA0AEE8F8, &HC083FFFF, &H458D5010, &H81E850C0, &H8D00002B, &HE850C045, &H4278&, &H8D087D8B, &HFE83FF70, &H8D297601, &H5050E045, &H3EC3E8, &H458D5600, &H19E850C0, &HB00004B, &H570B74C2, &H50E0458D
    pvAppendBuffer &H3E4AE850, &H834E0000, &HD77701FE, &H50E0458D, &H440AE857, &H5E5F0000, &HC25DE58B, &HCCCC0004, &H83EC8B55, &H565314EC, &HA042E857, &H5D8BFFFF, &H8B008B0C, &H8D0C8D0B, &H4&, &H8BD0FF51, &H89F88B0B, &HC8DF47D, &H48D&, &H57535100, &HFFFDABE8, &HCC483FF, &HFFA013E8, &H8758BFF, &HE8B008B, &H48D0C8D, &H51000000, &HE8BD0FF, &H4589D88B, &H8D0C8DFC, &H4&, &HE8535651
    pvAppendBuffer &HFFFFFD7C, &HE80CC483, &HFFFF9FE4, &H794B08D, &HD9E80000, &H8BFFFF9F, &H8D0C8D0E, &H4&, &HFF088B51, &H890E8BD1, &HC8DF845, &H48D&, &H50565100, &HFFFD47E8, &HCC483FF, &HFF9FAFE8, &H98B08DFF, &HE8000007, &HFFFF9FA4, &HC8D0E8B, &H48D&, &H88B5100, &H4589D1FF, &H8D068B08, &H48504, &H56500000, &H5608758B, &HFFFD0FE8, &HCC483FF, &H1F045C7, &HE8000000, &HFFFF9F70
    pvAppendBuffer &H79805, &HE8535000, &HFFFFB3B4, &H840FC085, &H161&, &HFF9F57E8, &H79405FF, &H53500000, &HFFB39BE8, &HFC085FF, &H22A84, &H8D3B8B00, &H4BD34, &H35E80000, &H56FFFF9F, &HD0FF008B, &H6AD88B56, &H5D895300, &HFCE2E8EC, &H3B89FFFF, &H8B0CC483, &H1B8BF45D, &H49D3C8D, &HE8000000, &HFFFF9F0C, &HFF008B57, &HF08B57D0, &HE856006A, &HFFFFFCBC, &H83F47D8B, &H1E890CC4, &H56EC5D8B
    pvAppendBuffer &HFC75FF53, &HAF26E857, &H3B83FFFF, &H90117601, &H3C83038B, &H8750083, &H83038948, &HF07701F8, &H76013E83, &H83068B10, &H7500863C, &H6894808, &H7701F883, &H8D078BF0, &H4850C, &HC9850000, &HC78B0D74, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9E9C, &H8408B57, &H458BD0FF, &H8BF88BFC, &H8951F84D, &H458BF445, &H89565008, &H4D89FC5D, &HF84589EC, &HFFB087E8, &H84589FF, &HF7F0458B
    pvAppendBuffer &HF04589D8, &H8BEC458B, &H8D0C8D08, &H4&, &HB74C985, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9E4C, &H8BEC75FF, &HD0FF0840, &HC8D068B, &H485&, &H74C98500, &HFC68B12, &H441F&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9E20, &H8408B56, &H15E8D0FF, &H5FFFF9E, &H798&, &H59E85350, &H8BFFFFB2, &HC085FC5D, &HFEA2850F, &H758BFFFF, &H8D038B08, &H4850C, &HC9850000
    pvAppendBuffer &HC38B0D74, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9DDC, &H8408B53, &H78BD0FF, &H4850C8D, &H85000000, &H8B1474C9, &H801F0FC7, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9DB0, &H8408B57, &H7D8BD0FF, &H8D078BF8, &H4850C, &HC9850000, &HC78B1574, &H841F0F, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9D80, &H8408B57, &H7D83D0FF, &H8D0F00F0, &H189&
    pvAppendBuffer &H8B0C458B, &HBD348D38, &H4&, &HFF9D5FE8, &H8B56FF, &H8B56D0FF, &HF04589F0, &HE856006A, &HFFFFFB0C, &H1B8&, &H833E8900, &H45890CC4, &H8BC933EC, &H8BF83BD8, &H8C0F087D, &H122&, &H8D0C458B, &HC72B0456, &H8BF44589, &H89C62BC7, &HBDE90845, &H8B000000, &H850C8D03, &H4&, &HD74C985, &HC6C38B, &H1408D00, &H7501E983, &H9CFAE8F5, &H8B53FFFF, &HD0FF0840, &HC8D078B
    pvAppendBuffer &H485&, &H74C98500, &HFC78B12, &H441F&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9CD0, &H8408B57, &H7D8BD0FF, &H8D078BF8, &H4850C, &HC9850000, &HC78B1574, &H841F0F, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9CA0, &H8408B57, &H758BD0FF, &H8D068B08, &H4850C, &HC9850000, &HC68B1574, &H841F0F, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9C70
    pvAppendBuffer &H8408B56, &H5E5FD0FF, &H8B5BC033, &H8C25DE5, &H8458B00, &H3B0C758B, &H3087F1E, &H348BF445, &H3302EB10, &H7F1F3BF6, &H8458B08, &HEB10048B, &H2BC03302, &H2BD0F7F0, &H853289F1, &H3B0774C9, &H41C91BF0, &HC63B06EB, &HD9F7C91B, &H85EC458B, &HF0758BF6, &H43C3450F, &H8904C283, &H1E3BEC45, &H689AF7E, &HC8D078B, &H485&, &H74C98500, &HC6C78B0D, &H408D0000, &H1E98301, &HEDE8F575
    pvAppendBuffer &H57FFFF9B, &HFF08408B, &HC68B5FD0, &HE58B5B5E, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H565324EC, &H5710758B, &H5D891E8B, &H9D3C8DF8, &H0&, &HE8DC7D89, &HFFFF9BB0, &HFF008B57, &H89D233D0, &HDB851045, &H9066117E, &HCA2B0E8B, &H898E0C8B, &H3B42900C, &H8BF17CD3, &H7D8B0875, &H8BCF8B0C, &HF073B06, &H98BCE47, &H8D084D89, &HC33B0904, &HC38B0D7F, &H8BC22B99
    pvAppendBuffer &H41F9D1C8, &H8D084D89, &H8D1C&, &H5D890000, &H9B5AE8F4, &H8B53FFFF, &H8BD0FF00, &H162B0855, &H85FC4589, &H8B0C7ED2, &HCA8BFC7D, &HABF3C033, &H8B0C7D8B, &H85C93306, &H8B1B7EC0, &H148DFC5D, &HF45D8B93, &H528DC12B, &H48B4104, &HFC428986, &HC83B068B, &H15E8EE7C, &H53FFFF9B, &HD0FF008B, &H8B08558B, &H89172BF0, &HD285EC75, &HCA8B0B7E, &HFE8BC033, &H7D8BABF3, &H33078B0C, &H7EC085C9
    pvAppendBuffer &H96148D15, &H528DC12B, &H48B4104, &HFC428987, &HC83B078B, &H7D8BEE7C, &HFD348D08, &H0&, &HE8E07589, &HFFFF9AC8, &HFF008B56, &H458957D0, &HD8AE808, &H348D0000, &H85&, &HE4758900, &HFF9AABE8, &H88B56FF, &H4D8BD1FF, &HF0458910, &H8956318B, &H85E80C75, &H89FFFFB0, &HC085E845, &HC88B2374, &H7589E6D3, &HF8758B0C, &H7E01FE83, &H20B917, &HC82B0000, &H8B10458B, &HE8D30440
    pvAppendBuffer &HEB0C4509, &HF8758B03, &H57F075FF, &HFF0875FF, &H75FFEC75, &HF1E6E8FC, &H75FFFFFF, &HFFFF03E8, &H7D890C75, &H16D6E8F8, &H6A500000, &H75FF5600, &H75FF5710, &HEF06E808, &HF73BFFFF, &H4C0FC78B, &HC4589C6, &H485348D, &HE8000000, &HFFFF9A20, &HFF008B56, &HF88B56D0, &HE857006A, &HFFFFF7D0, &H830C458B, &HD2330CC4, &HC0850789, &H4D8B277E, &HC5D8BF8, &H458BC82B, &H88348D08, &H401F0F
    pvAppendBuffer &H768D0F8B, &HFC468B04, &H8942CA2B, &HD33B8F04, &H5D8BEE7C, &H13F83F4, &HF661676, &H441F&, &H3C83078B, &H8750087, &H83078948, &HF07701F8, &H85F0758B, &H8B1474F6, &HC985E44D, &HC68B0D74, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF99A0, &H8408B56, &H758BD0FF, &H74F68508, &HE04D8B14, &HD74C985, &HC6C68B, &H1408D00, &H7501E983, &H997AE8F5, &H8B56FFFF, &HD0FF0840, &H85DC4D8B
    pvAppendBuffer &H8B0E74C9, &HC61045, &H1408D00, &H7501E983, &H995AE8F5, &H4D8BFFFF, &H408B5110, &H8BD0FF08, &HF685FC75, &HDB851574, &HCB8B1174, &H9066C68B, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9930, &H8408B56, &H758BD0FF, &H74F685EC, &H74DB8511, &HC6C68B0D, &H408D0000, &H1EB8301, &HDE8F575, &H56FFFF99, &HFF08408B, &H5FC78BD0, &HE58B5B5E, &HCC25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H8B5328EC, &H43F6105D, &H8D570104, &H1475047B, &HC75FF53, &HE80875FF, &H4D0&, &HE58B5B5F, &HCC25D, &H75FF5356, &HAA6EE808, &H1B8BFFFF, &HE1C1CB8B, &H45895105, &HAEDEE8E8, &HF08BFFFF, &H1075FF56, &HFFF853E8, &H1075FFFF, &H56084589, &HE8E875FF, &HFFFFFCC4, &H8BF04589, &H88BE845, &H48D148D, &H85000000, &H8B0D74D2, &H1C6C8, &H8301498D, &HF57501EA, &HFF986FE8
    pvAppendBuffer &HE875FFFF, &HFF08408B, &H1075FFD0, &HAA0EE856, &HE8BFFFF, &H8DEC4589, &H48D14, &HD2850000, &HCE8B0D74, &H8D0001C6, &HEA830149, &HE8F57501, &HFFFF983C, &H8408B56, &H348DD0FF, &H9D&, &HDC758900, &HFF9827E8, &H8B56FF, &H4589D0FF, &H7EDB85E8, &HFC4E8D1A, &HC803D38B, &H401F0F, &H7F8D078B, &H8D018904, &HEA83FC49, &HE8F17501, &HFFFF97FC, &HFF008B56, &H33F88BD0, &HF47D89C9
    pvAppendBuffer &H297EDB85, &H8D08458B, &HC083FC56, &H90D70304, &H3B087D8B, &H8B047D0F, &H3302EB38, &H413A89FF, &H8304C083, &HCB3B04EA, &H7D8BE67C, &H8558BF4, &HC8D028B, &H485&, &H74C98500, &H90C28B0E, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF97A0, &H8B0875FF, &HD0FF0840, &HFF9793E8, &H8B56FF, &HD08BD0FF, &H85E45589, &H8B0E7EDB, &HC1FA8BCE, &HC03302E9, &H7D8BABF3, &H575753F4, &HF2AEE852
    pvAppendBuffer &H7D8BFFFF, &H85C033F0, &H8B227EDB, &H4F8DE455, &HFCC28304, &H73BD603, &H318B047D, &HF63302EB, &H83403289, &HEA8304C1, &H7CC33B04, &H8D078BE9, &H4850C, &HC9850000, &HC78B1174, &H401F0F, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9720, &H8408B57, &H348DD0FF, &HDD&, &HE0758900, &HFF970BE8, &H8B56FF, &HF88BD0FF, &HE8087D89, &HFFFF96FC, &HFF088B56, &HF84589D1, &HDB85C033
    pvAppendBuffer &H4D8B257E, &HFCC683EC, &H304C183, &HEC7D8BF7, &H47D073B, &H2EB118B, &H1689D233, &H4C18340, &H3B04EE83, &H8BE97CC3, &H68BEC75, &H4850C8D, &H85000000, &H8B1474C9, &H801F0FC6, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF96A0, &H8408B56, &HE853D0FF, &H964&, &H35B0C8D, &H85348DC1, &H0&, &HE8D87589, &HFFFF9680, &HFF008B56, &H87D8BD0, &H458BD08B, &HB9F6330C
    pvAppendBuffer &H1F&, &H89FC5589, &H4D89EC75, &H7E3039F0, &H8D108B31, &H1B8903C, &HD3000000, &H750785E0, &H1E98312, &HB9460979, &H1F&, &H3B04EF83, &H8BE37CF2, &H558B087D, &HC458BFC, &H89EC7589, &H303BF04D, &H9F8D0F, &HF660000, &H441F&, &H880FC985, &H7D&, &H1BE&, &H90C6D300, &H75FF5352, &H9F048DF8, &H91E85050, &H53FFFFED, &HFFFC75FF, &H75FFF475, &HF875FFE8, &H58FE8
    pvAppendBuffer &HC4D8B00, &H452B018B, &H813485EC, &H75FF2674, &HF8458BFC, &H75FF5753, &H98048DE4, &HED5EE850, &HFF53FFFF, &H75FFFC75, &HE875FFF4, &H55EE857, &H8EB0000, &H7D8BC78B, &HF84589F8, &H8BF0458B, &HD148FC55, &HF04589CE, &H9479C085, &H8BEC758B, &H7D890C45, &H1FB94608, &H89000000, &H4D89EC75, &HF303BF0, &HFFFF678C, &HFF5253FF, &H75FFF475, &H19E857E8, &H8B000005, &H388B1045, &H4BD348D
    pvAppendBuffer &HE8000000, &HFFFF9568, &HFF008B56, &H6A56D0, &H10458950, &HFFF317E8, &H10558BFF, &H890CC483, &H7EDB853A, &H8458B2A, &H9D3C8D, &H3000000, &H89F633C7, &HA8B0C45, &HCE2B008B, &H8A048946, &H830C458B, &H458904C0, &H7CF33B0C, &H8B03EBE9, &H3A83DC7D, &H8758B01, &HF661676, &H441F&, &H3C83028B, &H8750082, &H83028948, &HF07701F8, &H85FC5D8B, &H8B1474DB, &HC985D84D, &HC38B0D74
    pvAppendBuffer &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF94E0, &H8408B53, &HF685D0FF, &H4D8B1474, &H74C985E0, &HC6C68B0D, &H408D0000, &H1E98301, &HBDE8F575, &H56FFFF94, &HFF08408B, &HF85D8BD0, &H74DB855E, &HE04D8B1A, &H1374C985, &HF66C38B, &H441F&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9490, &H8408B53, &H5D8BD0FF, &H74DB85F4, &H74FF8513, &H8BCF8B0F, &HC6C3&, &H8301408D, &HF57501E9
    pvAppendBuffer &HFF946BE8, &H408B53FF, &H8BD0FF08, &HDB85E85D, &HFF851974, &HCF8B1574, &HF66C38B, &H441F&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9440, &H8408B53, &H5D8BD0FF, &H74DB85E4, &H74FF8511, &HC6C38B0D, &H408D0000, &H1EF8301, &H1DE8F575, &H53FFFF94, &HFF08408B, &H10458BD0, &HE58B5B5F, &HCC25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H565330EC, &H107D8B57, &H875FF57
    pvAppendBuffer &HFFA59BE8, &H891F8BFF, &H5D89EC45, &H9D348DE0, &H0&, &HE8D47589, &HFFFF93D4, &HFF088B56, &H89D233D1, &HDB85F845, &HF66157E, &H441F&, &HCA2B0F8B, &H898F0C8B, &H3B42900C, &HE8F17CD3, &HFFFF93AC, &HFF008B56, &HEC7D8BD0, &HF08BD38B, &H2BD87589, &H7ED28517, &H33CA8B0B, &HF3FE8BC0, &HEC7D8BAB, &HC933078B, &H167EC085, &H9096148D, &H528DC12B, &H48B4104, &HFC428987, &HC83B078B
    pvAppendBuffer &H3C8DEE7C, &HDD&, &HDC7D8900, &HFF935FE8, &H8B57FF, &HF08BD0FF, &HE8087589, &HFFFF9350, &HFF088B57, &HFC4589D1, &H851B048D, &H8B0F7EC0, &H33FE8BC8, &H8DABF3C0, &HDD3C&, &HC7530000, &H1FC3744, &HE8000000, &H5F4&, &H85348D, &H89000000, &H15E8D075, &H56FFFF93, &HD0FF008B, &H33F04589, &HC458BF6, &H8DF47589, &H88B1F7E, &H85E44D89, &H8D2E7EC9, &HD98B8814, &H401F0F
    pvAppendBuffer &H1B8&, &HD3CF8B00, &H750285E0, &H1EF8312, &HBF460979, &H1F&, &H3B04EA83, &H8BE17CF3, &H7589E05D, &HF8458BF4, &H8950008B, &HADE8E845, &H8BFFFFA8, &H4589E855, &H74C085E0, &HD3C88B1A, &H1FB83E2, &H20B9117E, &H2B000000, &HF8458BC8, &HD304408B, &H52D00BE8, &HF23E8, &HE8458900, &HFE4753B, &HAE8D&, &HC458B00, &H401F0F, &H880FFF85, &H8E&, &H1BECF8B, &HD3000000
    pvAppendBuffer &H84D8BC6, &H8DF075FF, &HFF539904, &H5050FC75, &HFFE9EBE8, &HE075FFFF, &HFF1B048D, &H6AE875, &HF875FF53, &HFC75FF50, &HFFE713E8, &HC4D8BFF, &H452B018B, &H813485F4, &H75FF3374, &HFC458BF0, &H875FF53, &H8DD875FF, &HE8509804, &HFFFFE9B0, &H8DE075FF, &H75FF1B04, &H53006AE8, &H50F875FF, &HE80875FF, &HFFFFE6D8, &HEB084D8B, &H8458B0C, &H89FC4D8B, &H4589084D, &H83CED1FC, &H847901EF
    pvAppendBuffer &H8BF4758B, &HBF460C45, &H1F&, &H3BF47589, &H598C0F30, &H8BFFFFFF, &H3F8B107D, &H4BD348D, &HE8000000, &HFFFF91C8, &HFF008B56, &H6A56D0, &H10458950, &HFFEF77E8, &H10558BFF, &H890CC483, &H7EDB853A, &H8458B2A, &H9D3C8D, &H3000000, &H89F633C7, &HA8B0C45, &HCE2B008B, &H8A048946, &H830C458B, &H458904C0, &H7CF33B0C, &H8B03EBE9, &H3A83D47D, &H8758B01, &HF661676, &H441F&
    pvAppendBuffer &H3C83028B, &H8750082, &H83028948, &HF07701F8, &H1774F685, &H85DC4D8B, &H8B1074C9, &H1F0FC6, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9140, &H8408B56, &H5D8BD0FF, &H74DB85F0, &HD04D8B14, &HD74C985, &HC6C38B, &H1408D00, &H7501E983, &H911AE8F5, &H8B53FFFF, &HD0FF0840, &H85FC5D8B, &H8B1874DB, &HC985DC4D, &HC38B1174, &H401F0F, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF90F0
    pvAppendBuffer &H8408B53, &HFF85D0FF, &H458B1174, &H90CF8BF8, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF90D0, &H8BF875FF, &HD0FF0840, &H85D85D8B, &H851174DB, &H8B0D74FF, &HC6C3&, &H8301408D, &HF57501EF, &HFF90ABE8, &H408B53FF, &H8BD0FF08, &H68BEC75, &H4850C8D, &H85000000, &H8B1074C9, &H1F0FC6, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF9080, &H8408B56, &H458BD0FF, &H5B5E5F10, &HC25DE58B
    pvAppendBuffer &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H8B530855, &H8B56145D, &H8D571875, &H3C8D7604, &H53565783, &H8D1075FF, &H8950B204, &H19E81445, &H57FFFFEA, &H8D187D8B, &H5657B334, &H530C75FF, &HFFE7B7E8, &H85D8BFF, &H503F048D, &HE8535653, &HFFFFE418, &H4589C933, &H7EFF8518, &HBB348D1F, &H48BD68B, &H4528D8B, &HC7FC4289, &H8B04&, &H3B410000, &H8BEB7CCF, &H3EB1845
    pvAppendBuffer &H8514758B, &H853C75C0, &H8B387EFF, &HD68B0C5D, &HC3B0A8B, &H401A7583, &H3B04C283, &H57F17CC7, &HC75FF56, &HEB06E856, &H5E5FFFFF, &H14C25D5B, &H7DC73B00, &H8558B0E, &H8B380C8D, &HC3B8A0C, &H570B7683, &HC75FF56, &HEAE2E856, &H5E5FFFFF, &H14C25D5B, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &H100EC, &HC458B00, &H53C0570F, &H3CB95756, &H66000000, &H85130F, &H8DFFFFFF
    pvAppendBuffer &HFFFF00B5, &HFC45C7FF, &H10&, &HFF08BD8D, &HA5F3FFFF, &H8D104D8B, &HFFFF089D, &H10C183FF, &HC22BD38B, &H89F84D89, &HF660C45, &H441F&, &H45C7F98B, &H410&, &HFF38B00, &H441F&, &H41874FF, &HFF1834FF, &H77FFF477, &H9BFEE8F0, &H4601FFFF, &HC458BF8, &HFFFC5611, &HFF041874, &H77FF1834, &HF877FFFC, &HFF9BE3E8, &H8B0601FF, &H56110C45, &H1874FF04, &H1834FF04, &HFF0477FF
    pvAppendBuffer &H9BCAE837, &H4601FFFF, &HC458B08, &HFF0C5611, &HFF041874, &H77FF1834, &H877FF0C, &HFF9BAFE8, &H104601FF, &H8B207F8D, &H56110C45, &H20768D14, &H1106D83, &H4D8B8A75, &H8C383F8, &H1FC6D83, &HFF6A850F, &HF633FFFF, &H266A006A, &H84F574FF, &H80F574FF, &HFF9B77E8, &HF58401FF, &HFFFFFF00, &H9411006A, &HFFFF04F5, &HFF266AFF, &HFF8CF574, &HE888F574, &HFFFF9B58, &H8F58401, &H6AFFFFFF
    pvAppendBuffer &HF5941100, &HFFFFFF0C, &H74FF266A, &H74FF94F5, &H39E890F5, &H1FFFF9B, &HFF10F584, &H6AFFFF, &H14F59411, &H6AFFFFFF, &HF574FF26, &HF574FF9C, &H9B1AE898, &H8401FFFF, &HFFFF18F5, &H11006AFF, &HFF1CF594, &H266AFFFF, &HA4F574FF, &HA0F574FF, &HFF9AFBE8, &HF58401FF, &HFFFFFF20, &H24F59411, &H83FFFFFF, &HFE8305C6, &H59820F0F, &H8BFFFFFF, &HB58D085D, &HFFFFFF00, &H20B9&, &HF3FB8B00
    pvAppendBuffer &H69E853A5, &H53FFFFA4, &HFFA463E8, &H5B5E5FFF, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H565310EC, &H570C758B, &H6A187D8B, &H6A5600, &HE81475FF, &HFFFF9A94, &H6A56006A, &H45895700, &HE8DA8BF0, &HFFFF9A84, &H75FF006A, &HF4458910, &H6AF28B, &H9A72E857, &H6AFFFF, &H891075FF, &H6AFC45, &H891475FF, &H5DE8F855, &H8BFFFF9A, &HF4458BF8, &HD283FB03, &H13F80300
    pvAppendBuffer &H77D63BD6, &H3B04720E, &H830873F8, &H8300FC45, &H8B01F855, &HC9330845, &H89F04D0B, &H3C93308, &H7889FC55, &HF84D1304, &H50895E5F, &HC488908, &H5DE58B5B, &HCC0014C2, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H3356084D, &H32F983F6, &H9066147E, &H2B99C18B, &H2BF8D1C2, &H348D41C8, &H32F9838E, &HC68BEE7F, &H4C25D5E, &HCCCCCC00, &HCCCCCCCC, &H83EC8B55, &H565330EC, &H5708758B, &H570C7D8B
    pvAppendBuffer &H303AE856, &H206A0000, &HD0458D57, &H1AEEE850, &HD88B0000, &H8D0C5589, &HC683084E, &HD0458D38, &HE8515150, &H1568&, &H689C303, &H8B0C5513, &HC0830845, &H4568910, &HE8505057, &H1550&, &H89087D8B, &H458D4047, &H575750D0, &HE8445789, &H33CC&, &H30C4D8B, &H8BCA13D8, &HD32B3057, &H1B345F8B, &H345F3BD9, &H5772D72, &H7630573B, &HFF068326, &H5683068B, &H4623FF04, &HFFF88304
    pvAppendBuffer &H46831575, &H768DFF08, &H4568308, &H23068BFF, &HF8830446, &H89EB74FF, &H5789345F, &H5B5E5F30, &HC25DE58B, &HCCCC0008, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &H108EC, &H78858D00, &H53FFFFFF, &H75FF5756, &H65E8500C, &H8D00000B, &HFFFF7885, &H99E850FF, &H8DFFFFA2, &HFFFF7885, &H8DE850FF, &H8DFFFFA2, &HFFFF7885, &H81E850FF, &H8DFFFFA2, &HFFFEF8BD, &H2BBFF, &HF660000, &H441F&
    pvAppendBuffer &HFF788D8B, &H858BFFFF, &HFFFFFF7C, &HFFEDE981, &H8D890000, &HFFFFFEF8, &H8900D883, &HFFFEFC85, &H8B8FF, &H66660000, &H841F0F, &H0&, &HF807748B, &HFC074C8B, &H7805948B, &H89FFFFFF, &HAC0FF875, &H8C8B10CE, &HFFFF7C05, &H1E683FF, &HFC0744C7, &H0&, &HD983D62B, &HFFEA8100, &H890000FF, &HFEF80594, &HD983FFFF, &H58C8900, &HFFFFFEFC, &HF84DB70F, &HF8074C89, &H8308C083
    pvAppendBuffer &HAC7278F8, &HFF688D8B, &H858BFFFF, &HFFFFFF6C, &HFF0558B, &HF10C1AC, &HFF6885B7, &HE183FFFF, &H68858901, &H2BFFFFFF, &H6C85C7D1, &HFFFFFF, &H8B000000, &H1B8F44D, &H83000000, &HEA8100D9, &H7FFF&, &HFF709589, &HD983FFFF, &H748D8900, &HFFFFFFF, &H8310CAAC, &HF9C101E2, &H50C22B10, &HFEF8858D, &H8D50FFFF, &HFFFF7885, &H4DE850FF, &H83000009, &H850F01EB, &HFFFFFF04, &H3308758B
    pvAppendBuffer &HD5848AD2, &HFFFFFF78, &H78D58C8B, &H88FFFFFF, &H848B5604, &HFFFF7CD5, &HC1AC0FFF, &H564C8808, &HF8C14201, &H10FA8308, &H5E5FD772, &H5DE58B5B, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HD2330845, &H7D8B5756, &H8DF82B0C, &HC8B1172, &H4408D07, &H3FC4803, &HCAB60FD1, &H8908EAC1, &HEE83FC48, &H5FE77501, &H8C25D5E, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H56EC8B55, &H8B0C75FF, &HE8560875, &HFFFFFFB0, &H5044468D, &H1B6E856, &H5D5E0000, &HCC0008C2, &H83EC8B55, &H565344EC, &H5708758B, &H8B06100F, &H45894046, &H45110FFC, &H46100FBC, &H45110F10, &H46100FCC, &H45110F20, &H46100FDC, &H45110F30, &H8A1AE8EC, &H4405FFFF, &H50000005, &H50BC458D, &HFFFF5BE8, &HFC458BFF, &HF7BC7D8D, &HCC558DD0, &H8025&, &HB9FE2B00, &H2&, &HF7FF588D
    pvAppendBuffer &H1FE8C1D0, &H231FEBC1, &H2BDBF7D8, &HF7C38BD6, &H84589D0, &HC36E0F66, &HD0700F66, &H6E0F6600, &H66C68BC0, &HD8700F, &H841F0F, &H0&, &HF20408D, &HFE04010, &HE0074C10, &HC2DB0F66, &HCBDB0F66, &HC8EB0F66, &HE048110F, &HF040100F, &H24C100F, &HDB0F66E0, &HDB0F66C2, &HEB0F66CB, &H48110FC8, &H1E983F0, &H568DC675, &H1718D40, &H8D3A0C8B, &H4D230452, &H23C38B08, &HC80BFC42
    pvAppendBuffer &H83FC4A89, &HE87501EE, &H8B5B5E5F, &H4C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D44EC, &H446A56BC, &HE850006A, &HFFFFE6FC, &H8308758B, &HC0330CC4, &HA8968B, &HD2850000, &H66661B74, &H841F0F, &H0&, &H68CB60F, &H98&, &HBC854C89, &H72C23B40, &HBC458DEF, &HBC9544C7, &H1&, &H8DE85650, &H5EFFFFFE, &HC25DE58B, &HCCCC0004, &HCCCCCCCC
    pvAppendBuffer &H56EC8B55, &H3308758B, &HFD233C0, &H441F&, &HF960403, &HC89C8B6, &HE8C14296, &H10FA8308, &H4603EE7C, &HC1C88B40, &HE18302E8, &H89D23303, &HC8D404E, &H960C0380, &H89C1B60F, &HC1429604, &HFA8308E9, &H1EE7C10, &H5D5E404E, &HCC0004C2, &H83EC8B55, &H458B54EC, &HAC4D8D0C, &H758B5653, &H2BDB3308, &HF845C7C1, &H10&, &HF0458957, &HFF33D233, &H5589C033, &HFC558908, &H5178DB85
    pvAppendBuffer &H83014B8D, &H307C02F9, &H8DF04D8B, &HC8DAC55, &H8BD10399, &H528D860C, &H4AAF0FF8, &H84D0108, &H4864C8B, &HF02C083, &H1044AAF, &H4B8DFC4D, &H7EC13BFF, &H8558BDE, &HE7FC33B, &H8B0C7D8B, &H8BC82BCB, &HAF0F8F3C, &H458B863C, &H3C203FC, &H1438DF8, &H5589D233, &H89C88B08, &H4589FC55, &H11F883F4, &H7D83727D, &H437C02F8, &H8B0C558B, &H8DC12BC3, &HC2838214, &H801F0F40, &H0&
    pvAppendBuffer &H8D8E048B, &HAF0FF852, &H48D0C42, &H6E0C180, &H8B084501, &H83048E44, &HAF0F02C1, &H48D0842, &H6E0C180, &H83FC4501, &HD47C10F9, &H8308558B, &H1A7D11F9, &H8B0C558B, &H8BC12BC3, &HF448244, &H8B8E04AF, &H48D0855, &H6E0C180, &H458BF803, &H3C203FC, &HF4458BF8, &H49F84D8B, &HAC9D7C89, &H8BF84D89, &HFFF983D8, &HFF028F0F, &H458DFFFF, &H89E850AC, &HFFFFFFE, &H8BAC4510, &HF5FEC45
    pvAppendBuffer &H100F0611, &H110FBC45, &H100F1046, &H110FCC45, &H100F2046, &H110FDC45, &H46893046, &H8B5B5E40, &H8C25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HEC830C55, &HFC03344, &H441F&, &H100CB60F, &HBC854C89, &H10F88340, &H458DF27C, &HFC45C7BC, &H1&, &H875FF50, &HFFFC9FE8, &H5DE58BFF, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &H17CEC, &H57565300, &H75FF0C6A
    pvAppendBuffer &HE0458D0C, &HDC45C6, &HC7C0570F, &HE545&, &H66500000, &HDD45D60F, &HE945C766, &H45C60000, &H59E800EB, &H83FFFFE4, &H45C60CC4, &H458D00BC, &HD545C7DC, &H0&, &H66C0570F, &HD945C7, &H45110F00, &H50046ABD, &H75FF206A, &H30858D08, &H66FFFFFF, &HCD45D60F, &HDB45C650, &HABFEE800, &H206AFFFF, &H50BC458D, &H30858D50, &H50FFFFFF, &HFFA54BE8, &HCC458DFF, &HBC458D50, &H84858D50
    pvAppendBuffer &H50FFFFFE, &HFFB737E8, &HC0570FFF, &HBC45110F, &H6ABC458A, &HBC458D20, &H858D5050, &HFFFFFF30, &H45110F50, &HA516E8CC, &H758BFFFF, &HC0570F14, &H1075FF56, &HBC45110F, &H8DBC458A, &HFFFE8485, &HEC45C6FF, &H110F5000, &H45C7CC45, &HF5&, &HD60F6600, &HC766ED45, &HF945&, &HFB45C6, &HFFB7ABE8, &HF7C68BFF, &HFE083D8, &HEC458D50, &H84858D50, &H50FFFFFE, &HFFB793E8, &H247D83FF
    pvAppendBuffer &H207D8B01, &H531C5D8B, &HFF571475, &H858D1875, &HFFFFFF30, &HA4A6E850, &H5753FFFF, &H75FF03EB, &H84858D18, &H50FFFFFE, &HFFB763E8, &HF7C38BFF, &HFE083D8, &HEC458D50, &H84858D50, &H50FFFFFE, &HFFB74BE8, &H88D233FF, &HC68BF45D, &H88E85589, &HC88BEC45, &HAC0FC28B, &H106A08C1, &H8808E8C1, &HC28BED4D, &HAC0FCE8B, &HE8C110C1, &HEE4D8810, &HCE8BC28B, &H18C1AC0F, &HF18E8C1, &H4588C2B6
    pvAppendBuffer &HC1C28BF0, &H458808E8, &HC1C28BF1, &H458810E8, &H18EAC1F2, &H8BEF4D88, &HF35588CB, &HC28BD233, &HFE85589, &HC108C1AC, &H4D8808E8, &H8BC28BF5, &HC1AC0FCB, &H10E8C110, &H8BF64D88, &HFCB8BC2, &HC118C1AC, &HB60F18E8, &HF84588C2, &HE8C1C28B, &HF9458808, &HE8C1C28B, &HFA458810, &H50EC458D, &HFE84858D, &HEAC1FFFF, &H4D885018, &HFB5588F7, &HFFB69BE8, &H247D83FF, &HFF337501, &H858D2875
    pvAppendBuffer &HFFFFFE84, &HB536E850, &H7C6AFFFF, &HFF30858D, &H6AFFFF, &HE296E850, &H858AFFFF, &HFFFFFF30, &H330CC483, &H5B5E5FC0, &HC25DE58B, &H458D0024, &H858D50AC, &HFFFFFE84, &HB502E850, &H758BFFFF, &HAC4D8D28, &HDB32C18B, &H10BA&, &H90F02B00, &H8D0E048A, &H41320149, &H83D80AFF, &HF07501EA, &H841C458B, &H503F75DB, &H1875FF57, &HFF30858D, &HE850FFFF, &HFFFFA348, &H858D7C6A, &HFFFFFF30
    pvAppendBuffer &HE850006A, &HFFFFE228, &HFF30858A, &HC483FFFF, &HC0570F0C, &HAC45110F, &H5FAC458A, &H5BC0335E, &HC25DE58B, &HC0850024, &H6A500E74, &HFDE85700, &H8AFFFFE1, &HCC48307, &H858D7C6A, &HFFFFFF30, &HE850006A, &HFFFFE1E8, &HFF30858A, &HC483FFFF, &HC0570F0C, &HAC45110F, &H5FAC458A, &H1B85E, &H8B5B0000, &H24C25DE5, &HCCCCCC00, &HCCCCCCCC, &H56EC8B55, &H87D8B57, &H9907B60F, &HF28BC88B
    pvAppendBuffer &H147B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H247B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H347B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H447B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H547B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H647B60F, &H8CEA40F, &H8E1C199, &HC80BF20B, &H747B60F, &H8CEA40F, &H8E1C199, &HD60BC10B, &HC25D5E5F, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H53EC8B55, &H33085D8B, &HC1CB8BD2, &H564110E9, &HFF818D57, &HF77FFFFF, &HC1F08BF1, &HC68B10E6, &HC88BE3F7, &HD1F7CE03, &H3300D283, &H1C183C0, &HC013D2F7, &HE6F7C203, &H8B1FE8C1, &HBF603F2, &HF7C68BF0, &H8BC603E3, &HD783FA, &HFF81&, &H23728000, &HC933D3F7, &HC913C303, &H8301C083, &H34F00D1, &HFF814EF9, &H80000000, &H8B5FE873, &H5D5B5EC6, &H330004C2, &H13C303D2, &H81D703D2
    pvAppendBuffer &HFA&, &H33117380, &HC30346C9, &HD103C913, &HFA81&, &HEF728000, &H5EC68B5F, &H4C25D5B, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H53EC8B55, &H8758B56, &H75FF5657, &H978EE80C, &HFF56FFFF, &HF88B1075, &HE8087D89, &HFFFF9780, &H1475FF56, &H5D89D88B, &H9772E80C, &H5350FFFF, &H10458957, &HFFE997E8, &H85D88BFF, &H8B1774F6, &H564E187D, &H9676E853, &H788FFFF, &H85017F8D
    pvAppendBuffer &H8BEF75F6, &H78B087D, &H4850C8D, &H85000000, &H8B0D74C9, &HC6C7&, &H8301408D, &HF57501E9, &HFF824BE8, &H408B57FF, &H8BD0FF08, &H68B0C75, &H4850C8D, &H85000000, &H8B1074C9, &H1F0FC6, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF8220, &H8408B56, &H758BD0FF, &H8D068B10, &H4850C, &HC9850000, &HC68B1574, &H841F0F, &H0&, &H8D0000C6, &HE9830140, &HE8F57501, &HFFFF81F0
    pvAppendBuffer &H8408B56, &H38BD0FF, &H4850C8D, &H85000000, &H8B0D74C9, &HC6C3&, &H8301408D, &HF57501E9, &HFF81CBE8, &H408B53FF, &H5FD0FF08, &HC25D5B5E, &HCCCC0014, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458B08EC, &HD0F74810, &H5D8B5399, &HF8458908, &H890C458B, &HFF3FC55, &H8DF85D7E, &H3356784B, &H6C0F66F6, &H78508DDB, &H4B77C13B, &H4772D33B, &H45C7D82B, &H1010&, &H90665700
    pvAppendBuffer &H8D183C8B, &H748B0840, &H488BFC18, &HFC508BF8, &H4D23CF33, &H23D633F8, &HF933FC55, &H7C89F233, &H7489F818, &H4831FC18, &HFC5031F8, &H1106D83, &H5E5FCE75, &H5DE58B5B, &H8B000CC2, &H10488DD3, &H100FD02B, &H498DF30C, &H51100F20, &HEF0F66D0, &HDB0F66D1, &HC2280FD3, &HC1EF0F66, &HF304110F, &HF04C683, &H66D04110, &HFD0EF0F, &HFD05111, &HE00A4C10, &HE051100F, &HD1EF0F66, &HD3DB0F66
    pvAppendBuffer &H66C2280F, &HFC1EF0F, &HE00A4411, &HE041100F, &HC2EF0F66, &HE041110F, &H7210FE83, &H8B5B5EA5, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H458B0C55, &H56D02B08, &H10BE&, &H20C8B00, &H8908408D, &H4C8BF848, &H4889FC02, &H1EE83FC, &H5D5EEB75, &HCC0008C2, &HCCCCCCCC, &H8BEC8B55, &H57561045, &H7410F883, &H20F88339, &H758B5F75, &H87D8B0C, &H5756106A, &HFFDDFFE8
    pvAppendBuffer &H8D106AFF, &H8D501046, &HE8501047, &HFFFFDDF0, &HE818C483, &HFFFF8058, &H53105, &H30478900, &HC25D5E5F, &H758B000C, &H87D8B0C, &H5756106A, &HFFDDCBE8, &H8D106AFF, &H50561047, &HFFDDBFE8, &H18C483FF, &HFF8027E8, &H52005FF, &H47890000, &H5D5E5F30, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458B6CEC, &H94558D08, &HA0BB5653, &H33000001, &H4488BF6, &H8BF84D89, &H4D890848
    pvAppendBuffer &HC488BF4, &H8BE84D89, &H4D891048, &H14488BFC, &H8BF04D89, &H4D891848, &HC4D8BEC, &H8902C183, &H8B57DC75, &H8BD32B38, &H7D891C40, &HE44589E0, &H89D84D89, &H55890C5D, &H801F0FD4, &H0&, &H7310FE83, &H71B60F29, &H41B60FFE, &H8E6C1FF, &HB60FF00B, &H8E6C101, &HB60FF00B, &HE6C10141, &H83F00B08, &H348904C1, &HD84D891A, &H5E8D54EB, &HFE68301, &H83FD438D, &H7D8D0FE0, &HB73C8D94
    pvAppendBuffer &H94854C8B, &HE083C38B, &HC1F18B0F, &H548B0FC6, &HC18B9485, &H330DC0C1, &HAE9C1F0, &HC28BF133, &HC8C1CA8B, &HEC1C107, &HEAC1C833, &HF8438D03, &H5D8BCA33, &HFE0830C, &H7403F103, &H37039485, &H29E83789, &H8BFFFF7F, &HD78BFC7D, &H8B0BCAC1, &H7C1C1CF, &HCF8BD133, &HF706C9C1, &HEC7D23D7, &HC8BD133, &H4C38318, &H3F0458B, &HFC4523CA, &H758BCE03, &H8BF833E0, &HC5D89D6, &H8B0DCAC1
    pvAppendBuffer &HAC0C1C6, &H7D03F903, &H8BD033E4, &H2C8C1C6, &H458BD033, &H23C88BF8, &H23CE33C6, &HC833F44D, &H89EC458B, &HD103E445, &H8BF0458B, &H4589F84D, &HFC458BEC, &H8BF04589, &HC703E845, &H8BF87589, &HFA03DC75, &H46D4558B, &H8BFC4589, &H4D89F445, &HD84D8BF4, &H89E84589, &H7589E07D, &HA0FB81DC, &HF000002, &HFFFED782, &H8458BFF, &H8BF84D8B, &H4801FC55, &HF44D8B04, &H1084801, &H38011050
    pvAppendBuffer &H8BE84D8B, &H4801F055, &H1450010C, &H8BEC558B, &H5001E44D, &H1C480118, &H5F6040FF, &HE58B5B5E, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &HE0EC&, &H8B565300, &HA0BB0875, &H57000002, &H8BB85D89, &HEC458906, &H8904468B, &H468BF045, &H87E8B0C, &H8BE04589, &H45891046, &H14468BD4, &H8BD04589, &H45891846, &H1C468BB4, &H8BB04589, &H45892046, &H24468BE8, &H8BF44589
    pvAppendBuffer &H45892846, &H2C468BCC, &H8BC84589, &H45893046, &H34468BC4, &H8BC04589, &H45893846, &H3C468BAC, &H890C758B, &HBD8DD87D, &HFFFFFF20, &H33A84589, &H89FB2BC0, &H7D89DC45, &H801F0FA0, &H0&, &H7310F883, &HA5E8561F, &H8BFFFFF9, &H8C683C8, &H4D89C28B, &HE445890C, &H891F0C89, &HE9041F44, &H113&, &HC701508D, &HC45&, &H428D0000, &HFE083FD, &H20C58C8B, &H8BFFFFFF, &HFF24C584
    pvAppendBuffer &H4589FFFF, &H83C28BF8, &H4D890FE0, &H208D8DFC, &H8BFFFFFF, &HFF20C594, &HFA8BFFFF, &H24C59C8B, &H8BFFFFFF, &HE083DC45, &HBC55890F, &H8D18E7C1, &HCB8BC104, &H8BA44589, &HC8AC0FC2, &HC450908, &HC1BC458B, &HF90B08E9, &HAC0FCB8B, &H7D8901C8, &HD1FA8BE4, &HBD233E9, &H1FE7C1D0, &HB0C5531, &HBC458BF9, &HFE44D8B, &H3307D8AC, &HC4531CF, &HC1FC458B, &HCB3307EB, &H4D89DB33, &HF84D8BE4
    pvAppendBuffer &HA40FD18B, &HEAC103C1, &H3E0C11D, &H4D8BD90B, &H8BD00BF8, &HF88BFC45, &H13C8AC0F, &H33BC5589, &HC1D00BD2, &H458B13E9, &HC1C233BC, &H558B0DE7, &H8BF90BFC, &HDF33F84D, &H6CAAC0F, &HE9C1C233, &HC558B06, &H4D8BD933, &H8BD003E4, &HCB13DC45, &H83F9C083, &H94030FE0, &HFFFF20C5, &HC58C13FF, &HFFFFFF24, &H3A4458B, &HC558910, &H89044813, &HE44D8910, &HE8044889, &HFFFF7C74, &H33F4558B
    pvAppendBuffer &HE84D8BFF, &HA40FDA8B, &HEBC117CA, &HC1FA0B09, &H558B17E1, &H8BD90BF4, &H5D89E84D, &HFD98BFC, &H8912D1AC, &HFF33F87D, &HEAC1F90B, &HFC7D3112, &H4D8BFF33, &HEE3C1E8, &H558BDA0B, &HF85D31F4, &HAC0FD98B, &HE3C10ED1, &HC1F90B12, &H7D310EEA, &H8BDA0BFC, &H558BF84D, &H8BCB33B8, &H7D8BFC5D, &H3D7F7E8, &H4C13101C, &H7D230410, &HF4558BC4, &HF7C8458B, &HF44523D2, &H33C05523, &HF84D89D0
    pvAppendBuffer &H23CC4D8B, &H458BE84D, &H8BF933F8, &HDF03F04D, &H5D03C213, &HE445130C, &H89AC5D03, &H4513FC5D, &H89DB33A8, &H458BF845, &HFD08BEC, &HC11CC8AC, &HE9C104E2, &H8BD80B1C, &HD10BEC45, &H8BF04D8B, &HC1A40FF9, &HC55891E, &HEFC1D233, &HC1D10B02, &HF80B1EE0, &H5531DF33, &H8BD2330C, &HF98BF04D, &HFEC458B, &HC119C1A4, &HD10B07EF, &H3119E0C1, &HF80B0C55, &H33D84D8B, &HE0558BDF, &H7D33F98B
    pvAppendBuffer &HD47D23EC, &H33EC4D23, &HF933F055, &H8BD05523, &H4523E045, &HC44D8BF0, &H458BD033, &H8BDF030C, &HC213F87D, &H8BAC4D89, &H558BC04D, &HB45503FC, &H13A84D89, &H4D8BB07D, &HFC5D03CC, &H8BC44D89, &H4D89C84D, &HE84D8BC0, &H8BCC4D89, &H7D89F44D, &HD47D8BF4, &H8BB47D89, &H7D89D07D, &HD87D8BB0, &H8BD47D89, &H7D89E07D, &HEC7D8BD0, &H8BC84D89, &HF84D13C8, &H89DC458B, &H8B40EC5D, &H7D89B85D
    pvAppendBuffer &H8C383D8, &H89F07D8B, &H7D8BE07D, &HE85589A0, &H89F04D89, &H5D89DC45, &H20FB81B8, &HF000005, &HFFFD1B82, &H8758BFF, &H8BEC458B, &H601D87D, &H11E0458B, &HCA8B044E, &H8B087E01, &H4611B47D, &HD4458B0C, &H8B104601, &H4611D045, &H187E0114, &H11B0458B, &H4E011C46, &HF4458B20, &H8B244611, &H4601CC45, &HC8458B28, &H8B2C4611, &H4601C445, &HC0458B30, &H8B344611, &H4E01AC4D, &HA84D8B38
    pvAppendBuffer &HFF3C4E11, &HC086&, &H5B5E5F00, &HC25DE58B, &HCCCC0008, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H53EC8B55, &H56085D8B, &H7BB60F57, &H43B60F07, &H73B60F0A, &H53B60F0B, &H8E7C10F, &HB60FF80B, &HB60F034B, &HE7C10D43, &HC1F80B08, &HB60F08E6, &H8E7C103, &HE2C1F80B, &H43B60F08, &HC1F00B0E, &HB60F08E1, &HE6C10143, &HFF00B08, &HC10443B6, &HF00B08E6, &H243B60F, &HB60FD00B, &HE2C10543
    pvAppendBuffer &HFD00B08, &HC10843B6, &HD00B08E2, &H643B60F, &H7B89C80B, &H43B60F04, &H8E1C109, &H7389C80B, &H43B60F08, &H8E1C10C, &H89C80B5F, &H895E0C53, &HC25D5B0B, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H50500C45, &HE80875FF, &HFFFFEA10, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H56531045, &H8D08758B, &H8B577848, &H568D0C7D, &H77F13B78, &H73D03B04, &H784F8D0B
    pvAppendBuffer &H3077F13B, &H2C72D73B, &H10BBF82B, &H2B000000, &H38148BF0, &H4C8B102B, &H481B0438, &H8408D04, &HF8305489, &HFC304C89, &H7501EB83, &H5B5E5FE4, &HCC25D, &H488DD78B, &H2BDE8B10, &H2BD82BD0, &H4B8FE, &H768D0000, &H20498D20, &HD041100F, &H374C100F, &HFB0F66E0, &H4E110FC8, &H4C100FE0, &H100FE00A, &HF66E041, &H110FC8FB, &H83E00B4C, &HD27501E8, &H5D5B5E5F, &HCC000CC2, &HCCCCCCCC
    pvAppendBuffer &H56EC8B55, &HFF78E7E8, &H8758BFF, &H58805, &H36FF5000, &H57E8&, &HE8068900, &HFFFF78D0, &H58805, &H76FF5000, &H42E804, &H46890000, &H78BAE804, &H8805FFFF, &H50000005, &HE80876FF, &H2C&, &HE8084689, &HFFFF78A4, &H58805, &H76FF5000, &H16E80C, &H46890000, &HC25D5E0C, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H8B530C55, &HC38B085D, &H8B18E8C1
    pvAppendBuffer &HE9C156CB, &HC9B60F08, &H1034B60F, &HE8C1C38B, &HC0B60F10, &H110CB60F, &HF08E6C1, &HB1004B6, &H8E0C1C6, &HB60FC10B, &H8E0C1CB, &HB60F5B5E, &HC10B110C, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H8B530C4D, &H8356085D, &H45C710C3, &H40C&, &HC1835700, &H801F0F03, &H0&, &HFE41B60F, &H99205B8D, &H8B08498D, &HFFA8BF0, &HFF541B6, &H9908F7A4, &H308E6C1, &HD07389F0
    pvAppendBuffer &H7B89FA13, &H41B60FD4, &HF08B99F7, &HB60FFA8B, &HF99F841, &HC108C2A4, &HF00308E0, &H13D87389, &HDC7B89FA, &HFA41B60F, &H8BF08B99, &H41B60FFA, &HF7A40FF9, &HE6C19908, &H89F00308, &HFA13E073, &HFE47B89, &H99FC41B6, &HFA8BF08B, &HFB41B60F, &H8F7A40F, &H8E6C199, &H7389F003, &H83FA13E8, &H89010C6D, &H850FEC7B, &HFFFFFF74, &H5F084D8B, &H61815B5E, &H7FFF78, &H7C41C700, &H0&
    pvAppendBuffer &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H8B530CEC, &H570F0C5D, &H8B5756C0, &H138B107D, &H438BF28B, &H66C88B04, &HF445130F, &H4F133703, &H75F23B04, &H75C83B06, &H3B18EB04, &H720F77C8, &H73F23B04, &H1B809, &HD2330000, &HF660BEB, &H8BF44513, &H558BF445, &H87D8BF8, &H89044F89, &H84B8B37, &H8910758B, &H4B8BFC4D, &HF84D890C, &H3084E8B, &H4D89FC4D, &HC4E8B08
    pvAppendBuffer &H8BF84D13, &HD803085D, &H13085D89, &HFC5D3BCA, &H750C5D8B, &HC4B3B05, &H4B3B2374, &H7213770C, &H8438B08, &H73084539, &H1B809, &HD2330000, &HF660BEB, &H8BF44513, &H458BF855, &H8758BF4, &H890C4F89, &H4B8B0877, &H10758B10, &H8BFC4D89, &H4D89144B, &H104E8BF8, &H89FC4D03, &H4E8B084D, &HF84D1314, &H3085D8B, &H85D89D8, &H5D3BCA13, &HC5D8BFC, &H4B3B0575, &H3B237414, &H1377144B
    pvAppendBuffer &H438B0872, &H8453910, &H1B80973, &H33000000, &H660BEBD2, &HF445130F, &H8BF8558B, &H758BF445, &H144F8908, &H8B107789, &H758B184B, &HFC4D8910, &H891C4B8B, &H4E8BF84D, &HFC4D0318, &H8B084D89, &H4D131C4E, &H85D8BF8, &H5D89D803, &H3BCA1308, &H5D8BFC5D, &H3B05750C, &H23741C4B, &H771C4B3B, &H8B087213, &H45391843, &HB8097308, &H1&, &HBEBD233, &H45130F66, &HF8558BF4, &H8BF4458B
    pvAppendBuffer &H4F890875, &H1877891C, &H8B204B8B, &H4D891075, &H244B8BFC, &H8BF84D89, &H4D03204E, &H84D89FC, &H13244E8B, &H5D8BF84D, &H89D80308, &HCA13085D, &H8BFC5D3B, &H5750C5D, &H74244B3B, &H244B3B23, &H8721377, &H3920438B, &H9730845, &H1B8&, &HEBD23300, &H130F660B, &H558BF445, &HF4458BF8, &H8908758B, &H758B2077, &H244F8910, &H8B284B8B, &H768B2C5B, &H89F10328, &H4D8B0C4D, &H2C498B10
    pvAppendBuffer &HF003CB13, &H753BCA13, &H3B04750C, &H3B2C74CB, &H721D77CB, &HC753B05, &H77891673, &H1B828, &H4F890000, &H5FD2332C, &HE58B5B5E, &HCC25D, &H45130F66, &HF8558BF4, &H89F4458B, &H4F892877, &H5B5E5F2C, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &H83EC8B55, &H8B5308EC, &H570F0C5D, &H8B5756C0, &H138B107D, &H438BF28B, &H66C88B04, &HF845130F, &H4F133703, &H75F23B04, &H75C83B06, &H3B18EB04
    pvAppendBuffer &H720F77C8, &H73F23B04, &H1B809, &HD2330000, &HF660BEB, &H8BF84513, &H558BF845, &H87D8BFC, &H8B044F89, &H3789104D, &H308718B, &H498B0873, &HC4B130C, &HCA13F003, &H7508733B, &HC4B3B05, &H4B3B2074, &H7210770C, &H8733B05, &H1B80973, &H33000000, &H660BEBD2, &HF845130F, &H8BFC558B, &H4F89F845, &H104D8B0C, &H8B087789, &H73031071, &H14498B10, &H3144B13, &H3BCA13F0, &H5751073
    pvAppendBuffer &H74144B3B, &H144B3B20, &H5721077, &H7310733B, &H1B809, &HD2330000, &HF660BEB, &H8BF84513, &H458BFC55, &H144F89F8, &H8B107789, &H5B8B184B, &HC4D891C, &H8B104D8B, &H75031871, &H1C498B0C, &HF003CB13, &H753BCA13, &H3B04750C, &H3B2C74CB, &H721D77CB, &HC753B05, &H77891673, &H1B818, &H4F890000, &H5FD2331C, &HE58B5B5E, &HCC25D, &H45130F66, &HFC558BF8, &H89F8458B, &H4F891877
    pvAppendBuffer &H5B5E5F1C, &HC25DE58B, &HCCCC000C, &HCCCCCCCC, &H8BEC8B55, &H1C7084D, &H0&, &H441C7, &H8B000000, &H8418901, &H8904418B, &H418B0C41, &H10418908, &H890C418B, &H418B1441, &H18418910, &H8914418B, &H418B1C41, &H20418918, &H891C418B, &H418B2441, &H28418920, &H8924418B, &HC25D2C41, &HCCCC0004, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HC70845, &H0&, &H440C7
    pvAppendBuffer &HC7000000, &H840&, &H40C70000, &HC&, &H1040C700, &H0&, &H1440C7, &HC7000000, &H1840&, &H40C70000, &H1C&, &H4C25D00, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H5BA0C4D, &H53000000, &H56085D8B, &H418DD92B, &H5D895728, &H801F0F08, &H0&, &H8B03348B, &H8B04035C, &H88B0478, &H2E77DF3B, &HF13B2272, &HDF3B2877, &H4771A72, &H1472F13B
    pvAppendBuffer &H83085D8B, &HEA8308E8, &H5FD57901, &H5BC0335E, &H8C25D, &HC8835E5F, &HC25D5BFF, &H5E5F0008, &H1B8&, &HC25D5B00, &HCCCC0008, &HCCCCCCCC, &H8BEC8B55, &H3BA0C4D, &H53000000, &H56085D8B, &H418DD92B, &H5D895718, &H801F0F08, &H0&, &H8B03348B, &H8B04035C, &H88B0478, &H2E77DF3B, &HF13B2272, &HDF3B2877, &H4771A72, &H1472F13B, &H83085D8B, &HEA8308E8, &H5FD57901, &H5BC0335E
    pvAppendBuffer &H8C25D, &HC8835E5F, &HC25D5BFF, &H5E5F0008, &H1B8&, &HC25D5B00, &HCCCC0008, &HCCCCCCCC, &H8BEC8B55, &HC0330855, &H841F0F, &H0&, &HBC20C8B, &H7504C24C, &HF883400F, &HB8F17206, &H1&, &H4C25D, &HC25DC033, &HCCCC0004, &H8BEC8B55, &HC0330855, &H841F0F, &H0&, &HBC20C8B, &H7504C24C, &HF883400F, &HB8F17204, &H1&, &H4C25D, &HC25DC033, &HCCCC0004
    pvAppendBuffer &H83EC8B55, &H8B5310EC, &H40B9105D, &H56000000, &H2B08758B, &H7D8B57CB, &H6E0F660C, &H104D89C3, &H578B078B, &HF8458904, &HF3FC5589, &HF84D7E0F, &HC8F30F66, &HED60F66, &HFF7EE3E8, &H104D8BFF, &H8BF04589, &H55890847, &HC578BF4, &H89F84589, &HFF3FC55, &H66F84D7E, &H66C36E0F, &HF3C8F30F, &HF0457E0F, &HC8EB0F66, &H4ED60F66, &H7EAEE808, &H4D8BFFFF, &HF0458910, &H8910478B, &H578BF455
    pvAppendBuffer &HF8458914, &HF3FC5589, &HF84D7E0F, &HC36E0F66, &HC8F30F66, &H457E0FF3, &HEB0F66F0, &HD60F66C8, &H79E8104E, &H8BFFFF7E, &H4589104D, &H18478BF0, &H8BF45589, &H45891C57, &HFC5589F8, &H4D7E0FF3, &H6E0F66F8, &HF30F66C3, &H7E0FF3C8, &HF66F045, &HF66C8EB, &HE8184ED6, &HFFFF7E44, &H89104D8B, &H478BF045, &HF4558920, &H8924578B, &H5589F845, &H7E0FF3FC, &HF66F84D, &HF66C36E, &HFF3C8F3
    pvAppendBuffer &H66F0457E, &H66C8EB0F, &H204ED60F, &HFF7E0FE8, &HF04589FF, &H8928478B, &H578BF455, &HF845892C, &HF3FC5589, &HF84D7E0F, &H66104D8B, &H66C36E0F, &HF3C8F30F, &HF0457E0F, &HC8EB0F66, &H4ED60F66, &H7DDAE828, &H5E5FFFFF, &H5DE58B5B, &HCC000CC2, &H83EC8B55, &H8B5310EC, &H40B9105D, &H56000000, &H2B08758B, &H7D8B57CB, &H6E0F660C, &H104D89C3, &H578B078B, &HF8458904, &HF3FC5589, &HF84D7E0F
    pvAppendBuffer &HC8F30F66, &HED60F66, &HFF7D93E8, &H104D8BFF, &H8BF04589, &H55890847, &HC578BF4, &H89F84589, &HFF3FC55, &H66F84D7E, &H66C36E0F, &HF3C8F30F, &HF0457E0F, &HC8EB0F66, &H4ED60F66, &H7D5EE808, &H4D8BFFFF, &HF0458910, &H8910478B, &H578BF455, &HF8458914, &HF3FC5589, &HF84D7E0F, &HC36E0F66, &HC8F30F66, &H457E0FF3, &HEB0F66F0, &HD60F66C8, &H29E8104E, &H8BFFFF7D, &H4589104D, &H18478BF0
    pvAppendBuffer &H8BF45589, &H45891C57, &HFC5589F8, &H4D7E0FF3, &H6E0F66F8, &HF30F66C3, &H7E0FF3C8, &HF66F045, &HF66C8EB, &HE8184ED6, &HFFFF7CF4, &H8B5B5E5F, &HCC25DE5, &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H565368EC, &H8D0C758B, &HE853305E, &HFFFFFD4C, &H850FC085, &H264&, &H1F0F57, &HF98458D, &H6650C057, &HF845130F, &HFFFB9FE8, &HC8458DFF, &HFB96E850, &H8D53FFFF, &HE8509845
    pvAppendBuffer &HFFFFE26C, &HFB86E853, &H168BFFFF, &H7D03FA8B, &H4468B98, &H4D13C88B, &H75FA3B9C, &H75C83B06, &H3B1BEB04, &H720F77C8, &H73FA3B04, &H1B809, &HD2330000, &H570F0EEB, &H130F66C0, &H458BF845, &HFC558BF8, &H890C5E8B, &H87E8B3E, &H89A07D03, &HCB8B044E, &H3A44D13, &H3BCA13F8, &H475087E, &H2274CB3B, &H1077CB3B, &H7E3B0572, &HB8097308, &H1&, &HEEBD233, &H66C0570F, &HF845130F
    pvAppendBuffer &H8BFC558B, &H5E8BF845, &H87E8914, &H3107E8B, &H4E89A87D, &H13CB8B0C, &HF803AC4D, &H7E3BCA13, &H3B047510, &H3B2274CB, &H721077CB, &H107E3B05, &H1B80973, &H33000000, &HF0EEBD2, &HF66C057, &H8BF84513, &H458BFC55, &H1C5E8BF8, &H8B107E89, &H7D03187E, &H144E89B0, &H4D13CB8B, &H13F803B4, &H187E3BCA, &HCB3B0475, &HCB3B2274, &H5721077, &H73187E3B, &H1B809, &HD2330000, &H570F0EEB
    pvAppendBuffer &H130F66C0, &H558BF845, &HF8458BFC, &H89245E8B, &H7E8B187E, &HB87D0320, &H8B1C4E89, &HBC4D13CB, &HCA13F803, &H75207E3B, &H74CB3B04, &H77CB3B22, &H3B057210, &H973207E, &H1B8&, &HEBD23300, &HC0570F0E, &H45130F66, &HFC558BF8, &H8BF8458B, &H7E892C5E, &H287E8B20, &H89C07D03, &HCB8B244E, &H3C44D13, &H3BCA13F8, &H475287E, &H2274CB3B, &H1077CB3B, &H7E3B0572, &HB8097328, &H1&
    pvAppendBuffer &HEEBD233, &H66C0570F, &HF845130F, &H8BFC558B, &H5E8BF845, &H287E8934, &H3307E8B, &H4E89C87D, &H13CB8B2C, &HF803CC4D, &H7E3BCA13, &H3B047530, &H3B2274CB, &H721077CB, &H307E3B05, &H1B80973, &H33000000, &HF0EEBD2, &HF66C057, &H8BF84513, &H458BFC55, &H3C5E8BF8, &H8B307E89, &H7D03387E, &H344E89D0, &H4D13CB8B, &H13F803D4, &H387E3BCA, &HCB3B0475, &HCB3B2274, &H5721077, &H73387E3B
    pvAppendBuffer &H1B809, &HD2330000, &H570F0EEB, &H130F66C0, &H558BF845, &HF8458BFC, &H8D3C4E89, &H4D8B305E, &H89C803D8, &H458B387E, &H1C213DC, &H1153404E, &HE9E84446, &H85FFFFFA, &HA1840FC0, &H5FFFFFFD, &HFF6D2BE8, &HB005FF, &H56500000, &HFFF9EFE8, &H7EC085FF, &H6D16E827, &HB005FFFF, &H50000000, &H39E85656, &HE8000014, &HFFFF6D04, &HB005&, &HE8565000, &HFFFFF9C8, &HD97FC085, &H875FF56
    pvAppendBuffer &H103BE8, &H8B5B5E00, &H8C25DE5, &HCCCCCC00, &H83EC8B55, &H8B5328EC, &H5756085D, &H570C7D8B, &H107AE853, &H478B0000, &HC0570F2C, &H8BE44589, &H45893047, &H34478BE8, &H8BEC4589, &H45893847, &H3C478BF0, &H8DF44589, &H16AD845, &HF665050, &HC7D84513, &HE045&, &HF1E80000, &H8BFFFFFB, &HD8458DF0, &HE8535350, &HFFFFF764, &H3384F8B, &H30478BF0, &H893C578B, &HC033E445, &H8934470B
    pvAppendBuffer &H458DE845, &H50016AD8, &HE045C750, &H0&, &H89EC4D89, &H45C7F055, &HF4&, &HFBAEE800, &HF003FFFF, &H50D8458D, &H21E85353, &H3FFFFF7, &HE445C7F0, &H0&, &HF20478B, &H4589C057, &H24478BD8, &H8BDC4589, &H45892847, &H38478BE0, &H8BF04589, &H45893C47, &HD8458DF4, &H66535350, &HE845130F, &HFFF6E7E8, &H244F8BFF, &HC033F003, &HBD84D89, &H45892847, &H30478BDC, &H8B34578B
    pvAppendBuffer &HF84589CA, &H470BC033, &HE045892C, &H8938478B, &H478BE845, &HEC45893C, &H470BC033, &HF4458920, &H50D8458D, &H4D895353, &HF05589E4, &HFFF69FE8, &H2C4F8BFF, &H578BF003, &HBC03334, &H570F3047, &HDC4589C0, &H8920478B, &H458DF045, &HD84D89D8, &H4F0BC933, &H53535028, &HC7E05589, &HE445&, &HF660000, &H89E84513, &HC1E8F44D, &H8B000014, &HF02B2457, &HF30478B, &H4589C057, &H8B20B1D8
    pvAppendBuffer &H45893447, &H38478BDC, &H8BE04589, &H45893C47, &H20478BE4, &H45130F66, &H7882E8E8, &H570BFFFF, &HF045892C, &H50D8458D, &H55895353, &H147EE8F4, &H558B0000, &H8BF02B0C, &HC033344F, &H8B38470B, &H4589245F, &HBC033DC, &H4D893C47, &H204F8BD8, &H4589FF33, &H28428BE0, &H892C528B, &H20B1E44D, &HFF781BE8, &HC7D80BFF, &HF045&, &H5D890000, &H8BFA0BE8, &H7D890C5D, &H87D8BEC, &H8930438B
    pvAppendBuffer &H458DF445, &H575750D8, &H1423E8, &HC7F02B00, &HE045&, &H438B0000, &HD8458938, &H893C438B, &H438BDC45, &HE4458924, &H8928438B, &H438BE845, &HEC45892C, &H8934438B, &H458DF445, &H575750D8, &HF045C7, &HE8000000, &H13E4&, &H2079F02B, &HFF6A9BE8, &H10C083FF, &HE8575750, &HFFFFF570, &HEC78F003, &H8B5B5E5F, &H8C25DE5, &H1F0F00, &H1475F685, &H6A76E857, &HC083FFFF, &HADE85010
    pvAppendBuffer &H83FFFFF7, &HDC7401F8, &HFF6A63E8, &H10C083FF, &HE8575750, &H1398&, &HD4EBF02B, &HCCCCCCCC, &H56EC8B55, &H8B1075FF, &H75FF0875, &HDDE8560C, &HBFFFFF2, &HFF0D75C2, &HE8561475, &HFFFFF700, &HA78C085, &H561475FF, &H1152E856, &H5D5E0000, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H8B1075FF, &H75FF0875, &HDDE8560C, &HBFFFFF4, &HFF0D75C2, &HE8561475, &HFFFFF730
    pvAppendBuffer &HA78C085, &H561475FF, &H1322E856, &H5D5E0000, &HCC0010C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H81EC8B55, &HC8EC&, &H758B5600, &H6DE8560C, &H85FFFFF7, &HFF0F74C0, &HD1E80875, &H5EFFFFF5, &HC25DE58B, &H5657000C, &HFF38858D, &HE850FFFF, &HCEC&, &H8D107D8B, &HFFFF6885, &HE85057FF, &HCDC&, &H50C8458D, &HFFF5A3E8, &H98458DFF, &H1C845C7, &H50000000, &HCC45C7, &HE8000000
    pvAppendBuffer &HFFFFF58C, &HFF68858D, &H8D50FFFF, &HFFFF3885, &H29E850FF, &H8BFFFFF6, &HFD285D0, &H1BE84, &H8D8B5300, &HFFFFFF38, &H83C0570F, &HF6601E1, &H83F84513, &H2F7500C9, &HFF38858D, &HE850FFFF, &HBBC&, &H83C8458B, &HC88301E0, &HBF840F00, &H57000000, &H50C8458D, &HF1B2E850, &HF08BFFFF, &HB1E9DA8B, &H8B000000, &HFFFF6885, &H1E083FF, &H7500C883, &H68858D2F, &H50FFFFFF, &HB7FE8
    pvAppendBuffer &H98458B00, &H8301E083, &H840F00C8, &H111&, &H98458D57, &H75E85050, &H8BFFFFF1, &HE9DA8BF0, &H103&, &H8E0FD285, &H8F&, &HFF68858D, &H8D50FFFF, &HFFFF3885, &HE85050FF, &HFE0&, &HFF38858D, &HE850FFFF, &HB34&, &H5098458D, &H50C8458D, &HFFF567E8, &H79C085FF, &H458D570B, &HE85050C8, &HFFFFF128, &H5098458D, &H50C8458D, &HFAAE850, &H458B0000, &H1E083C8, &H7400C883
    pvAppendBuffer &H458D5711, &HE85050C8, &HFFFFF104, &HDA8BF08B, &H5D8B06EB, &HF8758BFC, &H50C8458D, &HADFE8, &HFF30B00, &H9884&, &HF0458B00, &HF44D81, &H89800000, &H86E9F045, &H8D000000, &HFFFF3885, &H858D50FF, &HFFFFFF68, &H51E85050, &H8D00000F, &HFFFF6885, &HA5E850FF, &H8D00000A, &H8D50C845, &HE8509845, &HFFFFF4D8, &HB79C085, &H98458D57, &H99E85050, &H8DFFFFF0, &H8D50C845, &H50509845
    pvAppendBuffer &HF1BE8, &H98458B00, &H8301E083, &H117400C8, &H98458D57, &H75E85050, &H8BFFFFF0, &HEBDA8BF0, &HFC5D8B06, &H8DF8758B, &HE8509845, &HA50&, &HD74F30B, &H81C0458B, &HC44D&, &H45898000, &H68858DC0, &H50FFFFFF, &HFF38858D, &HE850FFFF, &HFFFFF46C, &HD285D08B, &HFE44850F, &H8D5BFFFF, &HFF50C845, &HD5E80875, &H5F00000A, &H5DE58B5E, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H81EC8B55, &H88EC&, &H758B5600, &H3DE8560C, &H85FFFFF5, &HFF0F74C0, &HD1E80875, &H5EFFFFF3, &HC25DE58B, &H5657000C, &HFF78858D, &HE850FFFF, &HAEC&, &H8D107D8B, &H50579845, &HADFE8, &HD8458D00, &HF3A6E850, &H458DFFFF, &HD845C7B8, &H1&, &HDC45C750, &H0&, &HFFF38FE8, &H98458DFF, &H78858D50, &H50FFFFFF, &HFFF43FE8, &H85D08BFF, &HB0840FD2, &H53000001, &H401F0F
    pvAppendBuffer &HFF788D8B, &H570FFFFF, &H1E183C0, &H45130F66, &HC983F8, &H858D2F75, &HFFFFFF78, &H9BEE850, &H458B0000, &H1E083D8, &HF00C883, &HB684&, &H458D5700, &HE85050D8, &HFFFFF194, &HDA8BF08B, &HA8E9&, &H98458B00, &H8301E083, &H2C7500C8, &H5098458D, &H987E8, &HB8458B00, &H8301E083, &H840F00C8, &H108&, &HB8458D57, &H5DE85050, &H8BFFFFF1, &HE9DA8BF0, &HFA&, &H8E0FD285
    pvAppendBuffer &H8C&, &H5098458D, &HFF78858D, &H5050FFFF, &HF9BE8, &H78858D00, &H50FFFFFF, &H93FE8, &HB8458D00, &HD8458D50, &HF382E850, &HC085FFFF, &H8D570B79, &H5050D845, &HFFF113E8, &HB8458DFF, &HD8458D50, &H65E85050, &H8B00000F, &HE083D845, &HC88301, &H8D571174, &H5050D845, &HFFF0EFE8, &H8BF08BFF, &H8B06EBDA, &H758BFC5D, &HD8458DF8, &H8EAE850, &HF30B0000, &H92840F, &H458B0000
    pvAppendBuffer &HF44D81F0, &H80000000, &HE9F04589, &H80&, &HFF78858D, &H8D50FFFF, &H50509845, &HF0FE8, &H98458D00, &H8B6E850, &H458D0000, &H458D50D8, &HF9E850B8, &H85FFFFF2, &H570B79C0, &H50B8458D, &HF08AE850, &H458DFFFF, &H458D50D8, &HE85050B8, &HEDC&, &H83B8458B, &HC88301E0, &H57117400, &H50B8458D, &HF066E850, &HF08BFFFF, &H6EBDA8B, &H8BFC5D8B, &H458DF875, &H61E850B8, &HB000008
    pvAppendBuffer &H8B0D74F3, &H4D81D045, &HD4&, &HD0458980, &H5098458D, &HFF78858D, &HE850FFFF, &HFFFFF290, &HD285D08B, &HFE56850F, &H8D5BFFFF, &HFF50D845, &HE9E80875, &H5F000008, &H5DE58B5E, &HCC000CC2, &H81EC8B55, &HC0EC&, &H8B565300, &H56571475, &H6ABE8, &H1075FF00, &H858DD88B, &HFFFFFF40, &H500C75FF, &H3D7E8, &H70858D00, &H50FFFFFF, &H68BE8, &H85F88B00, &H810874FF, &H180C7
    pvAppendBuffer &H8D0EEB00, &HFFFF4085, &H71E850FF, &H8B000006, &H73FB3BF8, &H40858D18, &H50FFFFFF, &HE80875FF, &H81C&, &H8B5B5E5F, &H10C25DE5, &HA0458D00, &HF0DAE850, &H458DFFFF, &HD1E850D0, &H8BFFFFF0, &H8BC32BC7, &H6EBC1D8, &H743FE083, &H458D5018, &H48D56A0, &HA5E850D8, &H89FFFFF2, &H89D0DD44, &HEBD4DD54, &HA0458D0D, &HD8048D56, &H7CEE850, &H5D8B0000, &H95E85308, &HC7FFFFF0, &H103&
    pvAppendBuffer &H443C700, &H0&, &H180FF81, &H12770000, &HA0458D56, &HF126E850, &HC085FFFF, &H82880F, &H858D0000, &HFFFFFF70, &HD0458D50, &HF10EE850, &HC085FFFF, &H48751678, &HFF40858D, &H8D50FFFF, &HE850A045, &HFFFFF0F8, &H347FC085, &H50A0458D, &HFF40858D, &H5050FFFF, &HB43E8, &H74C20B00, &H858D530E, &HFFFFFF70, &H31E85050, &H8D00000B, &H8D50D045, &HFFFF7085, &HE85050FF, &HB20&
    pvAppendBuffer &H8DD0758B, &HC150D045, &H71E81FE6, &H8D000006, &HE850A045, &H668&, &H4FCC7509, &HE914758B, &HFFFFFF64, &HFF40858D, &H5350FFFF, &H70FE8, &H5B5E5F00, &HC25DE58B, &HCCCC0010, &HCCCCCCCC, &H81EC8B55, &H80EC&, &H8B565300, &H56571475, &H57BE8, &H1075FF00, &H458DD88B, &HC75FF80, &H3BAE850, &H458D0000, &H61E850A0, &H8B000005, &H74FF85F8, &HC78108, &HEB000001, &H80458D0B
    pvAppendBuffer &H54AE850, &HF88B0000, &H1573FB3B, &H5080458D, &HE80875FF, &H708&, &H8B5B5E5F, &H10C25DE5, &HC0458D00, &HEFC6E850, &H458DFFFF, &HBDE850E0, &H8BFFFFEF, &H8BC32BC7, &H6EBC1D8, &H743FE083, &H458D5018, &H48D56C0, &H81E850D8, &H89FFFFF2, &H89E0DD44, &HEBE4DD54, &HC0458D0D, &HD8048D56, &H6BAE850, &H5D8B0000, &H81E85308, &HC7FFFFEF, &H103&, &H443C700, &H0&, &H401F0F
    pvAppendBuffer &H100FF81, &HE770000, &HC0458D56, &HF01EE850, &HC085FFFF, &H458D7378, &H458D50A0, &HDE850E0, &H85FFFFF0, &H751378C0, &H80458D3C, &HC0458D50, &HEFFAE850, &HC085FFFF, &H458D2B7F, &H458D50C0, &HE8505080, &HBE8&, &HB74C20B, &HA0458D53, &HD9E85050, &H8D00000B, &H8D50E045, &H5050A045, &HBCBE8, &HE0758B00, &H50E0458D, &HE81FE6C1, &H56C&, &H50C0458D, &H563E8, &HDC750900
    pvAppendBuffer &H14758B4F, &HFFFF77E9, &H80458DFF, &HDE85350, &H5F000006, &HE58B5B5E, &H10C25D, &HCCCCCCCC, &H83EC8B55, &H458D60EC, &H1075FFA0, &H500C75FF, &H10BE8, &HA0458D00, &H875FF50, &HFFF27FE8, &H5DE58BFF, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D40EC, &H1075FFC0, &H500C75FF, &H23BE8, &HC0458D00, &H875FF50, &HFFF51FE8, &H5DE58BFF, &HCC000CC2, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H458D60EC, &HC75FFA0, &H5CEE850, &H458D0000, &H75FF50A0, &HF222E808, &HE58BFFFF, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &H458D40EC, &HC75FFC0, &H73EE850, &H458D0000, &H75FF50C0, &HF4C2E808, &HE58BFFFF, &H8C25D, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H8B1075FF, &H75FF0875, &HADE8560C, &HB000008, &HFF0A74C2, &H56561475, &HFFEA0FE8
    pvAppendBuffer &HC25D5EFF, &HCCCC0010, &HCCCCCCCC, &HCCCCCCCC, &H56EC8B55, &H8B1075FF, &H75FF0875, &H8DE8560C, &HB00000A, &HFF0A74C2, &H56561475, &HFFEC1FE8, &HC25D5EFF, &HCCCC0010, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &HF5360EC, &H6656C057, &HD845130F, &H57DC458B, &H45130F66, &H8BFF33D0, &H4589D45D, &H8DF633FC, &HFF83FB47, &HC0570F06, &H45130F66, &HF4558BF4, &H3BF0430F, &HD2870FF7, &H8B000000
    pvAppendBuffer &HC78B104D, &HD045100F, &H110FC62B, &H1C8DC045, &HF8458BC1, &H89F04589, &HF66F855, &H441F&, &HF06FE83, &HA383&, &H473FF00, &HFF0C458B, &HF074FF33, &HF034FF04, &H50B0458D, &HFFD2DFE8, &H10EC83FF, &HEC83CC8B, &H100F10, &H8B08100F, &H1110FC4, &HC045100F, &HE04D110F, &H8D00110F, &HE850A045, &HFFFF6E78, &HD9730F66, &H10100F0C, &HC87E0F66, &H66C2280F, &HCD8730F, &HC17E0F66
    pvAppendBuffer &HC055110F, &HFFC4D89, &H3BD05511, &H721377C8, &HD8458B08, &H73E8453B, &H1B809, &HC9330000, &H570F0EEB, &H130F66C0, &H4D8BE845, &HE8458BEC, &H3F8558B, &HF0458BD0, &H13F85589, &HEB8346C1, &HF0458908, &H860FF73B, &HFFFFFF54, &HEBD45D8B, &HF8458B03, &H8B084D8B, &H3489D075, &H8BF18BF9, &H89D08BCA, &H5C89DC55, &H8B4704FE, &H5D8BD875, &HD07589FC, &H89D45D89, &H5589D84D, &HBFF83FC
    pvAppendBuffer &HFEDB820F, &H458BFFFF, &H70895F08, &H58895E58, &HE58B5B5C, &HCC25D, &HCCCCCCCC, &HCCCCCCCC, &H83EC8B55, &HF5360EC, &H6656C057, &HD845130F, &H57DC458B, &H45130F66, &H8BFF33D0, &H4589D45D, &H8DF633FC, &HFF83FD47, &HC0570F04, &H45130F66, &HF4558BF4, &H3BF0430F, &HD2870FF7, &H8B000000, &HC78B104D, &HD045100F, &H110FC62B, &H1C8DC045, &HF8458BC1, &H89F04589, &HF66F855, &H441F&
    pvAppendBuffer &HF04FE83, &HA383&, &H473FF00, &HFF0C458B, &HF074FF33, &HF034FF04, &H50B0458D, &HFFD17FE8, &H10EC83FF, &HEC83CC8B, &H100F10, &H8B08100F, &H1110FC4, &HC045100F, &HE04D110F, &H8D00110F, &HE850A045, &HFFFF6D18, &HD9730F66, &H10100F0C, &HC87E0F66, &H66C2280F, &HCD8730F, &HC17E0F66, &HC055110F, &HFFC4D89, &H3BD05511, &H721377C8, &HD8458B08, &H73E8453B, &H1B809, &HC9330000
    pvAppendBuffer &H570F0EEB, &H130F66C0, &H4D8BE845, &HE8458BEC, &H3F8558B, &HF0458BD0, &H13F85589, &HEB8346C1, &HF0458908, &H860FF73B, &HFFFFFF54, &HEBD45D8B, &HF8458B03, &H8B084D8B, &H3489D075, &H8BF18BF9, &H89D08BCA, &H5C89DC55, &H8B4704FE, &H5D8BD875, &HD07589FC, &H89D45D89, &H5589D84D, &H7FF83FC, &HFEDB820F, &H458BFFFF, &H70895F08, &H58895E38, &HE58B5B3C, &HCC25D, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H56EC8B55, &H87D8B57, &H92E857, &HF08B0000, &H675F685, &HC25D5E5F, &H548B0004, &HCA8BF8F7, &HFCF7448B, &HC80BFF33, &HF661374, &H441F&, &H1C2AC0F, &H8BE8D147, &H75C80BCA, &H6E6C1F3, &H3C0468D, &H5D5E5FC7, &HCC0004C2, &HCCCCCCCC, &H56EC8B55, &H87D8B57, &H72E857, &HF08B0000, &H675F685, &HC25D5E5F, &H548B0004, &HCA8BF8F7, &HFCF7448B, &HC80BFF33, &HF661374, &H441F&
    pvAppendBuffer &H1C2AC0F, &H8BE8D147, &H75C80BCA, &H6E6C1F3, &H3C0468D, &H5D5E5FC7, &HCC0004C2, &HCCCCCCCC, &H8BEC8B55, &H5B80855, &HF000000, &H441F&, &HBC20C8B, &H7504C24C, &H1E88305, &H5D40F279, &HCC0004C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H3B80855, &HF000000, &H441F&, &HBC20C8B, &H7504C24C, &H1E88305, &H5D40F279, &HCC0004C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC
    pvAppendBuffer &H83EC8B55, &H458B08EC, &HC0570F08, &H66D88B53, &HF845130F, &H3B30C083, &H8B3876C3, &H5756F84D, &H89FC7D8B, &H708B084D, &H8E883F8, &H508BCE8B, &HD1AC0F04, &H84D0B01, &HD70BEAD1, &HFE8B0889, &HC1045089, &H45C71FE7, &H8&, &H77C33B00, &H5B5E5FD5, &HC25DE58B, &HCCCC0004, &HCCCCCCCC, &H83EC8B55, &H458B08EC, &HC0570F08, &H66D88B53, &HF845130F, &H3B20C083, &H8B3876C3, &H5756F84D
    pvAppendBuffer &H89FC7D8B, &H708B084D, &H8E883F8, &H508BCE8B, &HD1AC0F04, &H84D0B01, &HD70BEAD1, &HFE8B0889, &HC1045089, &H45C71FE7, &H8&, &H77C33B00, &H5B5E5FD5, &HC25DE58B, &HCCCC0004, &HCCCCCCCC, &H8BEC8B55, &H4D8B0C55, &H89028B08, &H4428B01, &H8B044189, &H41890842, &HC428B08, &H8B0C4189, &H41891042, &H14428B10, &H8B144189, &H41891842, &H1C428B18, &H8B1C4189, &H41892042, &H24428B20
    pvAppendBuffer &H8B244189, &H41892842, &H2C428B28, &H5D2C4189, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &H4D8B0C55, &H89028B08, &H4428B01, &H8B044189, &H41890842, &HC428B08, &H8B0C4189, &H41891042, &H14428B10, &H8B144189, &H41891842, &H1C428B18, &H5D1C4189, &HCC0008C2, &HCCCCCCCC, &H83EC8B55, &HF5360EC, &HC933C057, &H130F6656, &H458BD845, &HF6657DC, &H8BD04513, &H4D89D47D
    pvAppendBuffer &HF04589E8, &H418DF633, &H6F983FB, &H66C0570F, &HF845130F, &HFFC5D8B, &HF13BF043, &H119870F, &H558B0000, &HFC18B0C, &H2BD04510, &HF45D89C6, &HC045110F, &H8BC2048D, &H4589F855, &HFC5589EC, &HFE2BF98B, &H870FF73B, &HEA&, &HFF0470FF, &HC458B30, &H4F074FF, &H8DF034FF, &HE850B045, &HFFFFCDBC, &HF00100F, &H3BD04511, &H8B4373F7, &HC18BDC4D, &H8BD4558B, &H1FE8C1FA, &H8BFC4501
    pvAppendBuffer &HD383D845, &H1FEFC100, &H1C1A40F, &H33F45D89, &HBC003DB, &H89F80BD9, &H458BDC5D, &HC2A40FD0, &HD87D8901, &H5589C003, &HD04589D4, &HD045100F, &H5D8B06EB, &HD87D8BDC, &H8B10EC83, &H10EC83C4, &H8B00110F, &H45100FC4, &H110FC0, &H50A0458D, &HFF690BE8, &H8100FFF, &H66C1280F, &HCD8730F, &HC07E0F66, &HC04D110F, &HFF04589, &H3BD04D11, &H721077C3, &HD87D3905, &H1B80973, &H33000000
    pvAppendBuffer &HF0EEBC9, &HF66C057, &H8BE04513, &H458BE44D, &HFC558BE0, &H3F45D8B, &HEC458BD0, &H5589D913, &HE84D8BFC, &H8E88346, &H89F45D89, &HF13BEC45, &HFF0A860F, &H7D8BFFFF, &H8B03EBD4, &H758BF855, &HD0458B08, &H8BCE0489, &H7C89D845, &H8B4104CE, &H5589F07D, &H89D38BD8, &H7D89D045, &HF05589D4, &H89DC5589, &HF983E84D, &H95820F0B, &H89FFFFFE, &H895F5C7E, &H5B5E5846, &HC25DE58B, &HCCCC0008
    pvAppendBuffer &H83EC8B55, &HF5360EC, &HC933C057, &H130F6656, &H458BD845, &HF6657DC, &H8BD04513, &H4D89D47D, &HF04589E8, &H418DF633, &H4F983FD, &H66C0570F, &HF845130F, &HFFC5D8B, &HF13BF043, &H119870F, &H558B0000, &HFC18B0C, &H2BD04510, &HF45D89C6, &HC045110F, &H8BC2048D, &H4589F855, &HFC5589EC, &HFE2BF98B, &H870FF73B, &HEA&, &HFF0470FF, &HC458B30, &H4F074FF, &H8DF034FF, &HE850B045
    pvAppendBuffer &HFFFFCC1C, &HF00100F, &H3BD04511, &H8B4373F7, &HC18BDC4D, &H8BD4558B, &H1FE8C1FA, &H8BFC4501, &HD383D845, &H1FEFC100, &H1C1A40F, &H33F45D89, &HBC003DB, &H89F80BD9, &H458BDC5D, &HC2A40FD0, &HD87D8901, &H5589C003, &HD04589D4, &HD045100F, &H5D8B06EB, &HD87D8BDC, &H8B10EC83, &H10EC83C4, &H8B00110F, &H45100FC4, &H110FC0, &H50A0458D, &HFF676BE8, &H8100FFF, &H66C1280F, &HCD8730F
    pvAppendBuffer &HC07E0F66, &HC04D110F, &HFF04589, &H3BD04D11, &H721077C3, &HD87D3905, &H1B80973, &H33000000, &HF0EEBC9, &HF66C057, &H8BE04513, &H458BE44D, &HFC558BE0, &H3F45D8B, &HEC458BD0, &H5589D913, &HE84D8BFC, &H8E88346, &H89F45D89, &HF13BEC45, &HFF0A860F, &H7D8BFFFF, &H8B03EBD4, &H758BF855, &HD0458B08, &H8BCE0489, &H7C89D845, &H8B4104CE, &H5589F07D, &H89D38BD8, &H7D89D045, &HF05589D4
    pvAppendBuffer &H89DC5589, &HF983E84D, &H95820F07, &H89FFFFFE, &H895F3C7E, &H5B5E3846, &HC25DE58B, &HCCCC0008, &H83EC8B55, &H8B530CEC, &H570F0C5D, &H8B5756C0, &H138B107D, &H438BF28B, &H66C88B04, &HF445130F, &H4F1B372B, &H75F23B04, &H75C83B06, &H3B18EB04, &H770F72C8, &H76F23B04, &H1B809, &HD2330000, &HF660BEB, &H8BF44513, &H558BF445, &H87D8BF8, &H89044F89, &H8738B37, &H7589CE8B, &H10758BF8
    pvAppendBuffer &H89084E2B, &H4B8B084D, &HC4E1B0C, &H2B085D8B, &H85D89D8, &H5D3BCA1B, &HC5D8BF8, &H4B3B0575, &H3B23740C, &H13720C4B, &H438B0877, &H8453908, &H1B80976, &H33000000, &H660BEBD2, &HF445130F, &H8BF8558B, &H758BF445, &HC4F8908, &H8B087789, &HCE8B1073, &H8BFC7589, &H4E2B1075, &H84D8910, &H1B144B8B, &H5D8B144E, &H89D82B08, &HCA1B085D, &H8BFC5D3B, &H5750C5D, &H74144B3B, &H144B3B23
    pvAppendBuffer &H8771372, &H3910438B, &H9760845, &H1B8&, &HEBD23300, &H130F660B, &H558BF445, &HF4458BF8, &H8908758B, &H7789144F, &H18738B10, &H7589CE8B, &H10758BFC, &H89184E2B, &H4B8B084D, &H1C4E1B1C, &H2B085D8B, &H85D89D8, &H5D3BCA1B, &HC5D8BFC, &H4B3B0575, &H3B23741C, &H13721C4B, &H438B0877, &H8453918, &H1B80976, &H33000000, &H660BEBD2, &HF445130F, &H8BF8558B, &H758BF445, &H18778908
    pvAppendBuffer &H8910758B, &H4B8B1C4F, &H204E2B20, &H8B0C4D89, &H4E1B244B, &HC758B24, &HCA1BF02B, &H7520733B, &H244B3B05, &H4B3B2074, &H77107224, &H20733B05, &H1B80976, &H33000000, &H660BEBD2, &HF445130F, &H8BF8558B, &H7789F445, &H244F8920, &H8B28738B, &H5D8B2C4B, &H8758910, &H2B0C4D89, &H4B1B2873, &H8BF02B2C, &HCA1B0C5D, &H7508753B, &H74CB3B04, &H72CB3B2C, &H3B05771D, &H16760875, &HB8287789
    pvAppendBuffer &H1&, &H332C4F89, &H5B5E5FD2, &HC25DE58B, &HF66000C, &H8BF44513, &H458BF855, &H287789F4, &H5F2C4F89, &HE58B5B5E, &HCC25D, &HCCCCCCCC, &H83EC8B55, &H8B530CEC, &H570F0C5D, &H8B5756C0, &H138B107D, &H438BF28B, &H66C88B04, &HF445130F, &H4F1B372B, &H75F23B04, &H75C83B06, &H3B18EB04, &H770F72C8, &H76F23B04, &H1B809, &HD2330000, &HF660BEB, &H8BF44513, &H558BF445, &H87D8BF8
    pvAppendBuffer &H8B044F89, &H3789104D, &H8908738B, &H712BF875, &HC4B8B08, &H1B105D8B, &HF02B0C4B, &H1B0C5D8B, &HF8753BCA, &H4B3B0575, &H3B20740C, &H10720C4B, &H733B0577, &HB8097608, &H1&, &HBEBD233, &H45130F66, &HF8558BF4, &H89F4458B, &H4D8B0C4F, &H8778910, &H8910738B, &H712BFC75, &H144B8B10, &H1B105D8B, &HF02B144B, &H1B0C5D8B, &HFC753BCA, &H4B3B0575, &H3B207414, &H1072144B, &H733B0577
    pvAppendBuffer &HB8097610, &H1&, &HBEBD233, &H45130F66, &HF8558BF4, &H89F4458B, &H7789144F, &H184B8B10, &H7D8BF18B, &H1C5B8B10, &H8B0C4D89, &H712B104D, &H1BCB8B18, &HF02B1C4F, &H1B087D8B, &HC753BCA, &HCB3B0475, &HCB3B2C74, &H5771D72, &H760C753B, &H18778916, &H1B8&, &H1C4F8900, &H5E5FD233, &H5DE58B5B, &H66000CC2, &HF445130F, &H8BF8558B, &H7789F445, &H1C4F8918, &H8B5B5E5F, &HCC25DE5
    pvAppendBuffer &HCCCCCC00, &HCCCCCCCC, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HD233084D, &H7D8B5756, &H8BF6330C, &H3FE083C7, &H83C6AB0F, &H430F20F8, &H83F233D6, &H430F40F8, &H6EFC1D6, &H23F93423, &H8B04F954, &H5D5E5FC6, &HCC0008C2, &HCCCCCCCC, &HCCCCCCCC, &H8BEC8B55, &HEC831455, &H85C93310, &HC2840FD2, &H53000000, &H56105D8B, &H5708758B, &H830C7D8B, &H820F20FA, &H8B&, &H3FF438D, &H77F03BC2
    pvAppendBuffer &HFF468D09, &HC33BC203, &H478D7973, &H3BC203FF, &H8D0977F0, &HC203FF46, &H6773C73B, &HD78BC28B, &HE083D32B, &HFC5589E0, &HD32BD68B, &H89F04589, &HC38BF855, &H8BF85D8B, &HFC7D8BD7, &H5589D62B, &H10568DF4, &H8B00100F, &HC183F475, &H20408D20, &HF20528D, &HE0074C10, &HC8EF0F66, &H34C110F, &H4C100FE0, &H758BE016, &H40100F08, &HEF0F66F0, &H4A110FC8, &HF04D3BE0, &H558BCA72, &HC7D8B14
    pvAppendBuffer &H3B105D8B, &H2B1B73CA, &H19048DFB, &HD12BF32B, &H8D380C8A, &H48320140, &H304C88FF, &H1EA83FF, &H5E5FEE75, &H5DE58B5B, &H10C2&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&, &H0&
    '--- end thunk data
    ReDim baBuffer(0 To 44077 - 1) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub
