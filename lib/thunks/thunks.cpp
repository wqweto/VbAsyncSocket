#define IMPL_CURVE25519
#define IMPL_ECC256_THUNK
#define IMPL_ECC384_THUNK
#define IMPL_SHA256_THUNK
#define IMPL_SHA384_THUNK
#define IMPL_SHA512_THUNK
#define IMPL_CHACHA20_THUNK
#define IMPL_AESGCM_THUNK
//#define IMPL_GMPRSA_THUNK
#define IMPL_SSHRSA_THUNK
//#define IMPL_TINF_THUNK

#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include <windows.h>

#pragma intrinsic(memset, memcpy)
#pragma comment(lib, "crypt32")
#pragma comment(lib, "comctl32")
//#pragma nodefaultlib
//#pragma comment(linker, "/entry:main")
//#pragma comment(linker, "/INCLUDE:_mainCRTStartup")

#pragma code_seg(".supcode")

static LPWSTR __stdcall GetCurrentDateTime()
{
    static WCHAR szResult[50];
    SYSTEMTIME  st;
    DATE        dt;
    VARIANT     vdt = { VT_DATE, };
    VARIANT     vstr = { VT_EMPTY };

    GetLocalTime(&st);
    SystemTimeToVariantTime(&st, &dt);
    vdt.date = dt;
    VariantChangeType(&vstr, &vdt, 0, VT_BSTR);
    memcpy(szResult, vstr.bstrVal, sizeof szResult);
    VariantClear(&vstr);
    return szResult;
}


#define assert(e)
#define abort() { }
#define MIN(x, y) ((x) < (y) ? (x) : (y))

#ifdef IMPL_ECC256_THUNK
    #include "ecc.h"
#endif
#ifdef IMPL_ECC384_THUNK
    #include "ecc384.h"
#endif
#ifdef IMPL_SSHRSA_THUNK
    #include "sshbn.h"
#endif
typedef LPVOID (__stdcall *CoTaskMemAlloc_t)(SIZE_T cb);
typedef LPVOID (__stdcall *CoTaskMemRealloc_t)(LPVOID pv, SIZE_T cb);
typedef void (__stdcall *CoTaskMemFree_t)(LPVOID pv);

typedef struct {
#if defined(IMPL_SSHRSA_THUNK) || defined (IMPL_GMPRSA_THUNK)
    CoTaskMemAlloc_t m_CoTaskMemAlloc;
    CoTaskMemRealloc_t m_CoTaskMemRealloc;
    CoTaskMemFree_t m_CoTaskMemFree;
#endif
#ifdef IMPL_ECC256_THUNK
    uint64_t m_curve_p[NUM_ECC_DIGITS];
    uint64_t m_curve_b[NUM_ECC_DIGITS];
    EccPoint m_curve_G;
    uint64_t m_curve_n[NUM_ECC_DIGITS];
#endif
#ifdef IMPL_ECC384_THUNK
    uint64_t m_curve_p_384[NUM_ECC_DIGITS_384];
    uint64_t m_curve_b_384[NUM_ECC_DIGITS_384];
    EccPoint384 m_curve_G_384;
    uint64_t m_curve_n_384[NUM_ECC_DIGITS_384];
#endif
#ifdef IMPL_SHA256_THUNK
    uint32_t m_K256[64];
#endif
#if defined(IMPL_SHA384_THUNK) || defined(IMPL_SHA512_THUNK)
    uint64_t m_K512[80];
#endif
#ifdef IMPL_CHACHA20_THUNK
    uint8_t m_chacha20_tau[17];  // "expand 16-byte k";
    uint8_t m_chacha20_sigma[17]; // "expand 32-byte k";
    uint32_t m_negative_1305[17];
#endif
#ifdef IMPL_AESGCM_THUNK
    uint8_t m_S[256];
    uint8_t m_Rcon[11];
    uint8_t m_S_inv[256];
#endif
#ifdef IMPL_SSHRSA_THUNK
    BignumInt m_bnZero[1];
    BignumInt m_bnOne[2];
#endif
} thunk_context_t;

#define curve_p (getContext()->m_curve_p)
#define curve_b (getContext()->m_curve_b)
#define curve_G (getContext()->m_curve_G)
#define curve_n (getContext()->m_curve_n)
#define curve_p_384 (getContext()->m_curve_p_384)
#define curve_b_384 (getContext()->m_curve_b_384)
#define curve_G_384 (getContext()->m_curve_G_384)
#define curve_n_384 (getContext()->m_curve_n_384)
#define K256 (getContext()->m_K256)
#define K512 (getContext()->m_K512)
#define chacha20_tau (getContext()->m_chacha20_tau)
#define chacha20_sigma (getContext()->m_chacha20_sigma)
#define negative_1305 (getContext()->m_negative_1305)
#define S (getContext()->m_S)
#define Rcon (getContext()->m_Rcon)
#define S_inv (getContext()->m_S_inv)
#define bnZero (getContext()->m_bnZero)
#define bnOne (getContext()->m_bnOne)

#pragma code_seg(push, r1, ".mythunk")

static int beginOfThunk(int i) { 
    int a[] = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 }; return a[i]; 
}

__declspec(naked) static thunk_context_t *getContext() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + getContext
        add     eax, beginOfThunk
        mov     eax, [eax]
        ret
    }
}

__declspec(naked) static uint8_t *getThunk() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + getThunk
        add     eax, beginOfThunk
        ret
    }
}

#define DECLARE_PFN(t, f) const t pfn_##f = (t)(getThunk() + (((uint8_t *)f) - ((uint8_t *)beginOfThunk)))

#ifdef __cplusplus
extern "C" {
#endif

#include "cf_inlines.h"
#include "win32_crt.cpp"
#if defined(IMPL_CURVE25519) || defined(IMPL_ECC256_THUNK) || defined(IMPL_ECC384_THUNK) || defined(IMPL_SHA384_THUNK) || defined(IMPL_SHA512_THUNK)
    #include "win32_crt_float.cpp"
#endif
#ifdef IMPL_CURVE25519
    #include "curve25519.c"
#endif
#ifdef IMPL_ECC256_THUNK
    #include "ecc.c"
#endif
#ifdef IMPL_ECC384_THUNK
    #include "ecc384.c"
#endif
#include "blockwise.c"
#ifdef IMPL_SHA256_THUNK
    #include "sha256.c"
#endif
#if defined(IMPL_SHA384_THUNK) || defined(IMPL_SHA512_THUNK)
    #include "sha512.c"
#endif
#ifdef IMPL_CHACHA20_THUNK
    #include "chacha20.c"
    #include "poly1305.c"
    #include "chacha20poly1305.c"
#endif
#ifdef IMPL_AESGCM_THUNK
    #include "aes.c"
    #include "gf128.c"
    #include "modes.c"
    #include "gcm.c"
#endif
#ifdef IMPL_GMPRSA_THUNK
    #include "mini-gmp.c"
    #include "rsa.c"
#endif
#ifdef IMPL_SSHRSA_THUNK
    #include "sshbn.c"
    #include "rsa.c"
#endif
#ifdef IMPL_TINF_THUNK
    #include "tinflate.c"
#endif

#ifdef __cplusplus
}
#endif

#pragma code_seg(pop, r1)

#pragma code_seg(push, r1, ".endthunk")
static int endOfThunk() { return 0; }
#pragma code_seg(pop, r1)


#define THUNK_SIZE (((uint8_t *)endOfThunk - (uint8_t *)beginOfThunk))

#ifdef IMPL_CURVE25519
    typedef void (*cf_curve25519_mul_t)(uint8_t *q, const uint8_t *n, const uint8_t *p);
    typedef void (*cf_curve25519_mul_base_t)(uint8_t *q, const uint8_t *n);
#endif
#ifdef IMPL_ECC256_THUNK
    static int _getRandomNumber256(uint64_t *p_vli);
    typedef int (*ecc_make_key_t)(uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES]);
    typedef int (*ecdh_shared_secret_t)(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
#endif
#ifdef IMPL_ECC384_THUNK
    static int _getRandomNumber384(uint64_t *p_vli);
    typedef int (*ecc_make_key384_t)(uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384]);
    typedef int (*ecdh_shared_secret384_t)(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384], uint8_t p_secret[ECC_BYTES_384]);
#endif
#ifdef IMPL_SHA256_THUNK
    typedef void (*cf_sha256_init_t)(cf_sha256_context *ctx);
    typedef void (*cf_sha256_update_t)(cf_sha256_context *ctx, const void *data, size_t nbytes);
    typedef void (*cf_sha256_digest_final_t)(cf_sha256_context *ctx, uint8_t hash[CF_SHA256_HASHSZ]);
#endif
#if defined(IMPL_SHA384_THUNK)
    typedef void (*cf_sha384_init_t)(cf_sha512_context *ctx);
    typedef void (*cf_sha384_update_t)(cf_sha512_context *ctx, const void *data, size_t nbytes);
    typedef void (*cf_sha384_digest_final_t)(cf_sha512_context *ctx, uint8_t hash[CF_SHA384_HASHSZ]);
#endif
#if defined(IMPL_SHA512_THUNK)
    typedef void (*cf_sha512_init_t)(cf_sha512_context *ctx);
    typedef void (*cf_sha512_update_t)(cf_sha512_context *ctx, const void *data, size_t nbytes);
    typedef void (*cf_sha512_digest_final_t)(cf_sha512_context *ctx, uint8_t hash[CF_SHA512_HASHSZ]);
#endif
#ifdef IMPL_CHACHA20_THUNK
    typedef void (*cf_chacha20poly1305_encrypt_t)(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
                                                  const uint8_t *plaintext, size_t nbytes, uint8_t *ciphertext, uint8_t tag[16]);
    typedef int (*cf_chacha20poly1305_decrypt_t)(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
                                                 const uint8_t *ciphertext, size_t nbytes, const uint8_t tag[16], uint8_t *plaintext);
#endif
#ifdef IMPL_AESGCM_THUNK
    typedef void (*cf_aesgcm_encrypt_t)(uint8_t *c, uint8_t *mac, const uint8_t *m, const size_t mlen, const uint8_t *ad, const size_t adlen,
                                        const uint8_t *npub, const uint8_t *k, size_t klen);
    typedef int (*cf_aesgcm_decrypt_t)(uint8_t *m, const uint8_t *c, const size_t clen, const uint8_t *mac, const uint8_t *ad, const size_t adlen,
                                       const uint8_t *npub, const uint8_t *k, const size_t klen);
#endif
#if defined(IMPL_GMPRSA_THUNK) || defined(IMPL_SSHRSA_THUNK)
    typedef void (*rsa_modexp_t)(uint32_t maxbytes, void *b, void *e, void *m, void *r);
#endif

typedef struct _RSA_PUBLIC_KEY_XX
{
    #pragma pack(push, 1)
    PUBLICKEYSTRUC PublicKeyStruc;
    RSAPUBKEY RsaPubKey;
    BYTE RsaModulus[ANYSIZE_ARRAY];
} RSA_PUBLIC_KEY_XX;

void __cdecl main()
{
    RSA_PUBLIC_KEY_XX a;
    printf("pubexp offset=%d\n", ((uint8_t *)&a.RsaPubKey.pubexp) - ((uint8_t *)&a));
#ifdef IMPL_SHA256_THUNK
    printf("sizeof(cf_sha256_context)=%d\n", sizeof cf_sha256_context);
#endif
#if defined(IMPL_SHA384_THUNK) || defined(IMPL_SHA512_THUNK)
    printf("sizeof(cf_sha512_context)=%d\n", sizeof cf_sha512_context);
#endif
#ifdef IMPL_CHACHA20_THUNK
    printf("sizeof(g_chacha20_tau)=%d\n", sizeof g_chacha20_tau);
#endif
    static thunk_context_t ctx;
#if defined(IMPL_SSHRSA_THUNK) || defined (IMPL_GMPRSA_THUNK)
    ctx.m_CoTaskMemAlloc = (CoTaskMemAlloc_t)GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemAlloc");
    ctx.m_CoTaskMemRealloc = (CoTaskMemRealloc_t)GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemRealloc");
    ctx.m_CoTaskMemFree = (CoTaskMemFree_t)GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemFree");
#endif
#ifdef IMPL_ECC256_THUNK
    memcpy(&ctx.m_curve_p, &g_curve_p, sizeof g_curve_p);
    memcpy(&ctx.m_curve_b, &g_curve_b, sizeof g_curve_b);
    memcpy(&ctx.m_curve_G, &g_curve_G, sizeof g_curve_G);
    memcpy(&ctx.m_curve_n, &g_curve_n, sizeof g_curve_p);
#endif
#ifdef IMPL_ECC384_THUNK
    memcpy(&ctx.m_curve_p_384, &g_curve_p_384, sizeof g_curve_p_384);
    memcpy(&ctx.m_curve_b_384, &g_curve_b_384, sizeof g_curve_b_384);
    memcpy(&ctx.m_curve_G_384, &g_curve_G_384, sizeof g_curve_G_384);
    memcpy(&ctx.m_curve_n_384, &g_curve_n_384, sizeof g_curve_p_384);
#endif
#ifdef IMPL_SHA256_THUNK
    memcpy(&ctx.m_K256, &g_K256, sizeof g_K256);
#endif
#if defined(IMPL_SHA384_THUNK) || defined(IMPL_SHA512_THUNK)
    memcpy(&ctx.m_K512, &g_K512, sizeof g_K512);
#endif
#ifdef IMPL_CHACHA20_THUNK
    memcpy(&ctx.m_chacha20_tau, &g_chacha20_tau, sizeof g_chacha20_tau);
    memcpy(&ctx.m_chacha20_sigma, &g_chacha20_sigma, sizeof g_chacha20_sigma);
    memcpy(&ctx.m_negative_1305, &g_negative_1305, sizeof g_negative_1305);
#endif
#ifdef IMPL_AESGCM_THUNK
    memcpy(&ctx.m_S, &g_S, sizeof g_S);
    memcpy(&ctx.m_Rcon, &g_Rcon, sizeof g_Rcon);
    memcpy(&ctx.m_S_inv, &g_S_inv, sizeof g_S_inv);
#endif
#ifdef IMPL_SSHRSA_THUNK
    memcpy(&ctx.m_bnZero, &g_bnZero, sizeof g_bnZero);
    memcpy(&ctx.m_bnOne, &g_bnOne, sizeof g_bnOne);
#endif

    CoInitialize(0);
    DWORD dwDummy;
    VirtualProtect(beginOfThunk, 1024, PAGE_EXECUTE_READWRITE, &dwDummy);
    ((void **)beginOfThunk)[0] = &ctx;

    size_t thunkSize = THUNK_SIZE;
    while(thunkSize > 4 && ((uint8_t *)beginOfThunk)[thunkSize - 4] == 0)
        thunkSize--;
    void *hThunk = VirtualAlloc(0, 2*THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d -> %d\n", hThunk, THUNK_SIZE, thunkSize);
    memcpy(hThunk, beginOfThunk, THUNK_SIZE);
    memset(((uint8_t *)hThunk) + thunkSize, 0xCC, 2*THUNK_SIZE - thunkSize);

    // test thunks
#ifdef IMPL_CURVE25519
    DECLARE_PFN(cf_curve25519_mul_t, cf_curve25519_mul);
#endif
#ifdef IMPL_ECC256_THUNK
    DECLARE_PFN(ecc_make_key_t, ecc_make_key);
    DECLARE_PFN(ecdh_shared_secret_t, ecdh_shared_secret);
#endif
#ifdef IMPL_SHA256_THUNK
    DECLARE_PFN(cf_sha256_init_t, cf_sha256_init);
    DECLARE_PFN(cf_sha256_update_t, cf_sha256_update);
    DECLARE_PFN(cf_sha256_digest_final_t, cf_sha256_digest_final);
#endif
#if defined(IMPL_SHA384_THUNK)
    DECLARE_PFN(cf_sha384_init_t, cf_sha384_init);
    DECLARE_PFN(cf_sha384_update_t, cf_sha384_update);
    DECLARE_PFN(cf_sha384_digest_final_t, cf_sha384_digest_final);
#endif
#if defined(IMPL_SHA512_THUNK)
    DECLARE_PFN(cf_sha512_init_t, cf_sha512_init);
    DECLARE_PFN(cf_sha512_update_t, cf_sha512_update);
    DECLARE_PFN(cf_sha512_digest_final_t, cf_sha512_digest_final);
#endif
#ifdef IMPL_CHACHA20_THUNK
    DECLARE_PFN(cf_chacha20poly1305_encrypt_t, cf_chacha20poly1305_encrypt);
    DECLARE_PFN(cf_chacha20poly1305_decrypt_t, cf_chacha20poly1305_decrypt);
#endif
#ifdef IMPL_AESGCM_THUNK
    DECLARE_PFN(cf_aesgcm_encrypt_t, cf_aesgcm_encrypt);
    DECLARE_PFN(cf_aesgcm_decrypt_t, cf_aesgcm_decrypt);
#endif
#ifdef IMPL_GMPRSA_THUNK
    DECLARE_PFN(rsa_modexp_t, gmp_rsa_public_encrypt);
#endif
#ifdef IMPL_SSHRSA_THUNK
    DECLARE_PFN(rsa_modexp_t, rsa_modexp);
#endif

#ifdef IMPL_ECC256_THUNK
    uint8_t pubkey[ECC_BYTES+1] = { 0 };
    uint8_t privkey[ECC_BYTES] = { 0 };
    uint8_t secret[ECC_BYTES] = { 0 };
    do {
        _getRandomNumber256((uint64_t *)privkey);
    } while (!ecc_make_key(pubkey, privkey));
    pfn_ecc_make_key(pubkey, privkey);
    pfn_ecdh_shared_secret(pubkey, privkey, secret);
    #ifdef IMPL_CURVE25519
        pfn_cf_curve25519_mul(secret, privkey, pubkey);
    #endif
#endif
#ifdef IMPL_SHA256_THUNK
    cf_sha256_context sha256_ctx = { 0 };
    uint8_t hash256[CF_SHA256_HASHSZ] = { 0 };
    pfn_cf_sha256_init(&sha256_ctx);
    pfn_cf_sha256_update(&sha256_ctx, "123456", 6);
    pfn_cf_sha256_digest_final(&sha256_ctx, hash256);
#endif
#if defined(IMPL_SHA384_THUNK)
    cf_sha512_context sha384_ctx = { 0 };
    uint8_t hash384[CF_SHA384_HASHSZ] = { 0 };
    pfn_cf_sha384_init(&sha384_ctx);
    pfn_cf_sha384_update(&sha384_ctx, "123456", 6);
    pfn_cf_sha384_digest_final(&sha384_ctx, hash384);
#endif
    uint8_t key[32] = { 1, 2, 3, 4 };
    uint8_t nonce[12] = { 1, 2, 3, 4 };
    uint8_t tag[16];
    uint8_t aad[] = "header text";
    uint8_t plaintext[] = "this is a test 1234567890";
    uint8_t cyphertext[100] = { 0 };
#ifdef IMPL_CHACHA20_THUNK
    pfn_cf_chacha20poly1305_encrypt(key, nonce, aad, sizeof aad, plaintext, sizeof plaintext, cyphertext, tag);
    pfn_cf_chacha20poly1305_decrypt(key, nonce, aad, sizeof aad, cyphertext, sizeof plaintext, tag, cyphertext);
#endif
#ifdef IMPL_AESGCM_THUNK
    uint8_t *mac = cyphertext + sizeof plaintext;
    pfn_cf_aesgcm_encrypt(cyphertext, mac, plaintext, sizeof plaintext, aad, sizeof aad, nonce, key, sizeof key);
    pfn_cf_aesgcm_decrypt(cyphertext, cyphertext, sizeof plaintext, mac, aad, sizeof aad, nonce, key, sizeof key);
#endif
#ifdef IMPL_GMPRSA_THUNK
    {
    #define BUFFER_SIZE 32
    uint8_t m[BUFFER_SIZE] = { 0 };
    uint8_t e[BUFFER_SIZE] = { 0 };
    uint8_t from[BUFFER_SIZE] = { 0 };
    uint8_t to[BUFFER_SIZE] = { 33 }, to2[BUFFER_SIZE] = { 0 };
    m[BUFFER_SIZE-1] = 200;
    m[BUFFER_SIZE-2] = 199;
    e[BUFFER_SIZE-1] = 2;
    from[BUFFER_SIZE-1] = 123;
    pfn_gmp_rsa_public_encrypt(BUFFER_SIZE, from, e, m, to2);
    }
#endif
#ifdef IMPL_SSHRSA_THUNK
    {
    #define BUFFER_SIZE 32
    uint8_t m[BUFFER_SIZE] = { 0 };
    uint8_t e[BUFFER_SIZE] = { 0 };
    uint8_t from[BUFFER_SIZE] = { 0 };
    uint8_t to[BUFFER_SIZE] = { 33 }, to2[BUFFER_SIZE] = { 0 };
    m[BUFFER_SIZE-1] = 200;
    e[BUFFER_SIZE-1] = 2;
    from[BUFFER_SIZE-1] = 123;
    pfn_rsa_modexp(BUFFER_SIZE, from, e, m, to);
    }
#endif

    // init offsets at beginning of thunk
    int idx = 1;
#ifdef IMPL_CURVE25519
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_curve25519_mul - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_curve25519_mul_base - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_ECC256_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)ecc_make_key - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdh_shared_secret - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdh_uncompress_key - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdsa_sign - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdsa_verify - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_ECC384_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)ecc_make_key384 - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdh_shared_secret384 - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdh_uncompress_key384 - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdsa_sign384 - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)ecdsa_verify384 - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_SHA256_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha256_init - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha256_update - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha256_digest_final - (uint8_t *)beginOfThunk);
#endif
#if defined(IMPL_SHA384_THUNK)
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha384_init - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha384_update - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha384_digest_final - (uint8_t *)beginOfThunk);
#endif
#if defined(IMPL_SHA512_THUNK)
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha512_init - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha512_update - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_sha512_digest_final - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_CHACHA20_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_chacha20poly1305_encrypt - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_chacha20poly1305_decrypt - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_AESGCM_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_aesgcm_encrypt - (uint8_t *)beginOfThunk);
    ((int *)hThunk)[idx++] = ((uint8_t *)cf_aesgcm_decrypt - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_GMPRSA_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)gmp_rsa_public_encrypt - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_SSHRSA_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)rsa_modexp - (uint8_t *)beginOfThunk);
#endif
#ifdef IMPL_TINF_THUNK
    ((int *)hThunk)[idx++] = ((uint8_t *)tinf_uncompress - (uint8_t *)beginOfThunk);
#endif
    printf("i=%d, needed=0x%02X, allocated=0x%02X\n", idx, (idx*4 + 15) & -16, ((uint8_t *)getContext) - ((uint8_t *)beginOfThunk));

    WCHAR szBuffer[100000] = { 0 }, *pBuffer;
    DWORD dwBufSize, i, j, l;
    dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)&ctx, sizeof ctx, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; ) {
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
        if (j % 900 == 0) {
            memcpy(szBuffer + j, L"\" & _\n\t\t\t\t\t\t\t\t\t\t\t\t\t\"", 40);
            j += 20;
        }
    }
    printf("Private Const STR_GLOB                  As String = \"%S\" ' %d, %S\n", szBuffer, sizeof ctx, GetCurrentDateTime());
    dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)hThunk, thunkSize, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    pBuffer = szBuffer;
    for(i = 0, j = 0, l = 1; (szBuffer[j] = szBuffer[i]) != 0; ) {
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
        if (j % 900 == 0) {
            if (l % 14 == 0) {
                memcpy(szBuffer + j, L"", 2);
                printf("Private Const STR_THUNK%d                As String = \"%S\"\n", l / 14, pBuffer);
                pBuffer = szBuffer + j;
            }
            else {
                memcpy(szBuffer + j, L"\" & _\n\t\t\t\t\t\t\t\t\t\t\t\t\t\"", 40);
                j += 20;
            }
            l++;
        }
    }
    printf("Private Const STR_THUNK%d                As String = \"%S\" ' %d, %S\n", l / 14 + 1, pBuffer, thunkSize, GetCurrentDateTime());
}

#if defined(IMPL_ECC256_THUNK)
static int _getRandomNumber256(uint64_t *p_vli)
{
    HCRYPTPROV l_prov;

    if(!CryptAcquireContext(&l_prov, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT))
    {
        return 0;
    }

    CryptGenRandom(l_prov, ECC_BYTES, (BYTE *)p_vli);
    CryptReleaseContext(l_prov, 0);
    
    return 1;
}
#endif

#if defined(IMPL_ECC384_THUNK)
static int _getRandomNumber384(uint64_t *p_vli)
{
    HCRYPTPROV l_prov;

    if(!CryptAcquireContext(&l_prov, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT))
    {
        return 0;
    }

    CryptGenRandom(l_prov, ECC_BYTES_384, (BYTE *)p_vli);
    CryptReleaseContext(l_prov, 0);
    
    return 1;
}
#endif