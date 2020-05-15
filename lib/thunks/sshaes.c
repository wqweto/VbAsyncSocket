/*
 * Implementation of AES for PuTTY using AES-NI
 * instuction set expansion was made by:
 * @author Pavel Kryukov <kryukov@frtk.ru>
 * @author Maxim Kuznetsov <maks.kuznetsov@gmail.com>
 * @author Svyatoslav Kuzmich <svatoslav1@gmail.com>
 *
 * For Putty AES NI project
 * http://pavelkryukov.github.io/putty-aes-ni/
 */

#include <assert.h>
#include <stdlib.h>

typedef uint32_t word32;

#define MAX_NR 14		            /* max no of rounds */
#define NB 4                        /* no of words in cipher blk */

/*
 * Select appropriate inline keyword for the compiler
 */
#if defined __GNUC__ || defined __clang__
#    define INLINE __inline__
#elif defined (_MSC_VER)
#    define INLINE __forceinline
#else
#    define INLINE
#endif

typedef struct cf_aes_context_ni {
    word32 keysched_buf[(MAX_NR + 1) * NB + 3];
    word32 invkeysched_buf[(MAX_NR + 1) * NB + 3];
    word32 *keysched, *invkeysched;
    int Nr; /* number of rounds */
} cf_aes_context_ni;


/*
 * Check of compiler version
 */
#ifdef _FORCE_AES_NI
#   define COMPILER_SUPPORTS_AES_NI
#elif defined(__clang__)
#   if (__clang_major__ > 3 || (__clang_major__ == 3 && __clang_minor__ >= 8)) && (defined(__x86_64__) || defined(__i386))
#       define COMPILER_SUPPORTS_AES_NI
#   endif
#elif defined(__GNUC__)
#    if (__GNUC__ > 4 || (__GNUC__ == 4 && __GNUC_MINOR__ >= 4)) && (defined(__x86_64__) || defined(__i386))
#       define COMPILER_SUPPORTS_AES_NI
#    endif
#elif defined (_MSC_VER)
#   if (defined(_M_X64) || defined(_M_IX86)) && _MSC_FULL_VER >= 150030729
#      define COMPILER_SUPPORTS_AES_NI
#   endif
#endif

#ifdef COMPILER_SUPPORTS_AES_NI

/*
 * Set target architecture for Clang and GCC
 */
#if !defined(__clang__) && defined(__GNUC__)
#    pragma GCC target("aes")
#    pragma GCC target("sse4.1")
#endif

#if defined(__clang__) || (defined(__GNUC__) && (__GNUC__ > 4 || (__GNUC__ == 4 && __GNUC_MINOR__ >= 8)))
#    define FUNC_ISA __attribute__ ((target("sse4.1,aes")))
#else
#    define FUNC_ISA
#endif

#include <wmmintrin.h>
#include <smmintrin.h>

/*
 * Determinators of CPU type
 */
#if defined(__clang__) || defined(__GNUC__)

#include <cpuid.h>
INLINE static int supports_aes_ni()
{
    unsigned int CPUInfo[4];
    __cpuid(1, CPUInfo[0], CPUInfo[1], CPUInfo[2], CPUInfo[3]);
    return (CPUInfo[2] & (1 << 25)) && (CPUInfo[2] & (1 << 19)); /* Check AES and SSE4.1 */
}

#else /* defined(__clang__) || defined(__GNUC__) */

#include <intrin.h>

INLINE static int supports_aes_ni()
{
    int CPUInfo[4];
    __cpuid(CPUInfo, 1);
    return (CPUInfo[2] & (1 << 25)) && (CPUInfo[2] & (1 << 19)); /* Check AES and SSE4.1 */
}

#endif /* defined(__clang__) || defined(__GNUC__) */

/*
 * Wrapper of SHUFPD instruction for MSVC
 */
#ifdef _MSC_VER
INLINE static __m128i mm_shuffle_pd_i0(__m128i a, __m128i b)
{
    union {
        __m128i i;
        __m128d d;
    } au, bu, ru;
    au.i = a;
    bu.i = b;
    ru.d = _mm_shuffle_pd(au.d, bu.d, 0);
    return ru.i;
}

INLINE static __m128i mm_shuffle_pd_i1(__m128i a, __m128i b)
{
    union {
        __m128i i;
        __m128d d;
    } au, bu, ru;
    au.i = a;
    bu.i = b;
    ru.d = _mm_shuffle_pd(au.d, bu.d, 1);
    return ru.i;
}
#else
#define mm_shuffle_pd_i0(a, b) ((__m128i)_mm_shuffle_pd((__m128d)a, (__m128d)b, 0));
#define mm_shuffle_pd_i1(a, b) ((__m128i)_mm_shuffle_pd((__m128d)a, (__m128d)b, 1));
#endif

/*
 * AES-NI key expansion assist functions
 */
FUNC_ISA
INLINE static __m128i AES_128_ASSIST (__m128i temp1, __m128i temp2)
{
    __m128i temp3;
    temp2 = _mm_shuffle_epi32 (temp2 ,0xff);
    temp3 = _mm_slli_si128 (temp1, 0x4);
    temp1 = _mm_xor_si128 (temp1, temp3);
    temp3 = _mm_slli_si128 (temp3, 0x4);
    temp1 = _mm_xor_si128 (temp1, temp3);
    temp3 = _mm_slli_si128 (temp3, 0x4);
    temp1 = _mm_xor_si128 (temp1, temp3);
    temp1 = _mm_xor_si128 (temp1, temp2);
    return temp1;
}

#ifdef AES192_NI
FUNC_ISA
INLINE static void KEY_192_ASSIST(__m128i* temp1, __m128i * temp2, __m128i * temp3)
{
    __m128i temp4;
    *temp2 = _mm_shuffle_epi32 (*temp2, 0x55);
    temp4 = _mm_slli_si128 (*temp1, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    *temp1 = _mm_xor_si128 (*temp1, *temp2);
    *temp2 = _mm_shuffle_epi32(*temp1, 0xff);
    temp4 = _mm_slli_si128 (*temp3, 0x4);
    *temp3 = _mm_xor_si128 (*temp3, temp4);
    *temp3 = _mm_xor_si128 (*temp3, *temp2);
}
#endif

FUNC_ISA
INLINE static void KEY_256_ASSIST_1(__m128i* temp1, __m128i * temp2)
{
    __m128i temp4;
    *temp2 = _mm_shuffle_epi32(*temp2, 0xff);
    temp4 = _mm_slli_si128 (*temp1, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp1 = _mm_xor_si128 (*temp1, temp4);
    *temp1 = _mm_xor_si128 (*temp1, *temp2);
}

FUNC_ISA
INLINE static void KEY_256_ASSIST_2(__m128i* temp1, __m128i * temp3)
{
    __m128i temp2,temp4;
    temp4 = _mm_aeskeygenassist_si128 (*temp1, 0x0);
    temp2 = _mm_shuffle_epi32(temp4, 0xaa);
    temp4 = _mm_slli_si128 (*temp3, 0x4);
    *temp3 = _mm_xor_si128 (*temp3, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp3 = _mm_xor_si128 (*temp3, temp4);
    temp4 = _mm_slli_si128 (temp4, 0x4);
    *temp3 = _mm_xor_si128 (*temp3, temp4);
    *temp3 = _mm_xor_si128 (*temp3, temp2);
}

/*
 * AES-NI key expansion core
 */
FUNC_ISA
static void AES_128_Key_Expansion (const unsigned char *userkey, __m128i *key)
{
    __m128i temp1, temp2;
    temp1 = _mm_loadu_si128((__m128i*)userkey);
    key[0] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1 ,0x1);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[1] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x2);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[2] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x4);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[3] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x8);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[4] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x10);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[5] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x20);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[6] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x40);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[7] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x80);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[8] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x1b);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[9] = temp1;
    temp2 = _mm_aeskeygenassist_si128 (temp1,0x36);
    temp1 = AES_128_ASSIST(temp1, temp2);
    key[10] = temp1;
}

FUNC_ISA
static void AES_192_Key_Expansion (const unsigned char *userkey, __m128i *key)
{
#ifdef AES192_NI
    __m128i temp1, temp2, temp3;
    temp1 = _mm_loadu_si128((__m128i*)userkey);
    temp3 = _mm_loadu_si128((__m128i*)(userkey+16));
    key[0]=temp1;
    key[1]=temp3;
    temp2=_mm_aeskeygenassist_si128 (temp3,0x1);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[1] = mm_shuffle_pd_i0(key[1], temp1);
    key[2] = mm_shuffle_pd_i1(temp1, temp3);
    temp2=_mm_aeskeygenassist_si128 (temp3,0x2);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[3]=temp1;
    key[4]=temp3;
    temp2=_mm_aeskeygenassist_si128 (temp3,0x4);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[4] = mm_shuffle_pd_i0(key[4], temp1);
    key[5] = mm_shuffle_pd_i1(temp1, temp3);
    temp2=_mm_aeskeygenassist_si128 (temp3,0x8);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[6]=temp1;
    key[7]=temp3;
    temp2=_mm_aeskeygenassist_si128 (temp3,0x10);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[7] = mm_shuffle_pd_i0(key[7], temp1);
    key[8] = mm_shuffle_pd_i1(temp1, temp3);
    temp2=_mm_aeskeygenassist_si128 (temp3,0x20);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[9]=temp1;
    key[10]=temp3;
    temp2=_mm_aeskeygenassist_si128 (temp3,0x40);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[10] = mm_shuffle_pd_i0(key[10], temp1);
    key[11] = mm_shuffle_pd_i1(temp1, temp3);
    temp2=_mm_aeskeygenassist_si128 (temp3,0x80);
    KEY_192_ASSIST(&temp1, &temp2, &temp3);
    key[12]=temp1;
    key[13]=temp3;
#else
    assert(0);
#endif
}

FUNC_ISA
static void AES_256_Key_Expansion (const unsigned char *userkey, __m128i *key)
{
    __m128i temp1, temp2, temp3;
    temp1 = _mm_loadu_si128((__m128i*)userkey);
    temp3 = _mm_loadu_si128((__m128i*)(userkey+16));
    key[0] = temp1;
    key[1] = temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x01);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[2]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[3]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x02);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[4]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[5]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x04);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[6]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[7]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x08);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[8]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[9]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x10);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[10]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[11]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x20);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[12]=temp1;
    KEY_256_ASSIST_2(&temp1, &temp3);
    key[13]=temp3;
    temp2 = _mm_aeskeygenassist_si128 (temp3,0x40);
    KEY_256_ASSIST_1(&temp1, &temp2);
    key[14]=temp1;
}

/*
 * AES-NI encrypt/decrypt core
 */
FUNC_ISA
static void cf_aes_encrypt_ni(cf_aes_context_ni * ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    __m128i enc;
    __m128i* keysched = (__m128i*)ctx->keysched;

    enc  = _mm_loadu_si128((__m128i*)in);
    enc  = _mm_xor_si128(enc, *keysched);
    switch (ctx->Nr) {
    case 14:
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
    case 12:
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
    case 10:
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenc_si128(enc, *(++keysched));
        enc = _mm_aesenclast_si128(enc, *(++keysched));
        break;
    default:
        assert(0);
    }

    _mm_storeu_si128((__m128i*)out, enc);
}

FUNC_ISA
static void cf_aes_decrypt_ni(cf_aes_context_ni * ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    __m128i dec;
    __m128i* keysched = (__m128i*)ctx->invkeysched;

    dec = _mm_loadu_si128((__m128i*)in);
    dec = _mm_xor_si128(dec, *keysched);
    switch (ctx->Nr) {
    case 14:
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
    case 12:
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
    case 10:
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdec_si128(dec, *(++keysched));
        dec = _mm_aesdeclast_si128(dec, *(++keysched));
        break;
    default:
        assert(0);
    }
    _mm_storeu_si128((__m128i*)out, dec);
}

FUNC_ISA
static void aes_inv_key_10(cf_aes_context_ni * ctx)
{
    __m128i* keysched = (__m128i*)ctx->keysched;
    __m128i* invkeysched = (__m128i*)ctx->invkeysched;

    *(invkeysched + 10) = *(keysched + 0);
    *(invkeysched + 9) = _mm_aesimc_si128(*(keysched + 1));
    *(invkeysched + 8) = _mm_aesimc_si128(*(keysched + 2));
    *(invkeysched + 7) = _mm_aesimc_si128(*(keysched + 3));
    *(invkeysched + 6) = _mm_aesimc_si128(*(keysched + 4));
    *(invkeysched + 5) = _mm_aesimc_si128(*(keysched + 5));
    *(invkeysched + 4) = _mm_aesimc_si128(*(keysched + 6));
    *(invkeysched + 3) = _mm_aesimc_si128(*(keysched + 7));
    *(invkeysched + 2) = _mm_aesimc_si128(*(keysched + 8));
    *(invkeysched + 1) = _mm_aesimc_si128(*(keysched + 9));
    *(invkeysched + 0) = *(keysched + 10);
}

FUNC_ISA
static void aes_inv_key_12(cf_aes_context_ni * ctx)
{
#ifdef AES192_NI
    __m128i* keysched = (__m128i*)ctx->keysched;
    __m128i* invkeysched = (__m128i*)ctx->invkeysched;

    *(invkeysched + 12) = *(keysched + 0);
    *(invkeysched + 11) = _mm_aesimc_si128(*(keysched + 1));
    *(invkeysched + 10) = _mm_aesimc_si128(*(keysched + 2));
    *(invkeysched + 9) = _mm_aesimc_si128(*(keysched + 3));
    *(invkeysched + 8) = _mm_aesimc_si128(*(keysched + 4));
    *(invkeysched + 7) = _mm_aesimc_si128(*(keysched + 5));
    *(invkeysched + 6) = _mm_aesimc_si128(*(keysched + 6));
    *(invkeysched + 5) = _mm_aesimc_si128(*(keysched + 7));
    *(invkeysched + 4) = _mm_aesimc_si128(*(keysched + 8));
    *(invkeysched + 3) = _mm_aesimc_si128(*(keysched + 9));
    *(invkeysched + 2) = _mm_aesimc_si128(*(keysched + 10));
    *(invkeysched + 1) = _mm_aesimc_si128(*(keysched + 11));
    *(invkeysched + 0) = *(keysched + 12);
#else
    assert(0);
#endif
}

FUNC_ISA
static void aes_inv_key_14(cf_aes_context_ni * ctx)
{
    __m128i* keysched = (__m128i*)ctx->keysched;
    __m128i* invkeysched = (__m128i*)ctx->invkeysched;

    *(invkeysched + 14) = *(keysched + 0);
    *(invkeysched + 13) = _mm_aesimc_si128(*(keysched + 1));
    *(invkeysched + 12) = _mm_aesimc_si128(*(keysched + 2));
    *(invkeysched + 11) = _mm_aesimc_si128(*(keysched + 3));
    *(invkeysched + 10) = _mm_aesimc_si128(*(keysched + 4));
    *(invkeysched + 9) = _mm_aesimc_si128(*(keysched + 5));
    *(invkeysched + 8) = _mm_aesimc_si128(*(keysched + 6));
    *(invkeysched + 7) = _mm_aesimc_si128(*(keysched + 7));
    *(invkeysched + 6) = _mm_aesimc_si128(*(keysched + 8));
    *(invkeysched + 5) = _mm_aesimc_si128(*(keysched + 9));
    *(invkeysched + 4) = _mm_aesimc_si128(*(keysched + 10));
    *(invkeysched + 3) = _mm_aesimc_si128(*(keysched + 11));
    *(invkeysched + 2) = _mm_aesimc_si128(*(keysched + 12));
    *(invkeysched + 1) = _mm_aesimc_si128(*(keysched + 13));
    *(invkeysched + 0) = *(keysched + 14);
}

/*
 * Set up an cf_aes_context_ni. `keylen' is measured in
 * bytes; it can be either 16 (128-bit), 24 (192-bit), or 32
 * (256-bit).
 */
static int cf_aes_setup_ni(cf_aes_context_ni *ctx, const unsigned char *key, int keylen)
{
    size_t bufaddr;

    if (!supports_aes_ni())
        return 0;
    ctx->Nr = 6 + (keylen / 4); /* Number of rounds */

    /* Ensure the key schedule arrays are 16-byte aligned */
    bufaddr = (size_t)ctx->keysched_buf;
    ctx->keysched = ctx->keysched_buf +
        (0xF & (~bufaddr+1)) / sizeof(word32);
    assert((size_t)ctx->keysched % 16 == 0);
    bufaddr = (size_t)ctx->invkeysched_buf;
    ctx->invkeysched = ctx->invkeysched_buf +
        (0xF & (~bufaddr+1)) / sizeof(word32);
    assert((size_t)ctx->invkeysched % 16 == 0);


    __m128i *keysched = (__m128i*)ctx->keysched;

    /*
     * Now do the key setup itself.
     */
    switch (keylen) {
      case 16:
        AES_128_Key_Expansion (key, keysched);
        break;
      case 24:
        AES_192_Key_Expansion (key, keysched);
        break;
      case 32:
        AES_256_Key_Expansion (key, keysched);
        break;
      default:
        assert(0);
    }

    /*
     * Now prepare the modified keys for the inverse cipher.
     */
    switch (ctx->Nr) {
      case 10:
        aes_inv_key_10(ctx);
        break;
      case 12:
        aes_inv_key_12(ctx);
        break;
      case 14:
        aes_inv_key_14(ctx);
        break;
      default:
        assert(0);
    }
    return 1;
}

#else /* COMPILER_SUPPORTS_AES_NI */

static void aes_setup_ni(cf_aes_context_ni * ctx, unsigned char *key, int keylen)
{
    assert(0);
}

INLINE static int supports_aes_ni()
{
    return 0;
}

#endif /* COMPILER_SUPPORTS_AES_NI */
