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

#define NB (AES_BLOCKSZ / 4)                /* no of uint32_t in cipher blk */

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

typedef struct cf_aes_ni_context {
    uint32_t rounds;
    uint32_t ks_e[(CF_AES_MAXROUNDS + 1) * NB + 3];
    uint32_t ks_d[(CF_AES_MAXROUNDS + 1) * NB + 3];
    __m128i *keysched_e, *keysched_d;
} cf_aes_ni_context;


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
 * AES-NI encrypt/decrypt core
 */
FUNC_ISA
static void cf_aes_ni_encrypt(cf_aes_ni_context *ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    __m128i *keysched = (__m128i *)ctx->keysched_e;
    __m128i enc = _mm_xor_si128(_mm_loadu_si128((__m128i*)in), *keysched++);
    switch (ctx->rounds) {
    case 14:
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
    case 12:
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
    case 10:
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenc_si128(enc, *keysched++);
        enc = _mm_aesenclast_si128(enc, *keysched++);
        break;
    default:
        assert(0);
    }

    _mm_storeu_si128((__m128i*)out, enc);
}

FUNC_ISA
static void cf_aes_ni_decrypt(cf_aes_ni_context *ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    __m128i *keysched = (__m128i *)ctx->keysched_d;
    __m128i dec = _mm_xor_si128(_mm_loadu_si128((__m128i*)in), *keysched++);
    switch (ctx->rounds) {
    case 14:
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
    case 12:
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
    case 10:
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdec_si128(dec, *keysched++);
        dec = _mm_aesdeclast_si128(dec, *keysched++);
        break;
    default:
        assert(0);
    }
    _mm_storeu_si128((__m128i*)out, dec);
}

/*
 * The main key expansion.
 */
static FUNC_ISA void cf_aes_ni_key_expand(const unsigned char *key, size_t key_words, 
                                          __m128i *keysched_e, __m128i *keysched_d)
{
    const uint8_t key_setup_round_constants[] = {
        /* The first few powers of X in GF(2^8), used during key setup.
         * This can safely be a lookup table without side channel risks,
         * because key setup iterates through it once in a standard way
         * regardless of the key. */
        0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1b, 0x36,
    };
    size_t rounds = key_words + 6;
    size_t sched_words = (rounds + 1) * NB;

    /*
     * Store the key schedule as 32-bit integers during expansion, so
     * that it's easy to refer back to individual previous words. We
     * collect them into the final __m128i form at the end.
     */
    uint32_t sched[(CF_AES_MAXROUNDS + 1) * NB];

    unsigned rconpos = 0;

    for (size_t i = 0; i < sched_words; i++) {
        if (i < key_words) {
            sched[i] = read32_le(key + 4 * i);
        } else {
            uint32_t temp = sched[i - 1];

            bool rotate_and_round_constant = (i % key_words == 0);
            bool only_sub = (key_words == 8 && i % 8 == 4);

            if (rotate_and_round_constant) {
                __m128i v = _mm_setr_epi32(0,temp,0,0);
                v = _mm_aeskeygenassist_si128(v, 0);
                temp = _mm_extract_epi32(v, 1);

                assert(rconpos < _countof(key_setup_round_constants));
                temp ^= key_setup_round_constants[rconpos++];
            } else if (only_sub) {
                __m128i v = _mm_setr_epi32(0,temp,0,0);
                v = _mm_aeskeygenassist_si128(v, 0);
                temp = _mm_extract_epi32(v, 0);
            }

            sched[i] = sched[i - key_words] ^ temp;
        }
    }

    /*
     * Combine the key schedule words into __m128i vectors and store
     * them in the output context.
     */
    for (size_t round = 0; round <= rounds; round++)
        keysched_e[round] = _mm_setr_epi32(
            sched[4*round  ], sched[4*round+1],
            sched[4*round+2], sched[4*round+3]);

    memset(sched, 0, sizeof(sched));

    /*
     * Now prepare the modified keys for the inverse cipher.
     */
    for (size_t eround = 0; eround <= rounds; eround++) {
        size_t dround = rounds - eround;
        __m128i rkey = keysched_e[eround];
        if (eround && dround)      /* neither first nor last */
            rkey = _mm_aesimc_si128(rkey);
        keysched_d[dround] = rkey;
    }
}

/*
 * Set up an cf_aes_ni_context. `keylen' is measured in
 * bytes; it can be either 16 (128-bit), 24 (192-bit), or 32
 * (256-bit).
 */
static int cf_aes_ni_setup(cf_aes_ni_context *ctx, const unsigned char *key, int keylen)
{
    size_t bufaddr;

    if (!supports_aes_ni())
        return 0;
    ctx->rounds = 6 + (keylen / 4);

    /* Ensure the key schedule arrays are 16-byte aligned */
    bufaddr = (size_t)ctx->ks_e;
    ctx->keysched_e = (__m128i *)(ctx->ks_e + (0xF & (~bufaddr+1)) / sizeof(uint32_t));
    assert((size_t)ctx->keysched_e % 16 == 0);

    bufaddr = (size_t)ctx->ks_d;
    ctx->keysched_d = (__m128i *)(ctx->ks_d + (0xF & (~bufaddr+1)) / sizeof(uint32_t));
    assert((size_t)ctx->keysched_d % 16 == 0);

    cf_aes_ni_key_expand(key, keylen / sizeof(uint32_t), ctx->keysched_e, ctx->keysched_d);
    return 1;
}

#else /* COMPILER_SUPPORTS_AES_NI */

FUNC_ISA
static void cf_aes_ni_encrypt(cf_aes_ni_context * ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    assert(0);
}

FUNC_ISA
static void cf_aes_ni_decrypt(cf_aes_ni_context * ctx, 
                              const uint8_t in[AES_BLOCKSZ],
                              uint8_t out[AES_BLOCKSZ])
{
    assert(0);
}

static void cf_aes_ni_setup(cf_aes_ni_context * ctx, unsigned char *key, int keylen)
{
    assert(0);
}

static int supports_aes_ni()
{
    return 0;
}

#endif /* COMPILER_SUPPORTS_AES_NI */

#undef INLINE
