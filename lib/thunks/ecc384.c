#include "ecc384.h"

#include <string.h>

#ifndef MAX_TRIES
#define MAX_TRIES 16

typedef unsigned int uint;

#if defined(__SIZEOF_INT128__) || ((__clang_major__ * 100 + __clang_minor__) >= 302)
    #define SUPPORTS_INT128 1
#else
    #define SUPPORTS_INT128 0
#endif

#if SUPPORTS_INT128
typedef unsigned __int128 uint128_t;
#else
typedef struct
{
    uint64_t m_low;
    uint64_t m_high;
} uint128_t;
#endif

#define CONCAT1(a, b) a##b
#define CONCAT(a, b) CONCAT1(a, b)

#define Curve_P_16 {0xFFFFFFFFFFFFFFFF, 0xFFFFFFFDFFFFFFFF}
#define Curve_P_24 {0xFFFFFFFFFFFFFFFFull, 0xFFFFFFFFFFFFFFFEull, 0xFFFFFFFFFFFFFFFFull}
#define Curve_P_32 {0xFFFFFFFFFFFFFFFFull, 0x00000000FFFFFFFFull, 0x0000000000000000ull, 0xFFFFFFFF00000001ull}
#define Curve_P_48 {0x00000000FFFFFFFF, 0xFFFFFFFF00000000, 0xFFFFFFFFFFFFFFFE, 0xFFFFFFFFFFFFFFFF, 0xFFFFFFFFFFFFFFFF, 0xFFFFFFFFFFFFFFFF}

#define Curve_B_16 {0xD824993C2CEE5ED3, 0xE87579C11079F43D}
#define Curve_B_24 {0xFEB8DEECC146B9B1ull, 0x0FA7E9AB72243049ull, 0x64210519E59C80E7ull}
#define Curve_B_32 {0x3BCE3C3E27D2604Bull, 0x651D06B0CC53B0F6ull, 0xB3EBBD55769886BCull, 0x5AC635D8AA3A93E7ull}
#define Curve_B_48 {0x2A85C8EDD3EC2AEF, 0xC656398D8A2ED19D, 0x0314088F5013875A, 0x181D9C6EFE814112, 0x988E056BE3F82D19, 0xB3312FA7E23EE7E4}

#define Curve_G_16 { \
    {0x0C28607CA52C5B86, 0x161FF7528B899B2D}, \
    {0xC02DA292DDED7A83, 0xCF5AC8395BAFEB13}}

#define Curve_G_24 { \
    {0xF4FF0AFD82FF1012ull, 0x7CBF20EB43A18800ull, 0x188DA80EB03090F6ull}, \
    {0x73F977A11E794811ull, 0x631011ED6B24CDD5ull, 0x07192B95FFC8DA78ull}}
    
#define Curve_G_32 { \
    {0xF4A13945D898C296ull, 0x77037D812DEB33A0ull, 0xF8BCE6E563A440F2ull, 0x6B17D1F2E12C4247ull}, \
    {0xCBB6406837BF51F5ull, 0x2BCE33576B315ECEull, 0x8EE7EB4A7C0F9E16ull, 0x4FE342E2FE1A7F9Bull}}

#define Curve_G_48 { \
    {0x3A545E3872760AB7, 0x5502F25DBF55296C, 0x59F741E082542A38, 0x6E1D3B628BA79B98, 0x8EB1C71EF320AD74, 0xAA87CA22BE8B0537}, \
    {0x7A431D7C90EA0E5F, 0x0A60B1CE1D7E819D, 0xE9DA3113B5F0B8C0, 0xF8F41DBD289A147C, 0x5D9E98BF9292DC29, 0x3617DE4A96262C6F}}

#define Curve_N_16 {0x75A30D1B9038A115, 0xFFFFFFFE00000000}
#define Curve_N_24 {0x146BC9B1B4D22831ull, 0xFFFFFFFF99DEF836ull, 0xFFFFFFFFFFFFFFFFull}
#define Curve_N_32 {0xF3B9CAC2FC632551ull, 0xBCE6FAADA7179E84ull, 0xFFFFFFFFFFFFFFFFull, 0xFFFFFFFF00000000ull}
#define Curve_N_48 {0xECEC196ACCC52973, 0x581A0DB248B0A77A, 0xC7634D81F4372DDF, 0xFFFFFFFFFFFFFFFF, 0xFFFFFFFFFFFFFFFF, 0xFFFFFFFFFFFFFFFF}
#endif

static uint64_t g_curve_p_384[NUM_ECC_DIGITS_384] = CONCAT(Curve_P_, ECC_CURVE_384);
static uint64_t g_curve_b_384[NUM_ECC_DIGITS_384] = CONCAT(Curve_B_, ECC_CURVE_384);
static EccPoint384 g_curve_G_384 = CONCAT(Curve_G_, ECC_CURVE_384);
static uint64_t g_curve_n_384[NUM_ECC_DIGITS_384] = CONCAT(Curve_N_, ECC_CURVE_384);

#if (defined(_WIN32) || defined(_WIN64))
/* Windows */
/*
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <wincrypt.h>


static int getRandomNumber(uint64_t *p_vli)
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
*/

#else /* _WIN32 */

/* Assume that we are using a POSIX-like system with /dev/urandom or /dev/random. */
#include <sys/types.h>
#include <fcntl.h>
#include <unistd.h>

#ifndef O_CLOEXEC
    #define O_CLOEXEC 0
#endif

static int getRandomNumber(uint64_t *p_vli)
{
    int l_fd = open("/dev/urandom", O_RDONLY | O_CLOEXEC);
    if(l_fd == -1)
    {
        l_fd = open("/dev/random", O_RDONLY | O_CLOEXEC);
        if(l_fd == -1)
        {
            return 0;
        }
    }
    
    char *l_ptr = (char *)p_vli;
    size_t l_left = ECC_BYTES_384;
    while(l_left > 0)
    {
        int l_read = read(l_fd, l_ptr, l_left);
        if(l_read <= 0)
        { // read failed
            close(l_fd);
            return 0;
        }
        l_left -= l_read;
        l_ptr += l_read;
    }
    
    close(l_fd);
    return 1;
}

#endif /* _WIN32 */

static void vli_clear384(uint64_t *p_vli)
{
    uint i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        p_vli[i] = 0;
    }
}

/* Returns 1 if p_vli == 0, 0 otherwise. */
static int vli_isZero384(uint64_t *p_vli)
{
    uint i;
    for(i = 0; i < NUM_ECC_DIGITS_384; ++i)
    {
        if(p_vli[i])
        {
            return 0;
        }
    }
    return 1;
}

/* Returns nonzero if bit p_bit of p_vli is set. */
static uint64_t vli_testBit384(uint64_t *p_vli, uint p_bit)
{
    return (p_vli[p_bit/64] & ((uint64_t)1 << (p_bit % 64)));
}

/* Counts the number of 64-bit "digits" in p_vli. */
static uint vli_numDigits384(uint64_t *p_vli)
{
    int i;
    /* Search from the end until we find a non-zero digit.
       We do it in reverse because we expect that most digits will be nonzero. */
    for(i = NUM_ECC_DIGITS_384 - 1; i >= 0 && p_vli[i] == 0; --i)
    {
    }

    return (i + 1);
}

/* Counts the number of bits required for p_vli. */
static uint vli_numBits384(uint64_t *p_vli)
{
    uint i;
    uint64_t l_digit;
    
    uint l_numDigits = vli_numDigits384(p_vli);
    if(l_numDigits == 0)
    {
        return 0;
    }

    l_digit = p_vli[l_numDigits - 1];
    for(i=0; l_digit; ++i)
    {
        l_digit >>= 1;
    }
    
    return ((l_numDigits - 1) * 64 + i);
}

/* Sets p_dest = p_src. */
static void vli_set384(uint64_t *p_dest, uint64_t *p_src)
{
    uint i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        p_dest[i] = p_src[i];
    }
}

/* Returns sign of p_left - p_right. */
static int vli_cmp384(uint64_t *p_left, uint64_t *p_right)
{
    int i;
    for(i = NUM_ECC_DIGITS_384-1; i >= 0; --i)
    {
        if(p_left[i] > p_right[i])
        {
            return 1;
        }
        else if(p_left[i] < p_right[i])
        {
            return -1;
        }
    }
    return 0;
}

/* Computes p_result = p_in << c, returning carry. Can modify in place (if p_result == p_in). 0 < p_shift < 64. */
static uint64_t vli_lshift384(uint64_t *p_result, uint64_t *p_in, uint p_shift)
{
    uint64_t l_carry = 0;
    uint i;
    for(i = 0; i < NUM_ECC_DIGITS_384; ++i)
    {
        uint64_t l_temp = p_in[i];
        p_result[i] = (l_temp << p_shift) | l_carry;
        l_carry = l_temp >> (64 - p_shift);
    }
    
    return l_carry;
}

/* Computes p_vli = p_vli >> 1. */
static void vli_rshift1384(uint64_t *p_vli)
{
    uint64_t *l_end = p_vli;
    uint64_t l_carry = 0;
    
    p_vli += NUM_ECC_DIGITS_384;
    while(p_vli-- > l_end)
    {
        uint64_t l_temp = *p_vli;
        *p_vli = (l_temp >> 1) | l_carry;
        l_carry = l_temp << 63;
    }
}

/* Computes p_result = p_left + p_right, returning carry. Can modify in place. */
static uint64_t vli_add384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right)
{
    uint64_t l_carry = 0;
    uint i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        uint64_t l_sum = p_left[i] + p_right[i] + l_carry;
        if(l_sum != p_left[i])
        {
            l_carry = (l_sum < p_left[i]);
        }
        p_result[i] = l_sum;
    }
    return l_carry;
}

/* Computes p_result = p_left - p_right, returning borrow. Can modify in place. */
static uint64_t vli_sub384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right)
{
    uint64_t l_borrow = 0;
    uint i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        uint64_t l_diff = p_left[i] - p_right[i] - l_borrow;
        if(l_diff != p_left[i])
        {
            l_borrow = (l_diff > p_left[i]);
        }
        p_result[i] = l_diff;
    }
    return l_borrow;
}

#if SUPPORTS_INT128

/* Computes p_result = p_left * p_right. */
static void vli_mult384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right)
{
    uint128_t r01 = 0;
    uint64_t r2 = 0;
    
    uint i, k;
    
    /* Compute each digit of p_result in sequence, maintaining the carries. */
    for(k=0; k < NUM_ECC_DIGITS_384*2 - 1; ++k)
    {
        uint l_min = (k < NUM_ECC_DIGITS_384 ? 0 : (k + 1) - NUM_ECC_DIGITS_384);
        for(i=l_min; i<=k && i<NUM_ECC_DIGITS_384; ++i)
        {
            uint128_t l_product = (uint128_t)p_left[i] * p_right[k-i];
            r01 += l_product;
            r2 += (r01 < l_product);
        }
        p_result[k] = (uint64_t)r01;
        r01 = (r01 >> 64) | (((uint128_t)r2) << 64);
        r2 = 0;
    }
    
    p_result[NUM_ECC_DIGITS_384*2 - 1] = (uint64_t)r01;
}

/* Computes p_result = p_left^2. */
static void vli_square384(uint64_t *p_result, uint64_t *p_left)
{
    uint128_t r01 = 0;
    uint64_t r2 = 0;
    
    uint i, k;
    for(k=0; k < NUM_ECC_DIGITS_384*2 - 1; ++k)
    {
        uint l_min = (k < NUM_ECC_DIGITS_384 ? 0 : (k + 1) - NUM_ECC_DIGITS_384);
        for(i=l_min; i<=k && i<=k-i; ++i)
        {
            uint128_t l_product = (uint128_t)p_left[i] * p_left[k-i];
            if(i < k-i)
            {
                r2 += l_product >> 127;
                l_product *= 2;
            }
            r01 += l_product;
            r2 += (r01 < l_product);
        }
        p_result[k] = (uint64_t)r01;
        r01 = (r01 >> 64) | (((uint128_t)r2) << 64);
        r2 = 0;
    }
    
    p_result[NUM_ECC_DIGITS_384*2 - 1] = (uint64_t)r01;
}

#else /* #if SUPPORTS_INT128 */

static uint128_t mul_64_64_384(uint64_t p_left, uint64_t p_right)
{
    uint128_t l_result;
    
    uint64_t a0 = p_left & 0xffffffffull;
    uint64_t a1 = p_left >> 32;
    uint64_t b0 = p_right & 0xffffffffull;
    uint64_t b1 = p_right >> 32;
    
    uint64_t m0 = a0 * b0;
    uint64_t m1 = a0 * b1;
    uint64_t m2 = a1 * b0;
    uint64_t m3 = a1 * b1;
    
    m2 += (m0 >> 32);
    m2 += m1;
    if(m2 < m1)
    { // overflow
        m3 += 0x100000000ull;
    }
    
    l_result.m_low = (m0 & 0xffffffffull) | (m2 << 32);
    l_result.m_high = m3 + (m2 >> 32);
    
    return l_result;
}

static uint128_t add_128_128_384(uint128_t a, uint128_t b)
{
    uint128_t l_result;
    l_result.m_low = a.m_low + b.m_low;
    l_result.m_high = a.m_high + b.m_high + (l_result.m_low < a.m_low);
    return l_result;
}

static void vli_mult384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right)
{
    uint128_t r01 = {0, 0};
    uint64_t r2 = 0;
    
    uint i, k;
    
    /* Compute each digit of p_result in sequence, maintaining the carries. */
    for(k=0; k < NUM_ECC_DIGITS_384*2 - 1; ++k)
    {
        uint l_min = (k < NUM_ECC_DIGITS_384 ? 0 : (k + 1) - NUM_ECC_DIGITS_384);
        for(i=l_min; i<=k && i<NUM_ECC_DIGITS_384; ++i)
        {
            uint128_t l_product = mul_64_64_384(p_left[i], p_right[k-i]);
            r01 = add_128_128_384(r01, l_product);
            r2 += (r01.m_high < l_product.m_high);
        }
        p_result[k] = r01.m_low;
        r01.m_low = r01.m_high;
        r01.m_high = r2;
        r2 = 0;
    }
    
    p_result[NUM_ECC_DIGITS_384*2 - 1] = r01.m_low;
}

static void vli_square384(uint64_t *p_result, uint64_t *p_left)
{
    uint128_t r01 = {0, 0};
    uint64_t r2 = 0;
    
    uint i, k;
    for(k=0; k < NUM_ECC_DIGITS_384*2 - 1; ++k)
    {
        uint l_min = (k < NUM_ECC_DIGITS_384 ? 0 : (k + 1) - NUM_ECC_DIGITS_384);
        for(i=l_min; i<=k && i<=k-i; ++i)
        {
            uint128_t l_product = mul_64_64_384(p_left[i], p_left[k-i]);
            if(i < k-i)
            {
                r2 += l_product.m_high >> 63;
                l_product.m_high = (l_product.m_high << 1) | (l_product.m_low >> 63);
                l_product.m_low <<= 1;
            }
            r01 = add_128_128_384(r01, l_product);
            r2 += (r01.m_high < l_product.m_high);
        }
        p_result[k] = r01.m_low;
        r01.m_low = r01.m_high;
        r01.m_high = r2;
        r2 = 0;
    }
    
    p_result[NUM_ECC_DIGITS_384*2 - 1] = r01.m_low;
}

#endif /* SUPPORTS_INT128 */


/* Computes p_result = (p_left + p_right) % p_mod.
   Assumes that p_left < p_mod and p_right < p_mod, p_result != p_mod. */
static void vli_modAdd384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right, uint64_t *p_mod)
{
    uint64_t l_carry = vli_add384(p_result, p_left, p_right);
    if(l_carry || vli_cmp384(p_result, p_mod) >= 0)
    { /* p_result > p_mod (p_result = p_mod + remainder), so subtract p_mod to get remainder. */
        vli_sub384(p_result, p_result, p_mod);
    }
}

/* Computes p_result = (p_left - p_right) % p_mod.
   Assumes that p_left < p_mod and p_right < p_mod, p_result != p_mod. */
static void vli_modSub384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right, uint64_t *p_mod)
{
    uint64_t l_borrow = vli_sub384(p_result, p_left, p_right);
    if(l_borrow)
    { /* In this case, p_result == -diff == (max int) - diff.
         Since -x % d == d - x, we can get the correct result from p_result + p_mod (with overflow). */
        vli_add384(p_result, p_result, p_mod);
    }
}

#if ECC_CURVE_384 == secp128r1

/* Computes p_result = p_product % curve_p_384.
   See algorithm 5 and 6 from http://www.isys.uni-klu.ac.at/PDF/2001-0126-MT.pdf */
static void vli_mmod_fast384(uint64_t *p_result, uint64_t *p_product)
{
    uint64_t l_tmp[NUM_ECC_DIGITS_384];
    int l_carry;
    
    vli_set384(p_result, p_product);
    
    l_tmp[0] = p_product[2];
    l_tmp[1] = (p_product[3] & 0x1FFFFFFFFull) | (p_product[2] << 33);
    l_carry = vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = (p_product[2] >> 31) | (p_product[3] << 33);
    l_tmp[1] = (p_product[3] >> 31) | ((p_product[2] & 0xFFFFFFFF80000000ull) << 2);
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = (p_product[2] >> 62) | (p_product[3] << 2);
    l_tmp[1] = (p_product[3] >> 62) | ((p_product[2] & 0xC000000000000000ull) >> 29) | (p_product[3] << 35);
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = (p_product[3] >> 29);
    l_tmp[1] = ((p_product[3] & 0xFFFFFFFFE0000000ull) << 4);
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = (p_product[3] >> 60);
    l_tmp[1] = (p_product[3] & 0xFFFFFFFE00000000ull);
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = 0;
    l_tmp[1] = ((p_product[3] & 0xF000000000000000ull) >> 27);
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    while(l_carry || vli_cmp384(curve_p_384, p_result) != 1)
    {
        l_carry -= vli_sub384(p_result, p_result, curve_p_384);
    }
}

#elif ECC_CURVE_384 == secp192r1

/* Computes p_result = p_product % curve_p_384.
   See algorithm 5 and 6 from http://www.isys.uni-klu.ac.at/PDF/2001-0126-MT.pdf */
static void vli_mmod_fast384(uint64_t *p_result, uint64_t *p_product)
{
    uint64_t l_tmp[NUM_ECC_DIGITS_384];
    int l_carry;
    
    vli_set384(p_result, p_product);
    
    vli_set384(l_tmp, &p_product[3]);
    l_carry = vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = 0;
    l_tmp[1] = p_product[3];
    l_tmp[2] = p_product[4];
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    l_tmp[0] = l_tmp[1] = p_product[5];
    l_tmp[2] = 0;
    l_carry += vli_add384(p_result, p_result, l_tmp);
    
    while(l_carry || vli_cmp384(curve_p_384, p_result) != 1)
    {
        l_carry -= vli_sub384(p_result, p_result, curve_p_384);
    }
}

#elif ECC_CURVE_384 == secp256r1

/* Computes p_result = p_product % curve_p_384
   from http://www.nsa.gov/ia/_files/nist-routines.pdf */
static void vli_mmod_fast384(uint64_t *p_result, uint64_t *p_product)
{
    uint64_t l_tmp[NUM_ECC_DIGITS_384];
    int l_carry; // don't change to uint64_t as it stops working
    
    /* t */
    vli_set384(p_result, p_product);
    
    /* s1 */
    l_tmp[0] = 0;
    l_tmp[1] = p_product[5] & 0xffffffff00000000ull;
    l_tmp[2] = p_product[6];
    l_tmp[3] = p_product[7];
    l_carry = (int)vli_lshift384(l_tmp, l_tmp, 1);
    l_carry += (int)vli_add384(p_result, p_result, l_tmp);
    
    /* s2 */
    l_tmp[1] = p_product[6] << 32;
    l_tmp[2] = (p_product[6] >> 32) | (p_product[7] << 32);
    l_tmp[3] = p_product[7] >> 32;
    l_carry += (int)vli_lshift384(l_tmp, l_tmp, 1);
    l_carry += (int)vli_add384(p_result, p_result, l_tmp);
    
    /* s3 */
    l_tmp[0] = p_product[4];
    l_tmp[1] = p_product[5] & 0xffffffff;
    l_tmp[2] = 0;
    l_tmp[3] = p_product[7];
    l_carry += (int)vli_add384(p_result, p_result, l_tmp);
    
    /* s4 */
    l_tmp[0] = (p_product[4] >> 32) | (p_product[5] << 32);
    l_tmp[1] = (p_product[5] >> 32) | (p_product[6] & 0xffffffff00000000ull);
    l_tmp[2] = p_product[7];
    l_tmp[3] = (p_product[6] >> 32) | (p_product[4] << 32);
    l_carry += (int)vli_add384(p_result, p_result, l_tmp);
    
    /* d1 */
    l_tmp[0] = (p_product[5] >> 32) | (p_product[6] << 32);
    l_tmp[1] = (p_product[6] >> 32);
    l_tmp[2] = 0;
    l_tmp[3] = (p_product[4] & 0xffffffff) | (p_product[5] << 32);
    l_carry -= (int)vli_sub384(p_result, p_result, l_tmp);
    
    /* d2 */
    l_tmp[0] = p_product[6];
    l_tmp[1] = p_product[7];
    l_tmp[2] = 0;
    l_tmp[3] = (p_product[4] >> 32) | (p_product[5] & 0xffffffff00000000ull);
    l_carry -= (int)vli_sub384(p_result, p_result, l_tmp);
    
    /* d3 */
    l_tmp[0] = (p_product[6] >> 32) | (p_product[7] << 32);
    l_tmp[1] = (p_product[7] >> 32) | (p_product[4] << 32);
    l_tmp[2] = (p_product[4] >> 32) | (p_product[5] << 32);
    l_tmp[3] = (p_product[6] << 32);
    l_carry -= (int)vli_sub384(p_result, p_result, l_tmp);
    
    /* d4 */
    l_tmp[0] = p_product[7];
    l_tmp[1] = p_product[4] & 0xffffffff00000000ull;
    l_tmp[2] = p_product[5];
    l_tmp[3] = p_product[6] & 0xffffffff00000000ull;
    l_carry -= (int)vli_sub384(p_result, p_result, l_tmp);
    
    if(l_carry < 0)
    {
        do
        {
            l_carry += (int)vli_add384(p_result, p_result, curve_p_384);
        } while(l_carry < 0);
    }
    else
    {
        while(l_carry || vli_cmp384(curve_p_384, p_result) != 1)
        {
            l_carry -= (int)vli_sub384(p_result, p_result, curve_p_384);
        }
    }
}

#elif ECC_CURVE_384 == secp384r1

static void omega_mult384(uint64_t *p_result, uint64_t *p_right)
{
    uint64_t l_tmp[NUM_ECC_DIGITS_384];
    uint64_t l_carry, l_diff;
    
    /* Multiply by (2^128 + 2^96 - 2^32 + 1). */
    vli_set384(p_result, p_right); /* 1 */
    l_carry = vli_lshift384(l_tmp, p_right, 32);
    p_result[1 + NUM_ECC_DIGITS_384] = l_carry + vli_add384(p_result + 1, p_result + 1, l_tmp); /* 2^96 + 1 */
    p_result[2 + NUM_ECC_DIGITS_384] = vli_add384(p_result + 2, p_result + 2, p_right); /* 2^128 + 2^96 + 1 */
    l_carry += vli_sub384(p_result, p_result, l_tmp); /* 2^128 + 2^96 - 2^32 + 1 */
    l_diff = p_result[NUM_ECC_DIGITS_384] - l_carry;
    if(l_diff > p_result[NUM_ECC_DIGITS_384])
    { /* Propagate borrow if necessary. */
        uint i;
        for(i = 1 + NUM_ECC_DIGITS_384; ; ++i)
        {
            --p_result[i];
            if(p_result[i] != (uint64_t)-1)
            {
                break;
            }
        }
    }
    p_result[NUM_ECC_DIGITS_384] = l_diff;
}

/* Computes p_result = p_product % curve_p_384
    see PDF "Comparing Elliptic Curve Cryptography and RSA on 8-bit CPUs"
    section "Curve-Specific Optimizations" */
static void vli_mmod_fast384(uint64_t *p_result, uint64_t *p_product)
{
    uint64_t l_tmp[2*NUM_ECC_DIGITS_384];
     
    while(!vli_isZero384(p_product + NUM_ECC_DIGITS_384)) /* While c1 != 0 */
    {
        uint64_t l_carry = 0;
        uint i;
        
        vli_clear384(l_tmp);
        vli_clear384(l_tmp + NUM_ECC_DIGITS_384);
        omega_mult384(l_tmp, p_product + NUM_ECC_DIGITS_384); /* tmp = w * c1 */
        vli_clear384(p_product + NUM_ECC_DIGITS_384); /* p = c0 */
        
        /* (c1, c0) = c0 + w * c1 */
        for(i=0; i<NUM_ECC_DIGITS_384+3; ++i)
        {
            uint64_t l_sum = p_product[i] + l_tmp[i] + l_carry;
            if(l_sum != p_product[i])
            {
                l_carry = (l_sum < p_product[i]);
            }
            p_product[i] = l_sum;
        }
    }
    
    while(vli_cmp384(p_product, curve_p_384) > 0)
    {
        vli_sub384(p_product, p_product, curve_p_384);
    }
    vli_set384(p_result, p_product);
}

#endif

/* Computes p_result = (p_left * p_right) % curve_p_384. */
static void vli_modMult_fast384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right)
{
    uint64_t l_product[2 * NUM_ECC_DIGITS_384];
    vli_mult384(l_product, p_left, p_right);
    vli_mmod_fast384(p_result, l_product);
}

/* Computes p_result = p_left^2 % curve_p_384. */
static void vli_modSquare_fast384(uint64_t *p_result, uint64_t *p_left)
{
    uint64_t l_product[2 * NUM_ECC_DIGITS_384];
    vli_square384(l_product, p_left);
    vli_mmod_fast384(p_result, l_product);
}

#define EVEN(vli) (!(vli[0] & 1))
/* Computes p_result = (1 / p_input) % p_mod. All VLIs are the same size.
   See "From Euclid's GCD to Montgomery Multiplication to the Great Divide"
   https://labs.oracle.com/techrep/2001/smli_tr-2001-95.pdf */
static void vli_modInv384(uint64_t *p_result, uint64_t *p_input, uint64_t *p_mod)
{
    uint64_t a[NUM_ECC_DIGITS_384], b[NUM_ECC_DIGITS_384], u[NUM_ECC_DIGITS_384], v[NUM_ECC_DIGITS_384];
    uint64_t l_carry;
    int l_cmpResult;
    
    if(vli_isZero384(p_input))
    {
        vli_clear384(p_result);
        return;
    }

    vli_set384(a, p_input);
    vli_set384(b, p_mod);
    vli_clear384(u);
    u[0] = 1;
    vli_clear384(v);
    
    while((l_cmpResult = vli_cmp384(a, b)) != 0)
    {
        l_carry = 0;
        if(EVEN(a))
        {
            vli_rshift1384(a);
            if(!EVEN(u))
            {
                l_carry = vli_add384(u, u, p_mod);
            }
            vli_rshift1384(u);
            if(l_carry)
            {
                u[NUM_ECC_DIGITS_384-1] |= 0x8000000000000000ull;
            }
        }
        else if(EVEN(b))
        {
            vli_rshift1384(b);
            if(!EVEN(v))
            {
                l_carry = vli_add384(v, v, p_mod);
            }
            vli_rshift1384(v);
            if(l_carry)
            {
                v[NUM_ECC_DIGITS_384-1] |= 0x8000000000000000ull;
            }
        }
        else if(l_cmpResult > 0)
        {
            vli_sub384(a, a, b);
            vli_rshift1384(a);
            if(vli_cmp384(u, v) < 0)
            {
                vli_add384(u, u, p_mod);
            }
            vli_sub384(u, u, v);
            if(!EVEN(u))
            {
                l_carry = vli_add384(u, u, p_mod);
            }
            vli_rshift1384(u);
            if(l_carry)
            {
                u[NUM_ECC_DIGITS_384-1] |= 0x8000000000000000ull;
            }
        }
        else
        {
            vli_sub384(b, b, a);
            vli_rshift1384(b);
            if(vli_cmp384(v, u) < 0)
            {
                vli_add384(v, v, p_mod);
            }
            vli_sub384(v, v, u);
            if(!EVEN(v))
            {
                l_carry = vli_add384(v, v, p_mod);
            }
            vli_rshift1384(v);
            if(l_carry)
            {
                v[NUM_ECC_DIGITS_384-1] |= 0x8000000000000000ull;
            }
        }
    }
    
    vli_set384(p_result, u);
}

/* ------ Point operations ------ */

/* Returns 1 if p_point is the point at infinity, 0 otherwise. */
static int EccPoint_isZero384(EccPoint384 *p_point)
{
    return (vli_isZero384(p_point->x) && vli_isZero384(p_point->y));
}

/* Point multiplication algorithm using Montgomery's ladder with co-Z coordinates.
From http://eprint.iacr.org/2011/338.pdf
*/

/* Double in place */
static void EccPoint_double_jacobian384(uint64_t *X1, uint64_t *Y1, uint64_t *Z1)
{
    /* t1 = X, t2 = Y, t3 = Z */
    uint64_t t4[NUM_ECC_DIGITS_384];
    uint64_t t5[NUM_ECC_DIGITS_384];
    
    if(vli_isZero384(Z1))
    {
        return;
    }
    
    vli_modSquare_fast384(t4, Y1);   /* t4 = y1^2 */
    vli_modMult_fast384(t5, X1, t4); /* t5 = x1*y1^2 = A */
    vli_modSquare_fast384(t4, t4);   /* t4 = y1^4 */
    vli_modMult_fast384(Y1, Y1, Z1); /* t2 = y1*z1 = z3 */
    vli_modSquare_fast384(Z1, Z1);   /* t3 = z1^2 */
    
    vli_modAdd384(X1, X1, Z1, curve_p_384); /* t1 = x1 + z1^2 */
    vli_modAdd384(Z1, Z1, Z1, curve_p_384); /* t3 = 2*z1^2 */
    vli_modSub384(Z1, X1, Z1, curve_p_384); /* t3 = x1 - z1^2 */
    vli_modMult_fast384(X1, X1, Z1);    /* t1 = x1^2 - z1^4 */
    
    vli_modAdd384(Z1, X1, X1, curve_p_384); /* t3 = 2*(x1^2 - z1^4) */
    vli_modAdd384(X1, X1, Z1, curve_p_384); /* t1 = 3*(x1^2 - z1^4) */
    if(vli_testBit384(X1, 0))
    {
        uint64_t l_carry = vli_add384(X1, X1, curve_p_384);
        vli_rshift1384(X1);
        X1[NUM_ECC_DIGITS_384-1] |= l_carry << 63;
    }
    else
    {
        vli_rshift1384(X1);
    }
    /* t1 = 3/2*(x1^2 - z1^4) = B */
    
    vli_modSquare_fast384(Z1, X1);      /* t3 = B^2 */
    vli_modSub384(Z1, Z1, t5, curve_p_384); /* t3 = B^2 - A */
    vli_modSub384(Z1, Z1, t5, curve_p_384); /* t3 = B^2 - 2A = x3 */
    vli_modSub384(t5, t5, Z1, curve_p_384); /* t5 = A - x3 */
    vli_modMult_fast384(X1, X1, t5);    /* t1 = B * (A - x3) */
    vli_modSub384(t4, X1, t4, curve_p_384); /* t4 = B * (A - x3) - y1^4 = y3 */
    
    vli_set384(X1, Z1);
    vli_set384(Z1, Y1);
    vli_set384(Y1, t4);
}

/* Modify (x1, y1) => (x1 * z^2, y1 * z^3) */
static void apply_z384(uint64_t *X1, uint64_t *Y1, uint64_t *Z)
{
    uint64_t t1[NUM_ECC_DIGITS_384];

    vli_modSquare_fast384(t1, Z);    /* z^2 */
    vli_modMult_fast384(X1, X1, t1); /* x1 * z^2 */
    vli_modMult_fast384(t1, t1, Z);  /* z^3 */
    vli_modMult_fast384(Y1, Y1, t1); /* y1 * z^3 */
}

/* P = (x1, y1) => 2P, (x2, y2) => P' */
static void XYcZ_initial_double384(uint64_t *X1, uint64_t *Y1, uint64_t *X2, uint64_t *Y2, uint64_t *p_initialZ)
{
    uint64_t z[NUM_ECC_DIGITS_384];
    
    vli_set384(X2, X1);
    vli_set384(Y2, Y1);
    
    vli_clear384(z);
    z[0] = 1;
    if(p_initialZ)
    {
        vli_set384(z, p_initialZ);
    }

    apply_z384(X1, Y1, z);
    
    EccPoint_double_jacobian384(X1, Y1, z);
    
    apply_z384(X2, Y2, z);
}

/* Input P = (x1, y1, Z), Q = (x2, y2, Z)
   Output P' = (x1', y1', Z3), P + Q = (x3, y3, Z3)
   or P => P', Q => P + Q
*/
static void XYcZ_add384(uint64_t *X1, uint64_t *Y1, uint64_t *X2, uint64_t *Y2)
{
    /* t1 = X1, t2 = Y1, t3 = X2, t4 = Y2 */
    uint64_t t5[NUM_ECC_DIGITS_384];
    
    vli_modSub384(t5, X2, X1, curve_p_384); /* t5 = x2 - x1 */
    vli_modSquare_fast384(t5, t5);      /* t5 = (x2 - x1)^2 = A */
    vli_modMult_fast384(X1, X1, t5);    /* t1 = x1*A = B */
    vli_modMult_fast384(X2, X2, t5);    /* t3 = x2*A = C */
    vli_modSub384(Y2, Y2, Y1, curve_p_384); /* t4 = y2 - y1 */
    vli_modSquare_fast384(t5, Y2);      /* t5 = (y2 - y1)^2 = D */
    
    vli_modSub384(t5, t5, X1, curve_p_384); /* t5 = D - B */
    vli_modSub384(t5, t5, X2, curve_p_384); /* t5 = D - B - C = x3 */
    vli_modSub384(X2, X2, X1, curve_p_384); /* t3 = C - B */
    vli_modMult_fast384(Y1, Y1, X2);    /* t2 = y1*(C - B) */
    vli_modSub384(X2, X1, t5, curve_p_384); /* t3 = B - x3 */
    vli_modMult_fast384(Y2, Y2, X2);    /* t4 = (y2 - y1)*(B - x3) */
    vli_modSub384(Y2, Y2, Y1, curve_p_384); /* t4 = y3 */
    
    vli_set384(X2, t5);
}

/* Input P = (x1, y1, Z), Q = (x2, y2, Z)
   Output P + Q = (x3, y3, Z3), P - Q = (x3', y3', Z3)
   or P => P - Q, Q => P + Q
*/
static void XYcZ_addC384(uint64_t *X1, uint64_t *Y1, uint64_t *X2, uint64_t *Y2)
{
    /* t1 = X1, t2 = Y1, t3 = X2, t4 = Y2 */
    uint64_t t5[NUM_ECC_DIGITS_384];
    uint64_t t6[NUM_ECC_DIGITS_384];
    uint64_t t7[NUM_ECC_DIGITS_384];
    
    vli_modSub384(t5, X2, X1, curve_p_384); /* t5 = x2 - x1 */
    vli_modSquare_fast384(t5, t5);      /* t5 = (x2 - x1)^2 = A */
    vli_modMult_fast384(X1, X1, t5);    /* t1 = x1*A = B */
    vli_modMult_fast384(X2, X2, t5);    /* t3 = x2*A = C */
    vli_modAdd384(t5, Y2, Y1, curve_p_384); /* t4 = y2 + y1 */
    vli_modSub384(Y2, Y2, Y1, curve_p_384); /* t4 = y2 - y1 */

    vli_modSub384(t6, X2, X1, curve_p_384); /* t6 = C - B */
    vli_modMult_fast384(Y1, Y1, t6);    /* t2 = y1 * (C - B) */
    vli_modAdd384(t6, X1, X2, curve_p_384); /* t6 = B + C */
    vli_modSquare_fast384(X2, Y2);      /* t3 = (y2 - y1)^2 */
    vli_modSub384(X2, X2, t6, curve_p_384); /* t3 = x3 */
    
    vli_modSub384(t7, X1, X2, curve_p_384); /* t7 = B - x3 */
    vli_modMult_fast384(Y2, Y2, t7);    /* t4 = (y2 - y1)*(B - x3) */
    vli_modSub384(Y2, Y2, Y1, curve_p_384); /* t4 = y3 */
    
    vli_modSquare_fast384(t7, t5);      /* t7 = (y2 + y1)^2 = F */
    vli_modSub384(t7, t7, t6, curve_p_384); /* t7 = x3' */
    vli_modSub384(t6, t7, X1, curve_p_384); /* t6 = x3' - B */
    vli_modMult_fast384(t6, t6, t5);    /* t6 = (y2 + y1)*(x3' - B) */
    vli_modSub384(Y1, t6, Y1, curve_p_384); /* t2 = y3' */
    
    vli_set384(X1, t7);
}

static void EccPoint_mult384(EccPoint384 *p_result, EccPoint384 *p_point, uint64_t *p_scalar, uint64_t *p_initialZ)
{
    /* R0 and R1 */
    uint64_t Rx[2][NUM_ECC_DIGITS_384];
    uint64_t Ry[2][NUM_ECC_DIGITS_384];
    uint64_t z[NUM_ECC_DIGITS_384];
    
    int i, nb;
    
    vli_set384(Rx[1], p_point->x);
    vli_set384(Ry[1], p_point->y);

    XYcZ_initial_double384(Rx[1], Ry[1], Rx[0], Ry[0], p_initialZ);

    for(i = vli_numBits384(p_scalar) - 2; i > 0; --i)
    {
        nb = !vli_testBit384(p_scalar, i);
        XYcZ_addC384(Rx[1-nb], Ry[1-nb], Rx[nb], Ry[nb]);
        XYcZ_add384(Rx[nb], Ry[nb], Rx[1-nb], Ry[1-nb]);
    }

    nb = !vli_testBit384(p_scalar, 0);
    XYcZ_addC384(Rx[1-nb], Ry[1-nb], Rx[nb], Ry[nb]);
    
    /* Find final 1/Z value. */
    vli_modSub384(z, Rx[1], Rx[0], curve_p_384); /* X1 - X0 */
    vli_modMult_fast384(z, z, Ry[1-nb]);     /* Yb * (X1 - X0) */
    vli_modMult_fast384(z, z, p_point->x);   /* xP * Yb * (X1 - X0) */
    vli_modInv384(z, z, curve_p_384);            /* 1 / (xP * Yb * (X1 - X0)) */
    vli_modMult_fast384(z, z, p_point->y);   /* yP / (xP * Yb * (X1 - X0)) */
    vli_modMult_fast384(z, z, Rx[1-nb]);     /* Xb * yP / (xP * Yb * (X1 - X0)) */
    /* End 1/Z calculation */

    XYcZ_add384(Rx[nb], Ry[nb], Rx[1-nb], Ry[1-nb]);
    
    apply_z384(Rx[0], Ry[0], z);
    
    vli_set384(p_result->x, Rx[0]);
    vli_set384(p_result->y, Ry[0]);
}

static void ecc_bytes2native384(uint64_t p_native[NUM_ECC_DIGITS_384], const uint8_t p_bytes[ECC_BYTES_384])
{
    unsigned i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        const uint8_t *p_digit = p_bytes + 8 * (NUM_ECC_DIGITS_384 - 1 - i);
        p_native[i] = ((uint64_t)p_digit[0] << 56) | ((uint64_t)p_digit[1] << 48) | ((uint64_t)p_digit[2] << 40) | ((uint64_t)p_digit[3] << 32) |
            ((uint64_t)p_digit[4] << 24) | ((uint64_t)p_digit[5] << 16) | ((uint64_t)p_digit[6] << 8) | (uint64_t)p_digit[7];
    }
}

static void ecc_native2bytes384(uint8_t p_bytes[ECC_BYTES_384], const uint64_t p_native[NUM_ECC_DIGITS_384])
{
    unsigned i;
    for(i=0; i<NUM_ECC_DIGITS_384; ++i)
    {
        uint8_t *p_digit = p_bytes + 8 * (NUM_ECC_DIGITS_384 - 1 - i);
        p_digit[0] = (uint8_t)(p_native[i] >> 56);
        p_digit[1] = (uint8_t)(p_native[i] >> 48);
        p_digit[2] = (uint8_t)(p_native[i] >> 40);
        p_digit[3] = (uint8_t)(p_native[i] >> 32);
        p_digit[4] = (uint8_t)(p_native[i] >> 24);
        p_digit[5] = (uint8_t)(p_native[i] >> 16);
        p_digit[6] = (uint8_t)(p_native[i] >> 8);
        p_digit[7] = (uint8_t)(p_native[i]);
    }
}

/* Compute a = sqrt(a) (mod curve_p_384). */
static void mod_sqrt384(uint64_t a[NUM_ECC_DIGITS_384])
{
    unsigned i;
    uint64_t p1[NUM_ECC_DIGITS_384] = {1};
    uint64_t l_result[NUM_ECC_DIGITS_384] = {1};
    
    /* Since curve_p_384 == 3 (mod 4) for all supported curves, we can
       compute sqrt(a) = a^((curve_p_384 + 1) / 4) (mod curve_p_384). */
    vli_add384(p1, curve_p_384, p1); /* p1 = curve_p_384 + 1 */
    for(i = vli_numBits384(p1) - 1; i > 1; --i)
    {
        vli_modSquare_fast384(l_result, l_result);
        if(vli_testBit384(p1, i))
        {
            vli_modMult_fast384(l_result, l_result, a);
        }
    }
    vli_set384(a, l_result);
}

static void ecc_point_decompress384(EccPoint384 *p_point, const uint8_t p_compressed[ECC_BYTES_384+1])
{
    uint64_t _3[NUM_ECC_DIGITS_384] = {3}; /* -a = 3 */
    ecc_bytes2native384(p_point->x, p_compressed+1);
    if (p_compressed[0] == 0x04) {
        ecc_bytes2native384(p_point->y, p_compressed+1+ECC_BYTES_384);
        return;
    }
    
    vli_modSquare_fast384(p_point->y, p_point->x); /* y = x^2 */
    vli_modSub384(p_point->y, p_point->y, _3, curve_p_384); /* y = x^2 - 3 */
    vli_modMult_fast384(p_point->y, p_point->y, p_point->x); /* y = x^3 - 3x */
    vli_modAdd384(p_point->y, p_point->y, curve_b_384, curve_p_384); /* y = x^3 - 3x + b */
    
    mod_sqrt384(p_point->y);
    
    if((p_point->y[0] & 0x01) != (p_compressed[0] & 0x01))
    {
        vli_sub384(p_point->y, curve_p_384, p_point->y);
    }
}

static int ecc_make_key384(uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384])
{
    uint64_t l_private[NUM_ECC_DIGITS_384];
    EccPoint384 l_public;
    unsigned l_tries = 0;
    
    ecc_bytes2native384(l_private, p_privateKey);
    if(vli_isZero384(l_private))
    {
        return 0;
    }
    
    /* Make sure the private key is in the range [1, n-1].
        For the supported curves, n is always large enough that we only need to subtract once at most. */
    if(vli_cmp384(curve_n_384, l_private) != 1)
    {
        vli_sub384(l_private, l_private, curve_n_384);
    }

    EccPoint_mult384(&l_public, &curve_G_384, l_private, NULL);
    if (EccPoint_isZero384(&l_public))
        return 0;
    
    p_publicKey[0] = 2 + (l_public.y[0] & 0x01);
    ecc_native2bytes384(p_publicKey + 1, l_public.x);
    return 1;
}

static int ecdh_shared_secret384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384], uint8_t p_secret[ECC_BYTES_384])
{
    EccPoint384 l_public;
    uint64_t l_private[NUM_ECC_DIGITS_384];
    //uint64_t l_random[NUM_ECC_DIGITS_384];
    uint64_t *l_random = NULL;
    
    ecc_point_decompress384(&l_public, p_publicKey);
    ecc_bytes2native384(l_private, p_privateKey);
    
    EccPoint384 l_product;
    EccPoint_mult384(&l_product, &l_public, l_private, l_random);
    
    ecc_native2bytes384(p_secret, l_product.x);
    
    return !EccPoint_isZero384(&l_product);
}

static int ecdh_uncompress_key384(const uint8_t p_publicKey[ECC_BYTES_384 + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES_384 + 1])
{
    EccPoint384 l_public;
    ecc_point_decompress384(&l_public, p_publicKey);
    p_uncompressedKey[0] = 4;
    ecc_native2bytes384(p_uncompressedKey + 1, l_public.x);
    ecc_native2bytes384(p_uncompressedKey + ECC_BYTES_384 + 1, l_public.y);
    return 1;
}

/* -------- ECDSA code -------- */

/* Computes p_result = (p_left * p_right) % p_mod. */
static void vli_modMult384(uint64_t *p_result, uint64_t *p_left, uint64_t *p_right, uint64_t *p_mod)
{
    uint64_t l_product[2 * NUM_ECC_DIGITS_384];
    uint64_t l_modMultiple[2 * NUM_ECC_DIGITS_384];
    uint l_digitShift, l_bitShift;
    uint l_productBits;
    uint l_modBits = vli_numBits384(p_mod);
    
    vli_mult384(l_product, p_left, p_right);
    l_productBits = vli_numBits384(l_product + NUM_ECC_DIGITS_384);
    if(l_productBits)
    {
        l_productBits += NUM_ECC_DIGITS_384 * 64;
    }
    else
    {
        l_productBits = vli_numBits384(l_product);
    }
    
    if(l_productBits < l_modBits)
    { /* l_product < p_mod. */
        vli_set384(p_result, l_product);
        return;
    }
    
    /* Shift p_mod by (l_leftBits - l_modBits). This multiplies p_mod by the largest
       power of two possible while still resulting in a number less than p_left. */
    vli_clear384(l_modMultiple);
    vli_clear384(l_modMultiple + NUM_ECC_DIGITS_384);
    l_digitShift = (l_productBits - l_modBits) / 64;
    l_bitShift = (l_productBits - l_modBits) % 64;
    if(l_bitShift)
    {
        l_modMultiple[l_digitShift + NUM_ECC_DIGITS_384] = vli_lshift384(l_modMultiple + l_digitShift, p_mod, l_bitShift);
    }
    else
    {
        vli_set384(l_modMultiple + l_digitShift, p_mod);
    }

    /* Subtract all multiples of p_mod to get the remainder. */
    vli_clear384(p_result);
    p_result[0] = 1; /* Use p_result as a temp var to store 1 (for subtraction) */
    while(l_productBits > NUM_ECC_DIGITS_384 * 64 || vli_cmp384(l_modMultiple, p_mod) >= 0)
    {
        int l_cmp = vli_cmp384(l_modMultiple + NUM_ECC_DIGITS_384, l_product + NUM_ECC_DIGITS_384);
        if(l_cmp < 0 || (l_cmp == 0 && vli_cmp384(l_modMultiple, l_product) <= 0))
        {
            if(vli_sub384(l_product, l_product, l_modMultiple))
            { /* borrow */
                vli_sub384(l_product + NUM_ECC_DIGITS_384, l_product + NUM_ECC_DIGITS_384, p_result);
            }
            vli_sub384(l_product + NUM_ECC_DIGITS_384, l_product + NUM_ECC_DIGITS_384, l_modMultiple + NUM_ECC_DIGITS_384);
        }
        uint64_t l_carry = (l_modMultiple[NUM_ECC_DIGITS_384] & 0x01) << 63;
        vli_rshift1384(l_modMultiple + NUM_ECC_DIGITS_384);
        vli_rshift1384(l_modMultiple);
        l_modMultiple[NUM_ECC_DIGITS_384-1] |= l_carry;
        
        --l_productBits;
    }
    vli_set384(p_result, l_product);
}

static inline uint umax384(uint a, uint b)
{
    return (a > b ? a : b);
}

static int ecdsa_sign384(const uint8_t p_privateKey[ECC_BYTES_384], const uint8_t p_hash[ECC_BYTES_384], uint64_t k[NUM_ECC_DIGITS_384], uint8_t p_signature[ECC_BYTES_384*2])
{
    uint64_t l_tmp[NUM_ECC_DIGITS_384];
    uint64_t l_s[NUM_ECC_DIGITS_384];
    EccPoint384 p;
    unsigned l_tries = 0;
    
    if(vli_isZero384(k))
        return 0;
    
    if(vli_cmp384(curve_n_384, k) != 1)
    {
        vli_sub384(k, k, curve_n_384);
    }
    
    /* tmp = k * G */
    EccPoint_mult384(&p, &curve_G_384, k, NULL);
    
    /* r = x1 (mod n) */
    if(vli_cmp384(curve_n_384, p.x) != 1)
    {
        vli_sub384(p.x, p.x, curve_n_384);
    }

    if (vli_isZero384(p.x))
        return 0;

    ecc_native2bytes384(p_signature, p.x);
    
    ecc_bytes2native384(l_tmp, p_privateKey);
    vli_modMult384(l_s, p.x, l_tmp, curve_n_384); /* s = r*d */
    ecc_bytes2native384(l_tmp, p_hash);
    vli_modAdd384(l_s, l_tmp, l_s, curve_n_384); /* s = e + r*d */
    vli_modInv384(k, k, curve_n_384); /* k = 1 / k */
    vli_modMult384(l_s, l_s, k, curve_n_384); /* s = (e + r*d) / k */
    ecc_native2bytes384(p_signature + ECC_BYTES_384, l_s);
    
    return 1;
}

static int ecdsa_verify384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_hash[ECC_BYTES_384], const uint8_t p_signature[ECC_BYTES_384*2])
{
    uint64_t u1[NUM_ECC_DIGITS_384], u2[NUM_ECC_DIGITS_384];
    uint64_t z[NUM_ECC_DIGITS_384];
    EccPoint384 l_public, l_sum;
    uint64_t rx[NUM_ECC_DIGITS_384];
    uint64_t ry[NUM_ECC_DIGITS_384];
    uint64_t tx[NUM_ECC_DIGITS_384];
    uint64_t ty[NUM_ECC_DIGITS_384];
    uint64_t tz[NUM_ECC_DIGITS_384];
    
    uint64_t l_r[NUM_ECC_DIGITS_384], l_s[NUM_ECC_DIGITS_384];
    
    ecc_point_decompress384(&l_public, p_publicKey);
    ecc_bytes2native384(l_r, p_signature);
    ecc_bytes2native384(l_s, p_signature + ECC_BYTES_384);
    
    if(vli_isZero384(l_r) || vli_isZero384(l_s))
    { /* r, s must not be 0. */
        return 0;
    }
    
    if(vli_cmp384(curve_n_384, l_r) != 1 || vli_cmp384(curve_n_384, l_s) != 1)
    { /* r, s must be < n. */
        return 0;
    }

    /* Calculate u1 and u2. */
    vli_modInv384(z, l_s, curve_n_384); /* Z = s^-1 */
    ecc_bytes2native384(u1, p_hash);
    vli_modMult384(u1, u1, z, curve_n_384); /* u1 = e/s */
    vli_modMult384(u2, l_r, z, curve_n_384); /* u2 = r/s */
    
    /* Calculate l_sum = G + Q. */
    vli_set384(l_sum.x, l_public.x);
    vli_set384(l_sum.y, l_public.y);
    vli_set384(tx, curve_G_384.x);
    vli_set384(ty, curve_G_384.y);
    vli_modSub384(z, l_sum.x, tx, curve_p_384); /* Z = x2 - x1 */
    XYcZ_add384(tx, ty, l_sum.x, l_sum.y);
    vli_modInv384(z, z, curve_p_384); /* Z = 1/Z */
    apply_z384(l_sum.x, l_sum.y, z);
    
    /* Use Shamir's trick to calculate u1*G + u2*Q */
    EccPoint384 *l_points[4] = {NULL, &curve_G_384, &l_public, &l_sum};
    uint l_numBits = umax384(vli_numBits384(u1), vli_numBits384(u2));
    
    EccPoint384 *l_point = l_points[(!!vli_testBit384(u1, l_numBits-1)) | ((!!vli_testBit384(u2, l_numBits-1)) << 1)];
    vli_set384(rx, l_point->x);
    vli_set384(ry, l_point->y);
    vli_clear384(z);
    z[0] = 1;

    int i;
    for(i = l_numBits - 2; i >= 0; --i)
    {
        EccPoint_double_jacobian384(rx, ry, z);
        
        int l_index = (!!vli_testBit384(u1, i)) | ((!!vli_testBit384(u2, i)) << 1);
        EccPoint384 *l_point = l_points[l_index];
        if(l_point)
        {
            vli_set384(tx, l_point->x);
            vli_set384(ty, l_point->y);
            apply_z384(tx, ty, z);
            vli_modSub384(tz, rx, tx, curve_p_384); /* Z = x2 - x1 */
            XYcZ_add384(tx, ty, rx, ry);
            vli_modMult_fast384(z, z, tz);
        }
    }

    vli_modInv384(z, z, curve_p_384); /* Z = 1/Z */
    apply_z384(rx, ry, z);
    
    /* v = x1 (mod n) */
    if(vli_cmp384(curve_n_384, rx) != 1)
    {
        vli_sub384(rx, rx, curve_n_384);
    }

    /* Accept only if v == r. */
    return (vli_cmp384(rx, l_r) == 0);
}
