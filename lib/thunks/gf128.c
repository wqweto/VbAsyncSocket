/*
 * cifra - embedded cryptography library
 * Written in 2014 by Joseph Birr-Pixton <jpixton@gmail.com>
 *
 * To the extent possible under law, the author(s) have dedicated all
 * copyright and related and neighboring rights to this software to the
 * public domain worldwide. This software is distributed without any
 * warranty.
 *
 * You should have received a copy of the CC0 Public Domain Dedication
 * along with this software. If not, see
 * <http://creativecommons.org/publicdomain/zero/1.0/>.
 */

/**
 * @brief Operations in GF(2^128).
 *
 * These implementations are constant time, but relatively slow.
 */

typedef uint32_t cf_gf128[4];

static
void cf_gf128_tobytes_be(const cf_gf128 in, uint8_t out[16])
{
  write32_be(in[0], out + 0);
  write32_be(in[1], out + 4);
  write32_be(in[2], out + 8);
  write32_be(in[3], out + 12);
}

static
void cf_gf128_frombytes_be(const uint8_t in[16], cf_gf128 out)
{
  out[0] = read32_be(in + 0);
  out[1] = read32_be(in + 4);
  out[2] = read32_be(in + 8);
  out[3] = read32_be(in + 12);
}

/* out = 2 * in.  Arguments may alias. */
static
void cf_gf128_double_le(const cf_gf128 in, cf_gf128 out)
{
  uint8_t table[2] = { 0x00, 0xe1 };
  uint32_t borrow = 0;
  uint32_t inword;

  inword = in[0];   out[0] = (inword >> 1) | (borrow << 31);  borrow = inword & 1;
  inword = in[1];   out[1] = (inword >> 1) | (borrow << 31);  borrow = inword & 1;
  inword = in[2];   out[2] = (inword >> 1) | (borrow << 31);  borrow = inword & 1;
  inword = in[3];   out[3] = (inword >> 1) | (borrow << 31);  borrow = inword & 1;

#if CF_CACHE_SIDE_CHANNEL_PROTECTION
  out[0] ^= select_u8(borrow, table, 2) << 24;
#else
  out[0] ^= table[borrow] << 24;
#endif
}

/* out = x + y.  Arguments may alias. */
static
void cf_gf128_add(const cf_gf128 x, const cf_gf128 y, cf_gf128 out)
{
  out[0] = x[0] ^ y[0];
  out[1] = x[1] ^ y[1];
  out[2] = x[2] ^ y[2];
  out[3] = x[3] ^ y[3];
}

/* out = xy.  Arguments may alias. */
static
void cf_gf128_mul(const cf_gf128 x, const cf_gf128 y, cf_gf128 out)
{
#if CF_TIME_SIDE_CHANNEL_PROTECTION
  cf_gf128 zero = { 0 };
#endif
 
  /* Z_0 = 0^128
   * V_0 = Y */ 
  cf_gf128 Z, V;
  memset(Z, 0, sizeof Z);
  memcpy(V, y, sizeof V);

  for (int i = 0; i < 128; i++)
  {
    uint32_t word = x[i >> 5];
    uint8_t bit = (word >> (31 - (i & 31))) & 1;

#if CF_TIME_SIDE_CHANNEL_PROTECTION
    select_xor128(Z, zero, V, bit);
#else
    if (bit)
      xor_words(Z, V, 4);
#endif

    cf_gf128_double_le(V, V);
  }

  memcpy(out, Z, sizeof Z);
}

#include <immintrin.h>

/*
 *  From https://www.intel.com/content/www/us/en/processors/carry-less-multiplication-instruction-in-gcm-mode-paper.html
 */
static void gfmul(__m128i a, __m128i b, __m128i *res)
{
    __m128i tmp0, tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7, tmp8, tmp9;
    tmp3 = _mm_clmulepi64_si128(a, b, 0x00);
    tmp4 = _mm_clmulepi64_si128(a, b, 0x10);
    tmp5 = _mm_clmulepi64_si128(a, b, 0x01);
    tmp6 = _mm_clmulepi64_si128(a, b, 0x11);
    tmp4 = _mm_xor_si128(tmp4, tmp5);
    tmp5 = _mm_slli_si128(tmp4, 8);
    tmp4 = _mm_srli_si128(tmp4, 8);
    tmp3 = _mm_xor_si128(tmp3, tmp5);
    tmp6 = _mm_xor_si128(tmp6, tmp4);
    tmp7 = _mm_srli_epi32(tmp3, 31);
    tmp8 = _mm_srli_epi32(tmp6, 31);
    tmp3 = _mm_slli_epi32(tmp3, 1);
    tmp6 = _mm_slli_epi32(tmp6, 1);
    tmp9 = _mm_srli_si128(tmp7, 12);
    tmp8 = _mm_slli_si128(tmp8, 4);
    tmp7 = _mm_slli_si128(tmp7, 4);
    tmp3 = _mm_or_si128(tmp3, tmp7);
    tmp6 = _mm_or_si128(tmp6, tmp8);
    tmp6 = _mm_or_si128(tmp6, tmp9);
    tmp7 = _mm_slli_epi32(tmp3, 31);
    tmp8 = _mm_slli_epi32(tmp3, 30);
    tmp9 = _mm_slli_epi32(tmp3, 25);
    tmp7 = _mm_xor_si128(tmp7, tmp8);
    tmp7 = _mm_xor_si128(tmp7, tmp9);
    tmp8 = _mm_srli_si128(tmp7, 4);
    tmp7 = _mm_slli_si128(tmp7, 12);
    tmp3 = _mm_xor_si128(tmp3, tmp7);
    tmp2 = _mm_srli_epi32(tmp3, 1);
    tmp4 = _mm_srli_epi32(tmp3, 2);
    tmp5 = _mm_srli_epi32(tmp3, 7);
    tmp2 = _mm_xor_si128(tmp2, tmp4);
    tmp2 = _mm_xor_si128(tmp2, tmp5);
    tmp2 = _mm_xor_si128(tmp2, tmp8);
    tmp3 = _mm_xor_si128(tmp3, tmp2); 
    tmp6 = _mm_xor_si128(tmp6, tmp3);
    *res = tmp6;
}

static inline void cf_gf128_reflect(const cf_gf128 x, cf_gf128 out)
{
    out[3] = x[0];
    out[2] = x[1];
    out[1] = x[2];
    out[0] = x[3];    
}

static void cf_gf128_mul_fast(const cf_gf128 x, const cf_gf128 y, cf_gf128 out)
{
    cf_gf128 tmp;
    cf_gf128_reflect(x, tmp);
    const __m128i a = _mm_loadu_si128((const __m128i*)tmp);
    cf_gf128_reflect(y, tmp);
    const __m128i b = _mm_loadu_si128((const __m128i*)tmp);
    __m128i res;
    gfmul(a, b, &res);
    _mm_storeu_si128((__m128i*)tmp, res);
    cf_gf128_reflect(tmp, out);
}
