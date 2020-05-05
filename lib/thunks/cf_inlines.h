/** Encode v as a 64-bit big endian quantity into buf. */
static inline void write64_be(uint64_t v, uint8_t buf[8])
{
  *buf++ = (v >> 56) & 0xff;
  *buf++ = (v >> 48) & 0xff;
  *buf++ = (v >> 40) & 0xff;
  *buf++ = (v >> 32) & 0xff;
  *buf++ = (v >> 24) & 0xff;
  *buf++ = (v >> 16) & 0xff;
  *buf++ = (v >> 8) & 0xff;
  *buf   = v & 0xff;
}

/** Encode v as a 32-bit big endian quantity into buf. */
static inline void write32_be(uint32_t v, uint8_t buf[4])
{
  *buf++ = (v >> 24) & 0xff;
  *buf++ = (v >> 16) & 0xff;
  *buf++ = (v >> 8) & 0xff;
  *buf   = v & 0xff;
}

/** Read 4 bytes from buf, as a 32-bit big endian quantity. */
static inline uint32_t read32_be(const uint8_t buf[4])
{
  return (buf[0] << 24) |
         (buf[1] << 16) |
         (buf[2] << 8) |
         (buf[3]);
}

/** Circularly rotate right x by n bits.
 *  0 > n > 32. */
static inline uint32_t rotr32(uint32_t x, unsigned n)
{
  return (x >> n) | (x << (32 - n));
}

/** Read 4 bytes from buf, as a 32-bit little endian quantity. */
static inline uint32_t read32_le(const uint8_t buf[4])
{
  return (buf[3] << 24) |
         (buf[2] << 16) |
         (buf[1] << 8) |
         (buf[0]);
}

/** Read 8 bytes from buf, as a 64-bit big endian quantity. */
static inline uint64_t read64_be(const uint8_t buf[8])
{
  uint32_t hi = read32_be(buf),
           lo = read32_be(buf + 4);
  return ((uint64_t)hi) << 32 |
         lo;
}

/** Circularly rotate right x by n bits.
 *  0 > n > 64. */
static inline uint64_t rotr64(uint64_t x, unsigned n)
{
  return (x >> n) | (x << (64 - n));
}

/** Increments the integer stored at v (of non-zero length len)
 *  with the least significant byte first. */
static inline void incr_le(uint8_t *v, size_t len)
{
  size_t i = 0;
  while (1)
  {
    if (++v[i] != 0)
      return;
    i++;
    if (i == len)
      return;
  }
}

/** Circularly rotate left x by n bits.
 *  0 > n > 32. */
static inline uint32_t rotl32(uint32_t x, unsigned n)
{
  return (x << n) | (x >> (32 - n));
}

/** Encode v as a 32-bit little endian quantity into buf. */
static inline void write32_le(uint32_t v, uint8_t buf[4])
{
  *buf++ = v & 0xff;
  *buf++ = (v >> 8) & 0xff;
  *buf++ = (v >> 16) & 0xff;
  *buf   = (v >> 24) & 0xff;
}

/** Produce 0xffffffff if x == y, zero otherwise, without branching. */
static inline uint32_t mask_u32(uint32_t x, uint32_t y)
{
  uint32_t diff = x ^ y;
  uint32_t diff_is_zero = ~diff & (diff - 1);
  return - (diff_is_zero >> 31);
}

/** Like memset(ptr, 0, len), but not allowed to be removed by
 *  compilers. */
static inline void mem_clean(volatile void *v, size_t len)
{
  if (len)
  {
    memset((void *) v, 0, len);
    (void) *((volatile uint8_t *) v);
  }
}

/** Returns 1 if len bytes at va equal len bytes at vb, 0 if they do not.
 *  Does not leak length of common prefix through timing. */
static inline unsigned mem_eq(const void *va, const void *vb, size_t len)
{
  const volatile uint8_t *a = (const volatile uint8_t *)va;
  const volatile uint8_t *b = (const volatile uint8_t *)vb;
  uint8_t diff = 0;

  while (len--)
  {
    diff |= *a++ ^ *b++;
  }

  return !diff;
}

/** Encode v as a 64-bit little endian quantity into buf. */
static inline void write64_le(uint64_t v, uint8_t buf[8])
{
  *buf++ = v & 0xff;
  *buf++ = (v >> 8) & 0xff;
  *buf++ = (v >> 16) & 0xff;
  *buf++ = (v >> 24) & 0xff;
  *buf++ = (v >> 32) & 0xff;
  *buf++ = (v >> 40) & 0xff;
  *buf++ = (v >> 48) & 0xff;
  *buf   = (v >> 56) & 0xff;
}

/* out ^= x
 * out and x may alias. */
static inline void xor_words(uint32_t *out, const uint32_t *x, size_t nwords)
{
  for (size_t i = 0; i < nwords; i++)
    out[i] ^= x[i];
}

/** Increments the integer stored at v (of non-zero length len)
 *  with the most significant byte last. */
static inline void incr_be(uint8_t *v, size_t len)
{
  len--;
  while (1)
  {
    if (++v[len] != 0)
      return;
    if (len == 0)
      return;
    len--;
  }
}
