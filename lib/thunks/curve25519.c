/* This is based on tweetnacl.  Some typedefs have been
 * replaced with their stdint equivalents.
 *
 * Original code was public domain. */

typedef int64_t gf[16];

static void set25519(gf r, const gf a)
{
  for (size_t i = 0; i < 16; i++)
    r[i] = a[i];
}

static void car25519(gf o)
{
  int64_t c;

  for (size_t i = 0; i < 16; i++)
  {
    o[i] += (1LL << 16);
    c = o[i] >> 16;
    o[(i + 1) * (i < 15)] += c - 1 + 37 * (c - 1) * (i == 15);
    o[i] -= c << 16;
  }
}

static void sel25519(gf p, gf q, int b)
{
  int64_t tmp, mask = ~(b-1);
  for (size_t i = 0; i < 16; i++)
  {
    tmp = mask & (p[i] ^ q[i]);
    p[i] ^= tmp;
    q[i] ^= tmp;
  }
}

static void pack25519(uint8_t out[32], const gf n)
{
  int b;
  gf m, t;
  set25519(t, n);
  car25519(t);
  car25519(t);
  car25519(t);
  
  for(size_t j = 0; j < 2; j++)
  {
    m[0] = t[0] - 0xffed;
    for (size_t i = 1; i < 15; i++)
    {
      m[i] = t[i] - 0xffff - ((m[i - 1] >> 16) & 1);
      m[i - 1] &= 0xffff;
    }
    m[15] = t[15] - 0x7fff - ((m[14] >> 16) & 1);
    b = (m[15] >> 16) & 1;
    m[14] &= 0xffff;
    sel25519(t, m, 1 - b);
  }

  for (size_t i = 0; i < 16; i++)
  {
    out[2 * i] = t[i] & 0xff;
    out[2 * i + 1] = (uint8_t)(t[i] >> 8);
  }
}



static void unpack25519(gf o, const uint8_t *n)
{
  for (size_t i = 0; i < 16; i++)
    o[i] = n[2 * i] + ((int64_t) n[2 * i + 1] << 8);
  o[15] &= 0x7fff;
}

static void add(gf o, const gf a, const gf b)
{
  for (size_t i = 0; i < 16; i++)
    o[i] = a[i] + b[i];
}

static void sub(gf o, const gf a, const gf b)
{
  for (size_t i = 0; i < 16; i++)
    o[i] = a[i] - b[i];
}

static void mul(gf o, const gf a, const gf b)
{
  int64_t t[31];

  for (size_t i = 0; i < 31; i++)
    t[i] = 0;

  for (size_t i = 0; i < 16; i++)
    for (size_t j = 0; j < 16; j++)
      t[i + j] += a[i] * b[j];

  for (size_t i = 0; i < 15; i++)
    t[i] += 38 * t[i + 16];

  for (size_t i = 0; i < 16; i++)
    o[i] = t[i];

  car25519(o);
  car25519(o);
}

static void sqr(gf o, const gf a)
{
  mul(o, a, a);
}

static void inv25519(gf o, const gf i)
{
  gf c;
  for (int a = 0; a < 16; a++)
    c[a] = i[a];
  
  for (int a = 253; a >= 0; a--)
  {
    sqr(c, c);
    if(a != 2 && a != 4)
      mul(c, c, i);
  }

  for (int a = 0; a < 16; a++)
    o[a] = c[a];
}

static
void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
{
  uint8_t z[32];
  gf x;
  gf a, b, c, d, e, f;
  const gf _121665 = {0xDB41, 1};

  for (size_t i = 0; i < 31; i++)
    z[i] = priv[i];
  z[31] = (priv[31] & 127) | 64;
  z[0] &= 248;
  
  unpack25519(x, pub);

  for(size_t i = 0; i < 16; i++)
  {
    b[i] = x[i];
    d[i] = a[i] = c[i] = 0;
  }

  a[0] = d[0] = 1;

  for (int i = 254; i >= 0; i--)
  {
    int r = (z[i >> 3] >> (i & 7)) & 1;
    sel25519(a, b, r);
    sel25519(c, d, r);
    add(e, a, c);
    sub(a, a, c);
    add(c, b, d);
    sub(b, b, d);
    sqr(d, e);
    sqr(f, a);
    mul(a, c, a);
    mul(c, b, e);
    add(e, a, c);
    sub(a, a, c);
    sqr(b, a);
    sub(c, d, f);
    mul(a, c, _121665);
    add(a, a, d);
    mul(c, c, a);
    mul(a, d, f);
    mul(d, b, x);
    sqr(b, e);
    sel25519(a, b, r);
    sel25519(c, d, r);
  }

  inv25519(c, c);
  mul(a, a, c);
  pack25519(out, a);
}

static
void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
{
  const uint8_t _9[32] = {9};
  cf_curve25519_mul(out, priv, _9);
}
