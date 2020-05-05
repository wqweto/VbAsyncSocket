#ifdef IMPL_GMPRSA_THUNK

static size_t gmp_rsa_public_encrypt(const size_t flen, const unsigned char* from, const unsigned char* exp, const unsigned char* mod, unsigned char* to)
{
    size_t size;
    mpz_t n, e;
    mpz_t enc, dec;

    mpz_init(n);
    mpz_import(n, flen, 1, 1, 1, 0, mod);
    mpz_init(e);
    mpz_import(e, flen, 1, 1, 1, 0, exp);

    mpz_init(enc);
    mpz_init(dec);
    mpz_import(dec, flen, 1, 1, 1, 0, from);

    mpz_powm(enc, dec, e, n);

    mpz_clear(dec);
	uint8_t *r = (uint8_t *)mpz_export(NULL, &size, 1, 1, 1, 0, enc);
    mpz_clear(enc);
    if (flen < size) {
        memcpy(to, r+size-flen, flen);
    } else {
        memset(to, 0, flen-size);
        memcpy(to+flen-size, r, size);
    }
    gmp_default_free(r, size);

    return size;
}

#endif // IMPL_GMPRSA_THUNK

#ifdef IMPL_SSHRSA_THUNK

static void rsa_modexp(uint32_t maxbytes, const uint8_t *b, const uint8_t *e, const uint8_t *m, uint8_t *r)
{
    Bignum base, exponent, modulus, res;

    base = bignum_from_bytes(b, maxbytes);
    exponent = bignum_from_bytes(e, maxbytes);
    modulus = bignum_from_bytes(m, maxbytes);
    res = modpow(base, exponent, modulus);
    for (int i = maxbytes; i--;) {
	    *r++ = bignum_byte(res, i);
    }
    freebn(base);
    freebn(exponent);
    freebn(modulus);
    freebn(res);    
}

#endif // IMPL_SSHRSA_THUNK