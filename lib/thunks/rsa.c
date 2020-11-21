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

static void rsa_modexp(const uint32_t maxbytes, const uint8_t *base_in, const uint8_t *exp_in, const uint8_t *mod_in, uint8_t *ret_out)
{
    Bignum base, exp, mod, ret;

    base = bignum_from_bytes(base_in, maxbytes);
    exp = bignum_from_bytes(exp_in, maxbytes);
    mod = bignum_from_bytes(mod_in, maxbytes);
    ret = modpow(base, exp, mod);
    for (int i = maxbytes; i--;) {
	    *ret_out++ = bignum_byte(ret, i);
    }
    freebn(base);
    freebn(exp);
    freebn(mod);
    freebn(ret);
}

/*
 * Compute (base ^ exp) % mod, provided mod == p * q, with p,q
 * distinct primes, and iqmp is the multiplicative inverse of q mod p.
 * Uses Chinese Remainder Theorem to speed computation up over the
 * obvious implementation of a single big modpow.
 */
static void rsa_crt_modexp(const uint32_t maxbytes, const uint8_t *base_in, const uint8_t *exp_in, const uint8_t *mod_in, 
                           const uint8_t *p_in, const uint8_t *q_in, const uint8_t *iqmp_in, uint8_t *ret_out)
{
    Bignum base, exp, mod, p, q, iqmp;
    Bignum pm1, qm1, pexp, qexp, presult, qresult, diff, multiplier, ret0, ret;

    base = bignum_from_bytes(base_in, maxbytes);
    exp = bignum_from_bytes(exp_in, maxbytes);
    mod = bignum_from_bytes(mod_in, maxbytes);
    p = bignum_from_bytes(p_in, maxbytes / 2);
    q = bignum_from_bytes(q_in, maxbytes / 2);
    iqmp = bignum_from_bytes(iqmp_in, maxbytes / 2);

    /*
     * Reduce the exponent mod phi(p) and phi(q), to save time when
     * exponentiating mod p and mod q respectively. Of course, since p
     * and q are prime, phi(p) == p-1 and similarly for q.
     */
    pm1 = copybn(p);
    decbn(pm1);
    qm1 = copybn(q);
    decbn(qm1);
    pexp = bigmod(exp, pm1);
    qexp = bigmod(exp, qm1);

    /*
     * Do the two modpows.
     */
    presult = modpow(base, pexp, p);
    qresult = modpow(base, qexp, q);

    /*
     * Recombine the results. We want a value which is congruent to
     * qresult mod q, and to presult mod p.
     *
     * We know that iqmp * q is congruent to 1 * mod p (by definition
     * of iqmp) and to 0 mod q (obviously). So we start with qresult
     * (which is congruent to qresult mod both primes), and add on
     * (presult-qresult) * (iqmp * q) which adjusts it to be congruent
     * to presult mod p without affecting its value mod q.
     */
    if (bignum_cmp(presult, qresult) < 0) {
        /*
         * Can't subtract presult from qresult without first adding on
         * p.
         */
        Bignum tmp = presult;
        presult = bigadd(presult, p);
        freebn(tmp);
    }
    diff = bigsub(presult, qresult);
    multiplier = bigmul(iqmp, q);
    ret0 = bigmuladd(multiplier, diff, qresult);

    /*
     * Finally, reduce the result mod n.
     */
    ret = bigmod(ret0, mod);
    for (int i = maxbytes; i--;) {
	    *ret_out++ = bignum_byte(ret, i);
    }

    /*
     * Free all the intermediate results before returning.
     */
    freebn(pm1);
    freebn(qm1);
    freebn(pexp);
    freebn(qexp);
    freebn(presult);
    freebn(qresult);
    freebn(diff);
    freebn(multiplier);
    freebn(ret0);
    freebn(base);
    freebn(exp);
    freebn(mod);
    freebn(p);
    freebn(q);
    freebn(iqmp);
    freebn(ret);
}

#endif // IMPL_SSHRSA_THUNK