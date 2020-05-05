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
 * General block cipher description
 * ================================
 * This allows us to implement block cipher modes which can work
 * with different block ciphers.
 */

/* .. c:type:: cf_prp_block
 * Block processing function type.
 *
 * The `in` and `out` blocks may alias.
 *
 * :rtype: void
 * :param ctx: block cipher-specific context object.
 * :param in: input block.
 * :param out: output block.
 */
typedef void (*cf_prp_block)(void *ctx, const uint8_t *in, uint8_t *out);

/* .. c:type:: cf_prp
 * Describes an PRP in a general way.
 *
 * .. c:member:: cf_prp.blocksz
 * Block size in bytes. Must be no more than :c:macro:`CF_MAXBLOCK`.
 *
 * .. c:member:: cf_prp.encrypt
 * Block encryption function.
 *
 * .. c:member:: cf_prp.decrypt
 * Block decryption function.
 */
typedef struct
{
  size_t blocksz;
  cf_prp_block encrypt;
  cf_prp_block decrypt;
} cf_prp;

/* .. c:macro:: CF_MAXBLOCK
 * The maximum block cipher blocksize we support, in bytes.
 */
#define CF_MAXBLOCK 16

/**
 * CBC mode
 * --------
 * This implementation allows encryption or decryption of whole
 * blocks in CBC mode.  It does not offer a byte-wise incremental
 * interface, or do any padding.
 *
 * This mode provides no useful integrity and should not be used
 * directly.
 */

/* .. c:type:: cf_cbc
 * This structure binds together the things needed to encrypt/decrypt whole
 * blocks in CBC mode.
 *
 * .. c:member:: cf_cbc.prp
 * How to encrypt or decrypt blocks.  This could be, for example, :c:data:`cf_aes`.
 *
 * .. c:member:: cf_cbc.prpctx
 * Private data for prp functions.  For a `prp` of `cf_aes`, this would be a
 * pointer to a :c:type:`cf_aes_context` instance.
 *
 * .. c:member:: cf_cbc.block
 * The IV or last ciphertext block.
 */
typedef struct
{
  const cf_prp *prp;
  void *prpctx;
  uint8_t block[CF_MAXBLOCK];
} cf_cbc;

/**
 * Counter mode
 * ------------
 * This implementation allows incremental encryption/decryption of
 * messages.  Encryption and decryption are the same operation.
 *
 * The counter is always big-endian, but has configurable location
 * and size within the nonce block.  The counter wraps, so you
 * should make sure the length of a message with a given nonce
 * doesn't cause nonce reuse.
 *
 * This mode provides no integrity and should not be used directly.
 */

/* .. c:type:: cf_ctr
 *
 * .. c:member:: cf_ctr.prp
 * How to encrypt or decrypt blocks.  This could be, for example, :c:data:`cf_aes`.
 *
 * .. c:member:: cf_ctr.prpctx
 * Private data for prp functions.  For a `prp` of `cf_aes`, this would be a
 * pointer to a :c:type:`cf_aes_context` instance.
 *
 * .. c:member:: cf_ctr.nonce
 * The next block to encrypt to get another block of key stream.
 *
 * .. c:member:: cf_ctr.keymat
 * The current block of key stream.
 *
 * .. c:member:: cf_ctr.nkeymat
 * The number of bytes at the end of :c:member:`keymat` that are so-far unused.
 * If this is zero, all the bytes are used up and/or of undefined value.
 *
 * .. c:member:: cf_ctr.counter_offset
 * The offset (in bytes) of the counter block within the nonce.
 *
 * .. c:member:: cf_ctr.counter_width
 * The width (in bytes) of the counter block in the nonce.
 */
typedef struct
{
  const cf_prp *prp;
  void *prpctx;
  uint8_t nonce[CF_MAXBLOCK];
  uint8_t keymat[CF_MAXBLOCK];
  size_t nkeymat;
  size_t counter_offset;
  size_t counter_width;
} cf_ctr;

/* CBC */
static
void cf_cbc_init(cf_cbc *ctx, const cf_prp *prp, void *prpctx, const uint8_t iv[CF_MAXBLOCK])
{
  ctx->prp = prp;
  ctx->prpctx = prpctx;
  memcpy(ctx->block, iv, prp->blocksz);
}

static
void cf_cbc_encrypt(cf_cbc *ctx, const uint8_t *input, uint8_t *output, size_t blocks)
{
  uint8_t buf[CF_MAXBLOCK];
  size_t nblk = ctx->prp->blocksz;

  while (blocks--)
  {
    xor_bb(buf, input, ctx->block, nblk);
    ctx->prp->encrypt(ctx->prpctx, buf, ctx->block);
    memcpy(output, ctx->block, nblk);
    input += nblk;
    output += nblk;
  }
}

static
void cf_cbc_decrypt(cf_cbc *ctx, const uint8_t *input, uint8_t *output, size_t blocks)
{
  uint8_t buf[CF_MAXBLOCK];
  size_t nblk = ctx->prp->blocksz;

  while (blocks--)
  {
    ctx->prp->decrypt(ctx->prpctx, input, buf);
    xor_bb(output, buf, ctx->block, nblk);
    memcpy(ctx->block, input, nblk);
    input += nblk;
    output += nblk;
  }
}

/* CTR */
static
void cf_ctr_init(cf_ctr *ctx, const cf_prp *prp, void *prpctx, const uint8_t nonce[CF_MAXBLOCK])
{
  memset(ctx, 0, sizeof *ctx);
  ctx->counter_offset = 0;
  ctx->counter_width = prp->blocksz;
  ctx->prp = prp;
  ctx->prpctx = prpctx;
  ctx->nkeymat = 0;
  memcpy(ctx->nonce, nonce, prp->blocksz);
}

static
void cf_ctr_custom_counter(cf_ctr *ctx, size_t offset, size_t width)
{
  assert(ctx->prp->blocksz <= offset + width);
  ctx->counter_offset = offset;
  ctx->counter_width = width;
}

static void ctr_next_block(void *vctx, uint8_t *out)
{
  cf_ctr *ctx = (cf_ctr *)vctx;
  ctx->prp->encrypt(ctx->prpctx, ctx->nonce, out);
  incr_be(ctx->nonce + ctx->counter_offset, ctx->counter_width);
}

static
void cf_ctr_cipher(cf_ctr *ctx, const uint8_t *input, uint8_t *output, size_t bytes)
{
  DECLARE_PFN(cf_blockwise_out_fn, ctr_next_block);
  cf_blockwise_xor(ctx->keymat, &ctx->nkeymat,
                   ctx->prp->blocksz,
                   input, output, bytes,
                   pfn_ctr_next_block,
                   ctx);
}

static
void cf_ctr_discard_block(cf_ctr *ctx)
{
  ctx->nkeymat = 0;
}
