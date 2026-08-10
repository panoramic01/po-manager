// ─── AES-256-CBC (for Portal Storage password encryption) ───────────────────
// Apps Script's V8 runtime has no Web Crypto API and no built-in symmetric
// cipher, so AES is implemented here directly from the FIPS-197 definition.
// The S-box is *computed* from GF(2^8) multiplicative inverses rather than
// hardcoded, to avoid a transcription error in a 256-entry table. Verified
// against the official FIPS-197 AES-256 known-answer test -- see
// aesSelfTest_() below, which can be run manually from the Apps Script
// editor to re-confirm correctness.
//
// Public API: aesEncrypt_(plainText, keyBytes) / aesDecrypt_(payload, keyBytes)
// Key management: getPortalStorageKeyBytes_() (mirrors getSessionSecret_()).

var AES_NB = 4;  // block size in 32-bit words (fixed at 4 for AES)
var AES_NK = 8;  // key length in 32-bit words (8 = AES-256)
var AES_NR = 14; // number of rounds (14 for AES-256)

function aesGmul_(a, b) {
  var p = 0;
  for (var counter = 0; counter < 8; counter++) {
    if (b & 1) p ^= a;
    var hiBitSet = a & 0x80;
    a = (a << 1) & 0xFF;
    if (hiBitSet) a ^= 0x1B;
    b >>= 1;
  }
  return p & 0xFF;
}

function aesGInverse_(a) {
  if (a === 0) return 0;
  for (var b = 1; b < 256; b++) {
    if (aesGmul_(a, b) === 1) return b;
  }
  return 0;
}

function aesBuildSbox_() {
  var sbox = new Array(256);
  for (var i = 0; i < 256; i++) {
    var inv = aesGInverse_(i);
    var x = inv;
    var result = x;
    for (var shift = 1; shift <= 4; shift++) {
      x = ((x << 1) | (x >> 7)) & 0xFF;
      result ^= x;
    }
    sbox[i] = (result ^ 0x63) & 0xFF;
  }
  return sbox;
}

function aesGetSboxes_() {
  if (!aesGetSboxes_._cache) {
    var sbox = aesBuildSbox_();
    var invSbox = new Array(256);
    for (var i = 0; i < 256; i++) invSbox[sbox[i]] = i;
    aesGetSboxes_._cache = { sbox: sbox, invSbox: invSbox };
  }
  return aesGetSboxes_._cache;
}

function aesXtime_(a) {
  var hi = a & 0x80;
  a = (a << 1) & 0xFF;
  return hi ? (a ^ 0x1B) : a;
}

function aesRotWord_(w) { return [w[1], w[2], w[3], w[0]]; }
function aesSubWord_(w, sbox) { return [sbox[w[0]], sbox[w[1]], sbox[w[2]], sbox[w[3]]]; }

function aesKeyExpansion_(key, sbox) {
  var totalWords = AES_NB * (AES_NR + 1);
  var w = new Array(totalWords);
  for (var i = 0; i < AES_NK; i++) {
    w[i] = [key[4 * i], key[4 * i + 1], key[4 * i + 2], key[4 * i + 3]];
  }
  var rcon = 1;
  for (var i = AES_NK; i < totalWords; i++) {
    var temp = w[i - 1].slice();
    if (i % AES_NK === 0) {
      temp = aesSubWord_(aesRotWord_(temp), sbox);
      temp[0] ^= rcon;
      rcon = aesXtime_(rcon);
    } else if (AES_NK > 6 && i % AES_NK === 4) {
      temp = aesSubWord_(temp, sbox);
    }
    w[i] = [
      w[i - AES_NK][0] ^ temp[0],
      w[i - AES_NK][1] ^ temp[1],
      w[i - AES_NK][2] ^ temp[2],
      w[i - AES_NK][3] ^ temp[3]
    ];
  }
  return w;
}

function aesAddRoundKey_(state, w, round) {
  for (var c = 0; c < 4; c++) {
    var word = w[round * 4 + c];
    for (var r = 0; r < 4; r++) state[r][c] ^= word[r];
  }
}

function aesSubBytes_(state, sbox) {
  for (var r = 0; r < 4; r++)
    for (var c = 0; c < 4; c++)
      state[r][c] = sbox[state[r][c]];
}

function aesShiftRows_(state) {
  for (var r = 1; r < 4; r++) {
    var row = state[r];
    state[r] = row.slice(r).concat(row.slice(0, r));
  }
}

function aesMixColumns_(state) {
  for (var c = 0; c < 4; c++) {
    var a0 = state[0][c], a1 = state[1][c], a2 = state[2][c], a3 = state[3][c];
    state[0][c] = aesGmul_(a0, 2) ^ aesGmul_(a1, 3) ^ a2 ^ a3;
    state[1][c] = a0 ^ aesGmul_(a1, 2) ^ aesGmul_(a2, 3) ^ a3;
    state[2][c] = a0 ^ a1 ^ aesGmul_(a2, 2) ^ aesGmul_(a3, 3);
    state[3][c] = aesGmul_(a0, 3) ^ a1 ^ a2 ^ aesGmul_(a3, 2);
  }
}

function aesInvShiftRows_(state) {
  for (var r = 1; r < 4; r++) {
    var row = state[r];
    state[r] = row.slice(4 - r).concat(row.slice(0, 4 - r));
  }
}

function aesInvSubBytes_(state, invSbox) {
  for (var r = 0; r < 4; r++)
    for (var c = 0; c < 4; c++)
      state[r][c] = invSbox[state[r][c]];
}

function aesInvMixColumns_(state) {
  for (var c = 0; c < 4; c++) {
    var a0 = state[0][c], a1 = state[1][c], a2 = state[2][c], a3 = state[3][c];
    state[0][c] = aesGmul_(a0, 14) ^ aesGmul_(a1, 11) ^ aesGmul_(a2, 13) ^ aesGmul_(a3, 9);
    state[1][c] = aesGmul_(a0, 9)  ^ aesGmul_(a1, 14) ^ aesGmul_(a2, 11) ^ aesGmul_(a3, 13);
    state[2][c] = aesGmul_(a0, 13) ^ aesGmul_(a1, 9)  ^ aesGmul_(a2, 14) ^ aesGmul_(a3, 11);
    state[3][c] = aesGmul_(a0, 11) ^ aesGmul_(a1, 13) ^ aesGmul_(a2, 9)  ^ aesGmul_(a3, 14);
  }
}

function aesBytesToState_(input16) {
  var state = [[0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]];
  for (var i = 0; i < 16; i++) state[i % 4][Math.floor(i / 4)] = input16[i];
  return state;
}

function aesStateToBytes_(state) {
  var out = new Array(16);
  for (var i = 0; i < 16; i++) out[i] = state[i % 4][Math.floor(i / 4)];
  return out;
}

function aesCipherBlock_(input16, w, sbox) {
  var state = aesBytesToState_(input16);
  aesAddRoundKey_(state, w, 0);
  for (var round = 1; round < AES_NR; round++) {
    aesSubBytes_(state, sbox);
    aesShiftRows_(state);
    aesMixColumns_(state);
    aesAddRoundKey_(state, w, round);
  }
  aesSubBytes_(state, sbox);
  aesShiftRows_(state);
  aesAddRoundKey_(state, w, AES_NR);
  return aesStateToBytes_(state);
}

function aesInvCipherBlock_(input16, w, sbox, invSbox) {
  var state = aesBytesToState_(input16);
  aesAddRoundKey_(state, w, AES_NR);
  for (var round = AES_NR - 1; round >= 1; round--) {
    aesInvShiftRows_(state);
    aesInvSubBytes_(state, invSbox);
    aesAddRoundKey_(state, w, round);
    aesInvMixColumns_(state);
  }
  aesInvShiftRows_(state);
  aesInvSubBytes_(state, invSbox);
  aesAddRoundKey_(state, w, 0);
  return aesStateToBytes_(state);
}

function aesPkcs7Pad_(bytes) {
  var padLen = 16 - (bytes.length % 16);
  var out = bytes.slice();
  for (var i = 0; i < padLen; i++) out.push(padLen);
  return out;
}

function aesPkcs7Unpad_(bytes) {
  var padLen = bytes[bytes.length - 1];
  if (!padLen || padLen < 1 || padLen > 16 || padLen > bytes.length) throw new Error('Invalid padding');
  return bytes.slice(0, bytes.length - padLen);
}

function aesRandomIv_() {
  var hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').substring(0, 32);
  var bytes = [];
  for (var i = 0; i < 32; i += 2) bytes.push(parseInt(hex.substring(i, i + 2), 16));
  return bytes;
}

function aesToUtf8Bytes_(str) {
  return Utilities.newBlob(str).getBytes().map(function(b) { return b & 0xFF; });
}

function aesFromUtf8Bytes_(bytes) {
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

/** Encrypts plainText with a fresh random IV, returns "base64(iv):base64(cipherText)". */
function aesEncrypt_(plainText, keyBytes) {
  var sboxes = aesGetSboxes_();
  var w = aesKeyExpansion_(keyBytes, sboxes.sbox);
  var padded = aesPkcs7Pad_(aesToUtf8Bytes_(plainText));
  var iv = aesRandomIv_();
  var prev = iv;
  var cipherBytes = [];
  for (var i = 0; i < padded.length; i += 16) {
    var block = padded.slice(i, i + 16);
    var xored = block.map(function(b, idx) { return b ^ prev[idx]; });
    var enc = aesCipherBlock_(xored, w, sboxes.sbox);
    cipherBytes = cipherBytes.concat(enc);
    prev = enc;
  }
  return Utilities.base64Encode(iv) + ':' + Utilities.base64Encode(cipherBytes);
}

/** Reverses aesEncrypt_(); payload must be "base64(iv):base64(cipherText)". */
function aesDecrypt_(payload, keyBytes) {
  var parts = (payload || '').split(':');
  if (parts.length !== 2) throw new Error('Invalid encrypted payload');
  var iv = Utilities.base64Decode(parts[0]).map(function(b) { return b & 0xFF; });
  var cipherBytes = Utilities.base64Decode(parts[1]).map(function(b) { return b & 0xFF; });
  var sboxes = aesGetSboxes_();
  var w = aesKeyExpansion_(keyBytes, sboxes.sbox);
  var prev = iv;
  var plainBytes = [];
  for (var i = 0; i < cipherBytes.length; i += 16) {
    var block = cipherBytes.slice(i, i + 16);
    var dec = aesInvCipherBlock_(block, w, sboxes.sbox, sboxes.invSbox);
    var xored = dec.map(function(b, idx) { return b ^ prev[idx]; });
    plainBytes = plainBytes.concat(xored);
    prev = block;
  }
  return aesFromUtf8Bytes_(aesPkcs7Unpad_(plainBytes));
}

// ─── Key management (mirrors getSessionSecret_(), PO_Manager_Code.gs) ───────
// The raw secret is a random string kept in Script Properties; a SHA-256
// digest of it is used as the 256-bit AES key so the key never needs to be
// stored/transmitted at exactly 32 bytes itself.

function getPortalStorageKeyBytes_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('PORTAL_STORAGE_KEY');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PORTAL_STORAGE_KEY', secret);
  }
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret).map(function(b) { return b & 0xFF; });
}

/**
 * Manual verification only (not called automatically). Run from the Apps
 * Script editor after any change to this file. Checks the block cipher
 * against the official FIPS-197 AES-256 known-answer test, then checks a
 * full encrypt/decrypt/CBC/padding round trip.
 */
function aesSelfTest_() {
  var key = [
    0x00,0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,0x09,0x0a,0x0b,0x0c,0x0d,0x0e,0x0f,
    0x10,0x11,0x12,0x13,0x14,0x15,0x16,0x17,0x18,0x19,0x1a,0x1b,0x1c,0x1d,0x1e,0x1f
  ];
  var plain = [0x00,0x11,0x22,0x33,0x44,0x55,0x66,0x77,0x88,0x99,0xaa,0xbb,0xcc,0xdd,0xee,0xff];
  var expected = '8ea2b7ca516745bfeafc49904b496089'; // FIPS-197 Appendix C.3 AES-256 KAT ciphertext
  var sboxes = aesGetSboxes_();
  var w = aesKeyExpansion_(key, sboxes.sbox);
  var out = aesCipherBlock_(plain, w, sboxes.sbox);
  var hex = out.map(function(b) { return ('0' + b.toString(16)).slice(-2); }).join('');
  if (hex !== expected) throw new Error('AES KAT failed: got ' + hex + ' expected ' + expected);

  var roundTrip = aesDecrypt_(aesEncrypt_('hello portal storage!', getPortalStorageKeyBytes_()), getPortalStorageKeyBytes_());
  if (roundTrip !== 'hello portal storage!') throw new Error('AES round-trip failed: got ' + roundTrip);

  return 'AES self-test passed';
}
