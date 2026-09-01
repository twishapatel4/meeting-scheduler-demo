process.env.TOKEN_ENC_KEY = 'a'.repeat(64); // 32-byte hex key for tests
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { encrypt, decrypt } from './crypto';

describe('crypto', () => {
  it('round-trips a plaintext string', () => {
    const plaintext = 'super-secret-refresh-token-value';
    const ciphertext = encrypt(plaintext);
    expect(ciphertext).not.toBe(plaintext);
    expect(decrypt(ciphertext)).toBe(plaintext);
  });

  it('produces different ciphertext for the same plaintext on repeated calls', () => {
    const a = encrypt('same-value');
    const b = encrypt('same-value');
    expect(a).not.toBe(b);
    expect(decrypt(a)).toBe('same-value');
    expect(decrypt(b)).toBe('same-value');
  });

  it('throws when decrypting tampered ciphertext', () => {
    const ciphertext = encrypt('another-value');
    const tampered = ciphertext.slice(0, -2) + '00';
    expect(() => decrypt(tampered)).toThrow();
  });
});
