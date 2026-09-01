process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { StaffService } from './staff.service';
import { encrypt } from '@shared/crypto';
import { Staff } from './staff.entity';

describe('StaffService.getDecryptedRefreshToken', () => {
  it('decrypts a staff record refreshTokenEnc field', () => {
    const service = new StaffService();
    const encrypted = encrypt('my-refresh-token');
    const staff = { refreshTokenEnc: encrypted } as Staff;
    expect(service.getDecryptedRefreshToken(staff)).toBe('my-refresh-token');
  });

  it('returns null when refreshTokenEnc is null', () => {
    const service = new StaffService();
    const staff = { refreshTokenEnc: null } as Staff;
    expect(service.getDecryptedRefreshToken(staff)).toBeNull();
  });
});
