process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { StaffService } from '@modules/staff/staff.service';
import { TokenService, extractRefreshTokenForAccount } from './token.service';

describe('TokenService.isNearExpiry', () => {
  // isNearExpiry never touches staffService, so a stub is sufficient here.
  const service = new TokenService({} as StaffService);

  it('returns true when expiry is within 5 minutes', () => {
    const soon = new Date(Date.now() + 2 * 60 * 1000);
    expect(service.isNearExpiry(soon)).toBe(true);
  });

  it('returns false when expiry is well in the future', () => {
    const later = new Date(Date.now() + 60 * 60 * 1000);
    expect(service.isNearExpiry(later)).toBe(false);
  });

  it('returns true when expiry is null', () => {
    expect(service.isNearExpiry(null)).toBe(true);
  });
});

describe('extractRefreshTokenForAccount', () => {
  // Simulates the shared msalClient singleton's cache holding refresh tokens
  // for two different staff members at once (the exact scenario that caused
  // the "second staff member gets the wrong token" bug).
  const cacheJson = JSON.stringify({
    RefreshToken: {
      'entry-1': {
        secret: 'refresh-token-for-alice',
        home_account_id: 'alice-home-account-id',
      },
      'entry-2': {
        secret: 'refresh-token-for-bob',
        home_account_id: 'bob-home-account-id',
      },
    },
  });

  it('picks the entry matching the given homeAccountId', () => {
    const result = extractRefreshTokenForAccount(cacheJson, 'bob-home-account-id');
    expect(result?.secret).toBe('refresh-token-for-bob');
  });

  it('picks the correct entry regardless of cache ordering', () => {
    const result = extractRefreshTokenForAccount(cacheJson, 'alice-home-account-id');
    expect(result?.secret).toBe('refresh-token-for-alice');
  });

  it('returns undefined when no entry matches', () => {
    const result = extractRefreshTokenForAccount(cacheJson, 'unknown-home-account-id');
    expect(result).toBeUndefined();
  });

  it('returns undefined when the cache has no RefreshToken entries', () => {
    const result = extractRefreshTokenForAccount(JSON.stringify({}), 'alice-home-account-id');
    expect(result).toBeUndefined();
  });
});
