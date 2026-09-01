process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { BookingService } from './booking.service';

describe('BookingService.assertSameOrganization', () => {
  const service = new BookingService(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    { findById: undefined } as any,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    undefined as any,
  );

  it('does not throw when organization ids match', () => {
    expect(() => service.assertSameOrganization('org-1', 'org-1')).not.toThrow();
  });

  it('throws BadRequestError when organization ids differ', () => {
    expect(() => service.assertSameOrganization('org-1', 'org-2')).toThrow(
      'Swap-host is only allowed between staff in the same organization',
    );
  });
});
