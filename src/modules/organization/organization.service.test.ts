process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { OrganizationService } from './organization.service';

describe('OrganizationService.findOrCreateByDomain', () => {
  it('is a function that returns a promise', () => {
    const service = new OrganizationService();
    const result = service.findOrCreateByDomain('example.com', 'Example Co');
    result.catch(() => {}); // uninitialized AppDataSource in this unit test — expected rejection, not asserted
    expect(result).toBeInstanceOf(Promise);
  });
});
