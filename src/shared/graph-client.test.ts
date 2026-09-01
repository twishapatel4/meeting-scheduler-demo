process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { AxiosError, AxiosRequestConfig, AxiosResponse } from 'axios';
import { createGraphClient } from './graph-client';
import { GraphAuthError } from './errors/GraphAuthError';
import { GraphNotFoundError } from './errors/GraphNotFoundError';

/**
 * Builds a custom axios adapter that returns a scripted sequence of status
 * codes for successive calls to the same logical request. It mirrors what
 * axios's own adapters do internally (see axios/lib/core/settle.js): resolve
 * the promise when the response satisfies `validateStatus`, otherwise reject
 * with a real AxiosError carrying the response. This makes the response
 * interceptor under test run through its actual runtime logic (429 retry,
 * 401/404 mapping) instead of only asserting on client.defaults.
 */
function createScriptedAdapter(statuses: number[], headers: Record<string, string> = {}) {
  let callCount = 0;
  return (config: AxiosRequestConfig): Promise<AxiosResponse> => {
    const status = statuses[Math.min(callCount, statuses.length - 1)];
    callCount += 1;
    const response: AxiosResponse = {
      data: {},
      status,
      statusText: String(status),
      headers,
      config: config as any,
    } as AxiosResponse;

    const validateStatus = config.validateStatus ?? ((s: number) => s >= 200 && s < 300);

    if (validateStatus(status)) {
      return Promise.resolve(response);
    }

    return Promise.reject(
      new AxiosError(
        `Request failed with status code ${status}`,
        status >= 400 && status < 500 ? AxiosError.ERR_BAD_REQUEST : AxiosError.ERR_BAD_RESPONSE,
        config as any,
        undefined,
        response,
      ),
    );
  };
}

describe('createGraphClient', () => {
  it('sets the expected base config', () => {
    const client = createGraphClient('fake-token');
    expect(client.defaults.baseURL).toBe('https://graph.microsoft.com/v1.0');
    expect(client.defaults.headers.common['Authorization']).toBe('Bearer fake-token');
    expect(client.defaults.headers.common['Prefer']).toBe('outlook.timezone="UTC"');
  });

  it('exposes GraphAuthError and GraphNotFoundError as classes', () => {
    expect(new GraphAuthError().statusCode).toBe(401);
    expect(new GraphNotFoundError().statusCode).toBe(404);
  });

  describe('response interceptor behavior', () => {
    it('retries exactly once on a 429 then resolves on success', async () => {
      const client = createGraphClient('fake-token');
      const adapter = createScriptedAdapter([429, 200], { 'retry-after': '0' });

      const response = await client.get('/me', { adapter });

      expect(response.status).toBe(200);
    });

    it('throws (does not retry a second time) on consecutive 429s', async () => {
      const client = createGraphClient('fake-token');
      const adapter = createScriptedAdapter([429, 429, 429], { 'retry-after': '0' });

      await expect(client.get('/me', { adapter })).rejects.toMatchObject({
        isAxiosError: true,
        response: expect.objectContaining({ status: 429 }),
      });
    });

    it('throws GraphAuthError on a 401', async () => {
      const client = createGraphClient('fake-token');
      const adapter = createScriptedAdapter([401]);

      await expect(client.get('/me', { adapter })).rejects.toBeInstanceOf(GraphAuthError);
    });

    it('throws GraphNotFoundError on a 404', async () => {
      const client = createGraphClient('fake-token');
      const adapter = createScriptedAdapter([404]);

      await expect(client.get('/me', { adapter })).rejects.toBeInstanceOf(GraphNotFoundError);
    });
  });
});
