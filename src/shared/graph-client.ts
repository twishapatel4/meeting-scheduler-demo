import axios, { AxiosInstance, AxiosError } from 'axios';
import { GraphAuthError } from './errors/GraphAuthError';
import { GraphNotFoundError } from './errors/GraphNotFoundError';
import { logger } from './utils/logger';

export function createGraphClient(accessToken: string): AxiosInstance {
  const client = axios.create({
    baseURL: 'https://graph.microsoft.com/v1.0',
    headers: {
      common: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        Prefer: 'outlook.timezone="UTC"',
      },
    },
  });

  client.interceptors.response.use(
    (response) => response,
    async (error: AxiosError) => {
      const status = error.response?.status;

      if (status === 429 && error.config && !(error.config as any).__isRetry) {
        const retryAfterHeader = error.response?.headers['retry-after'];
        const retryAfterSeconds = retryAfterHeader ? parseInt(retryAfterHeader, 10) : 1;
        logger.warn('Graph 429 — retrying once', { retryAfterSeconds });
        await new Promise((resolve) => setTimeout(resolve, retryAfterSeconds * 1000));
        (error.config as any).__isRetry = true;
        return client.request(error.config);
      }

      if (status === 401) {
        throw new GraphAuthError();
      }

      if (status === 404) {
        throw new GraphNotFoundError();
      }

      throw error;
    },
  );

  return client;
}
