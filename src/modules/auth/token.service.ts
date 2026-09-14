import { Service } from 'typedi';
import axios from 'axios';
import * as msal from '@azure/msal-node';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';
import { StaffService } from '@modules/staff/staff.service';
import { UnauthorizedError } from '@shared/errors/UnauthorizedError';

const SCOPES = [
  'offline_access',
  'Calendars.ReadWrite',
  'OnlineMeetings.ReadWrite',
  'OnlineMeetingTranscript.Read.All',
  'OnlineMeetingAiInsight.Read.All',
  'User.Read',
];
const NEAR_EXPIRY_MS = 5 * 60 * 1000;

// MSAL Node's serialized token cache stores each RefreshToken entry with a
// snake_case `home_account_id` field (see @azure/msal-node's
// Serializer.serializeRefreshTokens), NOT the camelCase `homeAccountId` used
// on the in-memory AccountInfo/entity types. Confirmed by reading
// node_modules/@azure/msal-node/src/cache/serializer/Serializer.ts.
interface SerializedRefreshTokenCacheEntry {
  secret: string;
  home_account_id: string;
}

interface SerializedMsalCache {
  RefreshToken?: Record<string, SerializedRefreshTokenCacheEntry>;
}

// Pure/testable: pick the refresh-token cache entry belonging to the given
// homeAccountId. The msalClient is a shared singleton, so its cache can hold
// refresh tokens for multiple staff members at once — we must not just grab
// the first entry in the object.
export function extractRefreshTokenForAccount(
  cacheJson: string,
  homeAccountId: string,
): { secret: string } | undefined {
  const parsed = JSON.parse(cacheJson) as SerializedMsalCache;
  const entries = Object.values(parsed.RefreshToken ?? {});
  return entries.find((entry) => entry.home_account_id === homeAccountId);
}

@Service()
export class TokenService {
  private msalClient = new msal.ConfidentialClientApplication({
    auth: {
      clientId: env.MS_CLIENT_ID,
      authority: 'https://login.microsoftonline.com/common',
      clientSecret: env.MS_CLIENT_SECRET,
    },
  });

  constructor(private readonly staffService: StaffService) {}

  async getAuthCodeUrl(): Promise<string> {
    return this.msalClient.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri: env.MS_REDIRECT_URI,
    });
  }

  async exchangeCodeForTokens(code: string): Promise<{
    accessToken: string;
    refreshToken: string;
    expiresOn: Date;
    email: string;
  }> {
    const result = await this.msalClient.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: env.MS_REDIRECT_URI,
    });

    if (!result || !result.account) {
      throw new UnauthorizedError('Microsoft did not return an account for this code');
    }

    // MSAL Node's public API does not expose the raw refresh_token from
    // acquireTokenByCode; the token cache holds it internally. We pull it
    // back out of the cache immediately after exchange so we can encrypt
    // and persist it ourselves (StaffService owns the DB, not MSAL's cache).
    //
    // IMPORTANT: TokenService is a typedi singleton with one shared
    // msalClient, so its cache can accumulate refresh-token entries for
    // multiple staff members across the process lifetime. We must select
    // the entry for THIS account (by homeAccountId), not just the first one
    // in the cache object.
    const cache = this.msalClient.getTokenCache().serialize();
    const refreshTokenEntry = extractRefreshTokenForAccount(cache, result.account.homeAccountId);

    if (!refreshTokenEntry) {
      throw new UnauthorizedError('No refresh token returned — ensure offline_access scope is granted');
    }

    return {
      accessToken: result.accessToken,
      refreshToken: refreshTokenEntry.secret,
      expiresOn: result.expiresOn ?? new Date(Date.now() + 3600 * 1000),
      email: result.account.username,
    };
  }

  isNearExpiry(expiresAt: Date | null): boolean {
    if (!expiresAt) return true;
    return expiresAt.getTime() - Date.now() < NEAR_EXPIRY_MS;
  }

  // Returns a valid access token for the given staff, refreshing via the
  // Graph token endpoint (spec §3B) if the cached one is near expiry.
  async getValidAccessToken(staffId: string): Promise<string> {
    const staff = await this.staffService.findById(staffId);
    if (!staff || !staff.connected) {
      throw new UnauthorizedError('Staff is not connected — needs re-consent');
    }

    if (staff.accessToken && !this.isNearExpiry(staff.tokenExpiresAt)) {
      return staff.accessToken;
    }

    const refreshToken = this.staffService.getDecryptedRefreshToken(staff);
    if (!refreshToken) {
      await this.staffService.markDisconnected(staff.id);
      throw new UnauthorizedError('No refresh token on file — needs re-consent');
    }

    try {
      const response = await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        new URLSearchParams({
          client_id: env.MS_CLIENT_ID,
          client_secret: env.MS_CLIENT_SECRET,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          scope: SCOPES.join(' '),
        }),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } },
      );

      const { access_token, refresh_token, expires_in } = response.data;
      const expiresAt = new Date(Date.now() + expires_in * 1000);

      await this.staffService.updateTokens(staff.id, {
        accessToken: access_token,
        refreshToken: refresh_token,
        expiresAt,
      });

      return access_token;
    } catch (err) {
      const isInvalidGrant =
        axios.isAxiosError(err) && err.response?.data?.error === 'invalid_grant';

      if (isInvalidGrant) {
        logger.warn('Refresh token revoked — marking staff disconnected', { staffId: staff.id });
        await this.staffService.markDisconnected(staff.id);
        throw new UnauthorizedError('Staff refresh token revoked — needs re-consent');
      }

      throw err;
    }
  }
}
