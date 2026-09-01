import { Staff } from './staff.entity';

// Strips fields that must never reach the client: encrypted refresh token
// and raw access token. Everything else (email, connection status, org) is
// safe to expose to the admin dashboard.
export function sanitizeStaff(staff: Staff) {
  const { refreshTokenEnc: _refreshTokenEnc, accessToken: _accessToken, ...safe } = staff;
  return safe;
}
