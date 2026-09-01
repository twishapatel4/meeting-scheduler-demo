# Design: Teams Scheduler Demo (Day-1 Build)

Source spec: `teams-scheduling-spec-delegated.md`. Goal: working end-to-end demo by 7pm today, implementing the full delegated-permissions workflow (steps 1-5), single Azure App Registration, local-only hosting.

## Stack
- Backend: Node.js + Express (plain JS, no TypeScript — speed priority)
- ORM/DB: Prisma on existing PostgreSQL/pgAdmin instance
- Frontend: React (Vite) — two views: visitor booking form, staff/admin dashboard
- Auth: MSAL Node, authorization-code flow, `common` endpoint (multi-tenant + personal accounts per spec §2)
- Hosting: local only. Redirect URI: `http://localhost:5000/auth/staff/callback`

## Data Model (Prisma schema)

```
Organization
  id            Int    @id @default(autoincrement())
  name          String
  domain        String @unique

Staff
  id                Int      @id @default(autoincrement())
  orgId             Int
  organization      Organization @relation(fields: [orgId], references: [id])
  email             String   @unique
  refreshTokenEnc   String?  // AES-256-GCM encrypted
  accessToken       String?
  tokenExpiresAt    DateTime?
  accountType       String   // enterprise | small_business | personal
  connected         Boolean  @default(false)

Booking
  id              Int      @id @default(autoincrement())
  visitorEmail    String
  visitorName     String
  requestedStart  DateTime
  requestedEnd    DateTime
  status          String   // Requested | Scheduled | Rescheduled | Swapped | Cancelled
  staffId         Int?
  staff           Staff?   @relation(fields: [staffId], references: [id])
  msEventId       String?
  joinUrl         String?
  cancelledAt     DateTime?
  cancelledBy     String?
  createdAt       DateTime @default(now())
  updatedAt       DateTime @updatedAt
```

`Organization` is an addition beyond the source spec's table — needed so swap-host (spec §4 Step 5) can be restricted to staff within the same org, while still supporting multiple orgs in one deployment ("works for all organizations").

## Auth Flow
1. Staff clicks "Connect" on admin dashboard → redirected to Microsoft OAuth consent (`/authorize` via MSAL, `common` endpoint, scopes: `offline_access Calendars.ReadWrite OnlineMeetings.ReadWrite User.Read`).
2. Callback (`/auth/staff/callback`) exchanges code for tokens, encrypts `refresh_token` (AES-256-GCM, key from `TOKEN_ENC_KEY` env var), upserts `Staff` row, sets `connected = true`.
3. Background/on-demand refresh: before any Graph call, check `tokenExpiresAt`; if near expiry, POST to token endpoint (spec §3B) and update DB with rotated refresh_token.
4. `invalid_grant` on refresh → set `connected = false`, surface "needs re-consent" in admin UI. No silent retry.

## API Surface

| Route | Method | Purpose |
|---|---|---|
| `/auth/staff/login` | GET | Start OAuth for a staff member |
| `/auth/staff/callback` | GET | OAuth callback, stores tokens |
| `/api/visitor/request` | POST | Visitor submits booking request → `Booking` row, `status=Requested` |
| `/api/staff/:id/availability` | GET | Wraps Graph `getSchedule` (spec §4 Step 1) |
| `/api/bookings` | POST | Admin creates Teams meeting for chosen staff (spec §4 Step 2) → `status=Scheduled` |
| `/api/bookings/:id` | PATCH | Reschedule (spec §4 Step 3) → `status=Rescheduled` |
| `/api/bookings/:id/cancel` | POST | Cancel w/ visitor notice (spec §4 Step 4) → `status=Cancelled` |
| `/api/bookings/:id` | DELETE | Hard delete, no notice (spec §4 Step 4) |
| `/api/bookings/:id/swap` | POST | Cancel + rebook w/ different staff, same org only (spec §4 Step 5) → `status=Swapped` |
| `/api/admin/staff` | GET | List staff + connection status |

## Error Handling (spec §7)
- 429: read `Retry-After` header, backoff/retry once.
- 401 / `invalid_grant` on refresh: mark staff disconnected, no silent retry.
- 404 on event (already deleted in Outlook directly): catch, mark booking status accordingly, don't crash the reschedule/cancel/swap flow.
- Graph webhooks/change notifications: explicitly OUT of scope for today (spec §7 calls this fast-follow, not v1-blocking).

## Security (spec §9)
- `refresh_token` encrypted at rest via AES-256-GCM with env-var key — demo-grade stand-in for Key Vault/KMS. Flagged as a shortcut, not production-ready.
- Scopes limited to the four listed, no `.All`/app-only grants.
- All Graph timestamps sent/stored in UTC; convert only for display.

## Explicit Cuts (not silent)
- No real KMS/Key Vault integration — env-var key only.
- No Graph webhook subscriptions for drift detection.
- No production deploy — localhost only.

## Open Decision (carried from spec §10)
Swap-host does not preserve the same `joinUrl` (no shared service-account calendar). Flagged as a per-tier product decision, not a today-blocker.
