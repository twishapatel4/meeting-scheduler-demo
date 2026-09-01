# Design: Teams Scheduler Demo (Day-1 Build)

Source spec: `teams-scheduling-spec-delegated.md`. Goal: working end-to-end demo by 7pm today, implementing the full delegated-permissions workflow (steps 1-5), single Azure App Registration, local-only hosting.

## Stack
- Backend: Node.js + TypeScript + Express, following the **vms-backend** pattern (`../VMS/vms-backend`): module-per-feature folders (`entity/service/controller/routes/dto`), path aliases (`@modules/...`), abstract `BaseEntity`, migrations.
- ORM/DB: TypeORM (`DataSource`, `postgres` driver) on existing PostgreSQL/pgAdmin instance — replicates `vms-backend/src/config/database.ts` structure.
- Frontend: React (Vite) — two views: visitor booking form, staff/admin dashboard
- Auth: MSAL Node, authorization-code flow, `common` endpoint (multi-tenant + personal accounts per spec §2)
- Hosting: local only. Redirect URI: `http://localhost:5000/auth/staff/callback`

## Project Structure (mirrors vms-backend)

```
src/
  config/
    env.ts
    database.ts        // TypeORM DataSource, entities list, migrations glob
  shared/
    base.entity.ts      // abstract: uuid PK, createdAt, updatedAt, deletedAt (soft delete)
    crypto.ts            // AES-256-GCM helpers for refresh_token encryption
    graph-client.ts      // shared axios wrapper for Graph calls, 429/401/404 handling
  modules/
    organization/
      organization.entity.ts
      organization.service.ts
    staff/
      staff.entity.ts
      staff.service.ts
      staff.controller.ts
      staff.routes.ts
    auth/
      auth.controller.ts   // OAuth login/callback
      auth.routes.ts
      token.service.ts     // refresh flow, spec §3B
    booking/
      booking.entity.ts
      booking.service.ts   // availability/create/reschedule/cancel/swap, spec §4
      booking.controller.ts
      booking.routes.ts
      booking.dto.ts
  database/
    migrations/
  app.ts
  server.ts
```

## Data Model (TypeORM entities, extend `BaseEntity`)

```ts
// organization.entity.ts
@Entity()
class Organization extends BaseEntity {
  @Column() name: string;
  @Column({ unique: true }) domain: string;
  @OneToMany(() => Staff, (s) => s.organization) staff: Staff[];
}

// staff.entity.ts
@Entity()
class Staff extends BaseEntity {
  @ManyToOne(() => Organization, (o) => o.staff) organization: Organization;
  @Column({ unique: true }) email: string;
  @Column({ nullable: true }) refreshTokenEnc: string | null;
  @Column({ nullable: true }) accessToken: string | null;
  @Column({ type: 'timestamptz', nullable: true }) tokenExpiresAt: Date | null;
  @Column() accountType: string; // enterprise | small_business | personal
  @Column({ default: false }) connected: boolean;
}

// booking.entity.ts
@Entity()
class Booking extends BaseEntity {
  @Column() visitorEmail: string;
  @Column() visitorName: string;
  @Column({ type: 'timestamptz' }) requestedStart: Date;
  @Column({ type: 'timestamptz' }) requestedEnd: Date;
  @Column() status: string; // Requested | Scheduled | Rescheduled | Swapped | Cancelled
  @ManyToOne(() => Staff, { nullable: true }) staff: Staff | null;
  @Column({ nullable: true }) msEventId: string | null;
  @Column({ nullable: true }) joinUrl: string | null;
  @Column({ type: 'timestamptz', nullable: true }) cancelledAt: Date | null;
  @Column({ nullable: true }) cancelledBy: string | null;
}
```

`Organization` is an addition beyond the source spec's table — needed so swap-host (spec §4 Step 5) can be restricted to staff within the same org, while still supporting multiple orgs in one deployment ("works for all organizations"). `id` fields are uuid (matches `BaseEntity` from vms-backend) rather than the spec's implicit auto-increment — no functional difference, matches replicated pattern.

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
