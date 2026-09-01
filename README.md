# Teams Scheduler Demo

Calendly-style meeting scheduler: a visitor requests a meeting, an admin checks staff availability and books a real Microsoft Teams meeting on the staff member's behalf, with reschedule/cancel/swap-host support — all via Microsoft Graph delegated permissions. Works with enterprise, small-business, and personal Microsoft accounts.

Full requirements: [`teams-scheduling-spec-delegated.md`](./teams-scheduling-spec-delegated.md)
Design doc: [`docs/superpowers/specs/2026-09-01-teams-scheduler-demo-design.md`](./docs/superpowers/specs/2026-09-01-teams-scheduler-demo-design.md)
Implementation plan: [`docs/superpowers/plans/2026-09-01-teams-scheduler-demo.md`](./docs/superpowers/plans/2026-09-01-teams-scheduler-demo.md)
Manual smoke-test results: [`docs/superpowers/smoke-test.md`](./docs/superpowers/smoke-test.md)
Known limitations & production TODOs: [`docs/superpowers/demo-shortcuts-vs-production.md`](./docs/superpowers/demo-shortcuts-vs-production.md)

## Architecture

- **Backend** (`src/`): Node.js + TypeScript + Express, TypeORM on PostgreSQL. Module-per-feature (`entity`/`service`/`controller`/`routes`/`dto`), typedi for dependency injection, MSAL Node for OAuth, path aliases (`@modules/*`, `@shared/*`, `@config/*`, `@database/*`).
- **Frontend** (`frontend/`): Vite + React. Two pages — visitor booking form (`/`) and admin dashboard (`/admin`).
- **Database**: PostgreSQL, migrations under `src/database/migrations/`.
- **Auth**: MSAL Node authorization-code flow against `https://login.microsoftonline.com/common` (multi-tenant + personal Microsoft accounts). Each staff member connects once via OAuth; the backend stores their encrypted refresh token and refreshes access tokens on demand.

## Prerequisites

- Node.js 18+
- PostgreSQL running locally (or reachable via `DATABASE_URL`)
- An Azure App Registration (see below)

## 1. Azure App Registration

1. [Azure Portal](https://portal.azure.com) → **Entra ID** → **App registrations** → **New registration**.
2. **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts" (required — this is what makes personal + work/school accounts both work).
3. **Redirect URI**: platform **Web**, `http://localhost:5000/auth/staff/callback`.
4. **Certificates & secrets** → New client secret → copy the value (shown once).
5. **API permissions** → Add a permission → Microsoft Graph → **Delegated**: `offline_access`, `Calendars.ReadWrite`, `OnlineMeetings.ReadWrite`, `User.Read`.
6. Copy the **Application (client) ID** from the Overview page.

> Note: organizational (work/school) tenants with restrictive consent policies may require a tenant admin to grant admin consent for the app before staff can connect — see the known-limitations doc for details.

## 2. Database setup

```bash
createdb teams_scheduler
```

(Or create it via pgAdmin / any Postgres client — any empty database works, migrations create the schema.)

## 3. Backend setup

```bash
npm install
cp .env.example .env
```

Edit `.env`:

```
NODE_ENV=development
PORT=5000
DATABASE_URL=postgres://<user>:<password>@localhost:5432/teams_scheduler
TOKEN_ENC_KEY=<32-byte hex string>       # generate: node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
MS_CLIENT_ID=<your Azure app client ID>
MS_CLIENT_SECRET=<your Azure app client secret>
MS_REDIRECT_URI=http://localhost:5000/auth/staff/callback
FRONTEND_URL=http://localhost:5173
```

Run migrations:

```bash
npm run migration:run
```

Start the backend:

```bash
npm run dev
```

Runs on `http://localhost:5000`. Check `http://localhost:5000/health` for `{"status":"ok"}`.

## 4. Frontend setup

```bash
cd frontend
npm install
npm run dev
```

Runs on `http://localhost:5173`.

## 5. Using the app

1. Open `http://localhost:5173/admin` → click **"+ Connect Staff Account"** → sign in with a Microsoft account and consent. You're redirected back with the account listed as "Connected".
2. Open `http://localhost:5173/` → submit a visitor booking request (name, email, meeting subject, preferred time).
3. Back on `/admin`, click **Schedule**, pick a connected staff member → creates a real Teams meeting on their calendar (work/school accounts get a Teams join link; personal Microsoft accounts get a plain calendar invite — see the limitations doc).
4. Reschedule / Cancel / Swap / Delete are available from the bookings table once a meeting is scheduled. Swap requires a second connected staff account in the **same organization** (same email domain).

## Testing

```bash
npm run test          # backend: jest, from repo root
npx tsc --noEmit       # backend typecheck
cd frontend && npm run build   # frontend typecheck + production build
```

## Project structure

```
src/
  config/          env validation, TypeORM DataSource
  shared/          base entity, crypto (AES-256-GCM), errors, Graph client wrapper, middleware
  modules/
    organization/  org grouping for same-org swap-host restriction
    staff/         staff entities, OAuth-connected accounts
    auth/          MSAL token exchange/refresh
    booking/       core booking workflow (spec §4 steps 1-5)
  database/migrations/
  app.ts, server.ts

frontend/src/
  pages/           VisitorBookingPage, AdminDashboardPage
  api.ts           backend API client helpers
```
