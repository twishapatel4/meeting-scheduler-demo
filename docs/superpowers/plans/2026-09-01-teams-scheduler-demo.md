# Teams Scheduler Demo Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a working end-to-end demo of the delegated-permissions Teams scheduler (visitor request → admin checks staff availability → books/reschedules/cancels/swaps a Teams meeting via Microsoft Graph) by end of day, local-only.

**Architecture:** Node.js + TypeScript + Express backend replicating the `vms-backend` pattern (module-per-feature: entity/service/controller/routes/dto, typedi DI, TypeORM repositories, path aliases, Joi env validation, winston logging, centralized error handling) on PostgreSQL, plus a minimal Vite/React frontend (visitor booking form + admin dashboard).

**Tech Stack:** Express, TypeORM 0.3.x, PostgreSQL (`pg`), typedi, class-validator/class-transformer, Joi, winston, `@azure/msal-node`, axios, uuid, Vite + React. TypeScript throughout, path aliases via `tsconfig-paths`/`tsc-alias`.

## Global Constraints

- Graph scopes (exact, delegated only, no `.All`/app-only): `offline_access Calendars.ReadWrite OnlineMeetings.ReadWrite User.Read`
- OAuth authority: `https://login.microsoftonline.com/common` (multi-tenant + personal accounts — spec §2)
- Token refresh endpoint: `POST https://login.microsoftonline.com/common/oauth2/v2.0/token`, form-encoded, `grant_type=refresh_token`
- Redirect URI (registered in Azure App Registration): `http://localhost:5000/auth/staff/callback`
- All Graph request/response timestamps are UTC (`Prefer: outlook.timezone="UTC"` header on every Graph call); convert to local only for display
- `refresh_token` must be encrypted at rest (AES-256-GCM, key from `TOKEN_ENC_KEY` env var) — never logged in plaintext
- `invalid_grant` on any refresh attempt → mark staff `connected = false`, do not silently retry
- 429 responses → respect `Retry-After` header, one retry
- 404 on a Graph event (already deleted directly in Outlook) → handle gracefully, do not throw an unhandled error
- Booking `status` lifecycle (spec §8): `Requested → Scheduled → Rescheduled → Swapped → Cancelled` (any status can transition to `Cancelled`)
- Swap-host is restricted to staff within the same `Organization` (see design doc §Data Model)
- No production deploy, no Graph webhooks/change-notifications, no real KMS — these are explicit, documented cuts for today (design doc "Explicit Cuts")

---

## Task 1: Project Scaffold

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `tsconfig.build.json`
- Create: `nodemon.json`
- Create: `.env.example`
- Create: `.gitignore`
- Create: `src/config/paths.ts` (placeholder for later tasks, empty export for now)

**Interfaces:**
- Produces: npm scripts `dev`, `build`, `start`, `test`, `migration:run`. Path aliases `@modules/*`, `@shared/*`, `@config/*`, `@database/*` resolving to `src/modules/*`, `src/shared/*`, `src/config/*`, `src/database/*`.

- [ ] **Step 1: Create `package.json`**

```json
{
  "name": "teams-scheduler-backend",
  "version": "1.0.0",
  "description": "Teams Scheduler — delegated-permissions demo backend",
  "main": "dist/server.js",
  "scripts": {
    "dev": "nodemon",
    "build": "tsc -p tsconfig.build.json && tsc-alias -p tsconfig.build.json",
    "start": "node dist/server.js",
    "test": "jest",
    "migration:generate": "ts-node -r tsconfig-paths/register node_modules/typeorm/cli.js migration:generate -d src/config/database.ts",
    "migration:run": "ts-node -r tsconfig-paths/register node_modules/typeorm/cli.js migration:run -d src/config/database.ts",
    "migration:revert": "ts-node -r tsconfig-paths/register node_modules/typeorm/cli.js migration:revert -d src/config/database.ts"
  },
  "dependencies": {
    "@azure/msal-node": "^2.13.1",
    "axios": "^1.7.7",
    "class-transformer": "^0.5.1",
    "class-validator": "^0.14.1",
    "cors": "^2.8.5",
    "dotenv": "^16.4.5",
    "express": "^4.19.2",
    "helmet": "^7.1.0",
    "joi": "^17.13.3",
    "pg": "^8.12.0",
    "reflect-metadata": "^0.2.2",
    "typedi": "^0.10.0",
    "typeorm": "^0.3.20",
    "uuid": "^9.0.1",
    "winston": "^3.14.2"
  },
  "devDependencies": {
    "@types/cors": "^2.8.17",
    "@types/express": "^4.17.21",
    "@types/jest": "^29.5.12",
    "@types/node": "^20.14.15",
    "@types/supertest": "^6.0.2",
    "@types/uuid": "^9.0.8",
    "jest": "^29.7.0",
    "nodemon": "^3.1.4",
    "supertest": "^7.0.0",
    "ts-jest": "^29.2.4",
    "ts-node": "^10.9.2",
    "tsc-alias": "^1.8.10",
    "tsconfig-paths": "^4.2.0",
    "typescript": "^5.5.4"
  }
}
```

- [ ] **Step 2: Create `tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2021",
    "module": "commonjs",
    "lib": ["ES2021"],
    "outDir": "./dist",
    "rootDir": "./src",
    "strict": true,
    "strictPropertyInitialization": false,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "experimentalDecorators": true,
    "emitDecoratorMetadata": true,
    "resolveJsonModule": true,
    "baseUrl": ".",
    "paths": {
      "@modules/*": ["src/modules/*"],
      "@shared/*": ["src/shared/*"],
      "@config/*": ["src/config/*"],
      "@database/*": ["src/database/*"]
    }
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist"],
  "ts-node": {
    "files": true
  }
}
```

- [ ] **Step 3: Create `tsconfig.build.json`**

```json
{
  "extends": "./tsconfig.json",
  "exclude": ["node_modules", "dist", "src/**/*.test.ts"]
}
```

- [ ] **Step 4: Create `nodemon.json`**

```json
{
  "watch": ["src"],
  "ext": "ts",
  "exec": "ts-node -r tsconfig-paths/register src/server.ts"
}
```

- [ ] **Step 5: Create `.env.example`**

```
NODE_ENV=development
PORT=5000
DATABASE_URL=postgres://postgres:postgres@localhost:5432/teams_scheduler
TOKEN_ENC_KEY=replace_with_32_byte_hex_key_generated_via_crypto_randomBytes
MS_CLIENT_ID=replace_with_azure_app_client_id
MS_CLIENT_SECRET=replace_with_azure_app_client_secret
MS_REDIRECT_URI=http://localhost:5000/auth/staff/callback
FRONTEND_URL=http://localhost:5173
```

- [ ] **Step 6: Create `.gitignore`**

```
node_modules/
dist/
.env
*.log
```

- [ ] **Step 7: Create `src/config/paths.ts`**

```ts
export {};
```

- [ ] **Step 8: Install dependencies and verify TypeScript compiles an empty project**

Run: `npm install`
Expected: installs without errors.

Run: `npx tsc --noEmit`
Expected: no output (no source files yet besides `paths.ts`, which is valid).

- [ ] **Step 9: Commit**

```bash
git init
git add package.json tsconfig.json tsconfig.build.json nodemon.json .env.example .gitignore src/config/paths.ts
git commit -m "chore: scaffold TypeScript/Express project"
```

---

## Task 2: Shared Foundation (env, logger, errors, api response, base entity, crypto)

**Files:**
- Create: `src/config/env.ts`
- Create: `src/shared/utils/logger.ts`
- Create: `src/shared/utils/apiResponse.ts`
- Create: `src/shared/utils/asyncHandler.ts`
- Create: `src/shared/errors/AppError.ts`
- Create: `src/shared/errors/NotFoundError.ts`
- Create: `src/shared/errors/UnauthorizedError.ts`
- Create: `src/shared/errors/BadRequestError.ts`
- Create: `src/shared/errors/ConflictError.ts`
- Create: `src/shared/base.entity.ts`
- Create: `src/shared/crypto.ts`
- Test: `src/shared/crypto.test.ts`
- Create: `jest.config.js`

**Interfaces:**
- Consumes: none (foundation layer).
- Produces: `env: { NODE_ENV, PORT, DATABASE_URL, TOKEN_ENC_KEY, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_REDIRECT_URI, FRONTEND_URL }`; `logger.info/warn/error(msg, meta?)`; `sendSuccess(res, data, message?)`, `sendError(res, message, statusCode, errorCode?)`; `asyncHandler(fn)`; `AppError(message, statusCode, errorCode)` and subclasses `NotFoundError`, `UnauthorizedError`, `BadRequestError`, `ConflictError`; `abstract class BaseEntity { id: string; createdAt: Date; updatedAt: Date; deletedAt: Date | null }`; `encrypt(plaintext: string): string`, `decrypt(ciphertext: string): string`.

- [ ] **Step 1: Create `jest.config.js`**

```js
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  moduleNameMapper: {
    '^@modules/(.*)$': '<rootDir>/src/modules/$1',
    '^@shared/(.*)$': '<rootDir>/src/shared/$1',
    '^@config/(.*)$': '<rootDir>/src/config/$1',
    '^@database/(.*)$': '<rootDir>/src/database/$1',
  },
};
```

- [ ] **Step 2: Write the failing test for `crypto.ts`**

Create `src/shared/crypto.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64); // 32-byte hex key for tests
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { encrypt, decrypt } from './crypto';

describe('crypto', () => {
  it('round-trips a plaintext string', () => {
    const plaintext = 'super-secret-refresh-token-value';
    const ciphertext = encrypt(plaintext);
    expect(ciphertext).not.toBe(plaintext);
    expect(decrypt(ciphertext)).toBe(plaintext);
  });

  it('produces different ciphertext for the same plaintext on repeated calls', () => {
    const a = encrypt('same-value');
    const b = encrypt('same-value');
    expect(a).not.toBe(b);
    expect(decrypt(a)).toBe('same-value');
    expect(decrypt(b)).toBe('same-value');
  });

  it('throws when decrypting tampered ciphertext', () => {
    const ciphertext = encrypt('another-value');
    const tampered = ciphertext.slice(0, -2) + '00';
    expect(() => decrypt(tampered)).toThrow();
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npx jest src/shared/crypto.test.ts`
Expected: FAIL — `Cannot find module './crypto'`

- [ ] **Step 4: Create `src/config/env.ts`**

```ts
import Joi from 'joi';
import dotenv from 'dotenv';

dotenv.config();

const envSchema = Joi.object({
  NODE_ENV: Joi.string().valid('development', 'production', 'test').default('development'),
  PORT: Joi.number().default(5000),
  DATABASE_URL: Joi.string().uri().required().messages({
    'string.uri': 'DATABASE_URL must be a valid PostgreSQL URL',
    'any.required': 'DATABASE_URL is required',
  }),
  TOKEN_ENC_KEY: Joi.string()
    .length(64)
    .hex()
    .required()
    .messages({
      'string.length': 'TOKEN_ENC_KEY must be a 32-byte hex string (64 hex chars)',
      'any.required': 'TOKEN_ENC_KEY is required',
    }),
  MS_CLIENT_ID: Joi.string().required(),
  MS_CLIENT_SECRET: Joi.string().required(),
  MS_REDIRECT_URI: Joi.string().uri().default('http://localhost:5000/auth/staff/callback'),
  FRONTEND_URL: Joi.string().uri().default('http://localhost:5173'),
}).unknown(true);

const { error, value } = envSchema.validate(process.env, { abortEarly: false });

if (error) {
  console.error('Invalid environment variables:');
  console.error(error.details.map((d) => `  - ${d.message}`).join('\n'));
  process.exit(1);
}

export const env = value as {
  NODE_ENV: 'development' | 'production' | 'test';
  PORT: number;
  DATABASE_URL: string;
  TOKEN_ENC_KEY: string;
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  MS_REDIRECT_URI: string;
  FRONTEND_URL: string;
};
```

- [ ] **Step 5: Create `src/shared/crypto.ts`**

```ts
import crypto from 'crypto';
import { env } from '@config/env';

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;

// Format: iv:authTag:ciphertext, all hex-encoded.
export function encrypt(plaintext: string): string {
  const key = Buffer.from(env.TOKEN_ENC_KEY, 'hex');
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const authTag = cipher.getAuthTag();
  return `${iv.toString('hex')}:${authTag.toString('hex')}:${encrypted.toString('hex')}`;
}

export function decrypt(ciphertext: string): string {
  const key = Buffer.from(env.TOKEN_ENC_KEY, 'hex');
  const [ivHex, authTagHex, encryptedHex] = ciphertext.split(':');
  if (!ivHex || !authTagHex || !encryptedHex) {
    throw new Error('Malformed ciphertext');
  }
  const decipher = crypto.createDecipheriv(ALGORITHM, key, Buffer.from(ivHex, 'hex'));
  decipher.setAuthTag(Buffer.from(authTagHex, 'hex'));
  const decrypted = Buffer.concat([
    decipher.update(Buffer.from(encryptedHex, 'hex')),
    decipher.final(),
  ]);
  return decrypted.toString('utf8');
}
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx jest src/shared/crypto.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 7: Create `src/shared/utils/logger.ts`**

```ts
import winston from 'winston';
import { env } from '@config/env';

export const logger = winston.createLogger({
  level: env.NODE_ENV === 'development' ? 'debug' : 'info',
  format: winston.format.combine(winston.format.timestamp(), winston.format.json()),
  transports: [
    new winston.transports.Console({
      format:
        env.NODE_ENV === 'development'
          ? winston.format.combine(winston.format.colorize(), winston.format.simple())
          : winston.format.json(),
    }),
  ],
});
```

- [ ] **Step 8: Create `src/shared/errors/AppError.ts`**

```ts
export class AppError extends Error {
  constructor(
    message: string,
    public readonly statusCode: number,
    public readonly errorCode: string,
  ) {
    super(message);
    this.name = this.constructor.name;
    Error.captureStackTrace(this, this.constructor);
  }
}
```

- [ ] **Step 9: Create the four `AppError` subclasses**

`src/shared/errors/NotFoundError.ts`:

```ts
import { AppError } from './AppError';

export class NotFoundError extends AppError {
  constructor(message = 'Resource not found') {
    super(message, 404, 'NOT_FOUND');
  }
}
```

`src/shared/errors/UnauthorizedError.ts`:

```ts
import { AppError } from './AppError';

export class UnauthorizedError extends AppError {
  constructor(message = 'Unauthorized') {
    super(message, 401, 'UNAUTHORIZED');
  }
}
```

`src/shared/errors/BadRequestError.ts`:

```ts
import { AppError } from './AppError';

export class BadRequestError extends AppError {
  constructor(message = 'Bad request') {
    super(message, 400, 'BAD_REQUEST');
  }
}
```

`src/shared/errors/ConflictError.ts`:

```ts
import { AppError } from './AppError';

export class ConflictError extends AppError {
  constructor(message = 'Conflict') {
    super(message, 409, 'CONFLICT');
  }
}
```

- [ ] **Step 10: Create `src/shared/utils/apiResponse.ts`**

```ts
import { Response } from 'express';

export function sendSuccess<T>(res: Response, data: T, message = 'OK', statusCode = 200): void {
  res.status(statusCode).json({ success: true, message, data });
}

export function sendError(
  res: Response,
  message: string,
  statusCode = 500,
  errorCode = 'INTERNAL_SERVER_ERROR',
): void {
  res.status(statusCode).json({ success: false, message, errorCode });
}
```

- [ ] **Step 11: Create `src/shared/utils/asyncHandler.ts`**

```ts
import { Request, Response, NextFunction, RequestHandler } from 'express';

export function asyncHandler(
  fn: (req: Request, res: Response, next: NextFunction) => Promise<void>,
): RequestHandler {
  return (req, res, next) => {
    fn(req, res, next).catch(next);
  };
}
```

- [ ] **Step 12: Create `src/shared/base.entity.ts`**

```ts
import {
  PrimaryGeneratedColumn,
  CreateDateColumn,
  UpdateDateColumn,
  DeleteDateColumn,
} from 'typeorm';

export abstract class BaseEntity {
  @PrimaryGeneratedColumn('uuid')
  id: string;

  @CreateDateColumn({ type: 'timestamptz' })
  createdAt: Date;

  @UpdateDateColumn({ type: 'timestamptz' })
  updatedAt: Date;

  @DeleteDateColumn({ type: 'timestamptz', nullable: true })
  deletedAt: Date | null;
}
```

- [ ] **Step 13: Verify full test suite and typecheck pass**

Run: `npx jest`
Expected: PASS (3 tests)

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 14: Commit**

```bash
git add jest.config.js src/config/env.ts src/shared/crypto.ts src/shared/crypto.test.ts \
  src/shared/utils/logger.ts src/shared/utils/apiResponse.ts src/shared/utils/asyncHandler.ts \
  src/shared/errors/AppError.ts src/shared/errors/NotFoundError.ts src/shared/errors/UnauthorizedError.ts \
  src/shared/errors/BadRequestError.ts src/shared/errors/ConflictError.ts src/shared/base.entity.ts
git commit -m "feat: add shared foundation (env, logger, errors, crypto, base entity)"
```

---

## Task 3: Database Config Shell

**Files:**
- Create: `src/config/database.ts`

**Interfaces:**
- Consumes: `env` from Task 2.
- Produces: `AppDataSource: DataSource` — exported empty-entities `DataSource` that later tasks import and push their entity class into the `entities` array (this task defines the array as a named, exported, mutable-at-import-time list populated by direct imports added in later tasks — see Task 4 Step for how an entity is registered).

- [ ] **Step 1: Create `src/config/database.ts`**

```ts
import path from 'path';
import { DataSource } from 'typeorm';
import { env } from './env';

export const AppDataSource = new DataSource({
  type: 'postgres',
  url: env.DATABASE_URL,
  entities: [path.join(__dirname, '../modules/**/*.entity{.ts,.js}')],
  migrations: [path.join(__dirname, '../database/migrations/*{.ts,.js}')],
  logging: env.NODE_ENV === 'development',
  synchronize: false,
});
```

Using a glob (`**/*.entity.ts`) instead of an explicit array means each later task that adds `<name>.entity.ts` under `src/modules/` is picked up automatically — no edits to this file required in later tasks.

- [ ] **Step 2: Verify it typechecks (no DB connection yet, just static shape)**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 3: Commit**

```bash
git add src/config/database.ts
git commit -m "feat: add TypeORM DataSource config"
```

---

## Task 4: Organization Module

**Files:**
- Create: `src/modules/organization/organization.entity.ts`
- Create: `src/modules/organization/organization.service.ts`
- Test: `src/modules/organization/organization.service.test.ts`
- Create: `src/database/migrations/<timestamp>-CreateOrganization.ts` (generated, see Step 5)

**Interfaces:**
- Consumes: `BaseEntity` from Task 2, `AppDataSource` from Task 3.
- Produces: `class Organization extends BaseEntity { name: string; domain: string; staff: Staff[] }` (the `staff` relation is declared here as `Staff[]` via a string-based `() => require('@modules/staff/staff.entity').Staff` is avoided — instead Task 5 adds the reverse `@ManyToOne` on `Staff`, and this file's `@OneToMany` is added in Task 5's steps to avoid a circular import at file-creation time). `OrganizationService` with `create(data)`, `findById(id)`, `findByDomain(domain)`, `findOrCreateByDomain(domain, name)`.

- [ ] **Step 1: Create `src/modules/organization/organization.entity.ts` (no relation yet — added in Task 5)**

```ts
import { Entity, Column } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';

@Entity('organizations')
export class Organization extends BaseEntity {
  @Column()
  name: string;

  @Column({ unique: true })
  domain: string;
}
```

- [ ] **Step 2: Write the failing test for `OrganizationService`**

Create `src/modules/organization/organization.service.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { OrganizationService } from './organization.service';

describe('OrganizationService.findOrCreateByDomain', () => {
  it('is a function that returns a promise', () => {
    const service = new OrganizationService();
    const result = service.findOrCreateByDomain('example.com', 'Example Co');
    expect(result).toBeInstanceOf(Promise);
  });
});
```

This is a shallow smoke test — full DB-backed integration testing happens in the manual smoke test (Task 13), since spinning up a throwaway Postgres per unit test is out of scope for today's timeline.

- [ ] **Step 3: Run test to verify it fails**

Run: `npx jest src/modules/organization/organization.service.test.ts`
Expected: FAIL — `Cannot find module './organization.service'`

- [ ] **Step 4: Create `src/modules/organization/organization.service.ts`**

```ts
import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { Organization } from './organization.entity';

@Service()
export class OrganizationService {
  private get repo(): Repository<Organization> {
    return AppDataSource.getRepository(Organization);
  }

  create(data: Partial<Organization>): Promise<Organization> {
    return this.repo.save(this.repo.create(data));
  }

  findById(id: string): Promise<Organization | null> {
    return this.repo.findOne({ where: { id } });
  }

  findByDomain(domain: string): Promise<Organization | null> {
    return this.repo.findOne({ where: { domain } });
  }

  async findOrCreateByDomain(domain: string, name: string): Promise<Organization> {
    const existing = await this.findByDomain(domain);
    if (existing) return existing;
    return this.create({ domain, name });
  }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npx jest src/modules/organization/organization.service.test.ts`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/modules/organization/organization.entity.ts src/modules/organization/organization.service.ts \
  src/modules/organization/organization.service.test.ts
git commit -m "feat: add Organization entity and service"
```

(The migration for this table is generated in Task 5 Step, alongside `Staff`, once both entities and their relation exist — TypeORM's `migration:generate` diffs the whole schema at once, so generating per-entity here would produce a migration that immediately becomes stale.)

---

## Task 5: Staff Module

**Files:**
- Modify: `src/modules/organization/organization.entity.ts` (add `staff` relation)
- Create: `src/modules/staff/staff.entity.ts`
- Create: `src/modules/staff/staff.service.ts`
- Test: `src/modules/staff/staff.service.test.ts`
- Create: `src/database/migrations/<timestamp>-CreateOrganizationAndStaff.ts` (generated)

**Interfaces:**
- Consumes: `Organization` (Task 4), `BaseEntity`, `encrypt`/`decrypt` (Task 2).
- Produces: `class Staff extends BaseEntity { organization: Organization; email: string; refreshTokenEnc: string | null; accessToken: string | null; tokenExpiresAt: Date | null; accountType: string; connected: boolean }`. `StaffService` with `upsertByEmail(email, orgId, accountType)`, `findById(id)`, `findByEmail(email)`, `listAll()`, `markDisconnected(id)`, `updateTokens(id, { refreshToken, accessToken, expiresAt })` (encrypts `refreshToken` internally via `encrypt()` before saving), `getDecryptedRefreshToken(staff: Staff)`.

- [ ] **Step 1: Create `src/modules/staff/staff.entity.ts`**

```ts
import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Organization } from '@modules/organization/organization.entity';

@Entity('staff')
export class Staff extends BaseEntity {
  @ManyToOne(() => Organization, { nullable: false })
  @JoinColumn({ name: 'organization_id' })
  organization: Organization;

  @Column({ unique: true })
  email: string;

  @Column({ type: 'text', nullable: true })
  refreshTokenEnc: string | null;

  @Column({ type: 'text', nullable: true })
  accessToken: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  tokenExpiresAt: Date | null;

  @Column({ default: 'enterprise' })
  accountType: string; // 'enterprise' | 'small_business' | 'personal'

  @Column({ default: false })
  connected: boolean;
}
```

- [ ] **Step 2: Add the reverse relation to `Organization`**

Modify `src/modules/organization/organization.entity.ts`:

```ts
import { Entity, Column, OneToMany } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Staff } from '@modules/staff/staff.entity';

@Entity('organizations')
export class Organization extends BaseEntity {
  @Column()
  name: string;

  @Column({ unique: true })
  domain: string;

  @OneToMany(() => Staff, (staff) => staff.organization)
  staff: Staff[];
}
```

- [ ] **Step 3: Write the failing test for `StaffService`**

Create `src/modules/staff/staff.service.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { StaffService } from './staff.service';
import { encrypt } from '@shared/crypto';
import { Staff } from './staff.entity';

describe('StaffService.getDecryptedRefreshToken', () => {
  it('decrypts a staff record refreshTokenEnc field', () => {
    const service = new StaffService();
    const encrypted = encrypt('my-refresh-token');
    const staff = { refreshTokenEnc: encrypted } as Staff;
    expect(service.getDecryptedRefreshToken(staff)).toBe('my-refresh-token');
  });

  it('returns null when refreshTokenEnc is null', () => {
    const service = new StaffService();
    const staff = { refreshTokenEnc: null } as Staff;
    expect(service.getDecryptedRefreshToken(staff)).toBeNull();
  });
});
```

- [ ] **Step 4: Run test to verify it fails**

Run: `npx jest src/modules/staff/staff.service.test.ts`
Expected: FAIL — `Cannot find module './staff.service'`

- [ ] **Step 5: Create `src/modules/staff/staff.service.ts`**

```ts
import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { encrypt, decrypt } from '@shared/crypto';
import { Staff } from './staff.entity';

interface TokenUpdate {
  refreshToken: string;
  accessToken: string;
  expiresAt: Date;
}

@Service()
export class StaffService {
  private get repo(): Repository<Staff> {
    return AppDataSource.getRepository(Staff);
  }

  findById(id: string): Promise<Staff | null> {
    return this.repo.findOne({ where: { id }, relations: ['organization'] });
  }

  findByEmail(email: string): Promise<Staff | null> {
    return this.repo.findOne({ where: { email }, relations: ['organization'] });
  }

  listAll(): Promise<Staff[]> {
    return this.repo.find({ relations: ['organization'], order: { createdAt: 'DESC' } });
  }

  async upsertByEmail(
    email: string,
    organizationId: string,
    accountType: string,
  ): Promise<Staff> {
    const existing = await this.repo.findOne({ where: { email } });
    if (existing) return existing;
    return this.repo.save(
      this.repo.create({
        email,
        organization: { id: organizationId } as Staff['organization'],
        accountType,
        connected: false,
      }),
    );
  }

  async updateTokens(id: string, tokens: TokenUpdate): Promise<void> {
    await this.repo.update(
      { id },
      {
        refreshTokenEnc: encrypt(tokens.refreshToken),
        accessToken: tokens.accessToken,
        tokenExpiresAt: tokens.expiresAt,
        connected: true,
      },
    );
  }

  async markDisconnected(id: string): Promise<void> {
    await this.repo.update({ id }, { connected: false });
  }

  getDecryptedRefreshToken(staff: Staff): string | null {
    if (!staff.refreshTokenEnc) return null;
    return decrypt(staff.refreshTokenEnc);
  }
}
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx jest src/modules/staff/staff.service.test.ts`
Expected: PASS

- [ ] **Step 7: Generate the migration for `organizations` and `staff` tables**

Prerequisite: a real Postgres database named in `DATABASE_URL` must exist and be reachable (create it now: `createdb teams_scheduler` or via pgAdmin).

Run: `npm run migration:generate -- src/database/migrations/CreateOrganizationAndStaff`
Expected: creates `src/database/migrations/<timestamp>-CreateOrganizationAndStaff.ts` containing `CREATE TABLE "organizations"...` and `CREATE TABLE "staff"...`.

- [ ] **Step 8: Run the migration**

Run: `npm run migration:run`
Expected: output ends with `Migration CreateOrganizationAndStaff... has been executed successfully.`

- [ ] **Step 9: Commit**

```bash
git add src/modules/organization/organization.entity.ts src/modules/staff/staff.entity.ts \
  src/modules/staff/staff.service.ts src/modules/staff/staff.service.test.ts \
  src/database/migrations/
git commit -m "feat: add Staff entity/service, wire Organization<->Staff relation"
```

---

## Task 6: Graph Client Wrapper

**Files:**
- Create: `src/shared/graph-client.ts`
- Create: `src/shared/errors/GraphAuthError.ts`
- Create: `src/shared/errors/GraphNotFoundError.ts`
- Test: `src/shared/graph-client.test.ts`

**Interfaces:**
- Consumes: `AppError` hierarchy (Task 2).
- Produces: `createGraphClient(accessToken: string): AxiosInstance` — base URL `https://graph.microsoft.com/v1.0`, sets `Authorization: Bearer {token}`, `Content-Type: application/json`, `Prefer: outlook.timezone="UTC"`; response interceptor maps 429 (respects `Retry-After`, retries once) → returns retried response or throws; 401 → throws `GraphAuthError`; 404 → throws `GraphNotFoundError`. `GraphAuthError` and `GraphNotFoundError` extend `AppError`.

- [ ] **Step 1: Create `src/shared/errors/GraphAuthError.ts`**

```ts
import { AppError } from './AppError';

export class GraphAuthError extends AppError {
  constructor(message = 'Microsoft Graph rejected the access token') {
    super(message, 401, 'GRAPH_AUTH_ERROR');
  }
}
```

- [ ] **Step 2: Create `src/shared/errors/GraphNotFoundError.ts`**

```ts
import { AppError } from './AppError';

export class GraphNotFoundError extends AppError {
  constructor(message = 'Graph resource not found (may already be deleted)') {
    super(message, 404, 'GRAPH_NOT_FOUND');
  }
}
```

- [ ] **Step 3: Write the failing test for the Graph client's error mapping**

Create `src/shared/graph-client.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { createGraphClient } from './graph-client';
import { GraphAuthError } from './errors/GraphAuthError';
import { GraphNotFoundError } from './errors/GraphNotFoundError';

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
});
```

- [ ] **Step 4: Run test to verify it fails**

Run: `npx jest src/shared/graph-client.test.ts`
Expected: FAIL — `Cannot find module './graph-client'`

- [ ] **Step 5: Create `src/shared/graph-client.ts`**

```ts
import axios, { AxiosInstance, AxiosError } from 'axios';
import { GraphAuthError } from './errors/GraphAuthError';
import { GraphNotFoundError } from './errors/GraphNotFoundError';
import { logger } from './utils/logger';

export function createGraphClient(accessToken: string): AxiosInstance {
  const client = axios.create({
    baseURL: 'https://graph.microsoft.com/v1.0',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      Prefer: 'outlook.timezone="UTC"',
    },
  });

  client.interceptors.response.use(
    (response) => response,
    async (error: AxiosError) => {
      const status = error.response?.status;

      if (status === 429) {
        const retryAfterHeader = error.response?.headers['retry-after'];
        const retryAfterSeconds = retryAfterHeader ? parseInt(retryAfterHeader, 10) : 1;
        logger.warn('Graph 429 — retrying once', { retryAfterSeconds });
        await new Promise((resolve) => setTimeout(resolve, retryAfterSeconds * 1000));
        if (error.config) {
          return client.request(error.config);
        }
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
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx jest src/shared/graph-client.test.ts`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/shared/graph-client.ts src/shared/graph-client.test.ts \
  src/shared/errors/GraphAuthError.ts src/shared/errors/GraphNotFoundError.ts
git commit -m "feat: add Microsoft Graph client wrapper with 429/401/404 handling"
```

---

## Task 7: MSAL Auth Module (login, callback, token refresh)

**Files:**
- Create: `src/modules/auth/token.service.ts`
- Create: `src/modules/auth/auth.controller.ts`
- Create: `src/modules/auth/auth.routes.ts`
- Test: `src/modules/auth/token.service.test.ts`

**Interfaces:**
- Consumes: `StaffService`, `OrganizationService` (Tasks 4-5), `env` (Task 2).
- Produces: `TokenService.getAuthCodeUrl(): Promise<string>`, `TokenService.exchangeCodeForTokens(code: string): Promise<{ accessToken, refreshToken, expiresOn, account }>`, `TokenService.refreshAccessToken(staffId: string): Promise<string>` (returns a valid access token, refreshing via Graph token endpoint if `tokenExpiresAt` is within 5 minutes, calling `StaffService.markDisconnected` on `invalid_grant`), `authRouter` mounted with `GET /login`, `GET /callback`.

- [ ] **Step 1: Write the failing test for duration/expiry math in `TokenService`**

Create `src/modules/auth/token.service.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { TokenService } from './token.service';

describe('TokenService.isNearExpiry', () => {
  const service = new TokenService();

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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest src/modules/auth/token.service.test.ts`
Expected: FAIL — `Cannot find module './token.service'`

- [ ] **Step 3: Create `src/modules/auth/token.service.ts`**

```ts
import { Service } from 'typedi';
import axios from 'axios';
import * as msal from '@azure/msal-node';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';
import { StaffService } from '@modules/staff/staff.service';
import { UnauthorizedError } from '@shared/errors/UnauthorizedError';

const SCOPES = ['offline_access', 'Calendars.ReadWrite', 'OnlineMeetings.ReadWrite', 'User.Read'];
const NEAR_EXPIRY_MS = 5 * 60 * 1000;

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
    const cache = this.msalClient.getTokenCache().serialize();
    const parsed = JSON.parse(cache);
    const refreshTokenEntry = Object.values(parsed.RefreshToken ?? {})[0] as
      | { secret: string }
      | undefined;

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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest src/modules/auth/token.service.test.ts`
Expected: PASS

- [ ] **Step 5: Create `src/modules/auth/auth.controller.ts`**

```ts
import { Service } from 'typedi';
import { Request, Response } from 'express';
import { TokenService } from './token.service';
import { StaffService } from '@modules/staff/staff.service';
import { OrganizationService } from '@modules/organization/organization.service';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';

@Service()
export class AuthController {
  constructor(
    private readonly tokenService: TokenService,
    private readonly staffService: StaffService,
    private readonly organizationService: OrganizationService,
  ) {}

  login = async (_req: Request, res: Response): Promise<void> => {
    const url = await this.tokenService.getAuthCodeUrl();
    res.redirect(url);
  };

  callback = async (req: Request, res: Response): Promise<void> => {
    const code = req.query.code as string | undefined;
    if (!code) {
      res.redirect(`${env.FRONTEND_URL}/admin?connect=error`);
      return;
    }

    const { accessToken, refreshToken, expiresOn, email } =
      await this.tokenService.exchangeCodeForTokens(code);

    const domain = email.split('@')[1] ?? 'unknown.local';
    const accountType = domain === 'outlook.com' || domain === 'hotmail.com' ? 'personal' : 'enterprise';
    const organization = await this.organizationService.findOrCreateByDomain(domain, domain);

    const staff = await this.staffService.upsertByEmail(email, organization.id, accountType);
    await this.staffService.updateTokens(staff.id, { accessToken, refreshToken, expiresAt: expiresOn });

    logger.info('Staff connected', { email });
    res.redirect(`${env.FRONTEND_URL}/admin?connect=success`);
  };
}
```

- [ ] **Step 6: Create `src/modules/auth/auth.routes.ts`**

```ts
import { Router } from 'express';
import { Container } from 'typedi';
import { AuthController } from './auth.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';

const router = Router();
const controller = Container.get(AuthController);

router.get('/staff/login', asyncHandler(controller.login));
router.get('/staff/callback', asyncHandler(controller.callback));

export { router as authRouter };
```

- [ ] **Step 7: Commit**

```bash
git add src/modules/auth/token.service.ts src/modules/auth/auth.controller.ts \
  src/modules/auth/auth.routes.ts src/modules/auth/token.service.test.ts
git commit -m "feat: add MSAL auth module (login, callback, token refresh)"
```

---

## Task 8: Booking Entity + Migration

**Files:**
- Create: `src/modules/booking/booking.entity.ts`
- Create: `src/database/migrations/<timestamp>-CreateBooking.ts` (generated)

**Interfaces:**
- Consumes: `BaseEntity` (Task 2), `Staff` (Task 5).
- Produces: `class Booking extends BaseEntity { visitorEmail: string; visitorName: string; requestedStart: Date; requestedEnd: Date; status: string; staff: Staff | null; msEventId: string | null; joinUrl: string | null; cancelledAt: Date | null; cancelledBy: string | null }`.

- [ ] **Step 1: Create `src/modules/booking/booking.entity.ts`**

```ts
import { Entity, Column, ManyToOne, JoinColumn } from 'typeorm';
import { BaseEntity } from '@shared/base.entity';
import { Staff } from '@modules/staff/staff.entity';

export type BookingStatus = 'Requested' | 'Scheduled' | 'Rescheduled' | 'Swapped' | 'Cancelled';

@Entity('bookings')
export class Booking extends BaseEntity {
  @Column()
  visitorEmail: string;

  @Column()
  visitorName: string;

  @Column({ type: 'timestamptz' })
  requestedStart: Date;

  @Column({ type: 'timestamptz' })
  requestedEnd: Date;

  @Column({ default: 'Requested' })
  status: BookingStatus;

  @ManyToOne(() => Staff, { nullable: true })
  @JoinColumn({ name: 'staff_id' })
  staff: Staff | null;

  @Column({ type: 'text', nullable: true })
  msEventId: string | null;

  @Column({ type: 'text', nullable: true })
  joinUrl: string | null;

  @Column({ type: 'timestamptz', nullable: true })
  cancelledAt: Date | null;

  @Column({ nullable: true })
  cancelledBy: string | null;
}
```

- [ ] **Step 2: Generate and run the migration**

Run: `npm run migration:generate -- src/database/migrations/CreateBooking`
Expected: creates a migration containing `CREATE TABLE "bookings"...`.

Run: `npm run migration:run`
Expected: `Migration CreateBooking... has been executed successfully.`

- [ ] **Step 3: Verify typecheck passes**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 4: Commit**

```bash
git add src/modules/booking/booking.entity.ts src/database/migrations/
git commit -m "feat: add Booking entity and migration"
```

---

## Task 9: Booking Service (availability, create, reschedule, cancel, delete, swap)

**Files:**
- Create: `src/modules/booking/booking.service.ts`
- Test: `src/modules/booking/booking.service.test.ts`

**Interfaces:**
- Consumes: `Booking` (Task 8), `Staff`/`StaffService` (Task 5), `TokenService.getValidAccessToken` (Task 7), `createGraphClient` (Task 6), `GraphNotFoundError` (Task 6), `ConflictError`/`BadRequestError` (Task 2).
- Produces: `BookingService` with:
  - `createVisitorRequest(data: { visitorEmail, visitorName, requestedStart, requestedEnd }): Promise<Booking>`
  - `checkAvailability(staffId: string, start: Date, end: Date): Promise<{ availabilityView: string; scheduleItems: unknown[] }>`
  - `createBooking(bookingId: string, staffId: string, subject: string): Promise<Booking>`
  - `reschedule(bookingId: string, newStart: Date, newEnd: Date): Promise<Booking>`
  - `cancel(bookingId: string, comment: string, cancelledBy: string): Promise<Booking>`
  - `deleteHard(bookingId: string): Promise<void>`
  - `swapHost(bookingId: string, newStaffId: string): Promise<Booking>` (throws `BadRequestError` if `newStaffId`'s organization differs from the current staff's organization)
  - `listAll(): Promise<Booking[]>`

- [ ] **Step 1: Write the failing test for the same-org swap-host guard**

Create `src/modules/booking/booking.service.test.ts`:

```ts
process.env.TOKEN_ENC_KEY = 'a'.repeat(64);
process.env.DATABASE_URL = 'postgres://postgres:postgres@localhost:5432/teams_scheduler_test';
process.env.MS_CLIENT_ID = 'test-client-id';
process.env.MS_CLIENT_SECRET = 'test-client-secret';

import { BookingService } from './booking.service';

describe('BookingService.assertSameOrganization', () => {
  const service = new BookingService(
    // @ts-expect-error — only staffService is used by this method
    { findById: undefined },
    undefined,
    undefined,
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest src/modules/booking/booking.service.test.ts`
Expected: FAIL — `Cannot find module './booking.service'`

- [ ] **Step 3: Create `src/modules/booking/booking.service.ts`**

```ts
import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { createGraphClient } from '@shared/graph-client';
import { GraphNotFoundError } from '@shared/errors/GraphNotFoundError';
import { NotFoundError } from '@shared/errors/NotFoundError';
import { BadRequestError } from '@shared/errors/BadRequestError';
import { logger } from '@shared/utils/logger';
import { Booking } from './booking.entity';
import { StaffService } from '@modules/staff/staff.service';
import { TokenService } from '@modules/auth/token.service';

interface VisitorRequestInput {
  visitorEmail: string;
  visitorName: string;
  requestedStart: Date;
  requestedEnd: Date;
}

@Service()
export class BookingService {
  constructor(
    private readonly staffService: StaffService,
    private readonly tokenService: TokenService,
  ) {}

  private get repo(): Repository<Booking> {
    return AppDataSource.getRepository(Booking);
  }

  listAll(): Promise<Booking[]> {
    return this.repo.find({ relations: ['staff', 'staff.organization'], order: { createdAt: 'DESC' } });
  }

  private async findOrThrow(id: string): Promise<Booking> {
    const booking = await this.repo.findOne({
      where: { id },
      relations: ['staff', 'staff.organization'],
    });
    if (!booking) throw new NotFoundError('Booking not found');
    return booking;
  }

  createVisitorRequest(data: VisitorRequestInput): Promise<Booking> {
    return this.repo.save(this.repo.create({ ...data, status: 'Requested' }));
  }

  // spec §4 Step 1
  async checkAvailability(
    staffId: string,
    start: Date,
    end: Date,
  ): Promise<{ availabilityView: string; scheduleItems: unknown[] }> {
    const staff = await this.staffService.findById(staffId);
    if (!staff) throw new NotFoundError('Staff not found');

    const accessToken = await this.tokenService.getValidAccessToken(staffId);
    const graph = createGraphClient(accessToken);

    const response = await graph.post('/me/calendar/getSchedule', {
      schedules: [staff.email],
      startTime: { dateTime: start.toISOString().replace('Z', ''), timeZone: 'UTC' },
      endTime: { dateTime: end.toISOString().replace('Z', ''), timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });

    const result = response.data.value[0];
    return { availabilityView: result.availabilityView, scheduleItems: result.scheduleItems };
  }

  // spec §4 Step 2
  async createBooking(bookingId: string, staffId: string, subject: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    const staff = await this.staffService.findById(staffId);
    if (!staff) throw new NotFoundError('Staff not found');

    const accessToken = await this.tokenService.getValidAccessToken(staffId);
    const graph = createGraphClient(accessToken);

    const response = await graph.post('/me/events', {
      subject,
      body: { contentType: 'HTML', content: 'Discussion about organization inquiry.' },
      start: { dateTime: booking.requestedStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
      end: { dateTime: booking.requestedEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      attendees: [
        { emailAddress: { address: booking.visitorEmail, name: booking.visitorName }, type: 'required' },
      ],
    });

    booking.staff = staff;
    booking.msEventId = response.data.id;
    booking.joinUrl = response.data.onlineMeeting?.joinUrl ?? null;
    booking.status = 'Scheduled';
    return this.repo.save(booking);
  }

  // spec §4 Step 3
  async reschedule(bookingId: string, newStart: Date, newEnd: Date): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to reschedule');
    }

    const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const graph = createGraphClient(accessToken);

    try {
      await graph.patch(`/me/events/${booking.msEventId}`, {
        start: { dateTime: newStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
        end: { dateTime: newEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
      });
    } catch (err) {
      if (err instanceof GraphNotFoundError) {
        logger.warn('Event already deleted in Outlook — cannot reschedule', { bookingId });
      }
      throw err;
    }

    booking.requestedStart = newStart;
    booking.requestedEnd = newEnd;
    booking.status = 'Rescheduled';
    return this.repo.save(booking);
  }

  // spec §4 Step 4 (cancel — notifies visitor)
  async cancel(bookingId: string, comment: string, cancelledBy: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to cancel');
    }

    const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const graph = createGraphClient(accessToken);

    try {
      await graph.post(`/me/events/${booking.msEventId}/cancel`, { comment });
    } catch (err) {
      if (!(err instanceof GraphNotFoundError)) throw err;
      logger.warn('Event already deleted in Outlook — treating as cancelled', { bookingId });
    }

    booking.status = 'Cancelled';
    booking.cancelledAt = new Date();
    booking.cancelledBy = cancelledBy;
    return this.repo.save(booking);
  }

  // spec §4 Step 4 (delete — no notification)
  async deleteHard(bookingId: string): Promise<void> {
    const booking = await this.findOrThrow(bookingId);
    if (booking.staff && booking.msEventId) {
      const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
      const graph = createGraphClient(accessToken);
      try {
        await graph.delete(`/me/events/${booking.msEventId}`);
      } catch (err) {
        if (!(err instanceof GraphNotFoundError)) throw err;
      }
    }
    booking.status = 'Cancelled';
    booking.cancelledAt = new Date();
    booking.cancelledBy = 'system-delete';
    await this.repo.save(booking);
  }

  assertSameOrganization(orgIdA: string, orgIdB: string): void {
    if (orgIdA !== orgIdB) {
      throw new BadRequestError('Swap-host is only allowed between staff in the same organization');
    }
  }

  // spec §4 Step 5 — cancel with Person A's token, rebook with Person B's token
  async swapHost(bookingId: string, newStaffId: string): Promise<Booking> {
    const booking = await this.findOrThrow(bookingId);
    if (!booking.staff || !booking.msEventId) {
      throw new BadRequestError('Booking has no scheduled event to swap');
    }

    const newStaff = await this.staffService.findById(newStaffId);
    if (!newStaff) throw new NotFoundError('New staff not found');

    this.assertSameOrganization(booking.staff.organization.id, newStaff.organization.id);

    const oldAccessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
    const oldGraph = createGraphClient(oldAccessToken);
    try {
      await oldGraph.post(`/me/events/${booking.msEventId}/cancel`, {
        comment: 'Apologies, this meeting has been reassigned to a different host.',
      });
    } catch (err) {
      if (!(err instanceof GraphNotFoundError)) throw err;
      logger.warn('Old event already deleted in Outlook — proceeding with rebook', { bookingId });
    }

    const newAccessToken = await this.tokenService.getValidAccessToken(newStaff.id);
    const newGraph = createGraphClient(newAccessToken);
    const response = await newGraph.post('/me/events', {
      subject: 'SaaS Intro Meeting',
      body: { contentType: 'HTML', content: 'Discussion about organization inquiry.' },
      start: { dateTime: booking.requestedStart.toISOString().replace('Z', ''), timeZone: 'UTC' },
      end: { dateTime: booking.requestedEnd.toISOString().replace('Z', ''), timeZone: 'UTC' },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      attendees: [
        { emailAddress: { address: booking.visitorEmail, name: booking.visitorName }, type: 'required' },
      ],
    });

    booking.staff = newStaff;
    booking.msEventId = response.data.id;
    booking.joinUrl = response.data.onlineMeeting?.joinUrl ?? null;
    booking.status = 'Swapped';
    return this.repo.save(booking);
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest src/modules/booking/booking.service.test.ts`
Expected: PASS

- [ ] **Step 5: Run full test suite**

Run: `npx jest`
Expected: PASS (all tests across all modules)

- [ ] **Step 6: Commit**

```bash
git add src/modules/booking/booking.service.ts src/modules/booking/booking.service.test.ts
git commit -m "feat: add BookingService (availability, create, reschedule, cancel, delete, swap)"
```

---

## Task 10: Booking Controller, DTOs, Routes

**Files:**
- Create: `src/modules/booking/booking.dto.ts`
- Create: `src/shared/middleware/validate.ts`
- Create: `src/modules/booking/booking.controller.ts`
- Create: `src/modules/booking/booking.routes.ts`
- Create: `src/modules/staff/staff.controller.ts`
- Create: `src/modules/staff/staff.routes.ts`

**Interfaces:**
- Consumes: `BookingService` (Task 9), `StaffService` (Task 5), `sendSuccess`/`asyncHandler` (Task 2).
- Produces: `validateDto(DtoClass, source?)` middleware; routers `bookingRouter`, `staffRouter` mounted in Task 11.

- [ ] **Step 1: Create `src/shared/middleware/validate.ts`**

```ts
import { Request, Response, NextFunction } from 'express';
import { plainToInstance } from 'class-transformer';
import { validate } from 'class-validator';
import { BadRequestError } from '@shared/errors/BadRequestError';

export function validateDto(
  DtoClass: new () => object,
  source: 'body' | 'query' = 'body',
) {
  return async (req: Request, _res: Response, next: NextFunction): Promise<void> => {
    const instance = plainToInstance(DtoClass, req[source]);
    const errors = await validate(instance);
    if (errors.length > 0) {
      const message = errors
        .map((e) => Object.values(e.constraints ?? {}).join(', '))
        .join('; ');
      next(new BadRequestError(message));
      return;
    }
    req[source] = instance as never;
    next();
  };
}
```

- [ ] **Step 2: Create `src/modules/booking/booking.dto.ts`**

```ts
import { IsString, IsEmail, IsISO8601, IsUUID, IsOptional } from 'class-validator';

export class VisitorRequestDto {
  @IsEmail()
  visitorEmail: string;

  @IsString()
  visitorName: string;

  @IsISO8601()
  requestedStart: string;

  @IsISO8601()
  requestedEnd: string;
}

export class CreateBookingDto {
  @IsUUID()
  bookingId: string;

  @IsUUID()
  staffId: string;

  @IsString()
  subject: string;
}

export class RescheduleDto {
  @IsISO8601()
  newStart: string;

  @IsISO8601()
  newEnd: string;
}

export class CancelDto {
  @IsString()
  comment: string;

  @IsString()
  cancelledBy: string;
}

export class SwapHostDto {
  @IsUUID()
  newStaffId: string;
}

export class AvailabilityQueryDto {
  @IsISO8601()
  start: string;

  @IsISO8601()
  end: string;

  @IsOptional()
  @IsUUID()
  staffId?: string;
}
```

- [ ] **Step 3: Create `src/modules/booking/booking.controller.ts`**

```ts
import { Service } from 'typedi';
import { Request, Response } from 'express';
import { BookingService } from './booking.service';
import { sendSuccess } from '@shared/utils/apiResponse';
import {
  VisitorRequestDto,
  CreateBookingDto,
  RescheduleDto,
  CancelDto,
  SwapHostDto,
} from './booking.dto';

@Service()
export class BookingController {
  constructor(private readonly bookingService: BookingService) {}

  list = async (_req: Request, res: Response): Promise<void> => {
    const bookings = await this.bookingService.listAll();
    sendSuccess(res, bookings);
  };

  visitorRequest = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as VisitorRequestDto;
    const booking = await this.bookingService.createVisitorRequest({
      visitorEmail: dto.visitorEmail,
      visitorName: dto.visitorName,
      requestedStart: new Date(dto.requestedStart),
      requestedEnd: new Date(dto.requestedEnd),
    });
    sendSuccess(res, booking, 'Booking request received', 201);
  };

  create = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as CreateBookingDto;
    const booking = await this.bookingService.createBooking(dto.bookingId, dto.staffId, dto.subject);
    sendSuccess(res, booking, 'Meeting scheduled');
  };

  reschedule = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as RescheduleDto;
    const booking = await this.bookingService.reschedule(
      req.params.id,
      new Date(dto.newStart),
      new Date(dto.newEnd),
    );
    sendSuccess(res, booking, 'Booking rescheduled');
  };

  cancel = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as CancelDto;
    const booking = await this.bookingService.cancel(req.params.id, dto.comment, dto.cancelledBy);
    sendSuccess(res, booking, 'Booking cancelled');
  };

  deleteHard = async (req: Request, res: Response): Promise<void> => {
    await this.bookingService.deleteHard(req.params.id);
    sendSuccess(res, {}, 'Booking deleted', 204);
  };

  swap = async (req: Request, res: Response): Promise<void> => {
    const dto = req.body as unknown as SwapHostDto;
    const booking = await this.bookingService.swapHost(req.params.id, dto.newStaffId);
    sendSuccess(res, booking, 'Host swapped');
  };
}
```

- [ ] **Step 4: Create `src/modules/booking/booking.routes.ts`**

```ts
import { Router } from 'express';
import { Container } from 'typedi';
import { BookingController } from './booking.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';
import { validateDto } from '@shared/middleware/validate';
import { VisitorRequestDto, CreateBookingDto, RescheduleDto, CancelDto, SwapHostDto } from './booking.dto';

const router = Router();
const controller = Container.get(BookingController);

router.get('/', asyncHandler(controller.list));
router.post('/visitor-request', validateDto(VisitorRequestDto), asyncHandler(controller.visitorRequest));
router.post('/', validateDto(CreateBookingDto), asyncHandler(controller.create));
router.patch('/:id', validateDto(RescheduleDto), asyncHandler(controller.reschedule));
router.post('/:id/cancel', validateDto(CancelDto), asyncHandler(controller.cancel));
router.delete('/:id', asyncHandler(controller.deleteHard));
router.post('/:id/swap', validateDto(SwapHostDto), asyncHandler(controller.swap));

export { router as bookingRouter };
```

- [ ] **Step 5: Create `src/modules/staff/staff.controller.ts`**

```ts
import { Service } from 'typedi';
import { Request, Response } from 'express';
import { StaffService } from './staff.service';
import { BookingService } from '@modules/booking/booking.service';
import { sendSuccess } from '@shared/utils/apiResponse';
import { NotFoundError } from '@shared/errors/NotFoundError';

@Service()
export class StaffController {
  constructor(
    private readonly staffService: StaffService,
    private readonly bookingService: BookingService,
  ) {}

  list = async (_req: Request, res: Response): Promise<void> => {
    const staff = await this.staffService.listAll();
    sendSuccess(res, staff);
  };

  availability = async (req: Request, res: Response): Promise<void> => {
    const staff = await this.staffService.findById(req.params.id);
    if (!staff) throw new NotFoundError('Staff not found');

    const start = new Date(req.query.start as string);
    const end = new Date(req.query.end as string);
    const result = await this.bookingService.checkAvailability(staff.id, start, end);
    sendSuccess(res, result);
  };
}
```

- [ ] **Step 6: Create `src/modules/staff/staff.routes.ts`**

```ts
import { Router } from 'express';
import { Container } from 'typedi';
import { StaffController } from './staff.controller';
import { asyncHandler } from '@shared/utils/asyncHandler';

const router = Router();
const controller = Container.get(StaffController);

router.get('/', asyncHandler(controller.list));
router.get('/:id/availability', asyncHandler(controller.availability));

export { router as staffRouter };
```

- [ ] **Step 7: Verify typecheck passes**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 8: Commit**

```bash
git add src/shared/middleware/validate.ts src/modules/booking/booking.dto.ts \
  src/modules/booking/booking.controller.ts src/modules/booking/booking.routes.ts \
  src/modules/staff/staff.controller.ts src/modules/staff/staff.routes.ts
git commit -m "feat: add booking and staff controllers, DTOs, routes"
```

---

## Task 11: App Wiring, Server Entrypoint, Error Handler, Health Route

**Files:**
- Create: `src/shared/middleware/errorHandler.ts`
- Create: `src/health/health.routes.ts`
- Create: `src/app.ts`
- Create: `src/server.ts`

**Interfaces:**
- Consumes: `authRouter` (Task 7), `bookingRouter`, `staffRouter` (Task 10), `AppDataSource` (Task 3), `env`, `logger` (Task 2).
- Produces: `app: Express`, running server on `env.PORT` after `AppDataSource.initialize()`.

- [ ] **Step 1: Create `src/shared/middleware/errorHandler.ts`**

```ts
import { Request, Response, NextFunction } from 'express';
import { QueryFailedError } from 'typeorm';
import { AppError } from '@shared/errors/AppError';
import { logger } from '@shared/utils/logger';
import { sendError } from '@shared/utils/apiResponse';

export const errorHandler = (
  err: Error,
  req: Request,
  res: Response,
  _next: NextFunction,
): void => {
  if (err instanceof AppError) {
    sendError(res, err.message, err.statusCode, err.errorCode);
    return;
  }

  if (err instanceof QueryFailedError) {
    const dbErr = err as QueryFailedError & { code?: string };
    if (dbErr.code === '23505') {
      sendError(res, 'A record with the given data already exists', 409, 'CONFLICT');
      return;
    }
  }

  logger.error('Unhandled error', { path: req.path, message: err.message, stack: err.stack });
  sendError(res, 'An unexpected error occurred', 500, 'INTERNAL_SERVER_ERROR');
};
```

- [ ] **Step 2: Create `src/health/health.routes.ts`**

```ts
import { Router } from 'express';

const router = Router();

router.get('/', (_req, res) => {
  res.json({ status: 'ok' });
});

export { router as healthRouter };
```

- [ ] **Step 3: Create `src/app.ts`**

```ts
import 'reflect-metadata';
import express from 'express';
import helmet from 'helmet';
import cors from 'cors';
import { env } from '@config/env';
import { errorHandler } from '@shared/middleware/errorHandler';
import { healthRouter } from './health/health.routes';
import { authRouter } from '@modules/auth/auth.routes';
import { bookingRouter } from '@modules/booking/booking.routes';
import { staffRouter } from '@modules/staff/staff.routes';

const app = express();

app.use(helmet());
app.use(cors({ origin: env.FRONTEND_URL, credentials: true }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use('/health', healthRouter);
app.use('/auth', authRouter);
app.use('/api/v1/bookings', bookingRouter);
app.use('/api/v1/staff', staffRouter);

app.use(errorHandler);

export { app };
```

- [ ] **Step 4: Create `src/server.ts`**

```ts
import 'reflect-metadata';
import { app } from './app';
import { AppDataSource } from '@config/database';
import { env } from '@config/env';
import { logger } from '@shared/utils/logger';

const startServer = async (): Promise<void> => {
  await AppDataSource.initialize();
  logger.info('Database connected');

  const server = app.listen(env.PORT, () => {
    logger.info(`Server running on port ${env.PORT}`, { env: env.NODE_ENV });
  });

  const shutdown = (signal: string): void => {
    logger.info(`${signal} received — shutting down`);
    server.close(async () => {
      await AppDataSource.destroy();
      process.exit(0);
    });
  };

  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
};

startServer().catch((err: Error) => {
  logger.error('Failed to start server', { message: err.message, stack: err.stack });
  process.exit(1);
});
```

- [ ] **Step 5: Verify the server boots**

Run: `npm run dev`
Expected: log lines `Database connected` then `Server running on port 5000`.

Run (in a second terminal): `curl http://localhost:5000/health`
Expected: `{"status":"ok"}`

- [ ] **Step 6: Commit**

```bash
git add src/shared/middleware/errorHandler.ts src/health/health.routes.ts src/app.ts src/server.ts
git commit -m "feat: wire Express app, server entrypoint, error handler, health route"
```

---

## Task 12: Frontend Scaffold (Vite + React)

**Files:**
- Create: `frontend/package.json`
- Create: `frontend/vite.config.ts`
- Create: `frontend/index.html`
- Create: `frontend/src/main.tsx`
- Create: `frontend/src/App.tsx`
- Create: `frontend/src/api.ts`
- Create: `frontend/src/pages/VisitorBookingPage.tsx`
- Create: `frontend/src/pages/AdminDashboardPage.tsx`

**Interfaces:**
- Consumes: backend REST API from Tasks 10-11 (`/api/v1/bookings`, `/api/v1/staff`, `/auth/staff/login`).
- Produces: a runnable Vite dev server on port 5173 with two routes: `/` (visitor form) and `/admin` (dashboard).

This task is intentionally kept light — functional, not polished, since the goal is demoing the workflow, not the UI.

- [ ] **Step 1: Create `frontend/package.json`**

```json
{
  "name": "teams-scheduler-frontend",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "dev": "vite",
    "build": "vite build"
  },
  "dependencies": {
    "react": "^18.3.1",
    "react-dom": "^18.3.1",
    "react-router-dom": "^6.26.1"
  },
  "devDependencies": {
    "@types/react": "^18.3.3",
    "@types/react-dom": "^18.3.0",
    "@vitejs/plugin-react": "^4.3.1",
    "typescript": "^5.5.4",
    "vite": "^5.4.2"
  }
}
```

- [ ] **Step 2: Create `frontend/vite.config.ts`**

```ts
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: { port: 5173 },
});
```

- [ ] **Step 3: Create `frontend/index.html`**

```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title>Teams Scheduler Demo</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

- [ ] **Step 4: Create `frontend/src/api.ts`**

```ts
const API_BASE = 'http://localhost:5000';

export async function apiPost(path: string, body: unknown) {
  const res = await fetch(`${API_BASE}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return res.json();
}

export async function apiPatch(path: string, body: unknown) {
  const res = await fetch(`${API_BASE}${path}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return res.json();
}

export async function apiGet(path: string) {
  const res = await fetch(`${API_BASE}${path}`);
  return res.json();
}

export function staffLoginUrl(): string {
  return `${API_BASE}/auth/staff/login`;
}
```

- [ ] **Step 5: Create `frontend/src/pages/VisitorBookingPage.tsx`**

```tsx
import { useState } from 'react';
import { apiPost } from '../api';

export function VisitorBookingPage() {
  const [visitorName, setVisitorName] = useState('');
  const [visitorEmail, setVisitorEmail] = useState('');
  const [requestedStart, setRequestedStart] = useState('');
  const [requestedEnd, setRequestedEnd] = useState('');
  const [status, setStatus] = useState<string | null>(null);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    const result = await apiPost('/api/v1/bookings/visitor-request', {
      visitorName,
      visitorEmail,
      requestedStart: new Date(requestedStart).toISOString(),
      requestedEnd: new Date(requestedEnd).toISOString(),
    });
    setStatus(result.success ? 'Request submitted — we will confirm shortly.' : result.message);
  };

  return (
    <div style={{ maxWidth: 480, margin: '40px auto', fontFamily: 'sans-serif' }}>
      <h1>Book a Meeting</h1>
      <form onSubmit={submit}>
        <div>
          <label>Name</label>
          <input value={visitorName} onChange={(e) => setVisitorName(e.target.value)} required />
        </div>
        <div>
          <label>Email</label>
          <input
            type="email"
            value={visitorEmail}
            onChange={(e) => setVisitorEmail(e.target.value)}
            required
          />
        </div>
        <div>
          <label>Preferred start</label>
          <input
            type="datetime-local"
            value={requestedStart}
            onChange={(e) => setRequestedStart(e.target.value)}
            required
          />
        </div>
        <div>
          <label>Preferred end</label>
          <input
            type="datetime-local"
            value={requestedEnd}
            onChange={(e) => setRequestedEnd(e.target.value)}
            required
          />
        </div>
        <button type="submit">Request Meeting</button>
      </form>
      {status && <p>{status}</p>}
    </div>
  );
}
```

- [ ] **Step 6: Create `frontend/src/pages/AdminDashboardPage.tsx`**

```tsx
import { useEffect, useState } from 'react';
import { apiGet, apiPost, apiPatch, staffLoginUrl } from '../api';

interface Staff {
  id: string;
  email: string;
  connected: boolean;
  accountType: string;
  organization: { id: string; domain: string };
}

interface Booking {
  id: string;
  visitorName: string;
  visitorEmail: string;
  requestedStart: string;
  requestedEnd: string;
  status: string;
  staff: Staff | null;
  joinUrl: string | null;
}

export function AdminDashboardPage() {
  const [staff, setStaff] = useState<Staff[]>([]);
  const [bookings, setBookings] = useState<Booking[]>([]);

  const refresh = async () => {
    const staffRes = await apiGet('/api/v1/staff');
    const bookingRes = await apiGet('/api/v1/bookings');
    setStaff(staffRes.data ?? []);
    setBookings(bookingRes.data ?? []);
  };

  useEffect(() => {
    refresh();
  }, []);

  const scheduleFor = async (bookingId: string, staffId: string) => {
    await apiPost('/api/v1/bookings', { bookingId, staffId, subject: 'SaaS Intro Meeting' });
    refresh();
  };

  const cancelBooking = async (bookingId: string) => {
    await apiPost(`/api/v1/bookings/${bookingId}/cancel`, {
      comment: 'Apologies, this meeting has been cancelled.',
      cancelledBy: 'admin',
    });
    refresh();
  };

  const swapHost = async (bookingId: string, newStaffId: string) => {
    await apiPost(`/api/v1/bookings/${bookingId}/swap`, { newStaffId });
    refresh();
  };

  const reschedule = async (bookingId: string) => {
    const newStart = prompt('New start (ISO 8601, e.g. 2026-09-06T10:00:00Z)');
    const newEnd = prompt('New end (ISO 8601)');
    if (!newStart || !newEnd) return;
    await apiPatch(`/api/v1/bookings/${bookingId}`, { newStart, newEnd });
    refresh();
  };

  return (
    <div style={{ maxWidth: 900, margin: '40px auto', fontFamily: 'sans-serif' }}>
      <h1>Admin Dashboard</h1>

      <h2>Staff</h2>
      <a href={staffLoginUrl()}>+ Connect Staff Account</a>
      <ul>
        {staff.map((s) => (
          <li key={s.id}>
            {s.email} — {s.connected ? 'Connected' : 'Disconnected'} ({s.accountType}, org:{' '}
            {s.organization?.domain})
          </li>
        ))}
      </ul>

      <h2>Bookings</h2>
      <table border={1} cellPadding={6}>
        <thead>
          <tr>
            <th>Visitor</th>
            <th>Time</th>
            <th>Status</th>
            <th>Staff</th>
            <th>Join URL</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          {bookings.map((b) => (
            <tr key={b.id}>
              <td>
                {b.visitorName} ({b.visitorEmail})
              </td>
              <td>
                {b.requestedStart} → {b.requestedEnd}
              </td>
              <td>{b.status}</td>
              <td>{b.staff?.email ?? '-'}</td>
              <td>{b.joinUrl ? <a href={b.joinUrl}>Join</a> : '-'}</td>
              <td>
                {b.status === 'Requested' && staff.length > 0 && (
                  <button onClick={() => scheduleFor(b.id, staff[0].id)}>Schedule w/ {staff[0].email}</button>
                )}
                {b.status !== 'Cancelled' && b.status !== 'Requested' && (
                  <>
                    <button onClick={() => reschedule(b.id)}>Reschedule</button>
                    <button onClick={() => cancelBooking(b.id)}>Cancel</button>
                    {staff.length > 1 && (
                      <button onClick={() => swapHost(b.id, staff[1].id)}>Swap to {staff[1].email}</button>
                    )}
                  </>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
```

- [ ] **Step 7: Create `frontend/src/App.tsx`**

```tsx
import { BrowserRouter, Routes, Route, Link } from 'react-router-dom';
import { VisitorBookingPage } from './pages/VisitorBookingPage';
import { AdminDashboardPage } from './pages/AdminDashboardPage';

export function App() {
  return (
    <BrowserRouter>
      <nav style={{ padding: 12, fontFamily: 'sans-serif' }}>
        <Link to="/">Visitor Booking</Link> | <Link to="/admin">Admin Dashboard</Link>
      </nav>
      <Routes>
        <Route path="/" element={<VisitorBookingPage />} />
        <Route path="/admin" element={<AdminDashboardPage />} />
      </Routes>
    </BrowserRouter>
  );
}
```

- [ ] **Step 8: Create `frontend/src/main.tsx`**

```tsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
```

- [ ] **Step 9: Install and verify the dev server boots**

Run: `cd frontend && npm install`
Expected: installs without errors.

Run: `npm run dev`
Expected: Vite prints `Local: http://localhost:5173/`.

Visit `http://localhost:5173/` in a browser: visitor booking form renders.
Visit `http://localhost:5173/admin`: admin dashboard renders (staff/booking lists empty until backend has data).

- [ ] **Step 10: Commit**

```bash
git add frontend/
git commit -m "feat: add Vite/React frontend (visitor booking form, admin dashboard)"
```

---

## Task 13: Manual End-to-End Smoke Test

**Files:**
- Create: `docs/superpowers/smoke-test.md`

**Interfaces:**
- Consumes: a real Azure App Registration (client_id/secret registered with redirect URI `http://localhost:5000/auth/staff/callback`, delegated scopes `offline_access Calendars.ReadWrite OnlineMeetings.ReadWrite User.Read`) and at least one real Microsoft account to connect as staff.

This task has no automated steps — it is a checklist run once the Azure App Registration exists and both backend (`npm run dev`, port 5000) and frontend (`npm run dev`, port 5173) are running.

- [ ] **Step 1: Create `docs/superpowers/smoke-test.md`**

```markdown
# Manual Smoke Test — Teams Scheduler Demo

Prerequisites: backend running on :5000, frontend running on :5173, Azure App
Registration created with redirect URI `http://localhost:5000/auth/staff/callback`
and delegated scopes `offline_access Calendars.ReadWrite OnlineMeetings.ReadWrite User.Read`.
`.env` populated with real `MS_CLIENT_ID`/`MS_CLIENT_SECRET`.

1. **Connect staff**: visit `http://localhost:5173/admin`, click "Connect Staff
   Account", complete Microsoft OAuth consent with a real account. Expect
   redirect back to `/admin?connect=success` and the account listed as
   "Connected".
2. **Visitor request**: visit `http://localhost:5173/`, submit a booking
   request with a start/end time in the future. Expect it to appear on
   `/admin` with status "Requested".
3. **Schedule**: on `/admin`, click "Schedule w/ <email>". Expect status to
   become "Scheduled" and a "Join" link to appear. Open the link, confirm it
   is a valid Teams meeting join URL. Check the staff member's real Outlook
   calendar — the event should appear with the visitor as an attendee.
4. **Reschedule**: click "Reschedule", enter a new ISO 8601 start/end. Expect
   status "Rescheduled" and the Outlook event's time to have moved. Confirm
   the visitor would receive Graph's automatic notification (Graph handles
   this — no separate email step in this app).
5. **Cancel**: on a different booking (repeat steps 2-3 to create one), click
   "Cancel". Expect status "Cancelled" and the event removed from the staff
   member's Outlook calendar.
6. **Swap host**: connect a second staff account in the *same organization*
   (same email domain) via step 1. Create and schedule a third booking, then
   click "Swap to <second email>". Expect status "Swapped", the original
   staff member's event cancelled, a new event on the second staff member's
   calendar, and a new join URL returned.
7. **Cross-org swap rejection**: connect a third staff account with a
   *different* email domain. Attempt a swap to that staff member via a raw
   API call:
   `curl -X POST http://localhost:5000/api/v1/bookings/<id>/swap -H "Content-Type: application/json" -d '{"newStaffId":"<cross-org-staff-id>"}'`
   Expect a 400 response: "Swap-host is only allowed between staff in the
   same organization".
8. **Disconnected staff handling**: in Azure Portal, revoke the app's
   consent for one connected staff account (Enterprise Apps → the app →
   Users and groups → remove user, or have the user revoke via
   myaccount.microsoft.com). Attempt to schedule a booking with that staff
   member. Expect a 401 and the staff's `connected` flag to flip to `false`
   on the next `GET /api/v1/staff` call.

Record pass/fail for each step; any failure blocks calling the demo done.
```

- [ ] **Step 2: Run through the checklist against a real Azure App Registration and at least two staff accounts in the same org**

No expected output beyond what's in the checklist — this step is what proves the demo actually works end-to-end. Do not mark the plan complete until every checklist item passes.

- [ ] **Step 3: Commit**

```bash
git add docs/superpowers/smoke-test.md
git commit -m "docs: add manual end-to-end smoke test checklist"
```

---

## Self-Review Notes

**Spec coverage:** §2 (App Registration values) → Global Constraints + Task 13 prerequisites. §3 (token mgmt) → Task 7. §4 Steps 1-5 → Task 9. §5 (admin experience) → Task 12 dashboard. §6 (parameter table) → covered by `Booking`/`Staff` entity fields (Tasks 5, 8). §7 (error handling) → Task 6 (429/401/404), Task 9 (404-on-event handling), Task 7 (invalid_grant). §8 (DB schema) → Tasks 4/5/8 entities. §9 (security) → Task 2 crypto + Task 7 encrypted storage + UTC-only timestamps throughout Task 9. §10 (open decision) → carried verbatim into design doc, no action needed today.

**Placeholder scan:** no TBD/TODO markers; every step has complete code.

**Type consistency:** `Booking.status` type `BookingStatus` (Task 8) matches the string literals set in Task 9 (`'Requested'`, `'Scheduled'`, `'Rescheduled'`, `'Cancelled'`, `'Swapped'`). `StaffService.updateTokens` signature `{ refreshToken, accessToken, expiresAt }` (Task 5) matches call sites in `TokenService` (Task 7). `BookingService` constructor args `(staffService, tokenService)` (Task 9) match `Container.get` DI resolution used implicitly by typedi in controllers (Task 10) — `BookingController`/`StaffController` depend on `BookingService`/`StaffService` by type, consistent with typedi's constructor-injection pattern used throughout.
