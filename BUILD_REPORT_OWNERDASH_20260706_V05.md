# BUILD REPORT — OWNERDASH_20260706_V05

PROJECT: KALMA OWNER DASHBOARD — CENTRAL AUTH FULL MIGRATION

ROOM ROLE: BUILDER ONLY

OWNER ACTION: NONE

## Baseline

- ZIP baseline: OWNERDASH_20260627_V04A.zip
- SHA256 baseline: 75042d07ea58717acee92224d5c4cbc350dfe3498f8efb6debd042e43f70b6fa

## Scope

Auth/login/session only.
No approval execution.
No aggregator.
No deploy.
No live spreadsheet mutation.
No business reader mutation.

## Function → file map

| Function | File | Purpose |
|---|---|---|
| `kodLoginOwnerDashboardV1(payload)` | `07_CENTRAL_AUTH_SESSION.js` | Central Auth login via `KalmaCore.requireStaffRole` |
| `kodValidateOwnerDashboardSessionV1(payload)` | `07_CENTRAL_AUTH_SESSION.js` | Server-side session validation |
| `kodLogoutOwnerDashboardV1(payload)` | `07_CENTRAL_AUTH_SESSION.js` | Server-side session cleanup |
| `kodRequireCentralStaffRole_(pin, allowedRoles, actionName)` | `07_CENTRAL_AUTH_SESSION.js` | Central Auth wrapper; no local PIN compare |
| `kodRequireDashboardSession_(payload, allowedRoles, actionName)` | `07_CENTRAL_AUTH_SESSION.js` | Protected read/session gate |
| `kodRequireOwnerDashboardSession_(payload, actionName)` | `07_CENTRAL_AUTH_SESSION.js` | OWNER-only guard |
| `kodGetOwnerDashboardV0(payload)` | `01_DASHBOARD_SERVICE.js` | Protected read-only dashboard load |
| `kodAuthorizeOwnerDashboard(payload)` | `01_DASHBOARD_SERVICE.js` | OWNER-only auth probe |

## Before / after auth flow

| Surface | Before | After |
|---|---|---|
| Dashboard load | Public dashboard endpoint after spreadsheet auth probe | Requires Central Auth server session |
| Login | No Owner Dashboard login endpoint | `kodLoginOwnerDashboardV1` calls `KalmaCore.requireStaffRole` |
| Session | No project server session | CacheService token with expiry |
| Local storage | No session UI | Token cache/display only; not permission authority |

## Before / after role guard

| Surface | Before | After |
|---|---|---|
| OWNER | No project role guard | OWNER allowed by Central Auth |
| OPS_ADMIN | No project role guard | OPS_ADMIN allowed for staff-level dashboard read if Central Auth allows module |
| OPS | Not handled | Legacy alias only; normalized to OPS_ADMIN after Central Auth result |
| OWNER-only | No explicit helper | `kodRequireOwnerDashboardSession_` allows OWNER only |
| Client role | Not applicable | Never authority |

## Regression matrix

See `central_auth_migration_proof_ownerdash_20260706_v05.json` → `regressionMatrix`.

## Static result

failedChecks = 0

## No mutation proof

Changed files contain no Spreadsheet write methods and no finance/ledger/stock/purchasing/invoice/receiving/revenue mutation calls.
