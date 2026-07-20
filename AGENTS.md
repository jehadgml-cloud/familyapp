# AGENTS.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> Note: this content is named `AGENTS.md` instead of `CLAUDE.md` because a hook installed by
> the academic-research-skills plugin (`ars_write_scope_guard.py`) unconditionally blocks any
> write to a file named `CLAUDE.md` in any project, treating it as protected ARS infrastructure.
> If you disable that plugin/hook, feel free to rename this file to `CLAUDE.md`.

## Repository reality check (read this first)

This repo contains **two unrelated, non-integrated applications**. There is no build step connecting them — they share a directory but not code, data, or purpose.

1. **The actual live app** — a single-page vanilla HTML/CSS/JS app at the repo root: [index.html](index.html), [app.js](app.js), [data.js](data.js), [styles.css](styles.css). This is a family monthly-fund collection tracker ("صندوق تحصيل عائلة آل اطفيحة" — CEMS-ATFIHAH fund/collection portal). It is what actually runs; it has no build process and no server dependency.
2. **An unused Next.js/Supabase scaffold** under [src/app](src/app) plus [schema.sql](schema.sql) — a different concept (a report submission/review/approval portal with departments, profiles, and audit logs). `package.json`, `tsconfig.json`, and `tailwind.config.js` all target *this* scaffold, and [README.md](README.md) documents *this* scaffold's setup — but none of it is wired to the app described in point 1. Treat this as a separate, largely incomplete surface; do not assume it shares state, auth, or data with the root-level app unless you're specifically extending it.

When a task mentions "the app", "the fund tracker", "members", "families", or "monthly payments", it almost always means the root-level vanilla-JS app, not the Next.js scaffold. Confirm with the user if genuinely ambiguous.

## Running things

### The vanilla-JS fund tracker (root-level app)
No build/install needed. Serve the directory statically and open `index.html`, e.g.:
```bash
npx serve .
# or
python -m http.server 8080
```
Opening `index.html` directly via `file://` mostly works too, since it only depends on two CDN scripts (SheetJS `xlsx`, Chart.js) loaded in the `<head>`.

### The Next.js/Supabase scaffold (src/)
```bash
npm install
npm run dev     # next dev
npm run build   # next build
npm run start   # next start
npm run lint    # next lint
```
Requires a `.env` populated from [.env.example](.env.example) (`NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY`, `SUPABASE_SERVICE_ROLE_KEY`) and the schema in [schema.sql](schema.sql) applied to a Supabase project's SQL editor. There is no test runner configured in either app.

## Architecture — vanilla-JS fund tracker (app.js / data.js / index.html)

Everything lives in one global `state` object (`app.js:115`) and is persisted to **`localStorage`** — there is no backend/database for this app. Key `state` fields: `currentUser`, `members`, `families`, `monthsList`, `expenses`, `originalWorkbook` (the raw SheetJS workbook, for re-export).

- **Auth**: hardcoded admin accounts live in the `ADMINS` object at the top of `app.js`, keyed by role (committee head, deputy, treasurer, financial auditor, collection coordinator). `MASTER_ADMIN_EMAIL` cannot be removed. Passwords/emails/active-flags can be overridden per-admin via `localStorage` (`cems_admin_pass_<key>`, `cems_admin_email_<key>`, etc. — see `loadAdminPasswords()`). Non-admin users go through a request → pending-approval → dynamic-user flow (`cems_pending_users`, `cems_dynamic_users` in localStorage), approved by an admin. `findUserByEmail()` is the single lookup point across both built-in and dynamic users. Credentials and session (`cems_logged_email`) are entirely client-side/localStorage — there is no real backend auth, so don't treat this as a secure auth system when reasoning about changes.
- **Data model**: members belong to families via a `parent` name field; each member has a `payments` map keyed by Arabic month-name strings (e.g. `"أبريل 2026"`) from `DEFAULT_MONTHS`/`state.monthsList`, plus a running `sum`. `data.js` seeds `DEFAULT_MEMBERS` and `DEFAULT_FAMILIES` (minified JSON literals — treat as generated data, not hand-edited source). `recalculateFamilyTotals()` keeps family-level rollups in sync with member payments.
- **Persistence**: `saveToLocalMemory()` / `loadCachedData()` are the read/write pair for all app data (members, families, months, expenses, selected ledger months) against `localStorage` keys prefixed `cems_data_*` / `cems_*`. Any new persisted field needs both a save and a load path added here.
- **Excel import/export**: driven by SheetJS (`xlsx` global from CDN). `handleExcelFile()` → `processWorkbook()` parses an uploaded workbook (`findHeaderRow()` locates header rows by keyword matching since sheet layouts vary); the raw workbook is base64-round-tripped (`arrayBufferToBase64`/`base64ToUint8Array`) and cached so it can be re-exported later.
- **UI**: single-page tab system — `switchTab()` toggles `.nav-link`/`.tab-content` elements by `data-tab` id; there's no router. Major tab renderers: `renderLedgerTable()` (monthly ledger with search/filter via `ledgerSearch`/`ledgerFilterMonth`/`ledgerFilterStatus`, debounced), `renderFamiliesTable()`, `renderMonthlyReport()`, `renderExpensesTable()`, `renderPendingUsers()`/`renderActiveUsers()` (admin user management), plus a signature pad (`initSignaturePad()`/`saveSignature()`) for receipt sign-off and `updatePrintPreview()` for printable receipts. Charts render via Chart.js in `renderCharts()`.
- **Arabic-specific handling**: `normalizeArabicName()` normalizes Arabic name variants (diacritics/letter forms) for matching members across sheets/search; keep using it rather than raw string comparison when matching Arabic names.

## Architecture — Next.js/Supabase scaffold (src/)

- Next.js 14 App Router + TypeScript, Tailwind CSS, RTL Arabic UI (`dir="rtl"`, Cairo font via `next/font/google`), configured as a PWA (`public/manifest.json`, `public/offline.html`).
- Routes under `src/app`: `login`, `dashboard`, `reports` (list/create), `reports/[id]` (view/edit), `settings`.
- Supabase clients: `src/lib/supabase/client.ts` (browser) and `server.ts` (server components/route handlers); `src/middleware.ts` handles session/auth routing.
- `schema.sql` defines the intended Postgres schema: `departments`, `profiles`, `reports`, `report_attachments`, `report_comments`, `audit_logs`, `system_settings`, with Row Level Security policies restricting writes to department members and report approval to admins/auditors (see [README.md](README.md) for the RLS summary).
- `src/lib/utils/arabicNormalize.ts` is the TS equivalent of the vanilla app's `normalizeArabicName()` — if both apps ever need to interoperate, this is the natural shared boundary, but currently they are independent implementations.

## Misc scripts

- [search_utils.py](search_utils.py) and [read_sheets.ps1](read_sheets.ps1) are one-off developer utility scripts (grepping `app.js` for function definitions; dumping sheet names/rows from a local Excel file via COM automation) rather than part of the app runtime. `read_sheets.ps1` hardcodes a local `OneDrive` path — it's environment-specific and not portable.
