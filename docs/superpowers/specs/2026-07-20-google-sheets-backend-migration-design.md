# Google Sheets Backend Migration — Design

Date: 2026-07-20
Status: Approved by user, ready for implementation planning

## 1. Problem & Goal

The fund-tracker app (root-level vanilla HTML/JS/CSS app — `index.html`, `app.js`, `data.js`,
`styles.css`, described in [AGENTS.md](../../../AGENTS.md)) currently persists **everything**
in the browser's `localStorage`: members, families, monthly payments, expenses, admin
accounts/passwords, pending registration requests, dynamic users, and receipt signatures.
There is no shared backend — data does not survive a cleared browser, does not sync across
devices, and each browser/device has its own independent copy.

The goal: replace `localStorage` as the source of truth with a **Google Sheet**, accessed
through a **Google Apps Script (GAS) Web App** acting as a JSON API. Every piece of data the
app currently keeps (accounts, members, families, months, expenses, pending/dynamic users,
signatures) must live in the Sheet — nothing stays local-only except the current browser
session's login flag.

## 2. Architecture

```
 Browser (index.html + app.js)
        |  fetch() POST, text/plain body = JSON {action, payload}
        v
 Google Apps Script Web App (doPost)
        |  reads/writes
        v
 Google Sheet (bound spreadsheet, one tab per entity)
        |  attachments/signatures
        v
 Google Drive folder (auto-created by setup script)
```

- The Apps Script project is **bound to the Google Sheet** the user creates (Extensions →
  Apps Script). All `.gs` files in this spec are pasted into that script editor.
- The script is deployed as a **Web App** (Deploy → New deployment → Web app), executing as
  "Me", access **"Anyone"** (per user decision — no additional auth layer; the app's existing
  login screen is the only gate, and this is a known, accepted tradeoff, not an oversight).
- The frontend talks to the Web App exclusively through `doPost` (a single JSON-RPC-style
  endpoint dispatched by an `action` field), never directly to the Sheets API. No Google
  credentials are ever embedded in the frontend.
- `doGet` is implemented only as a trivial health-check (returns a small JSON status) — it is
  not part of the data path, since GAS `doGet` requests from `fetch` are equally viable but
  the app uses POST uniformly for simplicity and consistency (all actions, including reads,
  go through `doPost`).

### CORS / transport note

Apps Script Web Apps do not support custom CORS preflight handling for `doPost`. To avoid the
browser sending a preflighted `OPTIONS` request (which GAS does not answer), all frontend
requests use:

```js
fetch(API_URL, {
  method: "POST",
  headers: { "Content-Type": "text/plain;charset=utf-8" },
  body: JSON.stringify({ action, payload })
})
```

`text/plain` is a CORS-simple content type, so no preflight is triggered. The Apps Script side
parses `e.postData.contents` as JSON regardless of the declared content type.

## 3. Google Sheet data model

One spreadsheet, one tab per entity. All tabs are created (if missing) and header rows are
written by an idempotent `setupSpreadsheet()` function — see §5.

### Members
| id | name | parent | \<month columns, one per entry in Months\> | sum |
|---|---|---|---|---|

- `parent` = family head's name (matches `Families.headName`), same relationship as today's
  `state.members[].parent`.
- Month columns are added dynamically: adding a new month (via the API) appends a new column
  to this sheet and a new row to `Months`, matching current behavior where `DEFAULT_MONTHS`
  drives both the ledger table and the members' `payments` map.
- `sum` is stored (not only computed client-side) so the sheet itself is a readable ledger for
  someone opening it directly; the API recomputes and rewrites `sum` on every payment update.

### Families
| familyId | headName | memberCount | subscription | totalPaid | membersList |
|---|---|---|---|---|---|

- Mirrors `state.families` exactly. `totalPaid` and `memberCount` are recalculated
  server-side by the API whenever a member's payments change (equivalent to today's
  `recalculateFamilyTotals()`, moved server-side).

### Months
| key | id | selected |
|---|---|---|

- `key` = display label (e.g. `"أبريل 2026"`), `id` = the month-of-year sort key used today,
  `selected` = TRUE/FALSE, replacing `state.selectedLedgerMonths` (the ledger view's month
  filter) with a real synced column instead of a local-only preference.

### Expenses
| id | date | amount | reason | category | authorized | attachmentName | attachmentType | attachmentUrl |
|---|---|---|---|---|---|---|---|---|

- `attachmentUrl` points to a file in the Drive attachments folder (see §4) instead of an
  inline base64 blob. `attachmentName`/`attachmentType` are kept for display purposes
  (matching today's attachment modal, which switches rendering by MIME type).

### Users
| key | name | role | email | pass | isBuiltIn | isMaster | isActive | firstLoginDone |
|---|---|---|---|---|---|---|---|---|

- Replaces the hardcoded `ADMINS` object and the `cems_dynamic_users` localStorage array with
  one unified table. `isBuiltIn` distinguishes the five original committee accounts from users
  approved later through the registration flow; `isMaster` flags the one account
  (`jehadgml@gmail.com`) that can never be removed/deactivated, matching
  `MASTER_ADMIN_EMAIL`'s current invariant.
- Passwords are stored in plain text, matching the app's current security posture (plaintext
  in `app.js`/`localStorage` today). This is a pre-existing tradeoff, not a new one introduced
  by this migration, and is called out here for visibility rather than silently carried
  forward.

### PendingUsers
| id | firstName | lastName | email | pass | requestedAt |
|---|---|---|---|---|---|

- Replaces `cems_pending_users`. Approving a pending user moves the row into `Users` and
  deletes it from `PendingUsers`; rejecting just deletes it.

### Signatures
| month | slotIndex | signatureUrl | updatedAt |
|---|---|---|---|

- Replaces the `cems_report_sigs_<month>` localStorage entries (an object keyed `1..5` per
  month, one slot per committee signer). `signatureUrl` points to a PNG in the Drive
  attachments folder.

### Settings
| key | value |
|---|---|

- Generic key/value store for any remaining singleton preference (e.g. last imported Excel
  filename, previously `cems_excel_filename`). Keeps the "nothing stays local-only" property
  without inventing a dedicated sheet for every small setting.

### Dropped (intentionally not migrated)
- `cems_excel_binary` (the cached raw uploaded workbook, kept only so the original file could
  be re-downloaded byte-for-byte): dropped. Once the Sheet is the source of truth, the
  original uploaded file has no further purpose — the Sheet itself, or the existing
  "export to Excel" buttons (which already rebuild an `.xlsx` from `state` client-side), cover
  the export need.
- Printable receipts: confirmed not persisted anywhere today (generated on the fly from
  `state.families` + form inputs for printing); nothing to migrate.

## 4. Drive attachments folder

`setupSpreadsheet()` creates (or reuses, if already present) a Drive folder named
`"CEMS-ATFIHAH Attachments"` and stores its folder ID in Script Properties
(`PropertiesService.getScriptProperties()`), so subsequent runs don't create duplicate
folders. Two API actions handle uploads:

- `uploadExpenseAttachment(payload)` — payload carries the same `{name, type, size, data}`
  shape the frontend's existing `readAttachmentFile()` already produces (a data-URL string in
  `data`). The handler decodes the base64 payload, writes it as a Drive file in the
  attachments folder, sets sharing to "anyone with the link can view" (so the frontend can
  render/download it directly), and returns the file's URL.
- `uploadSignature(payload)` — same shape, for the signature-pad's `canvas.toDataURL()`
  output.

The 1.5MB client-side size cap on expense attachments (already enforced in `app.js`) and the
signature-pad's inherently small PNG output keep individual uploads well under Apps Script's
request-size limits.

## 5. Apps Script files

All pasted into the Sheet's bound Apps Script project (Extensions → Apps Script):

- **`Setup.gs`** — `setupSpreadsheet()`, run once manually from the script editor (Run button).
  Idempotent: creates each tab with its header row only if the tab doesn't already exist;
  seeds `Months` with the 12 default month rows and `Users` with the five default committee
  accounts **only if those sheets are empty** (so re-running it after real data exists is a
  no-op, never destructive). Also creates/reuses the Drive attachments folder.
- **`Api.gs`** — `doPost(e)` parses `{action, payload}` from `e.postData.contents`, dispatches
  through an `ACTIONS` map to the matching handler function, wraps the whole call in
  `LockService.getScriptLock()` for any mutating action (serializes concurrent writes so two
  simultaneous requests can't corrupt a row), and returns
  `ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON)`.
  A minimal `doGet(e)` returns `{status: "ok"}` for a manual health check when visiting the
  deployment URL directly in a browser.
- **`Members.gs`** — `getMembers()`, `addMember()`, `updateMember()`, `deleteMember()`,
  `bulkImportMembers()`, `setMemberPayment()` (updates one month cell + recalculates `sum` and
  triggers the owning family's `recalculateFamilyTotals` equivalent).
- **`Families.gs`** — `getFamilies()`, `addFamily()`, `updateFamily()`, `deleteFamily()`,
  `recalculateFamilyTotals(headName)` (internal helper called by `Members.gs`).
- **`Months.gs`** — `getMonths()`, `addMonth()` (appends a Members column too), `updateMonth()`,
  `deleteMonth()` (removes the Members column too), `setSelectedMonths(list)`.
- **`Expenses.gs`** — `getExpenses()`, `addExpense()`, `deleteExpense()`,
  `uploadExpenseAttachment()`.
- **`Users.gs`** — `login(email, pass)`, `getUsers()`, `addSupervisor()`, `updateUser()`,
  `setUserActive()`, `deleteUser()`, `changePassword()`.
- **`PendingUsers.gs`** — `getPendingUsers()`, `requestRegistration()`, `approveUser()`,
  `rejectUser()`.
- **`Signatures.gs`** — `getSignatures(month)`, `saveSignature()` (calls the Drive upload
  helper), `clearSignature()`.
- **`Settings.gs`** — `getSettings()`, `setSetting(key, value)`.
- **`Bootstrap.gs`** — `getAllData()`, a single aggregating call returning
  `{members, families, months, expenses, users, pendingUsers, signatures, settings}` in one
  round trip, used once on app load.

## 6. Frontend integration (`app.js`)

- Add a config constant near the top of `app.js`:
  ```js
  const API_URL = "PASTE_YOUR_WEB_APP_URL_HERE";
  ```
  (the user fills this in after deploying the Web App).
- Add a single `callApi(action, payload)` async helper implementing the `text/plain` POST
  pattern from §2, with a clear Arabic error message surfaced (e.g. via the existing
  `showConfirm`/alert patterns already used in the app) on network failure or a non-2xx/JSON
  error response — the app now requires connectivity, so failures must be visible, not silent.
- Replace `loadCachedData()` with `loadDataFromServer()`, calling `getAllData` once on
  `DOMContentLoaded` and populating `state` from the response, instead of reading the
  `cems_data_*`/`cems_*` localStorage keys.
- Remove `saveToLocalMemory()` entirely. Every call site that currently calls it (member
  add/edit/delete, family recalculation, month add/edit/delete, expense add/delete, pending
  user approve/reject, dynamic user add/toggle-active/delete, signature save/clear, selected
  months toggle — roughly 15–20 call sites across `app.js`) is changed to call the matching
  `callApi(...)` action instead, then update `state` from the response and re-render.
- `handleExcelFile()`/`processWorkbook()` keep parsing the uploaded workbook client-side as
  today, but instead of writing straight into `state` + `saveToLocalMemory()`, the parsed
  members/families arrays are sent to the `bulkImportMembers` action, and `state` is then
  refreshed from the API response.
- `doLogin()`/the login form now call `callApi("login", {email, pass})` instead of checking
  the local `ADMINS` object; registration requests call `requestRegistration`; admin approval
  screens call `approveUser`/`rejectUser`/`setUserActive`/`deleteUser`.
- `saveSignature()`/`clearSignature()` call `uploadSignature`/`saveSignature`/`clearSignature`
  actions instead of writing to `cems_report_sigs_<month>`.
- Remove the `cems_excel_binary` base64 caching code path (`arrayBufferToBase64`,
  `base64ToUint8Array`, and their call sites) per §3's "dropped" decision.
- `state.currentUser` and the fact that a login happened this session remain in
  `localStorage` (`cems_logged_email`) — this is a local UI convenience (auto-resume session
  on refresh), not app data, and re-validates against the server on load rather than trusting
  a stale cached role.

## 7. Error handling

- Every `callApi` failure (network error, HTTP error, or `{error: "..."}` in the JSON
  response) surfaces a clear Arabic message to the user and aborts the in-progress action
  without mutating local `state` — the UI must never show a change that wasn't actually
  confirmed by the Sheet.
- Buttons that trigger a `callApi` call show a loading state (spinner/disabled), matching the
  pattern already used for the "add expense" button's file-processing spinner, since network
  round trips are no longer instant like `localStorage` writes were.
- No offline fallback: per the user's explicit requirement that nothing stay unsynced, the app
  does not fall back to any local cache when the API is unreachable — it surfaces the error
  and lets the user retry.

## 8. Out of scope / explicitly deferred

- No polling/auto-refresh for changes made by other users or directly in the Sheet (confirmed
  with user: only the acting user's own changes need to reflect immediately; a manual page
  reload is an acceptable way to see others' changes).
- No additional API authentication layer (confirmed with user: accepted risk, Web App deployed
  as "Anyone").
- No change to the unrelated Next.js/Supabase scaffold under `src/` (see AGENTS.md) — this
  migration only touches the root-level vanilla-JS app.

## 9. Verification approach

No automated test suite exists for either app in this repo. Verification is manual:
1. Run `setupSpreadsheet()` from the Apps Script editor; confirm all 8 tabs appear with correct
   headers, the Drive folder is created, and `Users`/`Months` are seeded.
2. Deploy as Web App, paste the URL into `API_URL` in `app.js`.
3. Serve the app locally (`npx serve .`) and manually exercise every migrated flow — login,
   add/edit/delete member, add/edit/delete family via a new member, add/remove a month,
   add/delete an expense with an attachment, approve/reject a registration request,
   draw/save/clear a signature, import an Excel file — confirming each change lands correctly
   in the actual Google Sheet and Drive folder, and that a full page reload reloads the same
   state from the Sheet (proving nothing was relying on localStorage).
