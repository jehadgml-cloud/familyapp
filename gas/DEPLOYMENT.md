# Deploying the Google Sheets backend

These steps are manual — done once by whoever owns the Google account this
runs under (the family committee). An AI agent cannot perform them; they
require the Google Sheets/Apps Script UI and real Google credentials.

## 1. Create the Sheet and paste in the code

1. Go to https://sheets.google.com and create a new blank spreadsheet.
   Rename it (e.g. "صندوق تحصيل عائلة آل اطفيحة — قاعدة البيانات").
2. Extensions → Apps Script. This opens a bound Apps Script project.
3. Delete the default empty `Code.gs` file's contents (or delete the file).
4. For each file in this repo's `gas/` folder — in this exact order —
   create a matching file in the Apps Script editor (File → New → Script
   file, name it without the `.gs` extension, e.g. `Util`) and paste the
   contents in:
   1. `Util.gs`
   2. `Setup.gs`
   3. `DriveHelper.gs`
   4. `Months.gs`
   5. `Families.gs`
   6. `Members.gs`
   7. `Expenses.gs`
   8. `Users.gs`
   9. `PendingUsers.gs`
   10. `Signatures.gs`
   11. `Settings.gs`
   12. `Bootstrap.gs`
   13. `Api.gs`
5. Save the project (Ctrl+S / Cmd+S).

## 2. Run the one-time setup

1. In the function dropdown (top toolbar, next to Run/Debug), select
   `setupSpreadsheet`.
2. Click **Run**. The first run will prompt for authorization — review the
   permissions (it needs access to this Sheet and to create a Drive
   folder) and allow them.
3. Check the execution log (View → Logs, or Ctrl+Enter) — it should end
   with `setupSpreadsheet() finished successfully.`
4. Switch back to the Sheet tab: confirm 8 new tabs now exist — Members,
   Families, Months, Expenses, Users, PendingUsers, Signatures, Settings —
   each with a header row. Months should have 12 rows and Users should
   have 5 rows (the committee accounts); the rest start empty.
5. Check Google Drive for a new folder named "CEMS-ATFIHAH Attachments".

Re-running `setupSpreadsheet` later (e.g. after adding a new `.gs` file) is
always safe — it never overwrites existing sheet data.

## 3. Deploy as a Web App

1. In the Apps Script editor: Deploy → New deployment.
2. Click the gear icon next to "Select type" → Web app.
3. Description: anything (e.g. "CEMS API v1").
4. Execute as: **Me** (your account).
5. Who has access: **Anyone**.
6. Click Deploy. Authorize again if prompted.
7. Copy the **Web app URL** shown (ends in `/exec`).

## 4. Wire it into the app

1. Open `app.js` in this repo, find the `API_URL` constant near the top.
2. Paste the Web app URL in as its value.
3. Serve the app (`npx serve .`) and open it in a browser — the login
   screen should appear, and logging in with one of the 5 default accounts
   (email from `gas/Setup.gs`'s `DEFAULT_USERS_`, password `ABC12345`)
   should work end-to-end against the live Sheet.

## Redeploying after a code change

Apps Script Web App URLs are pinned to a specific deployment. After
editing any `.gs` file:
1. Deploy → Manage deployments.
2. Click the pencil/edit icon on the active deployment.
3. Version: **New version**. Deploy.
4. The URL stays the same — no need to update `app.js` again.

## 5. What "empty Members/Families" means on first run

`setupSpreadsheet()` deliberately does not seed a member roster (see the
note at the end of Backend Task 2 in the implementation plan). To load the
real ~134-member roster: open the app, log in, go to "بوابة تحميل وتخزين
قاعدة البيانات" (the Excel upload tab) and upload the same Excel file the
app used before — this now pushes the parsed data into the Sheet via
`bulkImportMembers` instead of `localStorage`.
