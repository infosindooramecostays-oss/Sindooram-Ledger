# Sindooram ID — what this is and how it works

A separate, simple app for the property manager to upload guest ID/documents
per booking. It's deliberately its own app with its own Apps Script — the
manager never sees the Ledger's money data.

## What it does

1. **Bookings tab** — pulls a read-only list of bookings from the Ledger
   Sheet (Guest name, Booking number, Check-in, Remarks), filtered to
   **Confirmed** bookings only.
2. Tap a booking → its details show read-only, plus one field the manager
   fills in: **Additional comments**.
3. **Take Photo** or **Choose File** adds files to a "Ready to upload" list —
   nothing uploads yet. Each has a **View** link to double-check it before
   sending. Add as many as needed (e.g. ID front + back).
4. **Submit** uploads everything at once, tagged with the manager's name
   (set once in Settings).

## Where the data goes

- **Files** → Google Drive, under
  `Sindooram ID Vault/<check-in date>/<Booking Number>/`.
- **A log row** → a sheet tab called `ID Vault Uploads`, auto-created inside
  the same Ledger spreadsheet. Columns: Booking Number, Guest Name, Check-in,
  Comments, File Name, Uploaded By, Uploaded At, File Link.
- Nothing is ever written back to the Ledger's own Transactions or Bookings
  tabs — this app only reads a few booking fields and appends upload logs.

## Architecture

- `index.html` — the whole app (styles, markup, JS), same single-file pattern
  as the main Ledger.
- `apps-script.gs` — a **separate, standalone** Apps Script (not bound to the
  Ledger Sheet). It uses `SpreadsheetApp.openById(LEDGER_SHEET_ID)` to read
  bookings and log uploads in the same spreadsheet the main Ledger uses, but
  it's a different deployment with a different URL and much narrower access:
  no amount, status (beyond the Confirmed filter), source, check-out date, or
  guest contact info is ever read or returned.
- `index.html` keeps an identical copy of `apps-script.gs` embedded as a
  string (`APPS_SCRIPT_CODE`), shown in the app's Settings tab with a
  "Copy script" button — so setup doesn't require this repo, just the app.
- `manifest.webmanifest` / `sw.js` / `assets/icons/` — same PWA install
  pattern as the main Ledger (see `docs/MOBILE-APP-SETUP.md` at the repo
  root), installable as a home-screen app on iPhone and Android.

## One-time setup (already done for this deployment)

1. New Apps Script project at script.google.com (kept separate from the
   Ledger's own script on purpose).
2. Paste in `apps-script.gs`, with `LEDGER_SHEET_ID` filled in.
3. Deploy → Web app → Execute as **Me** → Who has access **Anyone** → Deploy.
4. Paste the resulting `/exec` URL into the app's Settings tab.

## After changing apps-script.gs

Saving in the Apps Script editor is **not** enough — the live URL only
updates when you redeploy:

Deploy → Manage deployments → pencil (edit) icon → Version: **New version**
→ Deploy.

Same URL stays live, no need to reconnect the app in Settings.

## Installing on a phone

**iPhone (Safari):** open the app link → Share icon → **Add to Home Screen**.

**Android (Chrome):** open the app link → **⋮** menu → **Add to Home
screen** / **Install app**.

Renaming the app later (manifest/title changes) won't rename an
already-installed home screen icon — remove and re-add it to pick up the
new name.
