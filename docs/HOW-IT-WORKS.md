# How Sindooram Ledger works

A simple explanation of what this project is and how the pieces fit together.

## What it is

A single web page (no app-store app, no server we manage) where anyone in the
family/staff can log villa expenses, income, and guest bookings from their
phone, and everyone sees the same live numbers.

**Live link:** https://infosindooramecostays-oss.github.io/Sindooram-Ledger/

## Where the data actually lives

**A Google Sheet is the real database.** The page itself stores nothing
permanently except a small per-device cache (so it opens instantly) and your
name/connection settings.

- The Sheet lives in Google Drive, owned by whoever created it.
- A small script attached to the Sheet (**Google Apps Script**, in
  Settings → "How to connect a Google Sheet") turns it into a tiny API: the
  page can ask it "give me all the data" (`doGet`) or "save this data"
  (`doPost`).
- Every add / edit / delete on the page: fetches the latest data from the
  Sheet, applies your change, writes the whole thing back. This keeps
  concurrent edits from clobbering each other most of the time.
- The page also silently re-checks the Sheet every ~25 seconds so everyone's
  view stays current without refreshing.

## Connecting a device

The Apps Script URL is **not baked into the page** (the code is public on
GitHub; the Sheet must not be). Each person opens the link once, pastes the
same Apps Script URL into **Settings → Save & connect**, and their browser
remembers it from then on.

## Categories, sub-categories, recurrence

- Category and Sub-category are dropdowns with an **"Other…"** option that
  reveals a text box for anything not listed. Whatever you type becomes a
  real option for everyone next time, because it's read back from the Sheet.
- The pencil icon next to each field **renames it everywhere**, including
  past entries — it rewrites the category/sub-category on every matching row
  in the Sheet, not just the one you're editing.
- "Recurring?" is a plain label (One-off / Weekly / Monthly / Yearly /
  custom) for your own reference — it does **not** automatically create
  future entries.

## Installing it like an app

The page is a basic installable web app (PWA):
- **Android/Chrome:** open the link → browser menu → "Add to Home screen" /
  "Install app".
- **iPhone/Safari:** open the link → Share button → "Add to Home Screen".

Details on how this was built: [MOBILE-APP-SETUP.md](MOBILE-APP-SETUP.md).

Either way it gets a real icon and opens full-screen, no browser bar.

## Repo layout

```
index.html              the entire app (markup + styles + logic, one file)
manifest.webmanifest     tells phones how to install it (name, icon, colors)
sw.js                    minimal service worker, required for installability
assets/icons/            app icon source (icon.svg) and generated PNGs
docs/                    this file
```

`index.html` must stay at the repo root — GitHub Pages serves it from there.

## Making changes & deploying

1. Edit `index.html` (or ask Claude to).
2. `git add`, `git commit`, `git push origin main`.
3. GitHub Pages rebuilds automatically within about a minute — no manual step.

## If you change what data is saved (e.g. add a new field)

The Apps Script deployed on your Google account is a **copy** of the code
shown in Settings, not something that updates itself. If a new field stops
saving after an update:

1. Open the Sheet → Extensions → Apps Script.
2. Select all, delete, paste the latest code from Settings → "How to connect
   a Google Sheet".
3. Save. The existing deployment URL keeps working — no redeploy needed.
