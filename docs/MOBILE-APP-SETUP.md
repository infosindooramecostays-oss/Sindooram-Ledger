# How we made this a home-screen app

What was added so the ledger installs like a real app on phones, and why.

## The 3 pieces

1. **`manifest.webmanifest`** — a small JSON file telling the phone the app's
   name, icon, and colors, and that it should open `display: standalone`
   (full-screen, no browser bar) when launched from the home screen.
2. **App icon** — generated from `assets/icons/icon.svg` (a gradient square
   with a ledger glyph, matching the app's theme) into the sizes phones
   expect: `icon-192.png`, `icon-512.png`, and `apple-touch-icon.png` (180px,
   what iOS specifically looks for).
3. **`sw.js`** — a minimal service worker (just passes network requests
   through, no offline caching). Chrome/Android require a service worker to
   be present before it will offer the "Install app" prompt; it's registered
   at the bottom of `index.html`.

These, plus a `<link rel="manifest">` and `<link rel="apple-touch-icon">` in
`index.html`'s `<head>`, are the entire setup — no build tools, no app store.

## Why no offline support

The service worker deliberately does **not** cache anything. The app depends
on a live connection to the Google Sheet anyway, so caching the page for
offline use would only risk showing stale data or old code. It exists purely
to satisfy the install requirement.

## Installing it

**iPhone (Safari):** open the link → Share icon → **Add to Home Screen**.

**Android (Chrome):** open the link → **⋮** menu → **Add to Home screen** /
**Install app**.

Both give a real icon and a full-screen window — confirmed working on iPhone.

## Bonus: link previews

Added Open Graph meta tags (`og:title`, `og:description`, `og:image`) so
pasting the link into WhatsApp or similar shows a title, description, and
the app icon as a preview card instead of a bare URL. Uses a wider version
of the same icon (`assets/icons/og-image.png`).
