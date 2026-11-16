# Copilot Instructions — AUTO BAT\other\timezync

This subfolder is a **separate, unrelated side project** kept apart from the rest of the
`AUTO BAT` launcher collection. It lives at `other\timezync\` — a named subfolder under the
otherwise-empty `other\` area, which is reserved for holding one subfolder per future side
project. This one has its own purpose and its own conventions — do not merge these notes into
the root `.github/copilot-instructions.md`, and do not apply the root file's conventions (e.g.
`%USERPROFILE%\<Name>.ps1` external scripts, `RUN_ALL.bat` exclusion list) here; they don't
apply.

## What's here

- **`IST-Meeting-to-Calendar.hta`** — Windows desktop tool (`mshta.exe`) for converting meeting
  invites given in any timezone into the user's local Michigan/Eastern time and adding them to
  a calendar. Displayed app name is **"TimeZync"** (in `<title>`, `HTA:APPLICATION
  APPLICATIONNAME`, and the `<h1>`) — the file itself keeps its original name
  (`IST-Meeting-to-Calendar.hta`) so existing shortcuts/paths don't break; only the
  user-visible branding changed. Card-style UI with emoji icons/gradient buttons.
  - **Meeting title**: dropdown built dynamically at load by `renderTitleSelect()` from
    `DEFAULT_TITLES` (`Gyananjali Gita/Bhagvat`, `Gyananjali Stotra`, `Gyananjali Karmkand`,
    `Gyananjali Other`, `Jyotish Acharya`, `Jyotish Bhushan`, `Nakshatra`) plus any user-added
    custom titles, then `Other` (which reveals a free-text `<input id="title">` box via
    `toggleCustomTitle()`).
  - **Meeting timezone**: dropdown built by `renderTzSelect()` from `DEFAULT_TIMEZONES` (IST,
    UTC, US Eastern/Central/Mountain/Pacific, Gulf/Dubai — all fixed offsets, no DST) plus any
    user-added custom timezone.
  - **Conversion**: `computeEvent()` computes the true UTC instant using the *selected*
    timezone's fixed offset (via `findTimezone()`), then derives the local-equivalent
    date/time using this PC's own clock/timezone — this correctly rolls the calendar date
    backward/forward when needed (e.g. an IST morning meeting often lands on the previous US
    calendar day).
  - **Google Calendar** button: opens a prefilled "create event" link via
    `WScript.Shell`/`cmd /c start`.
  - **Outlook** button: creates a real `olAppointmentItem` via
    `ActiveXObject('Outlook.Application')` and calls `.Display()` so the user reviews/saves it
    in desktop Outlook.
  - **Reset** button: clears the title/notes fields, resets the title dropdown to its first
    option, and reapplies today's date + current time via `setDefaultsToNow()`, so the form
    can be reused for the next meeting without closing/reopening the app.
- **`IST-Meeting-to-Calendar.html`** — plain HTML/JS twin of the `.hta` above, for use on
  Android/iOS/any browser (no ActiveX/`WScript.Shell` available there). Also branded
  **"TimeZync"** for the same reason (file name kept stable for the deployed URL/QR
  code/installed PWA). Mobile-responsive, can stay open across multiple entries.
  - Same dynamic title/timezone dropdowns, custom-title toggle, and
    `computeEvent()`/`setDefaultsToNow()` logic as the `.hta`.
  - **Google Calendar** button: same prefilled link, opened via a normal
    `<a target="_blank">` click instead of `cmd /c start`.
  - **Outlook** button: browsers can't automate `Outlook.Application`, so instead this builds
    a `.ics` file client-side (`Blob`/`URL.createObjectURL`) and triggers a download via a
    hidden `<a download>` link; the user opens it and imports into Outlook (desktop or mobile
    app) or their phone's default calendar.
- **`Open-PWSH7.bat`** — `cd /d "%~dp0"` then opens `pwsh.exe -NoExit -Command "copilot"`, i.e.
  launches the standalone GitHub Copilot CLI (`copilot`, not the `gh copilot` extension used by
  the root `Open-GH-Copilot.bat`) with its working directory set to this `other\timezync\`
  folder.
- **Settings panel** (gear ⚙️ button in the header of both `.hta` and `.html`) — an in-page
  overlay (no separate route/URL, so it still works offline via the cached page) for
  personalizing the form without editing code:
  - **Meeting titles** and **Timezones** sections each show **one combined list** (built-in +
    custom) where every entry — including the original built-ins — has a **✕ remove button**.
    Removing a built-in doesn't delete it from the code; it just adds its
    name/id to a `hiddenTitles`/`hiddenTimezoneIds` array in settings so it's filtered out of
    both the settings list and the main dropdown (`visibleDefaultTitles()` /
    `visibleDefaultTimezones()`). Each section has a **"↺ Restore built-in ..."** link (only
    shown once something is hidden) that clears that hidden list. Removing a *custom* entry
    (added via the "New meeting title" / "Name e.g. Dubai" add-rows) deletes it outright from
    `settings.customTitles` / `settings.customTimezones`.
  - A safety check (`hideDefaultTitle`/`hideDefaultTimezone`) blocks removing the very last
    remaining visible title or timezone (via `alert(...)`) so the dropdown can never end up
    empty.
  - Custom timezones store `{ id: 'custom-<timestamp>', label, offsetMinutes }`; the offset is
    always a fixed value entered as sign + HH:MM (no DST), same philosophy as IST.
  - **This device** section is read-only info only — the conversion to the phone/PC's local
    timezone is automatic via the browser's own `Date`/`Intl` APIs and needs no user setting.
  - **Persistence and its limits**: all settings (custom titles/timezones, hidden built-ins,
    selected timezone) persist in `localStorage` under key `istMeetingSettings.v1`, scoped to
    that one browser + origin/device only:
    - They do **not** sync between the desktop PC and a phone, or between different browsers
      on the same device — each is independent.
    - **Clearing that browser's site data/cache, or uninstalling and reinstalling the PWA,
      permanently wipes all custom titles/timezones and any hidden built-ins** back to the
      original defaults. There is currently **no export/import or backup feature** — the user
      explicitly declined adding one (see chat: "No thanks, this is fine as-is"), so don't add
      one unless asked again.
  - When adding new built-in titles/timezones or changing the settings data shape, update the
    `DEFAULT_TITLES`/`DEFAULT_TIMEZONES` arrays and the settings-render functions
    (`renderSettingsLists`, `makeSettingsRow`, `hideDefaultTitle`/`hideDefaultTimezone`,
    `restoreDefaultTitles`/`restoreDefaultTimezones`) in **both**
    `IST-Meeting-to-Calendar.hta` and `IST-Meeting-to-Calendar.html` identically.
- **PWA wrapper for `IST-Meeting-to-Calendar.html`** — turns the mobile HTML twin into an
  installable Android/Chrome "app" (home-screen icon, standalone window, offline reload) with
  no native build tooling required:
  - **`manifest.json`** — Web App Manifest (`name`/`short_name: "TimeZync"`, `start_url`,
    `display: standalone`, `theme_color`/`background_color`, icon list). Linked from the HTML
    via `<link rel="manifest">`.
  - **`icons/icon-192.png`, `icons/icon-512.png`, `icons/icon-maskable-512.png`** — generated
    PNG app icons (calendar-card glyph on the app's blue gradient) referenced by
    `manifest.json` and by the HTML's `<link rel="apple-touch-icon">`/favicon. Generated once
    via a PowerShell + `System.Drawing` script (not saved as a repo file) — regenerate the
    same way if the icon design needs to change, no external image tool is required.
  - **`service-worker.js`** — minimal cache-first service worker that precaches the HTML,
    manifest, and icons so the page still loads offline after first visit; registered from a
    `<script>` block at the bottom of `IST-Meeting-to-Calendar.html`.
  - **Cache-busting**: `CACHE_NAME` in `service-worker.js` is currently
    `ist-meeting-to-calendar-v3`. **Bump this version string (`...-v4`, `...-v5`, etc.) every
    time `IST-Meeting-to-Calendar.html` changes** — otherwise phones that already installed the
    PWA (or even plain Chrome tabs that visited it before) keep serving the old cached page
    indefinitely, since the old service worker intercepts all fetches until a new cache name
    triggers an update.
  - **Important:** Service worker registration (and therefore the Chrome "install app" /
    "Add to Home Screen" prompt) only works when the file is served over `http(s)://` or
    `localhost` — opening it directly as `file://...` in the browser will show the page fine
    but will **not** register the service worker or offer installation.
  - **Troubleshooting a stale/old-looking install on a phone**: after deploying an update, if
    the phone (browser tab or installed home-screen icon) still shows the old version, the fix
    is to clear that site's cached data specifically — tap the 🔒/ⓘ icon in Chrome's address
    bar → Site settings → "Clear & reset" (or clear "Cached images and files" for that site) —
    then reload / reinstall. A normal page refresh is often not enough because the old service
    worker is still in control until its cache is invalidated.
  - **Live deployment**: hosted via GitHub Pages from a **separate GitHub repo** named
    `TimeZync` (owned by a different GitHub account than the one this machine's local
    `git`/`gh` CLI is authenticated as, and **not** this `AUTO BAT` repo). That repo's root mirrors
    this folder's `IST-Meeting-to-Calendar.html`, `manifest.json`, `service-worker.js`, and
    `icons/` (plus a stray flattened copy of the icon PNGs at the repo root from an early
    manual upload — harmless, ignore it; new icon changes only need to go in `icons/`). If the
    repo is renamed again, GitHub Pages/the `.io` URL changes too — regenerate the QR code
    (see below) and re-verify the live URL whenever that happens.
  - **⚠️ How to deploy future updates**: `git push` from this machine to that repo **fails with
    a 403** (permission denied) because the authenticated
        account differs from the repo owner. Until/unless the user adds that account as a
    collaborator on that repo, every update must be deployed manually: the user goes to the
    repo page → **"Add file" → "Upload files"** → drags in the updated
    `IST-Meeting-to-Calendar.html` (and `service-worker.js`/`manifest.json` if changed) →
    commits directly to `main`. Give the user that exact file list each time files here change.
  - **QR code**: `qr-timezync.png` points at the current permanent GitHub Pages URL of the `TimeZync` repo,
    generated locally via `npx qrcode "<url>" -o qr-timezync.png -w 400` (the `qrcode` npm
    package) — no third-party QR web service was used, keeping the URL private to this
    machine's tools. Regenerate it (same command, new URL) any time the Pages URL changes
    (e.g. after a repo rename).

## Maintenance

- Keep `IST-Meeting-to-Calendar.hta` and `IST-Meeting-to-Calendar.html` in sync: the title
  dropdown options, `computeEvent()`/`setDefaultsToNow()` conversion logic, and card UI
  styling should match between the two. When one changes, mirror the change in the other.
- When adding a new file to this `other\timezync\` folder, add a short bullet here describing
  its purpose — do not add it to the root `.github/copilot-instructions.md`. If starting a
  genuinely new/unrelated app, give it its own new subfolder under `other\` instead (with its
  own `copilot-instructions.md`), rather than adding it in here.
- After editing `IST-Meeting-to-Calendar.html`, `manifest.json`, or `service-worker.js`: bump
  `CACHE_NAME` in `service-worker.js`, then tell the user which files to re-upload to
  the `TimeZync` repo via the web "Upload files" UI (see above) since this
  machine cannot push there directly.
