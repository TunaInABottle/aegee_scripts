# email_reminder

A Google Apps Script that watches a Gmail inbox and sends a **daily digest of
unanswered email threads** to the observed address, while keeping a tracking
spreadsheet in sync.

## What it does

1. Searches the observed inbox with `in:inbox to:<observedMail>`, skipping every
   thread tagged with the skip label (`SCRIPT_IGNORE` by default).
2. Threads marked `Ignore = TRUE` in the tracking sheet get the skip label, so
   they are not reported again.
3. The still-unanswered threads are written back into the active sheet — one row
   per thread with ID, date, subject, last sender email and last sender name.
4. If any thread is still unanswered, an HTML digest is sent to the observed
   inbox listing date, subject and last sender, with a footer that links to a
   spreadsheet where receivers can mark threads as "no reply needed", plus an
   automatic-message disclaimer.

## Project layout

| File | Contents |
| ----------------- | ------------------------------------------------------- |
| `main.js`         | entry point `mainThreadFunc()`, sheet sync, digest email |
| `email.js`        | `EmailThread` / `EmailMessage` / `EmailLabel` prototype helpers and follow-up support |
| `date_methods.js` | `Date` prototype helpers (add days, workdays, weekend handling) |
| `constants.js`    | configuration constants and lazy access to your personal `Config.js` |

## Installation

### Option A — Apps Script UI

1. Create (or open) a Google Sheet — the script reads/writes its active sheet.
2. **Extensions → Apps Script** and add `constants.js`, `main.js`, `email.js`,
   `date_methods.js` and `appsscript.json` from this repository.
3. Add one more script file named `Config` with your personal values (see
   [Configuration](#configuration)).

### Option B — clasp

```bash
git clone https://github.com/<you>/email_reminder.git
cd email_reminder
npm install -g @google/clasp && clasp login
clasp create --type sheets --title "Email reminder"   # creates your own script
clasp push
```

## Configuration

Personal values live in a local `Config.js` file:

```bash
cp Config.example.js Config.js   # then edit Config.js
```

| Key | Meaning |
| ----------------------- | -------------------------------------------------------- |
| `observedMail`          | inbox checked for unanswered mail; also the digest recipient |
| `footnoteSpreadsheetUrl`| spreadsheet linked in the digest footer |
| `contactNote`           | contact note appended to every automatic email |

`Config.js` is **gitignored** (it holds your personal inbox and contact info)
but *is* pushed to your Apps Script project by `clasp push`, so the code works
at runtime. `.clasp.json` (your personal script ID) is gitignored for the same
reason, and `Config.example.js` is excluded from pushes via `.claspignore`.

## Scheduling

Add a **time-driven trigger** for `mainThreadFunc()` (e.g. daily) via
Apps Script → Triggers → Add Trigger.

## License

[MIT](LICENSE)
