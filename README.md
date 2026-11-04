# hrc_email_merger

A Gmail **mail merge** that runs entirely inside Google Sheets, with a sending-alias
picker, a safe "sample" mode, per-row sending with live progress, and `{{token}}`
template support driven by your sheet columns.

Adapted from Google's mail-merge sample by Martin Hawksey
(<https://developers.google.com/apps-script/apps-script-samples>) — see
[License](#license).

## Features

- **Send from any of your Gmail aliases** — a dialog lists your account's addresses
  and aliases, so merged mail goes out from the identity you choose.
- **Sample mode** — send each draft to yourself (with a `(SAMPLE)` subject suffix)
  before doing the real run. Duplicate subjects are only sampled once.
- **Live progress dialog** — recipient preview, progress bar, per-row error
  reporting, and a summary at the end.
- **Draft-based templates** — write your email as a Gmail draft (HTML, attachments
  and inline images all supported) whose *subject* matches the sheet's
  `Email subject` column.
- **`{{token}}` replacement** — any column name in the sheet can be used as a
  `{{Column Name}}` merge token; cell values are HTML-escaped automatically.
- **Sent tracking** — the `Email Sent` column is stamped with the send date, and
  already-sent rows are skipped on re-runs.
- **VLOOKUP helper** (`google_ss.js`) — an `onEdit` trigger that, when a cell in
  column A of the data sheet is set to `EB`, fills the `Recipient` column of every
  pending `EB` row with
  `=IFERROR(VLOOKUP(TRIM(H<n>),'EB body Vlookup'!A$1:B, 2, FALSE), "NO MATCH FOUND")`.
  Requires a sheet named `EB body Vlookup` and must be installed as an
  *installable* trigger (Triggers → Add Trigger → function `onEdit`, event source
  *From spreadsheet*, event type *On edit*).

## Sheet setup

Your data sheet needs these columns (names are configured at the top of
`dropdown.js`):

| Column         | Purpose                                        |
| -------------- | ---------------------------------------------- |
| `Recipient`    | Email address of the recipient                 |
| `Email subject`| Subject of the Gmail draft to use as template  |
| `cc`           | Optional CC address                            |
| `bcc`          | Optional BCC address                           |
| `Email Sent`   | Filled automatically with the send date        |

Any additional columns become `{{tokens}}` for the template.

## Installation

### Option A — Apps Script UI

1. Create (or open) a Google Sheet.
2. **Extensions → Apps Script**.
3. Add the files from this repository: `dropdown.js`, `google_ss.js`,
   `EmailAliasDialog.html` and `appsscript.json`.
4. Save, reload the sheet, and authorize the script on first run.

### Option B — clasp

```bash
git clone https://github.com/<you>/hrc_email_merger.git
cd hrc_email_merger
npm install -g @google/clasp && clasp login
clasp create --type sheets --title "Email Merger"   # creates your own script
clasp push
```

> `.clasp.json` (which contains your personal script ID) is deliberately
> **gitignored** — keep yours local, never commit it.

## Usage

1. Create a Gmail draft whose **subject** equals the `Email subject` value in the
   sheet; use `{{Column Name}}` tokens in the body.
2. In the sheet, open **Mail Merge → Send Email samples** to test, then
   **Mail Merge → Send Emails** to send to all pending rows.

Privacy: emails are read only from your own sheet/drafts and are sent through your
own Gmail account; nothing is sent anywhere else.

## License

[Apache License 2.0](LICENSE) — this project is an adaptation of Google's
mail-merge sample © 2022 Martin Hawksey, licensed under Apache 2.0.
