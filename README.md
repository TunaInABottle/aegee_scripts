# hrc_oc_data_splitter

A Google Apps Script that splits a spreadsheet full of application responses
(e.g. Google Form results) into **one spreadsheet per European Body**, inside a
dated subfolder of your Google Drive.

## What it does

1. Reads the `Data to move` sheet of the active spreadsheet.
2. Finds the "European Body" question columns (both
   `Which European Body do you wish to apply for?` and
   `To which European Body do you wish to apply for?` are recognized and merged
   into a single `European Body Application` column).
3. Asks for a Drive **folder ID**, a **month** and a **year** via a dialog.
4. Creates a child folder named `{year}-{MM} (Delete 30 {delMonth} {delYear})`.
5. Creates `Open Call <month> - All Data` (all matching rows, EB columns merged).
6. Creates one `Open Call <Month>: <Body>` spreadsheet per body with only its
   rows — headers bolded, first row frozen, fully empty columns removed.
7. On complete success, clears the source sheet; if any body failed, the source
   data is kept and the failures are listed in the dialog.

Cleanup applied to every output spreadsheet:

- leading `Timestamp` column removed,
- columns whose header contains `I agree` (privacy-consent columns) removed,
- columns with empty headers removed,
- rows whose EB columns disagree are reported as conflicts instead of being split.

## Installation

### Option A — Apps Script UI

1. Open the spreadsheet that contains the `Data to move` sheet.
2. **Extensions → Apps Script**, add `Code.js`, `Form.html` and
   `appsscript.json` from this repository, save and authorize.

### Option B — clasp

```bash
git clone https://github.com/<you>/hrc_oc_data_splitter.git
cd hrc_oc_data_splitter
npm install -g @google/clasp && clasp login
clasp create --type sheets --title "Data Splitter"   # creates your own script
clasp push
```

### Configuration

`Code.js` reads folder IDs from a `CONFIG` variable defined in a local
`Config.js` file:

```bash
cp Config.example.js Config.js   # then edit Config.js with your folder ID
```

- `Config.js` is **gitignored** (it holds your personal IDs) but *is* pushed to
  your Apps Script project by `clasp push`, so the code works at runtime.
- `Config.example.js` shows the expected shape and is excluded from pushes via
  `.claspignore`.
- `.clasp.json` (your personal script ID) is also gitignored — keep it local.

The normal flow asks for the folder ID in the dialog, so `CONFIG` is only needed
for the `myCreateSpreadsheet()` quick-test entry point.

## Usage

In the sheet, open **Scripts Menu → Create EB Spreadsheets**, then enter the
Drive folder ID (the long string in the folder's URL), the month and the year.

## License

[MIT](LICENSE)
