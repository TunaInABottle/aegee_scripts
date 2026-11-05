# aegee_scripts

A monorepo of Google Apps Script automation tools. I have used these scripts to automate boring tass and ease some processes within AEGEE association.

| Project | Description |
|---|---|
| [`email_reminder/`](email_reminder/) | Daily digest of unanswered Gmail threads, with a tracking spreadsheet. |
| [`hrc_email_merger/`](hrc_email_merger/) | Gmail mail merge inside Google Sheets: alias picker, safe "sample" mode, `{{token}}` templates. |
| [`hrc_oc_data_splitter/`](hrc_oc_data_splitter/) | Splits a spreadsheet of application responses into one spreadsheet per European Body in Google Drive. |

## Layout & conventions

- Every subdirectory is a self-contained clasp project with its own
  `appsscript.json`, `.claspignore`, and a gitignored `.clasp.json` holding the
  personal script ID.
- Personal settings belong in each project's `Config.js` (see the shipped
  `Config.example.js`); it is gitignored and never committed.
- Each project keeps its own `.gitignore`, `LICENSE`, and `README.md`.

## Working on a project

```bash
cd <project>
clasp login
clasp push
```