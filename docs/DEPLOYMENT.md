# GitHub Actions deployment

This project ships with a GitHub Actions workflow that pushes the Apps Script source to Google whenever the `main` branch is updated.

## Prerequisites

1. **Apps Script project**
   - Create or open the target Apps Script project from the Google Sheet.
   - Copy its Script ID (Apps Script editor URL).

2. **Google Cloud project**
   - Open [Google Cloud Console](https://console.cloud.google.com/).
   - Create/select a project and enable the **Google Apps Script API**.
   - Create a service account with the `Editor` role.
   - Generate a JSON key for that service account.
   - Share the Google Sheet with the service account email.

3. **GitHub repository secrets**
   - `GAS_SCRIPT_ID`: the script ID string.
   - `GAS_CREDENTIALS`: the full JSON key contents (paste the JSON verbatim).

## Deployment flow

1. Run `npm run link -- --scriptId 1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0` once (or adjust the Script ID if needed) so that `.clasp.json` points at the correct Apps Script project.
2. Ensure your local clone is connected to the GitHub repository (for example: `git remote add origin https://github.com/sub3mvcbambi-pixel/pokemon-card-management.git && git branch -M main`).
3. Push or merge to `main`.
4. GitHub Actions installs dependencies and runs `clasp push` using the secrets above.
5. The Apps Script project updates with the repository contents.
6. Reopen the spreadsheet to load the latest menu and code.

## Manual deployment

If you need to push from your machine:

```bash
npm install
npx clasp login --creds <(echo "$GAS_CREDENTIALS")
GAS_SCRIPT_ID="..." npx clasp push -f
```

Alternatively copy/paste the `.gs` files into the Apps Script editor.

## Troubleshooting

- **`The caller does not have permission`** → ensure the service account has access to the script and the spreadsheet.
- **`PERMISSION_DENIED: Apps Script API has not been used`** → enable the API in Google Cloud and wait a few minutes.
- **Workflow not triggered** → confirm your branch is `main` or adjust the workflow trigger.