# Pokemon Card Sales Operations Spreadsheet

This repository contains a Google Apps Script (GAS) codebase and automation assets for a Google Sheets driven operations system tailored to Pokémon card commerce. The solution follows the requirements in the "要件定義 V1.1" specification and covers ID generation, order & inventory processing, payment reconciliation, and Looker Studio exports.

## Repository structure

```text
src/                  # GAS source files (one file per domain)
appsscript.json       # GAS manifest
.github/workflows/    # GitHub Actions workflows
package.json          # Tooling definition (clasp)
docs/                 # Extended documentation (setup, deployment)
```

## Getting started

1. Create (or open) the Google Spreadsheet that will host the solution.
2. Enable the Google Apps Script API in the associated Google Cloud project.
3. Generate a `.clasp.json` that points to the Apps Script project with:
   ```bash
   npm install
   npm run link -- --scriptId 1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0
   ```
   Adjust the Script ID if you deploy to a different spreadsheet.
4. Deploy this repository using the GitHub workflow (see [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)) or run `npm run deploy` locally.
5. Reload the spreadsheet – the custom **SalesOps** menu will appear.
6. Run **SalesOps → システム初期化** once to create and prepare all sheets.

### Linking this repository to GitHub

If you already created the remote repository `sub3mvcbambi-pixel/pokemon-card-management`, connect it to this working copy before
pushing changes:

```bash
git remote add origin https://github.com/sub3mvcbambi-pixel/pokemon-card-management.git
git branch -M main
git push -u origin main
```

GitHub will ask for authentication. Generate a [Personal Access Token](https://github.com/settings/tokens) with the `repo`
scope, copy it, and use it as the password when prompted. Once the push succeeds, commits in this repository will appear in the
GitHub project and the Actions workflow can deploy to Google Apps Script.

## Core capabilities

- Automated sheet provisioning with validation rules and protected formula columns.
- Deterministic ID generators for customers, orders, and SKUs.
- Customer, order, and inventory CRUD helpers exposed through the SalesOps menu.
- TSV/CSV order line ingestion with flexible column mapping.
- SKU allocation and shipment workflows that synchronise the stock ledger.
- PayPal & Wise statement ingestion and order/payment reconciliation.
- Customer aggregate maintenance (purchase count, total, first/last purchase).
- Non-PII Looker Studio export rebuilt on demand.
- Optional backup routine and audit logging hooks.

## Local development

Although Apps Script runs in Google's servers, the source code is maintained here so it can be versioned and deployed through CI/CD.

- Clone the repository and edit the `.gs` files inside `src/`.
- Use `npm install` to install the clasp CLI if you need local pushes.
- The GitHub workflow performs the `clasp push` using repository secrets, so you rarely need to push manually.

## Support

If the workflow fails, check the Actions logs for details. Common causes include missing `GAS_SCRIPT_ID`, invalid service account credentials, or insufficient permissions for the Apps Script project.

## 日本語での利用手順

1. Google スプレッドシート（ID: `1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0`）を開き、共有先に GitHub Actions のサービスアカウントを追加します。
2. リポジトリで次のコマンドを実行して対象スクリプトを紐付けます。
   ```bash
   npm install
   npm run link -- --scriptId 1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0
   ```
3. `main` ブランチへプッシュするか、ローカルで `npm run deploy` を実行すると Apps Script にコードが反映されます。
4. スプレッドシートを再読み込みし、メニュー **SalesOps → システム初期化** を実行して各シートを自動生成します。
5. 以降は **新規顧客作成** や **新規注文作成** などのメニューから運用フローを開始できます。

より詳しい日本語ガイドは [docs/JP_GUIDE.md](docs/JP_GUIDE.md) を参照してください。