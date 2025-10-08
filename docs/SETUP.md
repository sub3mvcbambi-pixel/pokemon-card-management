# Spreadsheet initial setup

Follow these steps to get the Pokémon card sales operations spreadsheet ready after the Apps Script project has been deployed.

1. **Create a Google Sheet**
   - Name suggestion: `Pokémon Card Sales Ops`.
   - If you already have the sheet `1u8p0o9kCazC7_G45hmkxBSQFNnsEFvmArpuhfeADdg0`, open it instead.
   - Share it with the operations team and the Apps Script service account (so automated pushes can edit the sheet).

2. **Open the Apps Script editor**
   - `Extensions → Apps Script` from the spreadsheet.
   - Confirm that the files mirror the repository (`Config.gs`, `Customers.gs`, ...).

3. **Run the initialiser**
   - Back in the spreadsheet select `SalesOps → システム初期化`.
   - Grant the script the required scopes on the first execution.
   - The script creates and configures all sheets and seed data (lists, constants, counters).

4. **Review validation lists**
   - Open the `参照_リスト` sheet to confirm the dropdown values.
   - Adjust or append values to match the latest business rules.

5. **Lock formula columns (optional)**
   - If additional protection is required, you can extend the protected ranges created during initialisation using Google Sheets' UI.

6. **Test the workflow**
   - Run `SalesOps → 新規顧客作成` and create a dummy customer.
   - Run `SalesOps → 新規注文作成` to generate an order for the dummy customer.
   - Use `SalesOps → 明細貼り付け取り込み` with a sample TSV to populate the order lines.
   - Allocate inventory and mark the order as shipped to verify the onEdit/onOpen triggers.

Once the flow succeeds, the sheet is ready for production data entry.