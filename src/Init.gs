function initializeSystem() {
  SpreadsheetApp.getActive().toast('システム初期化を開始します...', 'SalesOps', 5);
  ensureSheet(CONFIG.sheets.settings, CONFIG.headers.settings);
  seedConstants();
  ensureSheet(CONFIG.sheets.lists, CONFIG.headers.lists);
  seedReferenceLists();
  ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  ensureSheet(CONFIG.sheets.orderLines, CONFIG.headers.orderLines);
  ensureSheet(CONFIG.sheets.sku, CONFIG.headers.sku);
  ensureSheet(CONFIG.sheets.stockLedger, CONFIG.headers.stockLedger);
  ensureSheet(CONFIG.sheets.paypal, CONFIG.headers.paypal);
  ensureSheet(CONFIG.sheets.wise, CONFIG.headers.wise);
  ensureSheet(CONFIG.sheets.looker, CONFIG.headers.looker);
  ensureSheet(CONFIG.sheets.audit, CONFIG.headers.audit);
  if (SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.dashboard) === null) {
    SpreadsheetApp.getActive().insertSheet(CONFIG.sheets.dashboard);
  }

  configureValidations();
  recordAudit('initializeSystem', 'システム初期化完了');
  SpreadsheetApp.getActive().toast('システム初期化が完了しました。', 'SalesOps', 5);
}

function seedConstants() {
  Object.keys(CONFIG.defaultConstants).forEach(function (key) {
    const existing = readSetting(key);
    if (existing === null || existing === undefined || existing === '') {
      writeSetting(key, CONFIG.defaultConstants[key], 'Initial default');
    }
  });
  if (!readSetting(CONFIG.counterKeys.customer)) {
    writeSetting(CONFIG.counterKeys.customer, 0, '顧客ID連番');
  }
  if (!readSetting(CONFIG.counterKeys.sku)) {
    writeSetting(CONFIG.counterKeys.sku, 0, 'SKU連番');
  }
  if (!readSetting(CONFIG.counterKeys.order)) {
    writeSetting(CONFIG.counterKeys.order, 0, '注文通番（顧客日別）');
  }
}

function seedReferenceLists() {
  const sheet = ensureSheet(CONFIG.sheets.lists, CONFIG.headers.lists);
  const map = getHeaderMap(sheet);
  const categoryCol = requireColumn(map, 'カテゴリ');
  const valueCol = requireColumn(map, '値');
  const existing = sheet
    .getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn())
    .getValues()
    .reduce(function (acc, row) {
      const key = row[categoryCol - 1] + '::' + row[valueCol - 1];
      if (row[categoryCol - 1] && row[valueCol - 1]) {
        acc[key] = true;
      }
      return acc;
    }, {});

  Object.keys(CONFIG.listSeeds).forEach(function (category) {
    CONFIG.listSeeds[category].forEach(function (value) {
      const key = category + '::' + value;
      if (!existing[key]) {
        sheet.appendRow([category, value]);
      }
    });
  });
}

function configureValidations() {
  const orders = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const orderMap = getHeaderMap(orders);
  addValidation(
    orders,
    orders.getRange(2, orderMap['注文ステータス'], orders.getMaxRows() - 1, 1).getA1Notation(),
    getListValues('注文ステータス')
  );
  addValidation(
    orders,
    orders.getRange(2, orderMap['支払い方法'], orders.getMaxRows() - 1, 1).getA1Notation(),
    getListValues('支払い方法')
  );

  const customers = ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  const customerMap = getHeaderMap(customers);
  addValidation(
    customers,
    customers.getRange(2, customerMap['集客先'], customers.getMaxRows() - 1, 1).getA1Notation(),
    getListValues('集客先')
  );
  addValidation(
    customers,
    customers.getRange(2, customerMap['顧客フラグ'], customers.getMaxRows() - 1, 1).getA1Notation(),
    getListValues('顧客フラグ')
  );

  const skuSheet = ensureSheet(CONFIG.sheets.sku, CONFIG.headers.sku);
  const skuMap = getHeaderMap(skuSheet);
  addValidation(
    skuSheet,
    skuSheet.getRange(2, skuMap['在庫ステータス'], skuSheet.getMaxRows() - 1, 1).getA1Notation(),
    getListValues('在庫ステータス')
  );
  addValidation(
    skuSheet,
    skuSheet.getRange(2, skuMap['還元率'], skuSheet.getMaxRows() - 1, 1).getA1Notation(),
    CONFIG.listSeeds.還元率
  );
}