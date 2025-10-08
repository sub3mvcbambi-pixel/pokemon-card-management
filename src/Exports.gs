function rebuildLookerExport() {
  const ordersSheet = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const ordersMap = getHeaderMap(ordersSheet);
  const ordersData = ordersSheet.getRange(2, 1, Math.max(ordersSheet.getLastRow() - 1, 0), ordersSheet.getLastColumn()).getValues();
  const orderMap = ordersData.reduce(function (acc, row) {
    const id = row[ordersMap['注文ID'] - 1];
    if (!id) {
      return acc;
    }
    acc[id] = row;
    return acc;
  }, {});

  const targetStatuses = getAggregateStatuses();
  const linesSheet = ensureSheet(CONFIG.sheets.orderLines, CONFIG.headers.orderLines);
  const linesMap = getHeaderMap(linesSheet);
  const linesData = linesSheet.getRange(2, 1, Math.max(linesSheet.getLastRow() - 1, 0), linesSheet.getLastColumn()).getValues();

  const exportSheet = ensureSheet(CONFIG.sheets.looker, CONFIG.headers.looker);
  exportSheet.clearContents();
  exportSheet.appendRow(CONFIG.headers.looker);

  const rows = [];
  linesData.forEach(function (line) {
    const orderId = line[linesMap['注文ID'] - 1];
    if (!orderId || !orderMap[orderId]) {
      return;
    }
    const orderRow = orderMap[orderId];
    const orderDate = parseDate(orderRow[ordersMap['注文日'] - 1]);
    const formattedDate = formatDate(orderDate);
    const status = orderRow[ordersMap['注文ステータス'] - 1];
    const fee = parseNumber(orderRow[ordersMap['決済手数料JPY'] - 1]);
    const chargedShipping = parseNumber(orderRow[ordersMap['請求送料JPY'] - 1]);
    const actualShipping = parseNumber(orderRow[ordersMap['実送料JPY'] - 1]);
    const aggregateFlag = targetStatuses.indexOf(status) !== -1 ? 1 : 0;

    const row = CONFIG.headers.looker.map(function (header) {
      switch (header) {
        case '受注日':
          return formattedDate;
        case '年':
          return orderDate ? orderDate.getFullYear() : '';
        case '月':
          return orderDate ? orderDate.getMonth() + 1 : '';
        case '週':
          if (!orderDate) {
            return '';
          }
          const firstJan = new Date(orderDate.getFullYear(), 0, 1);
          const week = Math.ceil(((orderDate - firstJan) / 86400000 + firstJan.getDay() + 1) / 7);
          return week;
        case '注文ID':
          return orderId;
        case '注文ステータス':
          return status;
        case '集計対象':
          return aggregateFlag;
        case '顧客ID':
          return orderRow[ordersMap['顧客ID'] - 1];
        case '国':
          const snapshot = orderRow[ordersMap['顧客ID'] - 1] ? getCustomerSnapshot(orderRow[ordersMap['顧客ID'] - 1]) : {};
          return snapshot['国'] || '';
        case '集客先':
          return snapshotFromOrder(orderRow, '集客先');
        case '顧客フラグ':
          return snapshotFromOrder(orderRow, '顧客フラグ');
        case '支払い方法':
          return orderRow[ordersMap['支払い方法'] - 1];
        case 'SKU':
          return line[linesMap['SKU'] - 1];
        case '商品名':
          return line[linesMap['商品名'] - 1];
        case '区分':
          return line[linesMap['区分'] - 1];
        case '数量':
          return line[linesMap['数量'] - 1];
        case '売上金額JPY':
          return line[linesMap['売上金額JPY'] - 1];
        case '原価_暫定JPY':
          return line[linesMap['原価_暫定JPY'] - 1];
        case '原価_還付後JPY':
          return line[linesMap['原価_還付後JPY'] - 1];
        case '行粗利_暫定JPY':
          return line[linesMap['行粗利_暫定JPY'] - 1];
        case '行粗利_還付後JPY':
          return line[linesMap['行粗利_還付後JPY'] - 1];
        case '決済手数料JPY':
          return fee;
        case '請求送料JPY':
          return chargedShipping;
        case '実送料JPY':
          return actualShipping;
        case '配送会社':
          return orderRow[ordersMap['配送会社'] - 1];
        case '還元率':
          return line[linesMap['還元率'] - 1];
        default:
          return '';
      }
    });
    rows.push(row);
  });

  if (rows.length) {
    exportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  recordAudit('rebuildLookerExport', '行数: ' + rows.length);
}

function getCustomerSnapshot(customerId) {
  const sheet = ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  const map = getHeaderMap(sheet);
  const rowNumber = findRowByValue(sheet, map['顧客ID'], customerId);
  if (rowNumber === -1) {
    return {};
  }
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    顧客ID: row[map['顧客ID'] - 1],
    国: row[map['国'] - 1],
    集客先: row[map['集客先'] - 1],
    顧客フラグ: row[map['顧客フラグ'] - 1]
  };
}

function snapshotFromOrder(orderRow, field) {
  const customerIdIndex = CONFIG.headers.orderHeader.indexOf('顧客ID');
  if (customerIdIndex === -1) {
    return '';
  }
  const customerId = orderRow[customerIdIndex];
  if (!customerId) {
    return '';
  }
  const snapshot = getCustomerSnapshot(customerId);
  return snapshot[field] || '';
}