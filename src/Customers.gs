function generateCustomerId() {
  const prefix = coalesce(readSetting(CONFIG.idPrefixes.customer), CONFIG.defaultConstants.顧客ID接頭辞);
  const digits = Number(coalesce(readSetting(CONFIG.idPrefixes.customerDigits), CONFIG.defaultConstants.顧客ID桁数));
  const next = incrementCounter(CONFIG.counterKeys.customer, 0);
  return formatId(prefix, next, digits);
}

function createCustomer(payload) {
  const sheet = ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  const map = getHeaderMap(sheet);
  const id = payload.customerId || generateCustomerId();
  const name = payload.name || payload['氏名'] || '';
  const rowValues = [];
  CONFIG.headers.customers.forEach(function (header) {
    switch (header) {
      case '顧客ID':
        rowValues.push(id);
        break;
      case '氏名':
        rowValues.push(name);
        break;
      case '購入回数':
      case '購入金額累計JPY':
        rowValues.push(0);
        break;
      default:
        if (payload.hasOwnProperty(header)) {
          rowValues.push(payload[header]);
        } else {
          rowValues.push('');
        }
    }
  });
  sheet.appendRow(rowValues);
  recordAudit('createCustomer', '顧客ID: ' + id);
  return { customerId: id, row: sheet.getLastRow() };
}

function getCustomerRow(customerId) {
  const sheet = ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  const map = getHeaderMap(sheet);
  const row = findRowByValue(sheet, map['顧客ID'], customerId);
  if (row === -1) {
    throw new Error('顧客IDが見つかりません: ' + customerId);
  }
  return { sheet: sheet, map: map, row: row };
}

function refreshCustomerAggregates() {
  const customersSheet = ensureSheet(CONFIG.sheets.customers, CONFIG.headers.customers);
  const customerMap = getHeaderMap(customersSheet);
  const ordersSheet = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const ordersMap = getHeaderMap(ordersSheet);
  const targetStatuses = getAggregateStatuses();

  const orderData = ordersSheet
    .getRange(2, 1, Math.max(ordersSheet.getLastRow() - 1, 0), ordersSheet.getLastColumn())
    .getValues()
    .filter(function (row) {
      const status = row[ordersMap['注文ステータス'] - 1];
      return status && targetStatuses.indexOf(status) !== -1;
    })
    .reduce(function (acc, row) {
      const customerId = row[ordersMap['顧客ID'] - 1];
      if (!customerId) {
        return acc;
      }
      const orderDate = parseDate(row[ordersMap['注文日'] - 1]);
      const sales = parseNumber(row[ordersMap['売上合計JPY'] - 1]);
      if (!acc[customerId]) {
        acc[customerId] = {
          count: 0,
          total: 0,
          first: null,
          last: null
        };
      }
      const bucket = acc[customerId];
      bucket.count += 1;
      bucket.total += sales;
      if (!bucket.first || orderDate < bucket.first) {
        bucket.first = orderDate;
      }
      if (!bucket.last || orderDate > bucket.last) {
        bucket.last = orderDate;
      }
      return acc;
    }, {});

  const customerRange = customersSheet.getRange(2, 1, Math.max(customersSheet.getLastRow() - 1, 0), customersSheet.getLastColumn());
  const values = customerRange.getValues();
  values.forEach(function (row) {
    const customerId = row[customerMap['顧客ID'] - 1];
    const aggregates = orderData[customerId];
    if (!aggregates) {
      row[customerMap['購入回数'] - 1] = 0;
      row[customerMap['購入金額累計JPY'] - 1] = 0;
      row[customerMap['初回購入日'] - 1] = '';
      row[customerMap['最終購入日'] - 1] = '';
    } else {
      row[customerMap['購入回数'] - 1] = aggregates.count;
      row[customerMap['購入金額累計JPY'] - 1] = aggregates.total;
      row[customerMap['初回購入日'] - 1] = formatDate(aggregates.first);
      row[customerMap['最終購入日'] - 1] = formatDate(aggregates.last);
    }
  });
  customerRange.setValues(values);
  recordAudit('refreshCustomerAggregates', '顧客集計を更新しました');
}