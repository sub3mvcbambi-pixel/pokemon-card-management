function importPaypalCsv(raw) {
  const sheet = ensureSheet(CONFIG.sheets.paypal, CONFIG.headers.paypal);
  const map = getHeaderMap(sheet);
  const existingIds = new Set();
  const idCol = map['Transaction ID'];
  if (sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function (row) {
      if (row[0]) {
        existingIds.add(String(row[0]));
      }
    });
  }
  const parsed = parseDelimited(raw);
  if (!parsed.length) {
    return 0;
  }
  const header = parsed[0];
  const hasHeader = header.includes('Transaction ID');
  const headerMap = hasHeader
    ? header.reduce(function (acc, value, idx) {
        acc[value] = idx;
        return acc;
      }, {})
    : {};
  const startIndex = hasHeader ? 1 : 0;
  let inserted = 0;
  for (let i = startIndex; i < parsed.length; i++) {
    const row = parsed[i];
    const txnIndex = hasHeader ? headerMap['Transaction ID'] : 10;
    const txnId = txnIndex !== undefined ? row[txnIndex] : null;
    if (!txnId || existingIds.has(String(txnId))) {
      continue;
    }
    const values = CONFIG.headers.paypal.map(function (col) {
      const idx = hasHeader ? headerMap[col] : CONFIG.headers.paypal.indexOf(col);
      return idx !== -1 && idx < row.length ? row[idx] : '';
    });
    sheet.appendRow(values);
    existingIds.add(String(txnId));
    inserted++;
  }
  recordAudit('importPaypalCsv', '行数: ' + inserted);
  return inserted;
}

function importWiseCsv(raw) {
  const sheet = ensureSheet(CONFIG.sheets.wise, CONFIG.headers.wise);
  const map = getHeaderMap(sheet);
  const existingIds = new Set();
  const idCol = map['Transfer ID'];
  if (sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function (row) {
      if (row[0]) {
        existingIds.add(String(row[0]));
      }
    });
  }
  const parsed = parseDelimited(raw);
  if (!parsed.length) {
    return 0;
  }
  const header = parsed[0];
  const hasHeader = header.includes('Transfer ID');
  const headerMap = hasHeader
    ? header.reduce(function (acc, value, idx) {
        acc[value] = idx;
        return acc;
      }, {})
    : {};
  const startIndex = hasHeader ? 1 : 0;
  let inserted = 0;
  for (let i = startIndex; i < parsed.length; i++) {
    const row = parsed[i];
    const txnIndex = hasHeader ? headerMap['Transfer ID'] : 6;
    const txnId = txnIndex !== undefined ? row[txnIndex] : null;
    if (!txnId || existingIds.has(String(txnId))) {
      continue;
    }
    const values = CONFIG.headers.wise.map(function (col) {
      const idx = hasHeader ? headerMap[col] : CONFIG.headers.wise.indexOf(col);
      return idx !== -1 && idx < row.length ? row[idx] : '';
    });
    sheet.appendRow(values);
    existingIds.add(String(txnId));
    inserted++;
  }
  recordAudit('importWiseCsv', '行数: ' + inserted);
  return inserted;
}

function matchPaymentsToOrders() {
  const payments = loadPayments();
  const orders = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const map = getHeaderMap(orders);
  let matched = 0;
  const unmatched = [];
  const dataRange = orders.getRange(2, 1, Math.max(orders.getLastRow() - 1, 0), orders.getLastColumn());
  const values = dataRange.getValues();

  values.forEach(function (row, index) {
    const orderId = row[map['注文ID'] - 1];
    if (!orderId) {
      return;
    }
    const reference = row[map['決済参照ID'] - 1];
    const amount = parseNumber(row[map['売上合計JPY'] - 1]) + parseNumber(row[map['請求送料JPY'] - 1]);
    const orderDate = parseDate(row[map['注文日'] - 1]);
    let payment = null;
    if (reference) {
      payment = payments.find(function (p) {
        return p.id === reference;
      });
    }
    if (!payment && orderDate) {
      payment = payments.find(function (p) {
        if (!p.date) {
          return false;
        }
        const diff = Math.abs((p.date - orderDate) / (1000 * 60 * 60 * 24));
        return diff <= 3 && Math.round(p.net) === Math.round(amount);
      });
    }
    if (payment) {
      applyPaymentToOrder(orderId, payment);
      matched++;
    } else {
      unmatched.push(orderId);
    }
  });
  recordAudit('matchPaymentsToOrders', 'matched: ' + matched + ', unmatched: ' + unmatched.join(','));
  return { matched: matched, unmatched: unmatched };
}

function loadPayments() {
  const payments = [];
  const paypalSheet = ensureSheet(CONFIG.sheets.paypal, CONFIG.headers.paypal);
  if (paypalSheet.getLastRow() > 1) {
    const map = getHeaderMap(paypalSheet);
    const rows = paypalSheet.getRange(2, 1, paypalSheet.getLastRow() - 1, paypalSheet.getLastColumn()).getValues();
    rows.forEach(function (row) {
      const currency = row[map['Currency'] - 1];
      if (currency && currency !== 'JPY') {
        return;
      }
      const fee = parseNumber(row[map['Fee'] - 1]);
      const net = parseNumber(row[map['Net'] - 1]);
      payments.push({
        id: row[map['Transaction ID'] - 1],
        date: parseDate(row[map['Date'] - 1]),
        fee: fee,
        net: net,
        source: 'PayPal'
      });
    });
  }
  const wiseSheet = ensureSheet(CONFIG.sheets.wise, CONFIG.headers.wise);
  if (wiseSheet.getLastRow() > 1) {
    const map = getHeaderMap(wiseSheet);
    const rows = wiseSheet.getRange(2, 1, wiseSheet.getLastRow() - 1, wiseSheet.getLastColumn()).getValues();
    rows.forEach(function (row) {
      const currency = row[map['Currency'] - 1];
      if (currency && currency !== 'JPY') {
        return;
      }
      payments.push({
        id: row[map['Transfer ID'] - 1],
        date: parseDate(row[map['Date'] - 1]),
        fee: parseNumber(row[map['Fee'] - 1]),
        net: parseNumber(row[map['Total'] - 1]),
        source: 'Wise'
      });
    });
  }
  return payments;
}

function applyPaymentToOrder(orderId, payment) {
  const order = getOrderRow(orderId);
  const sheet = order.sheet;
  sheet.getRange(order.row, order.map['決済参照ID']).setValue(payment.id);
  sheet.getRange(order.row, order.map['決済手数料JPY']).setValue(Math.abs(payment.fee));
  sheet.getRange(order.row, order.map['入金日']).setValue(formatDate(payment.date) || today());
  if (!sheet.getRange(order.row, order.map['注文ステータス']).getValue()) {
    sheet.getRange(order.row, order.map['注文ステータス']).setValue('支払い完了');
  }
  recalculateOrderTotals(orderId);
}