function generateOrderId(customerId, dateString) {
  if (!customerId) {
    throw new Error('顧客IDが必要です');
  }
  const date = dateString ? parseDate(dateString) : new Date();
  if (!date) {
    throw new Error('注文日の形式が不正です');
  }
  const formattedDate = Utilities.formatDate(date, CONFIG.timezone, 'yyyyMMdd');
  const type = coalesce(readSetting(CONFIG.idPrefixes.orderType), CONFIG.defaultConstants.注文ID_既定種別);
  const counterKey = CONFIG.counterKeys.order + '_' + customerId + '_' + formattedDate;
  const sequence = incrementCounter(counterKey, 0);
  const seqFormatted = Utilities.formatString('%02d', sequence);
  return [customerId, type, formattedDate, seqFormatted].join('-');
}

function createOrderHeader(customerId, orderDate) {
  const customer = getCustomerRow(customerId);
  const orders = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const map = getHeaderMap(orders);
  const id = generateOrderId(customerId, orderDate);
  const date = orderDate ? formatDate(orderDate) : today();

  const row = CONFIG.headers.orderHeader.map(function (header) {
    switch (header) {
      case '注文ID':
        return id;
      case '顧客ID':
        return customerId;
      case '顧客名':
        return customer.sheet.getRange(customer.row, customer.map['氏名']).getValue();
      case '注文日':
        return date;
      case '注文ステータス':
        return CONFIG.orderStatusFlow[0];
      case '売上合計JPY':
      case '原価合計_暫定JPY':
      case '原価合計_還付後JPY':
      case '粗利_暫定JPY':
      case '粗利_還付後JPY':
      case '決済手数料JPY':
      case '請求送料JPY':
      case '実送料JPY':
        return 0;
      default:
        return '';
    }
  });
  orders.appendRow(row);
  recordAudit('createOrderHeader', '注文ID: ' + id);
  return id;
}

function getOrderRow(orderId) {
  const sheet = ensureSheet(CONFIG.sheets.orderHeader, CONFIG.headers.orderHeader);
  const map = getHeaderMap(sheet);
  const row = findRowByValue(sheet, map['注文ID'], orderId);
  if (row === -1) {
    throw new Error('注文IDが見つかりません: ' + orderId);
  }
  return { sheet: sheet, map: map, row: row };
}

function getOrderLines(orderId) {
  const sheet = ensureSheet(CONFIG.sheets.orderLines, CONFIG.headers.orderLines);
  const map = getHeaderMap(sheet);
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).getValues();
  const lines = [];
  data.forEach(function (row, index) {
    if (row[map['注文ID'] - 1] === orderId) {
      lines.push({ index: index + 2, values: row });
    }
  });
  return { sheet: sheet, map: map, lines: lines };
}

function importOrderLines(orderId, rawData, defaults) {
  defaults = defaults || {};
  const order = getOrderRow(orderId);
  const parsed = parseDelimited(rawData);
  if (!parsed.length) {
    return { added: 0, errors: ['入力データが空です'] };
  }
  const linesSheet = ensureSheet(CONFIG.sheets.orderLines, CONFIG.headers.orderLines);
  const map = getHeaderMap(linesSheet);
  let headerDetected = false;
  const knownHeaders = CONFIG.headers.orderLines.reduce(function (acc, header) {
    acc[header] = true;
    return acc;
  }, {});

  if (parsed.length && parsed[0].some(function (cell) { return knownHeaders[cell]; })) {
    headerDetected = true;
  }
  const headerIndex = headerDetected
    ? parsed[0].reduce(function (acc, value, idx) {
        acc[value] = idx;
        return acc;
      }, {})
    : {};

  const rowsToAppend = [];
  const errors = [];
  const startLineNumber = getOrderLines(orderId).lines.length + 1;
  parsed.forEach(function (cols, idx) {
    if (headerDetected && idx === 0) {
      return;
    }
    try {
      const lineNumber = rowsToAppend.length + startLineNumber;
      const quantity = parseNumber(resolveValue(cols, headerIndex, '数量', [4, 2, 3]));
      const price = parseNumber(resolveValue(cols, headerIndex, '単価_売価JPY', [5, 3, 4]));
      const discount = parseNumber(resolveValue(cols, headerIndex, '行割引JPY', [6]));
      const cost = parseNumber(coalesce(resolveValue(cols, headerIndex, '原価_暫定JPY', [7]), defaults.cost || 0));
      const rateRaw = coalesce(resolveValue(cols, headerIndex, '還元率', [8]), defaults.rate || '0%');
      const rate = normaliseRate(rateRaw);
      const revenue = quantity * price - discount;
      const rebate = cost * rate;
      const netCost = cost - rebate;
      const grossMargin = revenue - cost;
      const netMargin = revenue - netCost;
      const sku = resolveValue(cols, headerIndex, 'SKU', [0]) || '';
      const productName = resolveValue(cols, headerIndex, '商品名', [1]) || '';
      const category = resolveValue(cols, headerIndex, '区分', [2]) || defaults.category || '';

      const row = CONFIG.headers.orderLines.map(function (header) {
        switch (header) {
          case '注文ID':
            return orderId;
          case '行No':
            return lineNumber;
          case 'SKU':
            return sku;
          case '商品名':
            return productName;
          case '区分':
            return category;
          case '数量':
            return quantity;
          case '単価_売価JPY':
            return price;
          case '行割引JPY':
            return discount;
          case '売上金額JPY':
            return revenue;
          case '原価_暫定JPY':
            return cost;
          case '還元率':
            return Utilities.formatString('%.2f%%', rate * 100);
          case '還付額JPY':
            return rebate;
          case '原価_還付後JPY':
            return netCost;
          case '行粗利_暫定JPY':
            return grossMargin;
          case '行粗利_還付後JPY':
            return netMargin;
          case '在庫引当ステータス':
            return CONFIG.inventoryAllocationStatuses.none;
          default:
            return '';
        }
      });
      rowsToAppend.push(row);
    } catch (error) {
      errors.push('行 ' + (idx + 1) + ': ' + error.message);
    }
  });

  rowsToAppend.forEach(function (row) {
    linesSheet.appendRow(row);
  });
  if (rowsToAppend.length) {
    recalculateOrderTotals(orderId);
  }
  recordAudit('importOrderLines', '注文ID: ' + orderId + ' 行数: ' + rowsToAppend.length);
  return { added: rowsToAppend.length, errors: errors };
}

function resolveValue(columns, headerIndex, headerName, fallbackIndexes) {
  if (headerIndex && headerIndex.hasOwnProperty(headerName)) {
    const idx = headerIndex[headerName];
    if (idx !== undefined && idx < columns.length) {
      return columns[idx];
    }
  }
  if (fallbackIndexes) {
    for (let i = 0; i < fallbackIndexes.length; i++) {
      const idx = fallbackIndexes[i];
      if (idx !== undefined && idx < columns.length && columns[idx] !== undefined) {
        return columns[idx];
      }
    }
  }
  return '';
}

function recalculateOrderTotals(orderId) {
  const order = getOrderRow(orderId);
  const lines = getOrderLines(orderId);
  const totals = {
    revenue: 0,
    cost: 0,
    costNet: 0,
    grossMargin: 0,
    netMargin: 0
  };
  lines.lines.forEach(function (entry) {
    const row = entry.values;
    totals.revenue += parseNumber(row[lines.map['売上金額JPY'] - 1]);
    totals.cost += parseNumber(row[lines.map['原価_暫定JPY'] - 1]);
    totals.costNet += parseNumber(row[lines.map['原価_還付後JPY'] - 1]);
    totals.grossMargin += parseNumber(row[lines.map['行粗利_暫定JPY'] - 1]);
    totals.netMargin += parseNumber(row[lines.map['行粗利_還付後JPY'] - 1]);
  });

  const sheet = order.sheet;
  sheet.getRange(order.row, order.map['売上合計JPY']).setValue(totals.revenue);
  sheet.getRange(order.row, order.map['原価合計_暫定JPY']).setValue(totals.cost);
  sheet.getRange(order.row, order.map['原価合計_還付後JPY']).setValue(totals.costNet);
  const fee = parseNumber(sheet.getRange(order.row, order.map['決済手数料JPY']).getValue());
  const shippingActual = parseNumber(sheet.getRange(order.row, order.map['実送料JPY']).getValue());
  const shippingCharged = parseNumber(sheet.getRange(order.row, order.map['請求送料JPY']).getValue());
  const gross = totals.revenue - totals.cost - fee - shippingActual + shippingCharged;
  const net = totals.revenue - totals.costNet - fee - shippingActual + shippingCharged;
  sheet.getRange(order.row, order.map['粗利_暫定JPY']).setValue(gross);
  sheet.getRange(order.row, order.map['粗利_還付後JPY']).setValue(net);
}

function updateOrderStatus(orderId, status, options) {
  if (!isStatusAllowed(status)) {
    throw new Error('不正なステータスです: ' + status);
  }
  const order = getOrderRow(orderId);
  const sheet = order.sheet;
  sheet.getRange(order.row, order.map['注文ステータス']).setValue(status);
  const todayStr = today();
  switch (status) {
    case '支払い完了':
      if (!sheet.getRange(order.row, order.map['入金日']).getValue()) {
        sheet.getRange(order.row, order.map['入金日']).setValue(todayStr);
      }
      break;
    case '仕入れ完了':
      break;
    case '発送済':
      if (!sheet.getRange(order.row, order.map['発送日']).getValue()) {
        sheet.getRange(order.row, order.map['発送日']).setValue(todayStr);
      }
      break;
    case '追跡番号送付済':
      if (!sheet.getRange(order.row, order.map['追跡番号送付日']).getValue()) {
        sheet.getRange(order.row, order.map['追跡番号送付日']).setValue(todayStr);
      }
      break;
    case '到着':
      if (!sheet.getRange(order.row, order.map['到着日']).getValue()) {
        sheet.getRange(order.row, order.map['到着日']).setValue(todayStr);
      }
      break;
    default:
      break;
  }
  if (!options || !options.silent) {
    recordAudit('updateOrderStatus', '注文ID: ' + orderId + ' → ' + status);
  }
}