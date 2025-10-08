function generateSku() {
  const prefix = coalesce(readSetting(CONFIG.idPrefixes.sku), CONFIG.defaultConstants.SKU接頭辞 || 'SKU');
  const next = incrementCounter(CONFIG.counterKeys.sku, 0);
  return formatId(prefix, next, 5);
}

function getSkuRow(sku) {
  const sheet = ensureSheet(CONFIG.sheets.sku, CONFIG.headers.sku);
  const map = getHeaderMap(sheet);
  const row = findRowByValue(sheet, map['SKU'], sku);
  if (row === -1) {
    throw new Error('SKUが見つかりません: ' + sku);
  }
  return { sheet: sheet, map: map, row: row };
}

function updateSkuQuantities(sku, deltaAvailable, deltaReserved) {
  const skuRow = getSkuRow(sku);
  const sheet = skuRow.sheet;
  const map = skuRow.map;
  const availableRange = sheet.getRange(skuRow.row, map['在庫数量_可用']);
  const reservedRange = sheet.getRange(skuRow.row, map['引当数量']);
  const available = parseNumber(availableRange.getValue()) + deltaAvailable;
  const reserved = parseNumber(reservedRange.getValue()) + deltaReserved;
  availableRange.setValue(available);
  reservedRange.setValue(reserved);
}

function createSku(payload) {
  const sheet = ensureSheet(CONFIG.sheets.sku, CONFIG.headers.sku);
  const id = payload.sku || generateSku();
  const row = CONFIG.headers.sku.map(function (header) {
    switch (header) {
      case 'SKU':
        return id;
      case '在庫数量_可用':
        return payload.available || 0;
      case '引当数量':
        return payload.reserved || 0;
      default:
        return payload[header] || '';
    }
  });
  sheet.appendRow(row);
  if (payload.available) {
    recordInventoryMovement(CONFIG.stockMovementTypes.inbound, id, payload.available, payload.unitCost || 0, 'Initial stock');
  }
  recordAudit('createSku', 'SKU: ' + id);
  return id;
}

function allocateInventory(orderId) {
  const lines = getOrderLines(orderId);
  const order = getOrderRow(orderId);
  const map = lines.map;
  let allocatedCount = 0;
  const errors = [];
  lines.lines.forEach(function (line) {
    const status = line.values[map['在庫引当ステータス'] - 1] || CONFIG.inventoryAllocationStatuses.none;
    if (status !== CONFIG.inventoryAllocationStatuses.none) {
      return;
    }
    let sku = line.values[map['SKU'] - 1];
    const productName = line.values[map['商品名'] - 1];
    const quantity = parseNumber(line.values[map['数量'] - 1]);
    if (!sku) {
      sku = findSkuByName(productName);
    }
    if (!sku) {
      errors.push('SKU未設定: 行No ' + line.values[map['行No'] - 1]);
      return;
    }
    try {
      const skuRow = getSkuRow(sku);
      const available = parseNumber(skuRow.sheet.getRange(skuRow.row, skuRow.map['在庫数量_可用']).getValue());
      if (available < quantity) {
        errors.push('在庫不足: SKU ' + sku + ' 必要数量 ' + quantity + ' / 在庫 ' + available);
        return;
      }
      updateSkuQuantities(sku, -quantity, quantity);
      const statusCell = lines.sheet.getRange(line.index, map['在庫引当ステータス']);
      statusCell.setValue(CONFIG.inventoryAllocationStatuses.allocated);
      if (!line.values[map['SKU'] - 1]) {
        lines.sheet.getRange(line.index, map['SKU']).setValue(sku);
      }
      recordInventoryMovement(CONFIG.stockMovementTypes.reserved, orderId, sku, quantity, line.values[map['原価_暫定JPY'] - 1], '引当');
      allocatedCount++;
    } catch (error) {
      errors.push(error.message);
    }
  });
  recordAudit('allocateInventory', '注文ID: ' + orderId + ' 引当数: ' + allocatedCount);
  return { allocated: allocatedCount, errors: errors };
}

function findSkuByName(name) {
  if (!name) {
    return null;
  }
  const sheet = ensureSheet(CONFIG.sheets.sku, CONFIG.headers.sku);
  const map = getHeaderMap(sheet);
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][map['商品名'] - 1] === name) {
      return data[i][map['SKU'] - 1];
    }
  }
  return null;
}

function shipOrder(orderId, carrier, tracking, shipDate) {
  const order = getOrderRow(orderId);
  const lines = getOrderLines(orderId);
  const map = lines.map;
  const date = shipDate ? formatDate(shipDate) : today();
  const sheet = order.sheet;
  if (carrier) {
    sheet.getRange(order.row, order.map['配送会社']).setValue(carrier);
  }
  if (tracking) {
    sheet.getRange(order.row, order.map['追跡番号']).setValue(tracking);
  }
  sheet.getRange(order.row, order.map['発送日']).setValue(date);
  updateOrderStatus(orderId, '発送済', { silent: true });

  lines.lines.forEach(function (line) {
    const status = line.values[map['在庫引当ステータス'] - 1];
    if (status === CONFIG.inventoryAllocationStatuses.allocated) {
      const sku = line.values[map['SKU'] - 1];
      const quantity = parseNumber(line.values[map['数量'] - 1]);
      updateSkuQuantities(sku, 0, -quantity);
      lines.sheet.getRange(line.index, map['在庫引当ステータス']).setValue(CONFIG.inventoryAllocationStatuses.shipped);
      recordInventoryMovement(CONFIG.stockMovementTypes.outbound, orderId, sku, quantity, line.values[map['原価_暫定JPY'] - 1], '出庫');
    }
  });
  recordAudit('shipOrder', '注文ID: ' + orderId);
}

function recordInventoryMovement(type, reference, sku, quantity, unitCost, memo) {
  const sheet = ensureSheet(CONFIG.sheets.stockLedger, CONFIG.headers.stockLedger);
  sheet.appendRow([
    type,
    reference,
    today(),
    sku,
    quantity,
    unitCost,
    memo || ''
  ]);
}