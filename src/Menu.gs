function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('SalesOps');
  menu.addItem('システム初期化', 'initializeSystem');
  menu.addSeparator();
  menu.addItem('新規顧客作成', 'menuCreateCustomer');
  menu.addItem('新規注文作成', 'menuCreateOrder');
  menu.addItem('明細貼り付け取り込み', 'menuImportOrderLines');
  menu.addItem('在庫引当', 'menuAllocateInventory');
  menu.addItem('発送処理', 'menuShipOrder');
  menu.addItem('注文ステータス更新', 'menuUpdateOrderStatus');
  menu.addSeparator();
  menu.addItem('顧客集計更新', 'refreshCustomerAggregates');
  menu.addItem('Export Looker 再生成', 'rebuildLookerExport');
  menu.addItem('バックアップ作成', 'createSpreadsheetBackup');
  menu.addToUi();
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }
  const sheet = e.range.getSheet();
  if (sheet.getName() === CONFIG.sheets.orderHeader) {
    const map = getHeaderMap(sheet);
    const statusCol = requireColumn(map, '注文ステータス');
    if (e.range.getColumn() === statusCol && e.value) {
      updateOrderStatus(String(sheet.getRange(e.range.getRow(), requireColumn(map, '注文ID')).getValue()), e.value, {
        silent: true
      });
    }
    const paymentRefCol = requireColumn(map, '決済参照ID');
    if (e.range.getColumn() === paymentRefCol && e.value) {
      matchPaymentsToOrders();
    }
  }
}

function menuCreateCustomer() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('新規顧客', '顧客名を入力してください', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const name = response.getResponseText();
  const row = createCustomer({ name: name });
  ui.alert('顧客を作成しました: ' + row.customerId);
}

function menuCreateOrder() {
  const ui = SpreadsheetApp.getUi();
  const customerId = ui.prompt('新規注文', '顧客IDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (customerId.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const dateResponse = ui.prompt('新規注文', '注文日 (YYYY-MM-DD) を入力してください。空欄の場合は本日になります。', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const orderDate = dateResponse.getResponseText();
  const orderId = createOrderHeader(customerId.getResponseText(), orderDate);
  ui.alert('注文を作成しました: ' + orderId);
}

function menuImportOrderLines() {
  const ui = SpreadsheetApp.getUi();
  const orderIdResponse = ui.prompt('明細取り込み', '注文IDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const orderId = orderIdResponse.getResponseText();
  const dataResponse = ui.prompt('明細取り込み', 'TSVまたはCSV形式で明細を貼り付けてください。', ui.ButtonSet.OK_CANCEL);
  if (dataResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const result = importOrderLines(orderId, dataResponse.getResponseText(), {});
  ui.alert('明細を取り込みました。追加: ' + result.added + '\nエラー: ' + result.errors.join('\n'));
}

function menuAllocateInventory() {
  const ui = SpreadsheetApp.getUi();
  const orderIdResponse = ui.prompt('在庫引当', '注文IDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const result = allocateInventory(orderIdResponse.getResponseText());
  ui.alert('在庫引当完了: ' + JSON.stringify(result));
}

function menuShipOrder() {
  const ui = SpreadsheetApp.getUi();
  const orderIdResponse = ui.prompt('発送処理', '注文IDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const carrier = ui.prompt('発送処理', '配送会社を入力してください', ui.ButtonSet.OK_CANCEL);
  if (carrier.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const tracking = ui.prompt('発送処理', '追跡番号を入力してください', ui.ButtonSet.OK_CANCEL);
  if (tracking.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const shipDate = ui.prompt('発送処理', '発送日 (YYYY-MM-DD)。空欄の場合は本日。', ui.ButtonSet.OK_CANCEL);
  if (shipDate.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  shipOrder(orderIdResponse.getResponseText(), carrier.getResponseText(), tracking.getResponseText(), shipDate.getResponseText());
  ui.alert('発送処理が完了しました');
}

function menuUpdateOrderStatus() {
  const ui = SpreadsheetApp.getUi();
  const orderIdResponse = ui.prompt('注文ステータス更新', '注文IDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const statusResponse = ui.prompt('注文ステータス更新', '新しいステータスを入力してください', ui.ButtonSet.OK_CANCEL);
  if (statusResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  updateOrderStatus(orderIdResponse.getResponseText(), statusResponse.getResponseText());
  ui.alert('ステータスを更新しました');
}