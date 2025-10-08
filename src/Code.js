/**
 * ポケモンカード販売管理システム
 * メイン処理とユーティリティ関数
 */

// グローバル定数
const CONFIG = {
  CUSTOMER_ID_PREFIX: 'PK',
  CUSTOMER_ID_DIGITS: 2,
  ORDER_TYPE: 'ASK',
  TIMEZONE: 'Asia/Tokyo',
  SALES_TARGET_STATUSES: ['支払い完了', '仕入れ完了', '発送済', '追跡番号送付済', '到着']
};

// シート名定数
const SHEET_NAMES = {
  CONFIG: '設定_定数',
  REFERENCE: '参照_リスト',
  CUSTOMER: 'マスター_顧客',
  ORDER_HEADER: '取引_注文ヘッダ',
  ORDER_DETAIL: '取引_注文明細',
  SKU: '在庫_SKU',
  INVENTORY: '在庫_入出庫',
  PAYPAL: '決済_取込_PayPal',
  WISE: '決済_取込_Wise',
  DASHBOARD: '集計_ダッシュボード',
  LOOKER_EXPORT: 'Export_Looker',
  AUDIT_LOG: '監査ログ_操作'
};

/**
 * スプレッドシート初期化
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SalesOps')
    .addItem('新規顧客作成', 'createNewCustomer')
    .addItem('新規注文作成', 'createNewOrder')
    .addItem('明細取り込み', 'importOrderLines')
    .addItem('在庫引当', 'allocateInventory')
    .addItem('発送処理', 'shipOrder')
    .addItem('注文ステータス更新', 'updateOrderStatus')
    .addItem('決済CSV取込', 'importPaymentCsv')
    .addItem('決済紐付け', 'matchPaymentsToOrders')
    .addItem('顧客集計更新', 'refreshCustomerAggregates')
    .addItem('Looker Export更新', 'rebuildLookerExport')
    .addSeparator()
    .addItem('バックアップ作成', 'createBackup')
    .addToUi();
}

/**
 * 顧客ID生成
 */
function generateCustomerId() {
  const sheet = getSheet(SHEET_NAMES.CONFIG);
  const lastId = sheet.getRange('B1').getValue() || 0; // 連番カウンタ
  const newId = lastId + 1;
  const customerId = CONFIG.CUSTOMER_ID_PREFIX + String(newId).padStart(CONFIG.CUSTOMER_ID_DIGITS, '0');
  
  // カウンタ更新
  sheet.getRange('B1').setValue(newId);
  
  return customerId;
}

/**
 * 注文ID生成
 */
function generateOrderId(customerId, date) {
  const orderDate = Utilities.formatDate(date || new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
  const sheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const data = sheet.getDataRange().getValues();
  
  // 該当顧客の該当日付の既存注文数をカウント
  const existingOrders = data.filter(row => 
    row[0] && row[0].toString().startsWith(`${customerId}-${CONFIG.ORDER_TYPE}-${orderDate}`)
  ).length;
  
  const sequence = String(existingOrders + 1).padStart(2, '0');
  return `${customerId}-${CONFIG.ORDER_TYPE}-${orderDate}-${sequence}`;
}

/**
 * SKU生成
 */
function generateSku() {
  const sheet = getSheet(SHEET_NAMES.SKU);
  const data = sheet.getDataRange().getValues();
  const lastSku = data[data.length - 1] && data[data.length - 1][0];
  
  if (!lastSku || !lastSku.toString().startsWith('SKU')) {
    return 'SKU00001';
  }
  
  const lastNumber = parseInt(lastSku.toString().replace('SKU', ''));
  const newNumber = lastNumber + 1;
  return 'SKU' + String(newNumber).padStart(5, '0');
}

/**
 * 新規顧客作成
 */
function createNewCustomer() {
  const ui = SpreadsheetApp.getUi();
  const customerId = generateCustomerId();
  
  const result = ui.prompt('新規顧客作成', `顧客ID: ${customerId}\n\n氏名を入力してください:`, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const name = result.getResponseText();
    if (name) {
      const sheet = getSheet(SHEET_NAMES.CUSTOMER);
      const newRow = [
        customerId,
        name,
        '', // メール
        '', // 電話
        '', // 住所1
        '', // 住所2
        '', // 郵便番号
        '', // 国
        '', // 集客先
        '', // 顧客フラグ
        false, // 興味_BOX
        false, // 興味_PSA
        false, // 興味_シングル
        false, // 興味_バルク
        false, // 支払_Wise
        false, // 支払_PayPal
        false, // PII同意
        '', // 同意日
        false, // 削除依頼
        '', // 削除対応日
        false, // 連絡不可
        '', // メモ
        0, // 購入回数
        0, // 購入金額累計
        '', // 初回購入日
        '' // 最終購入日
      ];
      
      sheet.appendRow(newRow);
      ui.alert('顧客作成完了', `顧客ID: ${customerId}\n氏名: ${name}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 新規注文作成
 */
function createNewOrder() {
  const ui = SpreadsheetApp.getUi();
  const customerSheet = getSheet(SHEET_NAMES.CUSTOMER);
  const customers = customerSheet.getDataRange().getValues().slice(1); // ヘッダー除外
  
  if (customers.length === 0) {
    ui.alert('エラー', '顧客が登録されていません。先に顧客を作成してください。', ui.ButtonSet.OK);
    return;
  }
  
  // 顧客選択
  let customerList = '顧客を選択してください:\n';
  customers.forEach((customer, index) => {
    customerList += `${index + 1}. ${customer[0]} - ${customer[1]}\n`;
  });
  
  const result = ui.prompt('新規注文作成', customerList, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText());
    if (selection >= 1 && selection <= customers.length) {
      const selectedCustomer = customers[selection - 1];
      const customerId = selectedCustomer[0];
      const customerName = selectedCustomer[1];
      
      const orderId = generateOrderId(customerId, new Date());
      
      // 注文ヘッダー作成
      const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
      const newOrderRow = [
        orderId,
        customerId,
        customerName,
        new Date(), // 注文日
        '', // 支払い方法
        '支払い前', // 注文ステータス
        '', // 入金日
        '', // 決済参照ID
        0, // 決済手数料
        0, // 請求送料
        0, // 実送料
        '', // 配送会社
        '', // 追跡番号
        '', // 発送日
        '', // 追跡番号送付日
        '', // 到着日
        0, // 売上合計
        0, // 原価合計_暫定
        0, // 原価合計_還付後
        0, // 粗利_暫定
        0, // 粗利_還付後
        '' // メモ
      ];
      
      orderSheet.appendRow(newOrderRow);
      
      // 明細テンプレート行作成
      const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
      const templateRow = [
        orderId,
        1, // 行No
        '', // SKU
        '', // 商品名
        '', // 区分
        0, // 数量
        0, // 単価_売価
        0, // 行割引
        0, // 売上金額
        0, // 原価_暫定
        0, // 還元率
        0, // 還付額
        0, // 原価_還付後
        0, // 行粗利_暫定
        0, // 行粗利_還付後
        '未', // 在庫引当ステータス
        '' // メモ
      ];
      
      detailSheet.appendRow(templateRow);
      
      ui.alert('注文作成完了', `注文ID: ${orderId}\n顧客: ${customerName}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * シート取得（存在しない場合は作成）
 */
function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet(sheetName);
  }
  
  return sheet;
}

/**
 * シート初期化
 */
function initializeSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  switch (sheetName) {
    case SHEET_NAMES.CONFIG:
      sheet.getRange('A1:B10').setValues([
        ['顧客ID接頭辞', CONFIG.CUSTOMER_ID_PREFIX],
        ['顧客ID桁数', CONFIG.CUSTOMER_ID_DIGITS],
        ['注文ID_既定種別', CONFIG.ORDER_TYPE],
        ['売上集計対象ステータス', CONFIG.SALES_TARGET_STATUSES.join(',')],
        ['タイムゾーン', CONFIG.TIMEZONE],
        ['連番カウンタ_顧客', 0],
        ['連番カウンタ_注文', 0],
        ['連番カウンタ_SKU', 0],
        ['PayPal手数料レート', ''],
        ['Wise手数料レート', '']
      ]);
      break;
      
    case SHEET_NAMES.CUSTOMER:
      sheet.getRange('A1:Z1').setValues([[
        '顧客ID', '氏名', 'メール', '電話', '住所1', '住所2', '郵便番号', '国',
        '集客先', '顧客フラグ', '興味_BOX', '興味_PSA', '興味_シングル', '興味_バルク',
        '支払_Wise', '支払_PayPal', 'PII同意', '同意日', '削除依頼', '削除対応日', '連絡不可',
        'メモ', '購入回数', '購入金額累計JPY', '初回購入日', '最終購入日'
      ]]);
      break;
      
    case SHEET_NAMES.ORDER_HEADER:
      sheet.getRange('A1:V1').setValues([[
        '注文ID', '顧客ID', '顧客名', '注文日', '支払い方法', '注文ステータス',
        '入金日', '決済参照ID', '決済手数料JPY', '請求送料JPY', '実送料JPY',
        '配送会社', '追跡番号', '発送日', '追跡番号送付日', '到着日',
        '売上合計JPY', '原価合計_暫定JPY', '原価合計_還付後JPY',
        '粗利_暫定JPY', '粗利_還付後JPY', 'メモ'
      ]]);
      break;
      
    case SHEET_NAMES.ORDER_DETAIL:
      sheet.getRange('A1:R1').setValues([[
        '注文ID', '行No', 'SKU', '商品名', '区分', '数量', '単価_売価JPY', '行割引JPY',
        '売上金額JPY', '原価_暫定JPY', '還元率', '還付額JPY', '原価_還付後JPY',
        '行粗利_暫定JPY', '行粗利_還付後JPY', '在庫引当ステータス', 'メモ'
      ]]);
      break;
      
    case SHEET_NAMES.SKU:
      sheet.getRange('A1:L1').setValues([[
        'SKU', '商品名', '区分', '規格', 'コンディション', '仕入日', '仕入数量',
        '仕入単価JPY', '仕入送料JPY', '還元率', '仕入元', '在庫ステータス',
        'ロケーション', '在庫数量_可用', '引当数量', 'メモ'
      ]]);
      break;
      
    case SHEET_NAMES.INVENTORY:
      sheet.getRange('A1:F1').setValues([[
        '取引種別', '参照ID', '日付', 'SKU', '数量', '単価JPY', '備考'
      ]]);
      break;
      
    case SHEET_NAMES.LOOKER_EXPORT:
      sheet.getRange('A1:U1').setValues([[
        '受注日', '年', '月', '週', '注文ID', '注文ステータス', '集計対象フラグ',
        '顧客ID', '国', '集客先', '顧客フラグ', '支払い方法', 'SKU', '商品名', '区分', '数量',
        '売上金額JPY', '原価_暫定JPY', '原価_還付後JPY', '行粗利_暫定JPY', '行粗利_還付後JPY',
        '決済手数料JPY', '請求送料JPY', '実送料JPY', '配送会社', '還元率'
      ]]);
      break;
  }
}

/**
 * 明細取り込み機能
 */
function importOrderLines() {
  const ui = SpreadsheetApp.getUi();
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const orders = orderSheet.getDataRange().getValues().slice(1);
  
  if (orders.length === 0) {
    ui.alert('エラー', '注文が登録されていません。先に注文を作成してください。', ui.ButtonSet.OK);
    return;
  }
  
  // 注文選択
  let orderList = '注文を選択してください:\n';
  orders.forEach((order, index) => {
    orderList += `${index + 1}. ${order[0]} - ${order[2]} (${order[5]})\n`;
  });
  
  const result = ui.prompt('明細取り込み', orderList, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText());
    if (selection >= 1 && selection <= orders.length) {
      const selectedOrder = orders[selection - 1];
      const orderId = selectedOrder[0];
      
      // TSV/CSVデータの入力
      const dataResult = ui.prompt('明細データ入力', 
        '以下の形式で明細データを入力してください:\n' +
        '商品名\t区分\t数量\t単価\t行割引\t原価\t還元率\n' +
        '例: ポケモンカード\tシングル\t1\t1000\t0\t800\t8%', 
        ui.ButtonSet.OK_CANCEL);
      
      if (dataResult.getSelectedButton() === ui.Button.OK) {
        const dataText = dataResult.getResponseText();
        const lines = dataText.split('\n');
        
        if (lines.length < 2) {
          ui.alert('エラー', 'データが正しく入力されていません。', ui.ButtonSet.OK);
          return;
        }
        
        const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
        const existingDetails = detailSheet.getDataRange().getValues();
        const orderDetails = existingDetails.filter(row => row[0] === orderId);
        
        // 既存の明細を削除（テンプレート行を除く）
        if (orderDetails.length > 1) {
          const startRow = existingDetails.findIndex(row => row[0] === orderId) + 1;
          const endRow = startRow + orderDetails.length - 2;
          if (endRow >= startRow) {
            detailSheet.deleteRows(startRow + 1, endRow - startRow + 1);
          }
        }
        
        // 新しい明細を追加
        let lineNumber = 1;
        let totalSales = 0;
        let totalCostTemporary = 0;
        let totalCostFinal = 0;
        
        for (let i = 1; i < lines.length; i++) {
          const columns = lines[i].split('\t');
          if (columns.length >= 6) {
            const productName = columns[0];
            const category = columns[1];
            const quantity = parseFloat(columns[2]) || 0;
            const unitPrice = parseFloat(columns[3]) || 0;
            const lineDiscount = parseFloat(columns[4]) || 0;
            const costTemporary = parseFloat(columns[5]) || 0;
            const rebateRate = parseFloat(columns[6]) || 0;
            
            const salesAmount = quantity * unitPrice - lineDiscount;
            const rebateAmount = costTemporary * (rebateRate / 100);
            const costFinal = costTemporary - rebateAmount;
            const lineProfitTemporary = salesAmount - costTemporary;
            const lineProfitFinal = salesAmount - costFinal;
            
            const newDetailRow = [
              orderId,
              lineNumber,
              '', // SKU（後で設定）
              productName,
              category,
              quantity,
              unitPrice,
              lineDiscount,
              salesAmount,
              costTemporary,
              rebateRate,
              rebateAmount,
              costFinal,
              lineProfitTemporary,
              lineProfitFinal,
              '未', // 在庫引当ステータス
              '' // メモ
            ];
            
            detailSheet.appendRow(newDetailRow);
            
            totalSales += salesAmount;
            totalCostTemporary += costTemporary;
            totalCostFinal += costFinal;
            lineNumber++;
          }
        }
        
        // 注文ヘッダーの合計を更新
        const orderRowIndex = orders.findIndex(row => row[0] === orderId) + 2; // +2 for header and 0-based index
        orderSheet.getRange(orderRowIndex, 17).setValue(totalSales); // 売上合計
        orderSheet.getRange(orderRowIndex, 18).setValue(totalCostTemporary); // 原価合計_暫定
        orderSheet.getRange(orderRowIndex, 19).setValue(totalCostFinal); // 原価合計_還付後
        
        // 粗利計算（決済手数料と送料は後で設定）
        const grossProfitTemporary = totalSales - totalCostTemporary;
        const grossProfitFinal = totalSales - totalCostFinal;
        orderSheet.getRange(orderRowIndex, 20).setValue(grossProfitTemporary); // 粗利_暫定
        orderSheet.getRange(orderRowIndex, 21).setValue(grossProfitFinal); // 粗利_還付後
        
        ui.alert('明細取り込み完了', 
          `注文ID: ${orderId}\n` +
          `明細行数: ${lineNumber - 1}\n` +
          `売上合計: ¥${totalSales.toLocaleString()}\n` +
          `粗利（暫定）: ¥${grossProfitTemporary.toLocaleString()}`, 
          ui.ButtonSet.OK);
      }
    }
  }
}

/**
 * 在庫引当機能
 */
function allocateInventory() {
  const ui = SpreadsheetApp.getUi();
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const orders = orderSheet.getDataRange().getValues().slice(1);
  
  // 支払い完了以降の注文のみ表示
  const eligibleOrders = orders.filter(order => 
    CONFIG.SALES_TARGET_STATUSES.includes(order[5]) && order[5] !== '支払い前'
  );
  
  if (eligibleOrders.length === 0) {
    ui.alert('エラー', '在庫引当可能な注文がありません。', ui.ButtonSet.OK);
    return;
  }
  
  // 注文選択
  let orderList = '在庫引当する注文を選択してください:\n';
  eligibleOrders.forEach((order, index) => {
    orderList += `${index + 1}. ${order[0]} - ${order[2]} (${order[5]})\n`;
  });
  
  const result = ui.prompt('在庫引当', orderList, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText());
    if (selection >= 1 && selection <= eligibleOrders.length) {
      const selectedOrder = eligibleOrders[selection - 1];
      const orderId = selectedOrder[0];
      
      // 明細を取得
      const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
      const allDetails = detailSheet.getDataRange().getValues();
      const orderDetails = allDetails.filter(row => row[0] === orderId);
      
      if (orderDetails.length === 0) {
        ui.alert('エラー', '注文明細が見つかりません。', ui.ButtonSet.OK);
        return;
      }
      
      // 在庫引当処理
      const skuSheet = getSheet(SHEET_NAMES.SKU);
      const inventorySheet = getSheet(SHEET_NAMES.INVENTORY);
      const skuData = skuSheet.getDataRange().getValues();
      
      let allocatedCount = 0;
      let errorMessages = [];
      
      for (let i = 1; i < orderDetails.length; i++) { // ヘッダーをスキップ
        const detail = orderDetails[i];
        const productName = detail[3];
        const category = detail[4];
        const quantity = parseFloat(detail[5]) || 0;
        
        if (quantity <= 0) continue;
        
        // 商品名でSKUを検索
        const matchingSku = skuData.find(sku => 
          sku[1] === productName && sku[2] === category && sku[11] === '在庫'
        );
        
        if (matchingSku) {
          const sku = matchingSku[0];
          const availableStock = parseFloat(matchingSku[13]) || 0;
          const allocatedStock = parseFloat(matchingSku[14]) || 0;
          const actualAvailable = availableStock - allocatedStock;
          
          if (actualAvailable >= quantity) {
            // 在庫引当可能
            const skuRowIndex = skuData.findIndex(row => row[0] === sku) + 1;
            const newAllocatedStock = allocatedStock + quantity;
            
            // 引当数量を更新
            skuSheet.getRange(skuRowIndex, 15).setValue(newAllocatedStock);
            
            // 明細のSKUと引当ステータスを更新
            const detailRowIndex = allDetails.findIndex(row => 
              row[0] === orderId && row[1] === detail[1]
            ) + 1;
            detailSheet.getRange(detailRowIndex, 3).setValue(sku); // SKU
            detailSheet.getRange(detailRowIndex, 16).setValue('引当'); // 在庫引当ステータス
            
            // 入出庫台帳に記録
            inventorySheet.appendRow([
              '予約',
              orderId,
              new Date(),
              sku,
              -quantity, // 出庫なので負の値
              0, // 単価（引当時は0）
              `注文引当: ${productName}`
            ]);
            
            allocatedCount++;
          } else {
            errorMessages.push(`${productName}: 在庫不足 (必要: ${quantity}, 可用: ${actualAvailable})`);
          }
        } else {
          errorMessages.push(`${productName}: SKUが見つかりません`);
        }
      }
      
      // 結果表示
      let message = `在庫引当完了\n引当件数: ${allocatedCount}`;
      if (errorMessages.length > 0) {
        message += `\n\nエラー:\n${errorMessages.join('\n')}`;
      }
      
      ui.alert('在庫引当結果', message, ui.ButtonSet.OK);
    }
  }
}

/**
 * 発送処理機能
 */
function shipOrder() {
  const ui = SpreadsheetApp.getUi();
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const orders = orderSheet.getDataRange().getValues().slice(1);
  
  // 仕入れ完了以降の注文のみ表示
  const eligibleOrders = orders.filter(order => 
    ['仕入れ完了', '発送済', '追跡番号送付済', '到着'].includes(order[5])
  );
  
  if (eligibleOrders.length === 0) {
    ui.alert('エラー', '発送可能な注文がありません。', ui.ButtonSet.OK);
    return;
  }
  
  // 注文選択
  let orderList = '発送する注文を選択してください:\n';
  eligibleOrders.forEach((order, index) => {
    orderList += `${index + 1}. ${order[0]} - ${order[2]} (${order[5]})\n`;
  });
  
  const result = ui.prompt('発送処理', orderList, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText());
    if (selection >= 1 && selection <= eligibleOrders.length) {
      const selectedOrder = eligibleOrders[selection - 1];
      const orderId = selectedOrder[0];
      
      // 配送会社と追跡番号の入力
      const carrierResult = ui.prompt('配送情報入力', 
        '配送会社を入力してください:\n1. Japan Post\n2. DHL\n3. FedEx\n4. UPS\n5. Other', 
        ui.ButtonSet.OK_CANCEL);
      
      if (carrierResult.getSelectedButton() === ui.Button.OK) {
        const carrierSelection = parseInt(carrierResult.getResponseText());
        const carriers = ['Japan Post', 'DHL', 'FedEx', 'UPS', 'Other'];
        const carrier = carriers[carrierSelection - 1] || 'Other';
        
        const trackingResult = ui.prompt('追跡番号入力', 
          '追跡番号を入力してください:', 
          ui.ButtonSet.OK_CANCEL);
        
        if (trackingResult.getSelectedButton() === ui.Button.OK) {
          const trackingNumber = trackingResult.getResponseText();
          const shipDate = new Date();
          
          // 注文ヘッダーを更新
          const orderRowIndex = orders.findIndex(row => row[0] === orderId) + 2;
          orderSheet.getRange(orderRowIndex, 5).setValue('発送済'); // 注文ステータス
          orderSheet.getRange(orderRowIndex, 11).setValue(carrier); // 配送会社
          orderSheet.getRange(orderRowIndex, 12).setValue(trackingNumber); // 追跡番号
          orderSheet.getRange(orderRowIndex, 13).setValue(shipDate); // 発送日
          
          // 明細の在庫引当ステータスを更新
          const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
          const allDetails = detailSheet.getDataRange().getValues();
          const orderDetails = allDetails.filter(row => row[0] === orderId);
          
          for (let i = 1; i < orderDetails.length; i++) {
            const detailRowIndex = allDetails.findIndex(row => 
              row[0] === orderId && row[1] === orderDetails[i][1]
            ) + 1;
            detailSheet.getRange(detailRowIndex, 16).setValue('出庫済'); // 在庫引当ステータス
          }
          
          // 在庫出庫処理
          const skuSheet = getSheet(SHEET_NAMES.SKU);
          const inventorySheet = getSheet(SHEET_NAMES.INVENTORY);
          const skuData = skuSheet.getDataRange().getValues();
          
          for (let i = 1; i < orderDetails.length; i++) {
            const detail = orderDetails[i];
            const sku = detail[2];
            const quantity = parseFloat(detail[5]) || 0;
            
            if (sku && quantity > 0) {
              // SKUの在庫数量を更新
              const skuRowIndex = skuData.findIndex(row => row[0] === sku) + 1;
              if (skuRowIndex > 0) {
                const currentStock = parseFloat(skuData[skuRowIndex - 1][13]) || 0;
                const currentAllocated = parseFloat(skuData[skuRowIndex - 1][14]) || 0;
                const newStock = currentStock - quantity;
                const newAllocated = Math.max(0, currentAllocated - quantity);
                
                skuSheet.getRange(skuRowIndex, 14).setValue(newStock); // 在庫数量_可用
                skuSheet.getRange(skuRowIndex, 15).setValue(newAllocated); // 引当数量
                
                // 入出庫台帳に記録
                inventorySheet.appendRow([
                  '出庫',
                  orderId,
                  shipDate,
                  sku,
                  -quantity, // 出庫なので負の値
                  0, // 単価（出庫時は0）
                  `発送出庫: ${detail[3]}`
                ]);
              }
            }
          }
          
          ui.alert('発送処理完了', 
            `注文ID: ${orderId}\n` +
            `配送会社: ${carrier}\n` +
            `追跡番号: ${trackingNumber}\n` +
            `発送日: ${Utilities.formatDate(shipDate, CONFIG.TIMEZONE, 'yyyy/MM/dd')}`, 
            ui.ButtonSet.OK);
        }
      }
    }
  }
}

/**
 * 注文ステータス更新機能
 */
function updateOrderStatus() {
  const ui = SpreadsheetApp.getUi();
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const orders = orderSheet.getDataRange().getValues().slice(1);
  
  if (orders.length === 0) {
    ui.alert('エラー', '注文が登録されていません。', ui.ButtonSet.OK);
    return;
  }
  
  // 注文選択
  let orderList = 'ステータスを更新する注文を選択してください:\n';
  orders.forEach((order, index) => {
    orderList += `${index + 1}. ${order[0]} - ${order[2]} (現在: ${order[5]})\n`;
  });
  
  const result = ui.prompt('注文ステータス更新', orderList, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const selection = parseInt(result.getResponseText());
    if (selection >= 1 && selection <= orders.length) {
      const selectedOrder = orders[selection - 1];
      const orderId = selectedOrder[0];
      const currentStatus = selectedOrder[5];
      
      // ステータス選択
      const statusOptions = [
        '支払い前', '支払い完了', '仕入れ完了', '発送済', 
        '追跡番号送付済', '到着', 'キャンセル'
      ];
      
      let statusList = '新しいステータスを選択してください:\n';
      statusOptions.forEach((status, index) => {
        const marker = status === currentStatus ? ' (現在)' : '';
        statusList += `${index + 1}. ${status}${marker}\n`;
      });
      
      const statusResult = ui.prompt('ステータス選択', statusList, ui.ButtonSet.OK_CANCEL);
      
      if (statusResult.getSelectedButton() === ui.Button.OK) {
        const statusSelection = parseInt(statusResult.getResponseText());
        if (statusSelection >= 1 && statusSelection <= statusOptions.length) {
          const newStatus = statusOptions[statusSelection - 1];
          
          // 注文ヘッダーを更新
          const orderRowIndex = orders.findIndex(row => row[0] === orderId) + 2;
          orderSheet.getRange(orderRowIndex, 6).setValue(newStatus); // 注文ステータス
          
          // ステータスに応じた自動処理
          const now = new Date();
          
          switch (newStatus) {
            case '支払い完了':
              orderSheet.getRange(orderRowIndex, 7).setValue(now); // 入金日
              break;
              
            case '仕入れ完了':
              // 明細の原価チェック
              const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
              const allDetails = detailSheet.getDataRange().getValues();
              const orderDetails = allDetails.filter(row => row[0] === orderId);
              
              let hasUnsetCost = false;
              for (let i = 1; i < orderDetails.length; i++) {
                const costTemporary = parseFloat(orderDetails[i][9]) || 0;
                if (costTemporary === 0) {
                  hasUnsetCost = true;
                  break;
                }
              }
              
              if (hasUnsetCost) {
                ui.alert('警告', '未設定の原価があります。明細を確認してください。', ui.ButtonSet.OK);
              }
              break;
              
            case '発送済':
              // 発送日が未設定の場合は現在日時を設定
              const shipDate = orderSheet.getRange(orderRowIndex, 14).getValue();
              if (!shipDate) {
                orderSheet.getRange(orderRowIndex, 14).setValue(now); // 発送日
              }
              break;
              
            case '追跡番号送付済':
              // 追跡番号送付日を設定
              orderSheet.getRange(orderRowIndex, 15).setValue(now); // 追跡番号送付日
              break;
              
            case '到着':
              // 到着日を設定
              orderSheet.getRange(orderRowIndex, 16).setValue(now); // 到着日
              break;
          }
          
          // 顧客集計を更新（売上集計対象ステータスの場合）
          if (CONFIG.SALES_TARGET_STATUSES.includes(newStatus)) {
            refreshCustomerAggregates();
          }
          
          ui.alert('ステータス更新完了', 
            `注文ID: ${orderId}\n` +
            `旧ステータス: ${currentStatus}\n` +
            `新ステータス: ${newStatus}`, 
            ui.ButtonSet.OK);
        }
      }
    }
  }
}

function importPaymentCsv() {
  SpreadsheetApp.getUi().alert('決済CSV取込機能', 'この機能は今後実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function matchPaymentsToOrders() {
  SpreadsheetApp.getUi().alert('決済紐付け機能', 'この機能は今後実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 顧客集計更新機能
 */
function refreshCustomerAggregates() {
  const ui = SpreadsheetApp.getUi();
  const customerSheet = getSheet(SHEET_NAMES.CUSTOMER);
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  
  const customers = customerSheet.getDataRange().getValues();
  const orders = orderSheet.getDataRange().getValues();
  
  let updatedCount = 0;
  
  // 各顧客の集計を更新
  for (let i = 1; i < customers.length; i++) { // ヘッダーをスキップ
    const customerId = customers[i][0];
    
    // 該当顧客の売上集計対象注文を取得
    const customerOrders = orders.filter(order => 
      order[1] === customerId && CONFIG.SALES_TARGET_STATUSES.includes(order[5])
    );
    
    // 購入回数
    const purchaseCount = customerOrders.length;
    
    // 購入金額累計
    const totalAmount = customerOrders.reduce((sum, order) => {
      return sum + (parseFloat(order[16]) || 0); // 売上合計JPY
    }, 0);
    
    // 初回購入日
    const firstPurchaseDate = customerOrders.length > 0 
      ? customerOrders.reduce((earliest, order) => {
          const orderDate = new Date(order[3]);
          return orderDate < earliest ? orderDate : earliest;
        }, new Date(customerOrders[0][3]))
      : '';
    
    // 最終購入日
    const lastPurchaseDate = customerOrders.length > 0
      ? customerOrders.reduce((latest, order) => {
          const orderDate = new Date(order[3]);
          return orderDate > latest ? orderDate : latest;
        }, new Date(customerOrders[0][3]))
      : '';
    
    // 顧客シートを更新
    customerSheet.getRange(i + 1, 23).setValue(purchaseCount); // 購入回数
    customerSheet.getRange(i + 1, 24).setValue(totalAmount); // 購入金額累計JPY
    customerSheet.getRange(i + 1, 25).setValue(firstPurchaseDate); // 初回購入日
    customerSheet.getRange(i + 1, 26).setValue(lastPurchaseDate); // 最終購入日
    
    updatedCount++;
  }
  
  ui.alert('顧客集計更新完了', 
    `更新対象顧客数: ${updatedCount}人\n` +
    `売上集計対象ステータス: ${CONFIG.SALES_TARGET_STATUSES.join(', ')}`, 
    ui.ButtonSet.OK);
}

/**
 * Looker Studio連携用Export更新機能
 */
function rebuildLookerExport() {
  const ui = SpreadsheetApp.getUi();
  const exportSheet = getSheet(SHEET_NAMES.LOOKER_EXPORT);
  const orderSheet = getSheet(SHEET_NAMES.ORDER_HEADER);
  const detailSheet = getSheet(SHEET_NAMES.ORDER_DETAIL);
  const customerSheet = getSheet(SHEET_NAMES.CUSTOMER);
  
  // 既存のデータをクリア（ヘッダーを除く）
  const lastRow = exportSheet.getLastRow();
  if (lastRow > 1) {
    exportSheet.getRange(2, 1, lastRow - 1, exportSheet.getLastColumn()).clear();
  }
  
  // 注文データを取得
  const orders = orderSheet.getDataRange().getValues();
  const details = detailSheet.getDataRange().getValues();
  const customers = customerSheet.getDataRange().getValues();
  
  let exportData = [];
  
  // 各注文の明細を処理
  for (let i = 1; i < orders.length; i++) { // ヘッダーをスキップ
    const order = orders[i];
    const orderId = order[0];
    const customerId = order[1];
    const orderDate = new Date(order[3]);
    const orderStatus = order[5];
    const paymentMethod = order[4];
    const carrier = order[11];
    const salesTotal = parseFloat(order[16]) || 0;
    const costTemporary = parseFloat(order[17]) || 0;
    const costFinal = parseFloat(order[18]) || 0;
    const grossProfitTemporary = parseFloat(order[19]) || 0;
    const grossProfitFinal = parseFloat(order[20]) || 0;
    const paymentFee = parseFloat(order[8]) || 0;
    const requestedShipping = parseFloat(order[9]) || 0;
    const actualShipping = parseFloat(order[10]) || 0;
    
    // 顧客情報を取得
    const customer = customers.find(c => c[0] === customerId);
    const country = customer ? customer[7] : '';
    const acquisitionSource = customer ? customer[8] : '';
    const customerFlag = customer ? customer[9] : '';
    
    // 集計対象フラグ
    const isTargetForSales = CONFIG.SALES_TARGET_STATUSES.includes(orderStatus) ? 1 : 0;
    
    // 日付情報
    const year = orderDate.getFullYear();
    const month = orderDate.getMonth() + 1;
    const week = Math.ceil((orderDate.getDate() + new Date(year, month - 1, 1).getDay()) / 7);
    
    // 該当注文の明細を取得
    const orderDetails = details.filter(detail => detail[0] === orderId);
    
    for (let j = 1; j < orderDetails.length; j++) { // ヘッダーをスキップ
      const detail = orderDetails[j];
      const sku = detail[2];
      const productName = detail[3];
      const category = detail[4];
      const quantity = parseFloat(detail[5]) || 0;
      const salesAmount = parseFloat(detail[8]) || 0;
      const costTemporaryDetail = parseFloat(detail[9]) || 0;
      const rebateRate = parseFloat(detail[10]) || 0;
      const rebateAmount = parseFloat(detail[11]) || 0;
      const costFinalDetail = parseFloat(detail[12]) || 0;
      const lineProfitTemporary = parseFloat(detail[13]) || 0;
      const lineProfitFinal = parseFloat(detail[14]) || 0;
      
      // 決済手数料と送料を明細に按分（数量比で配分）
      const totalQuantity = orderDetails.slice(1).reduce((sum, d) => sum + (parseFloat(d[5]) || 0), 0);
      const allocatedPaymentFee = totalQuantity > 0 ? (paymentFee * quantity / totalQuantity) : 0;
      const allocatedRequestedShipping = totalQuantity > 0 ? (requestedShipping * quantity / totalQuantity) : 0;
      const allocatedActualShipping = totalQuantity > 0 ? (actualShipping * quantity / totalQuantity) : 0;
      
      const exportRow = [
        orderDate, // 受注日
        year, // 年
        month, // 月
        week, // 週
        orderId, // 注文ID
        orderStatus, // 注文ステータス
        isTargetForSales, // 集計対象フラグ
        customerId, // 顧客ID
        country, // 国
        acquisitionSource, // 集客先
        customerFlag, // 顧客フラグ
        paymentMethod, // 支払い方法
        sku, // SKU
        productName, // 商品名
        category, // 区分
        quantity, // 数量
        salesAmount, // 売上金額JPY
        costTemporaryDetail, // 原価_暫定JPY
        costFinalDetail, // 原価_還付後JPY
        lineProfitTemporary, // 行粗利_暫定JPY
        lineProfitFinal, // 行粗利_還付後JPY
        allocatedPaymentFee, // 決済手数料JPY
        allocatedRequestedShipping, // 請求送料JPY
        allocatedActualShipping, // 実送料JPY
        carrier, // 配送会社
        rebateRate // 還元率
      ];
      
      exportData.push(exportRow);
    }
  }
  
  // データをシートに書き込み
  if (exportData.length > 0) {
    exportSheet.getRange(2, 1, exportData.length, exportData[0].length).setValues(exportData);
  }
  
  ui.alert('Looker Export更新完了', 
    `更新件数: ${exportData.length}行\n` +
    `注文数: ${orders.length - 1}件\n` +
    `集計対象: ${exportData.filter(row => row[6] === 1).length}行`, 
    ui.ButtonSet.OK);
}

/**
 * バックアップ作成機能
 */
function createBackup() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const timestamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
  
  try {
    // スプレッドシートをコピー
    const backupName = `ポケモンカード販売管理_バックアップ_${timestamp}`;
    const backupFile = DriveApp.getFileById(spreadsheet.getId()).makeCopy(backupName);
    
    ui.alert('バックアップ作成完了', 
      `バックアップ名: ${backupName}\n` +
      `作成日時: ${Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss')}\n` +
      `ファイルID: ${backupFile.getId()}`, 
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('バックアップ作成エラー', 
      `エラーが発生しました: ${error.message}`, 
      ui.ButtonSet.OK);
  }
}