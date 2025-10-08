const CONFIG = Object.freeze({
  timezone: 'Asia/Tokyo',
  idPrefixes: {
    customer: '顧客ID接頭辞',
    customerDigits: '顧客ID桁数',
    sku: 'SKU接頭辞',
    orderType: '注文ID_既定種別'
  },
  counterKeys: {
    customer: '連番_顧客',
    sku: '連番_SKU',
    order: '連番_注文'
  },
  sheets: {
    settings: '設定_定数',
    lists: '参照_リスト',
    customers: 'マスター_顧客',
    orderHeader: '取引_注文ヘッダ',
    orderLines: '取引_注文明細',
    sku: '在庫_SKU',
    stockLedger: '在庫_入出庫',
    paypal: '決済_取込_PayPal',
    wise: '決済_取込_Wise',
    dashboard: '集計_ダッシュボード',
    looker: 'Export_Looker',
    audit: '監査ログ_操作'
  },
  defaultConstants: {
    顧客ID接頭辞: 'PK',
    顧客ID桁数: 2,
    SKU接頭辞: 'SKU',
    注文ID_既定種別: 'ASK',
    売上集計対象ステータス: '支払い完了,仕入れ完了,発送済,追跡番号送付済,到着',
    タイムゾーン: 'Asia/Tokyo',
    PayPal手数料割合: 0,
    PayPal手数料固定額: 0
  },
  listSeeds: {
    集客先: ['Instagram', 'Facebook', 'WhatsApp', 'TikTok', 'Other'],
    顧客フラグ: ['ストリーマー', 'カードショップ', 'リセラー', '卸'],
    興味カテゴリ: ['BOX', 'PSA', 'シングル', 'バルク'],
    支払い方法: ['Wise', 'PayPal', 'Other'],
    注文ステータス: ['支払い前', '支払い完了', '仕入れ完了', '発送済', '追跡番号送付済', '到着', 'キャンセル'],
    在庫ステータス: ['在庫', '予約', '発注中', '出庫済', '欠品'],
    配送会社: ['Japan Post', 'DHL', 'FedEx', 'UPS', 'Other'],
    還元率: ['0%', '8%', '10%']
  },
  headers: {
    settings: ['キー', '値', '説明'],
    lists: ['カテゴリ', '値'],
    customers: [
      '顧客ID', '氏名', 'メール', '電話', '住所1', '住所2', '郵便番号', '国',
      '集客先', '顧客フラグ', '興味_BOX', '興味_PSA', '興味_シングル', '興味_バルク',
      '支払_Wise', '支払_PayPal',
      'PII同意', '同意日', '削除依頼', '削除対応日', '連絡不可', 'メモ',
      '購入回数', '購入金額累計JPY', '初回購入日', '最終購入日'
    ],
    orderHeader: [
      '注文ID', '顧客ID', '顧客名', '注文日', '支払い方法', '注文ステータス',
      '入金日', '決済参照ID', '決済手数料JPY', '請求送料JPY', '実送料JPY',
      '配送会社', '追跡番号', '発送日', '追跡番号送付日', '到着日',
      '売上合計JPY', '原価合計_暫定JPY', '原価合計_還付後JPY',
      '粗利_暫定JPY', '粗利_還付後JPY', 'メモ'
    ],
    orderLines: [
      '注文ID', '行No', 'SKU', '商品名', '区分', '数量', '単価_売価JPY', '行割引JPY',
      '売上金額JPY', '原価_暫定JPY', '還元率', '還付額JPY', '原価_還付後JPY',
      '行粗利_暫定JPY', '行粗利_還付後JPY', '在庫引当ステータス', 'メモ'
    ],
    sku: [
      'SKU', '商品名', '区分', '規格', 'コンディション', '仕入日', '仕入数量', '仕入単価JPY',
      '仕入送料JPY', '還元率', '仕入元', '在庫ステータス', 'ロケーション',
      '在庫数量_可用', '引当数量', 'メモ'
    ],
    stockLedger: [
      '取引種別', '参照ID', '日付', 'SKU', '数量', '単価JPY', '備考'
    ],
    paypal: [
      'Date', 'Time', 'Time Zone', 'Name', 'Type', 'Status', 'Gross', 'Fee', 'Net', 'Currency',
      'Transaction ID', 'From Email', 'To Email', 'Memo'
    ],
    wise: [
      'Date', 'Time', 'Amount', 'Fee', 'Total', 'Currency', 'Transfer ID', 'Counterparty'
    ],
    looker: [
      '受注日', '年', '月', '週', '注文ID', '注文ステータス', '集計対象',
      '顧客ID', '国', '集客先', '顧客フラグ', '支払い方法',
      'SKU', '商品名', '区分', '数量',
      '売上金額JPY', '原価_暫定JPY', '原価_還付後JPY', '行粗利_暫定JPY', '行粗利_還付後JPY',
      '決済手数料JPY', '請求送料JPY', '実送料JPY', '配送会社', '還元率'
    ],
    audit: ['タイムスタンプ', 'ユーザー', 'イベント', '詳細']
  },
  inventoryAllocationStatuses: {
    none: '未',
    allocated: '引当',
    shipped: '出庫済'
  },
  orderStatusFlow: ['支払い前', '支払い完了', '仕入れ完了', '発送済', '追跡番号送付済', '到着'],
  stockMovementTypes: {
    inbound: '入庫',
    reserved: '予約',
    outbound: '出庫',
    returned: '返品'
  }
});