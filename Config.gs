/**
 * 設定定数
 * アプリケーション全体で使用される設定値を定義
 */
const CONFIG = {
  // シート名
  SHEETS: {
    INPUT: '入力',
    TEMPLATE: 'テンプレート',
    HISTORY: '送信履歴',
    COMPANY_HISTORY: '宛名履歴'
  },
  
  // 入力シートのセル位置
  CELLS: {
    DOCUMENT_TYPE: 'B2',     // 書類種別
    ISSUE_DATE: 'B3',        // 発行日
    DOCUMENT_NUMBER: 'B4',   // 書類番号（3桁）
    COMPANY_NAME: 'B5',      // 宛先会社名
    CONTACT_NAME: 'B6',      // 担当者名
    ADDRESS: 'B7',           // 住所
    EMAIL: 'B8',             // メールアドレス
    REMARKS: 'B9',           // 備考
    TOTAL_AMOUNT: 'D30',     // 小計
    TAX: 'D31',              // 消費税
    GRAND_TOTAL: 'D32'       // 合計金額
  },
  
  // 商品明細設定
  ITEMS_CONFIG: {
    MAX_ROWS: 20,            // 最大行数
    DEFAULT_VISIBLE_ROWS: 10, // デフォルト表示行数
    START_ROW: 10            // 開始行
  },
  
  // 入力シートの範囲
  RANGES: {
    ITEMS: 'A10:D29'         // 商品明細（品目、数量、単価、小計）- 最大20行（動的に計算される）
  },
  
  // テンプレートシートのセル位置
  TEMPLATE_CELLS: {
    DOCUMENT_TYPE: 'A1',     // 書類種別
    ISSUE_DATE: 'D2',        // 発行日
    COMPANY_NAME: 'A4',      // 宛先会社名
    CONTACT_NAME: 'A5',      // 担当者名
    ADDRESS: 'A6',           // 住所
    REMARKS: 'A34',          // 備考
    TOTAL_AMOUNT: 'D30',     // 小計
    TAX: 'D31',              // 消費税
    GRAND_TOTAL: 'D32'       // 合計金額
  },
  
  // テンプレートシートの範囲
  TEMPLATE_RANGES: {
    ITEMS_START_ROW: 10,     // 商品明細開始行
    ITEMS_MAX_ROWS: 20       // 商品明細最大行数（ITEMS_CONFIG.MAX_ROWSと同期）
  },
  
  // フォルダ名
  FOLDERS: {
    ESTIMATES: '見積書',
    INVOICES: '請求書',
    BACKUP: 'バックアップ'
  },
  
  // メール設定
  EMAIL: {
    SENDER_COMPANY: '株式会社サンプル',
    SENDER_DEPARTMENT: '営業部',
    SENDER_NAME: '山田太郎'
  },
  
  // 送信履歴シートのヘッダー
  HISTORY_HEADERS: [
    '送信日時',
    '書類種別',
    '宛先会社',
    'メールアドレス',
    'ファイル名',
    'ファイルURL',
    'メール送信'
  ],
  
  // 宛名履歴シートのヘッダー
  COMPANY_HISTORY_HEADERS: [
    '会社名',
    '最終使用日時',
    '使用回数'
  ]
};