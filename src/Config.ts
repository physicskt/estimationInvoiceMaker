/**
 * 設定定数
 * アプリケーション全体で使用される設定値を定義
 */

// シート設定の型定義
interface SheetConfig {
  readonly INPUT: string;
  readonly TEMPLATE: string;
  readonly HISTORY: string;
}

// セル位置設定の型定義
interface CellConfig {
  readonly DOCUMENT_TYPE: string;
  readonly ISSUE_DATE: string;
  readonly COMPANY_NAME: string;
  readonly CONTACT_NAME: string;
  readonly ADDRESS: string;
  readonly EMAIL: string;
  readonly REMARKS: string;
  readonly TOTAL_AMOUNT: string;
  readonly TAX: string;
  readonly GRAND_TOTAL: string;
}

// テンプレートシートのセル位置設定の型定義
interface TemplateCellConfig {
  readonly DOCUMENT_TYPE: string;
  readonly ISSUE_DATE: string;
  readonly COMPANY_NAME: string;
  readonly CONTACT_NAME: string;
  readonly ADDRESS: string;
  readonly REMARKS: string;
  readonly TOTAL_AMOUNT: string;
  readonly TAX: string;
  readonly GRAND_TOTAL: string;
}

// 範囲設定の型定義
interface RangeConfig {
  readonly ITEMS: string;
}

// テンプレートシート範囲設定の型定義
interface TemplateRangeConfig {
  readonly ITEMS_START_ROW: number;
  readonly ITEMS_MAX_ROWS: number;
}

// フォルダ設定の型定義
interface FolderConfig {
  readonly ESTIMATES: string;
  readonly INVOICES: string;
  readonly BACKUP: string;
}

// メール設定の型定義
interface EmailConfig {
  readonly SENDER_COMPANY: string;
  readonly SENDER_DEPARTMENT: string;
  readonly SENDER_NAME: string;
}

// 全体設定の型定義
interface AppConfig {
  readonly SHEETS: SheetConfig;
  readonly CELLS: CellConfig;
  readonly RANGES: RangeConfig;
  readonly TEMPLATE_CELLS: TemplateCellConfig;
  readonly TEMPLATE_RANGES: TemplateRangeConfig;
  readonly FOLDERS: FolderConfig;
  readonly EMAIL: EmailConfig;
  readonly HISTORY_HEADERS: readonly string[];
}

const CONFIG: AppConfig = {
  // シート名
  SHEETS: {
    INPUT: '入力',
    TEMPLATE: 'テンプレート',
    HISTORY: '送信履歴'
  },
  
  // 入力シートのセル位置
  CELLS: {
    DOCUMENT_TYPE: 'B2',     // 書類種別
    ISSUE_DATE: 'B3',        // 発行日
    COMPANY_NAME: 'B4',      // 宛先会社名
    CONTACT_NAME: 'B5',      // 担当者名
    ADDRESS: 'B6',           // 住所
    EMAIL: 'B7',             // メールアドレス
    REMARKS: 'B8',           // 備考
    TOTAL_AMOUNT: 'F15',     // 小計
    TAX: 'F16',              // 消費税
    GRAND_TOTAL: 'F17'       // 合計金額
  },
  
  // 入力シートの範囲
  RANGES: {
    ITEMS: 'A10:D14'         // 商品明細（品目、数量、単価、小計）
  },
  
  // テンプレートシートのセル位置
  TEMPLATE_CELLS: {
    DOCUMENT_TYPE: 'A1',     // 書類種別
    ISSUE_DATE: 'F2',        // 発行日
    COMPANY_NAME: 'A4',      // 宛先会社名
    CONTACT_NAME: 'A5',      // 担当者名
    ADDRESS: 'A6',           // 住所
    REMARKS: 'A20',          // 備考
    TOTAL_AMOUNT: 'F15',     // 小計
    TAX: 'F16',              // 消費税
    GRAND_TOTAL: 'F17'       // 合計金額
  },
  
  // テンプレートシートの範囲
  TEMPLATE_RANGES: {
    ITEMS_START_ROW: 10,     // 商品明細開始行
    ITEMS_MAX_ROWS: 5        // 商品明細最大行数
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
    'ファイルURL'
  ] as const
};

// 書類種別の型定義
type DocumentType = '見積書' | '請求書';

// 商品明細の型定義
interface ItemData {
  name: string;
  quantity: number;
  unitPrice: number;
  subtotal: number;
}

// 入力データの型定義
interface InputData {
  documentType: DocumentType;
  issueDate: Date;
  companyName: string;
  contactName: string;
  address: string;
  email: string;
  remarks: string;
  items: ItemData[];
  totalAmount: number;
  tax: number;
  grandTotal: number;
}