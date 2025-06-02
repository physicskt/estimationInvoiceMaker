/**
 * ユーティリティ関数
 * 共通で使用される補助的な関数を定義
 */

/**
 * フォルダを取得または作成
 * @param {GoogleAppsScript.Drive.Folder} parentFolder 親フォルダ
 * @param {string} folderName フォルダ名
 * @return {GoogleAppsScript.Drive.Folder} フォルダオブジェクト
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

/**
 * 送信履歴シートを取得または作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 送信履歴シート
 */
function getOrCreateHistorySheet(spreadsheet) {
  let historySheet = spreadsheet.getSheetByName(CONFIG.SHEETS.HISTORY);
  
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet(CONFIG.SHEETS.HISTORY);
    
    // ヘッダーを設定
    const headers = CONFIG.HISTORY_HEADERS;
    for (let i = 0; i < headers.length; i++) {
      historySheet.getRange(1, i + 1).setValue(headers[i]);
    }
    
    // ヘッダー行のフォーマット
    const headerRange = historySheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e6f3ff');
  }
  
  return historySheet;
}

/**
 * スプレッドシートに入力シートを作成
 */
function createInputSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let inputSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
  
  if (!inputSheet) {
    inputSheet = spreadsheet.insertSheet(CONFIG.SHEETS.INPUT);
    
    // 入力フォームのレイアウトを作成
    setupInputSheetLayout(inputSheet);
  }
  
  return inputSheet;
}

/**
 * 入力シートのレイアウトを設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 入力シート
 */
function setupInputSheetLayout(sheet) {
  // ヘッダー
  sheet.getRange('A1').setValue('見積書・請求書 作成システム');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  
  // 基本情報入力欄
  const labels = [
    ['書類種別', 'B2'],
    ['発行日', 'B3'],
    ['宛先会社名', 'B4'],
    ['担当者名', 'B5'],
    ['住所', 'B6'],
    ['メールアドレス', 'B7'],
    ['備考', 'B8']
  ];
  
  labels.forEach(([label, cell], index) => {
    sheet.getRange(`A${index + 2}`).setValue(label);
    sheet.getRange(`A${index + 2}`).setFontWeight('bold');
  });
  
  // 商品明細ヘッダー
  sheet.getRange('A9').setValue('商品明細');
  sheet.getRange('A9').setFontSize(14).setFontWeight('bold');
  
  const itemHeaders = ['品目', '数量', '単価', '小計'];
  itemHeaders.forEach((header, index) => {
    sheet.getRange(10, index + 1).setValue(header);
    sheet.getRange(10, index + 1).setFontWeight('bold').setBackground('#f0f0f0');
  });
  
  // 合計欄
  sheet.getRange('E15').setValue('小計');
  sheet.getRange('E16').setValue('消費税');
  sheet.getRange('E17').setValue('合計');
  
  sheet.getRange('E15:E17').setFontWeight('bold');
  
  // 送信ボタン用のセル
  sheet.getRange('A19').setValue('送信ボタン');
  sheet.getRange('B19').setValue('ここにボタンを配置してください');
  sheet.getRange('B19').setBackground('#ffdddd');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 120); // A列
  sheet.setColumnWidth(2, 200); // B列
  sheet.setColumnWidth(3, 80);  // C列
  sheet.setColumnWidth(4, 100); // D列
  sheet.setColumnWidth(5, 80);  // E列
  sheet.setColumnWidth(6, 100); // F列
}

/**
 * スプレッドシートにテンプレートシートを作成
 */
function createTemplateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  if (!templateSheet) {
    templateSheet = spreadsheet.insertSheet(CONFIG.SHEETS.TEMPLATE);
    
    // テンプレートのレイアウトを作成
    setupTemplateSheetLayout(templateSheet);
  }
  
  return templateSheet;
}

/**
 * テンプレートシートのレイアウトを設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet テンプレートシート
 */
function setupTemplateSheetLayout(sheet) {
  // 書類タイトル
  sheet.getRange('A1').setValue('見積書');
  sheet.getRange('A1').setFontSize(24).setFontWeight('bold');
  
  // 発行日
  sheet.getRange('E2').setValue('発行日：');
  sheet.getRange('E2').setFontWeight('bold');
  
  // 宛先情報
  sheet.getRange('A4').setValue('宛先会社名');
  sheet.getRange('A5').setValue('担当者名');
  sheet.getRange('A6').setValue('住所');
  
  // 発行元情報（右上）
  sheet.getRange('E4').setValue('株式会社サンプル');
  sheet.getRange('E5').setValue('営業部 山田太郎');
  sheet.getRange('E6').setValue('〒000-0000 東京都●●区●●');
  sheet.getRange('E7').setValue('TEL: 03-0000-0000');
  sheet.getRange('E8').setValue('EMAIL: sample@example.com');
  
  // 明細ヘッダー
  const itemHeaders = ['品目', '数量', '単価', '小計'];
  itemHeaders.forEach((header, index) => {
    sheet.getRange(9, index + 1).setValue(header);
    sheet.getRange(9, index + 1).setFontWeight('bold').setBackground('#e6f3ff');
  });
  
  // 罫線を追加
  const itemRange = sheet.getRange('A9:D14');
  itemRange.setBorder(true, true, true, true, true, true);
  
  // 合計欄
  sheet.getRange('E15').setValue('小計');
  sheet.getRange('E16').setValue('消費税');
  sheet.getRange('E17').setValue('合計');
  
  sheet.getRange('E15:F17').setFontWeight('bold');
  sheet.getRange('F15:F17').setBorder(true, true, true, true, false, false);
  
  // 備考欄
  sheet.getRange('A19').setValue('備考：');
  sheet.getRange('A19').setFontWeight('bold');
}

/**
 * 初期セットアップを実行
 * 新しいスプレッドシートにシートを作成する
 */
function initialSetup() {
  try {
    createInputSheet();
    createTemplateSheet();
    
    SpreadsheetApp.getUi().alert('初期セットアップが完了しました。\n\n入力シートとテンプレートシートが作成されました。\n送信ボタンを配置して、sendDocument関数を割り当ててください。');
    
  } catch (error) {
    console.error('初期セットアップエラー:', error);
    SpreadsheetApp.getUi().alert(`初期セットアップでエラーが発生しました: ${error.message}`);
  }
}

/**
 * 数値フォーマット関数
 * @param {number} value 数値
 * @return {string} フォーマットされた文字列
 */
function formatCurrency(value) {
  if (!value || isNaN(value)) return '¥0';
  return `¥${value.toLocaleString()}`;
}

/**
 * 日付フォーマット関数
 * @param {Date} date 日付
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年MM月dd日');
}