/**
 * ユーティリティ関数
 * 共通で使用される補助的な関数を定義
 */

/**
 * スプレッドシートから必要なシートを一括取得してキャッシュ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 * @return {Object} シートオブジェクトのキャッシュ
 */
function getSheetCache(spreadsheet) {
  return {
    input: spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT),
    template: spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE),
    history: spreadsheet.getSheetByName(CONFIG.SHEETS.HISTORY),
    companyHistory: spreadsheet.getSheetByName(CONFIG.SHEETS.COMPANY_HISTORY)
  };
}

/**
 * 商品明細の範囲を動的に取得
 * @return {string} 範囲の文字列表現（例: 'A10:D29'）
 */
function getItemsRangeString() {
  const startRow = CONFIG.ITEMS_CONFIG.START_ROW;
  const endRow = startRow + CONFIG.ITEMS_CONFIG.MAX_ROWS - 1;
  return `A${startRow}:D${endRow}`;
}

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
 * 宛名履歴シートを取得または作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 宛名履歴シート
 */
function getOrCreateCompanyHistorySheet(spreadsheet) {
  let companyHistorySheet = spreadsheet.getSheetByName(CONFIG.SHEETS.COMPANY_HISTORY);
  
  if (!companyHistorySheet) {
    companyHistorySheet = spreadsheet.insertSheet(CONFIG.SHEETS.COMPANY_HISTORY);
    
    // ヘッダーを設定
    const headers = CONFIG.COMPANY_HISTORY_HEADERS;
    for (let i = 0; i < headers.length; i++) {
      companyHistorySheet.getRange(1, i + 1).setValue(headers[i]);
    }
    
    // ヘッダー行のフォーマット
    const headerRange = companyHistorySheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e6f3ff');
    
    // 列幅を調整
    companyHistorySheet.setColumnWidth(1, 200); // 会社名
    companyHistorySheet.setColumnWidth(2, 150); // 最終使用日時
    companyHistorySheet.setColumnWidth(3, 100); // 使用回数
  }
  
  return companyHistorySheet;
}

/**
 * 宛名履歴を更新
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 * @param {string} companyName 会社名
 * @param {GoogleAppsScript.Spreadsheet.Sheet} companyHistorySheet 宛名履歴シート（オプション）
 */
function updateCompanyHistory(spreadsheet, companyName, companyHistorySheet = null) {
  if (!companyName) return;
  
  if (!companyHistorySheet) {
    companyHistorySheet = getOrCreateCompanyHistorySheet(spreadsheet);
  }
  
  const lastRow = companyHistorySheet.getLastRow();
  const currentTime = new Date();
  
  // 既存の会社名を検索
  let foundRow = -1;
  if (lastRow > 1) {
    const companies = companyHistorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < companies.length; i++) {
      if (companies[i][0] === companyName) {
        foundRow = i + 2; // 2行目から開始なので+2
        break;
      }
    }
  }
  
  if (foundRow > 0) {
    // 既存の会社名の場合：最終使用日時と使用回数を更新（バッチ操作）
    const currentUsageCount = companyHistorySheet.getRange(foundRow, 3).getValue() || 0;
    const updateValues = [[currentTime, currentUsageCount + 1]];
    companyHistorySheet.getRange(foundRow, 2, 1, 2).setValues(updateValues);
  } else {
    // 新しい会社名の場合：新しい行を追加（バッチ操作）
    const newRow = lastRow + 1;
    const newValues = [[companyName, currentTime, 1]];
    companyHistorySheet.getRange(newRow, 1, 1, 3).setValues(newValues);
  }
  
  // 最終使用日時でソート（降順）
  if (companyHistorySheet.getLastRow() > 2) {
    const dataRange = companyHistorySheet.getRange(2, 1, companyHistorySheet.getLastRow() - 1, 3);
    dataRange.sort({column: 2, ascending: false});
  }
}

/**
 * 宛名履歴一覧を取得
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 * @param {number} limit 取得する件数の上限（デフォルト: 10）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} companyHistorySheet 宛名履歴シート（オプション）
 * @return {Array} 会社名の配列（最近使用した順）
 */
function getCompanyHistory(spreadsheet, limit = 10, companyHistorySheet = null) {
  if (!companyHistorySheet) {
    companyHistorySheet = spreadsheet.getSheetByName(CONFIG.SHEETS.COMPANY_HISTORY);
  }
  
  if (!companyHistorySheet || companyHistorySheet.getLastRow() <= 1) {
    return [];
  }
  
  const lastRow = companyHistorySheet.getLastRow();
  const actualLimit = Math.min(limit, lastRow - 1);
  
  if (actualLimit <= 0) return [];
  
  const companies = companyHistorySheet.getRange(2, 1, actualLimit, 1).getValues();
  return companies.map(row => row[0]).filter(name => name); // 空の値を除外
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
    // 新しいシートを作成
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
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#e6f3ff');
  sheet.getRange('A1:F1').merge();
  
  // 基本情報入力欄
  const labels = [
    ['書類種別', 'B2', 'お見積書 または ご請求書'],
    ['発行日', 'B3', '例: 2024/06/01'],
    ['書類番号', 'B4', '3桁の数字（例: 001）'],
    ['宛先会社名', 'B5', '必須（ドロップダウンから履歴選択可能）'],
    ['担当者名', 'B6', '任意'],
    ['住所', 'B7', '任意'],
    ['メールアドレス', 'B8', '必須'],
    ['備考', 'B9', '任意']
  ];
  
  labels.forEach(([label, cell, note], index) => {
    const row = index + 2;
    sheet.getRange(`A${row}`).setValue(label);
    sheet.getRange(`A${row}`).setFontWeight('bold').setBackground('#f0f0f0');
    if (note) {
      sheet.getRange(`C${row}`).setValue(note);
      sheet.getRange(`C${row}`).setFontStyle('italic').setFontColor('#666666');
    }
  });
  
  // 商品明細ヘッダー
  sheet.getRange('A9').setValue('商品明細');
  sheet.getRange('A9').setFontSize(14).setFontWeight('bold').setBackground('#ffe6cc');
  sheet.getRange('A9:D9').merge();
  
  const itemHeaders = ['品目', '数量', '単価', '小計'];
  itemHeaders.forEach((header, index) => {
    sheet.getRange(CONFIG.ITEMS_CONFIG.START_ROW, index + 1).setValue(header);
    sheet.getRange(CONFIG.ITEMS_CONFIG.START_ROW, index + 1).setFontWeight('bold').setBackground('#f0f0f0');
  });
  
  // 明細エリアに罫線 - 動的に範囲を計算
  const itemsEndRow = CONFIG.ITEMS_CONFIG.START_ROW + CONFIG.ITEMS_CONFIG.MAX_ROWS - 1;
  sheet.getRange(`A${CONFIG.ITEMS_CONFIG.START_ROW}:D${itemsEndRow}`).setBorder(true, true, true, true, true, true);
  
  // 合計欄
  sheet.getRange('C30').setValue('小計');
  sheet.getRange('C31').setValue('消費税');
  sheet.getRange('C32').setValue('合計');
  
  sheet.getRange('C30:C32').setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange('C30:D32').setBorder(true, true, true, true, false, false);
  
  // ボタン説明
  sheet.getRange('A34').setValue('操作ボタン');
  sheet.getRange('A34').setFontSize(14).setFontWeight('bold').setBackground('#ffcccc');
  
  // ボタン配置エリア
  sheet.getRange('A35').setValue('計算ボタン');
  sheet.getRange('B35').setValue('calculateTotals関数を割り当て');
  sheet.getRange('B35').setBackground('#e6ffe6');

  sheet.getRange('A36').setValue('送信ボタン');
  sheet.getRange('B36').setValue('sendDocument関数を割り当て');
  sheet.getRange('B36').setBackground('#ffe6e6');

  sheet.getRange('A37').setValue('クリアボタン');
  sheet.getRange('B37').setValue('clearInputData関数を割り当て');
  sheet.getRange('B37').setBackground('#e6e6ff');

  sheet.getRange('A38').setValue('宛名履歴ボタン');
  sheet.getRange('B38').setValue('showCompanyHistory関数を割り当て');
  sheet.getRange('B38').setBackground('#fff2e6');

  // エクスポート対象シート選択エリア
  sheet.getRange('E1').setValue('PDFエクスポート対象シート');
  sheet.getRange('E1').setFontSize(14).setFontWeight('bold').setBackground('#ffe6cc');
  sheet.getRange('E1:G1').merge();

  // エクスポート対象シート選択のヘッダー
  sheet.getRange('F2').setValue('シート名');
  sheet.getRange('G2').setValue('エクスポート');
  sheet.getRange('F2:G2').setFontWeight('bold').setBackground('#f0f0f0');

  // シート選択エリアを設定
  setupSheetSelectionArea(sheet);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 120); // A列
  sheet.setColumnWidth(2, 200); // B列
  sheet.setColumnWidth(3, 150); // C列
  sheet.setColumnWidth(4, 100); // D列
  sheet.setColumnWidth(5, 80);  // E列
  sheet.setColumnWidth(6, 150); // F列（シート名）
  sheet.setColumnWidth(7, 100); // G列（エクスポート）
  
  // データ検証の設定
  // 書類種別にドロップダウンを設定
  const documentTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['お見積書', 'ご請求書'])
    .setAllowInvalid(false)
    .setHelpText('お見積書またはご請求書を選択してください')
    .build();
  sheet.getRange(CONFIG.CELLS.DOCUMENT_TYPE).setDataValidation(documentTypeRule);
  
  // デフォルト日付を今日の日付に設定
  sheet.getRange(CONFIG.CELLS.ISSUE_DATE).setValue(new Date());
  
  // 書類番号をテキスト形式に設定し、ドロップダウンを追加（001-100）
  const documentNumberOptions = [];
  for (let i = 1; i <= 100; i++) {
    documentNumberOptions.push(i.toString().padStart(3, '0'));
  }
  const documentNumberRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(documentNumberOptions)
    .setAllowInvalid(true)
    .setHelpText('書類番号を選択または入力してください')
    .build();
  sheet.getRange(CONFIG.CELLS.DOCUMENT_NUMBER).setDataValidation(documentNumberRule);
  sheet.getRange(CONFIG.CELLS.DOCUMENT_NUMBER).setNumberFormat('@'); // テキスト形式
  
  // 宛先会社名のドロップダウンを設定（履歴から）
  setupCompanyNameDropdown(sheet);
}

/**
 * 宛先会社名のドロップダウンを設定
 * 宛名履歴シートから過去に使用した会社名を取得してドロップダウンに設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 入力シート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} companyHistorySheet 宛名履歴シート（オプション）
 */
function setupCompanyNameDropdown(sheet, companyHistorySheet = null) {
  try {
    if (!companyHistorySheet) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      companyHistorySheet = getOrCreateCompanyHistorySheet(spreadsheet);
    }
    
    let companyNames = [];
    
    // 履歴から会社名を取得（最新順）
    const lastRow = companyHistorySheet.getLastRow();
    if (lastRow > 1) {
      const companyData = companyHistorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
      companyNames = companyData.map(row => row[0]).filter(name => name && name.toString().trim() !== '');
    }
    
    // デフォルトの会社名を追加（履歴がない場合のみ）
    if (companyNames.length === 0) {
      companyNames = ['株式会社サンプル', '有限会社テスト', '個人事業主様'];
    }
    
    // 重複を除去し、最新順を維持
    companyNames = [...new Set(companyNames)];
    
    // 最大50件に制限（ドロップダウンの性能を考慮）
    if (companyNames.length > 50) {
      companyNames = companyNames.slice(0, 50);
    }
    
    // ドロップダウンを設定
    const helpText = companyNames.length > 0 
      ? `宛名履歴から選択または新しい会社名を入力してください（履歴: ${companyNames.length}件）`
      : '宛先会社名を選択または入力してください';
    
    const companyNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(companyNames)
      .setAllowInvalid(true)
      .setHelpText(helpText)
      .build();
    sheet.getRange(CONFIG.CELLS.COMPANY_NAME).setDataValidation(companyNameRule);
    
    console.log(`会社名ドロップダウンを更新しました（${companyNames.length}件の履歴）`);
    
  } catch (error) {
    console.error('会社名ドロップダウン設定エラー:', error);
    // エラーが発生した場合は基本的なドロップダウンを設定
    const defaultCompanies = ['株式会社サンプル', '有限会社テスト', '個人事業主様'];
    const companyNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(defaultCompanies)
      .setAllowInvalid(true)
      .setHelpText('宛先会社名を選択または入力してください（履歴読み込みエラー）')
      .build();
    sheet.getRange(CONFIG.CELLS.COMPANY_NAME).setDataValidation(companyNameRule);
  }
}

/**
 * スプレッドシートにテンプレートシートを作成
 */
function createTemplateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  if (!templateSheet) {
    // 新しいシートを作成
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
  sheet.getRange('A1').setValue('お見積書');
  sheet.getRange('A1').setFontSize(24).setFontWeight('bold');
  
  // 発行日
  sheet.getRange('D2').setValue('発行日：');
  sheet.getRange('D2').setFontWeight('bold');
  
  // 書類番号
  sheet.getRange('D3').setValue('書類番号：');
  sheet.getRange('D3').setFontWeight('bold');
  
  // 宛先情報
  sheet.getRange('A4').setValue('宛先会社名');
  sheet.getRange('A5').setValue('担当者名');
  sheet.getRange('A6').setValue('住所');
  
  // 発行元情報（D列に移動）
  sheet.getRange('D4').setValue('株式会社サンプル');
  sheet.getRange('D5').setValue('営業部 山田太郎');
  sheet.getRange('D6').setValue('〒000-0000 東京都●●区●●');
  sheet.getRange('D7').setValue('TEL: 03-0000-0000');
  sheet.getRange('D8').setValue('EMAIL: sample@example.com');
  
  // 明細ヘッダー
  const itemHeaders = ['品目', '数量', '単価', '小計'];
  const headerRow = CONFIG.ITEMS_CONFIG.START_ROW - 1; // ヘッダーは明細開始行の1行上
  itemHeaders.forEach((header, index) => {
    sheet.getRange(headerRow, index + 1).setValue(header);
    sheet.getRange(headerRow, index + 1).setFontWeight('bold').setBackground('#e6f3ff');
  });
  
  // 罫線を追加 - 動的に範囲を計算
  const itemsEndRow = CONFIG.ITEMS_CONFIG.START_ROW + CONFIG.ITEMS_CONFIG.MAX_ROWS - 1;
  const itemRange = sheet.getRange(`A${headerRow}:D${itemsEndRow}`);
  itemRange.setBorder(true, true, true, true, true, true);
  
  // 合計欄
  sheet.getRange('C30').setValue('小計');
  sheet.getRange('C31').setValue('消費税');
  sheet.getRange('C32').setValue('合計');
  
  sheet.getRange('C30:D32').setFontWeight('bold');
  sheet.getRange('D30:D32').setBorder(true, true, true, true, false, false);
  
  // 備考欄
  sheet.getRange('A33').setValue('備考：');
  sheet.getRange('A33').setFontWeight('bold');
  
  // 備考エリアの範囲設定（33〜47行）
  const remarksRange = sheet.getRange(`A${CONFIG.TEMPLATE_RANGES.REMARKS_START_ROW}:F${CONFIG.TEMPLATE_RANGES.REMARKS_END_ROW}`);
  remarksRange.setBorder(true, true, true, true, false, false);
}

/**
 * 初期セットアップを実行
 * 新しいスプレッドシートにシートを作成する
 * 既存シートがある場合はバックアップしてから新しいシートを作成する
 */
function initialSetup() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存テンプレートシートから書式情報を保存
    const existingTemplateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
    let templateFormatting = null;
    if (existingTemplateSheet) {
      templateFormatting = inheritTemplateFormatting(existingTemplateSheet);
    }
    
    // 既存シートのバックアップ
    backupExistingSheetsIfNeeded(spreadsheet);
    
    // 列の大きさを継承するための参照シート（シート1）を取得
    const referenceSheet = spreadsheet.getSheets()[0]; // 最初のシート
    
    const inputSheet = createInputSheet();
    const templateSheet = createTemplateSheet();
    
    // 保存した書式を新しいテンプレートシートに適用
    if (templateFormatting && templateSheet) {
      applyTemplateFormatting(templateSheet, templateFormatting);
    }
    
    // 列幅を継承
    if (referenceSheet && inputSheet) {
      inheritColumnWidths(referenceSheet, inputSheet);
    }
    if (referenceSheet && templateSheet) {
      inheritColumnWidths(referenceSheet, templateSheet);
    }
    
    // デフォルトの表示行数を設定
    if (inputSheet) {
      adjustItemRowsVisibility(inputSheet, CONFIG.ITEMS_CONFIG.DEFAULT_VISIBLE_ROWS);
      // シート選択エリアを設定
      setupSheetSelectionArea(inputSheet);
    }
    if (templateSheet) {
      adjustItemRowsVisibility(templateSheet, CONFIG.ITEMS_CONFIG.DEFAULT_VISIBLE_ROWS);
    }
    
    SpreadsheetApp.getUi().alert('初期セットアップが完了しました。\n\n入力シートとテンプレートシートが作成されました。\n既存シートはバックアップされました。\n送信ボタンを配置して、sendDocument関数を割り当ててください。');
    
  } catch (error) {
    console.error('初期セットアップエラー:', error);
    SpreadsheetApp.getUi().alert(`初期セットアップでエラーが発生しました: ${error.message}`);
  }
}

/**
 * 既存シートのバックアップ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet スプレッドシート
 */
function backupExistingSheetsIfNeeded(spreadsheet) {
  const inputSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
  const templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  if (inputSheet || templateSheet) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    
    if (inputSheet) {
      const newName = `${CONFIG.SHEETS.INPUT}_backup_${timestamp}`;
      inputSheet.setName(newName);
      console.log(`入力シートを ${newName} にバックアップしました`);
    }
    
    if (templateSheet) {
      const newName = `${CONFIG.SHEETS.TEMPLATE}_backup_${timestamp}`;
      templateSheet.setName(newName);
      console.log(`テンプレートシートを ${newName} にバックアップしました`);
    }
  }
}

/**
 * 列幅を継承
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet 参照元シート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet 適用先シート
 */
function inheritColumnWidths(sourceSheet, targetSheet) {
  try {
    // 最初の10列の幅を継承
    for (let col = 1; col <= 10; col++) {
      const width = sourceSheet.getColumnWidth(col);
      targetSheet.setColumnWidth(col, width);
    }
    console.log(`${targetSheet.getName()} の列幅を ${sourceSheet.getName()} から継承しました`);
  } catch (error) {
    console.error('列幅継承エラー:', error);
  }
}

/**
 * テンプレートシートの書式を保存
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet 既存のテンプレートシート
 * @return {Object} 書式情報
 */
function inheritTemplateFormatting(templateSheet) {
  try {
    const formatting = {};
    
    // 重要な範囲の書式を保存
    const ranges = [
      'A1:D50', // 全体的な範囲
    ];
    
    ranges.forEach(rangeAddress => {
      const range = templateSheet.getRange(rangeAddress);
      formatting[rangeAddress] = {
        backgrounds: range.getBackgrounds(),
        fontColors: range.getFontColors(),
        fontFamilies: range.getFontFamilies(),
        fontSizes: range.getFontSizes(),
        fontWeights: range.getFontWeights(),
        borders: getBorders(range)
      };
    });
    
    console.log('テンプレートシートの書式を保存しました');
    return formatting;
  } catch (error) {
    console.error('書式保存エラー:', error);
    return null;
  }
}

/**
 * 保存した書式をテンプレートシートに適用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet 新しいテンプレートシート
 * @param {Object} formatting 書式情報
 */
function applyTemplateFormatting(templateSheet, formatting) {
  try {
    Object.keys(formatting).forEach(rangeAddress => {
      const range = templateSheet.getRange(rangeAddress);
      const format = formatting[rangeAddress];
      
      if (format.backgrounds) range.setBackgrounds(format.backgrounds);
      if (format.fontColors) range.setFontColors(format.fontColors);
      if (format.fontFamilies) range.setFontFamilies(format.fontFamilies);
      if (format.fontSizes) range.setFontSizes(format.fontSizes);
      if (format.fontWeights) range.setFontWeights(format.fontWeights);
      if (format.borders) applyBorders(range, format.borders);
    });
    
    console.log('テンプレートシートに書式を適用しました');
  } catch (error) {
    console.error('書式適用エラー:', error);
  }
}

/**
 * 範囲の罫線情報を取得（簡易版）
 * @param {GoogleAppsScript.Spreadsheet.Range} range 範囲
 * @return {Object} 罫線情報
 */
function getBorders(range) {
  // Google Apps Scriptでは罫線情報を直接取得できないため、簡易的な実装
  return { simplified: true };
}

/**
 * 罫線を適用（簡易版）
 * @param {GoogleAppsScript.Spreadsheet.Range} range 範囲
 * @param {Object} borders 罫線情報
 */
function applyBorders(range, borders) {
  // 簡易的な罫線適用（既存の罫線設定を維持）
  if (borders.simplified) {
    // 基本的な罫線を設定
    range.setBorder(true, true, true, true, false, false);
  }
}

/**
 * 入力シートのドロップダウンを更新
 * 宛名履歴シートから最新の会社名リストを取得してドロップダウンを更新
 */
function refreshInputSheetDropdowns() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = getSheetCache(spreadsheet);
    
    if (!sheetCache.input) {
      SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 宛名履歴から会社名数を取得
    let companyCount = 0;
    if (sheetCache.companyHistory && sheetCache.companyHistory.getLastRow() > 1) {
      companyCount = sheetCache.companyHistory.getLastRow() - 1;
    }
    
    // ドロップダウンを更新（キャッシュされたシートを使用）
    setupCompanyNameDropdown(sheetCache.input, sheetCache.companyHistory);
    
    const message = companyCount > 0 
      ? `宛先会社名のドロップダウンを更新しました。\n\n宛名履歴: ${companyCount}件の会社名が利用可能です。\n\nB5セルのドロップダウンから選択できます。`
      : `宛先会社名のドロップダウンを更新しました。\n\n宛名履歴がまだありません。\n新しい会社名を入力すると、次回から履歴に表示されます。`;
    
    SpreadsheetApp.getUi().alert('ドロップダウン更新完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('ドロップダウン更新エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `ドロップダウンの更新中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
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

/**
 * 明細の自動計算を実行
 * 小計、消費税、合計を自動計算する
 */
function calculateTotals() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = getSheetCache(spreadsheet);
    
    if (!sheetCache.input) {
      SpreadsheetApp.getUi().alert('入力シートが見つかりません。');
      return;
    }
    
    // 明細の小計を計算
    const itemsRange = sheetCache.input.getRange(getItemsRangeString());
    const values = itemsRange.getValues();
    
    let subtotal = 0;
    const updatedValues = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (row[0] && row[1] && row[2]) { // 品目、数量、単価が入力されている場合
        const quantity = parseFloat(row[1]) || 0;
        const unitPrice = parseFloat(row[2]) || 0;
        const itemSubtotal = quantity * unitPrice;
        
        // 更新する値を配列に追加
        updatedValues.push({ row: CONFIG.ITEMS_CONFIG.START_ROW + i, value: itemSubtotal });
        subtotal += itemSubtotal;
      }
    }
    
    // 小計をバッチで更新
    updatedValues.forEach(update => {
      sheetCache.input.getRange(update.row, 4).setValue(update.value);
    });
    
    // 消費税率（10%）
    const taxRate = 0.1;
    const tax = Math.floor(subtotal * taxRate);
    const grandTotal = subtotal + tax;
    
    // 合計金額をバッチで設定
    const totalValues = [[subtotal], [tax], [grandTotal]];
    sheetCache.input.getRange(`${CONFIG.CELLS.TOTAL_AMOUNT}:${CONFIG.CELLS.GRAND_TOTAL}`).setValues(totalValues);
    
    SpreadsheetApp.getUi().alert('計算完了', `合計金額を計算しました。\n\n小計: ${formatCurrency(subtotal)}\n消費税: ${formatCurrency(tax)}\n合計: ${formatCurrency(grandTotal)}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('計算エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `計算中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * データ入力のクリア
 * 入力フォームをクリアする
 */
function clearInputData() {
  try {
    const result = SpreadsheetApp.getUi().alert(
      '入力データクリア',
      '入力されたデータをすべてクリアします。よろしいですか？',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (result !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = getSheetCache(spreadsheet);
    
    if (!sheetCache.input) {
      SpreadsheetApp.getUi().alert('入力シートが見つかりません。');
      return;
    }
    
    // 基本情報をバッチでクリア
    const basicCells = [
      CONFIG.CELLS.DOCUMENT_TYPE,
      CONFIG.CELLS.ISSUE_DATE,
      CONFIG.CELLS.DOCUMENT_NUMBER,
      CONFIG.CELLS.COMPANY_NAME,
      CONFIG.CELLS.CONTACT_NAME,
      CONFIG.CELLS.ADDRESS,
      CONFIG.CELLS.EMAIL,
      CONFIG.CELLS.REMARKS
    ];
    
    basicCells.forEach(cellAddress => {
      sheetCache.input.getRange(cellAddress).clearContent();
    });
    
    // 明細をクリア
    sheetCache.input.getRange(getItemsRangeString()).clearContent();
    
    // 合計金額をバッチでクリア
    sheetCache.input.getRange(`${CONFIG.CELLS.TOTAL_AMOUNT}:${CONFIG.CELLS.GRAND_TOTAL}`).clearContent();
    
    // シート選択をリセット（テンプレートシートのみ選択状態に）
    setupSheetSelectionArea(sheetCache.input);
    
    SpreadsheetApp.getUi().alert('入力データをクリアしました。');
    
  } catch (error) {
    console.error('クリアエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `クリア中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 宛名履歴一覧を表示
 * 過去に使用した宛先会社名を表示して選択可能にする
 */
function showCompanyHistory() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = getSheetCache(spreadsheet);
    const companyHistory = getCompanyHistory(spreadsheet, 20); // 最大20件取得
    
    if (companyHistory.length === 0) {
      SpreadsheetApp.getUi().alert('宛名履歴', '過去に使用した宛先会社名がありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 履歴一覧を表示
    let message = '📋 宛名履歴（最近使用した順）\n\n';
    message += '以下の会社名をコピーして入力シートの「宛先会社名」欄に貼り付けできます：\n\n';
    
    companyHistory.forEach((company, index) => {
      message += `${index + 1}. ${company}\n`;
    });
    
    message += '\n※会社名をクリップボードにコピーするには、この後表示される入力欄に番号を入力してください。';
    
    // 番号選択のプロンプト
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '宛名履歴',
      message + '\n\n会社名をクリップボードにコピーしたい場合は、番号を入力してください（1-' + companyHistory.length + '）：',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const input = response.getResponseText().trim();
      const selectedIndex = parseInt(input) - 1;
      
      if (selectedIndex >= 0 && selectedIndex < companyHistory.length) {
        const selectedCompany = companyHistory[selectedIndex];
        
        // 入力シートの会社名欄に自動入力（キャッシュされたシートを使用）
        if (sheetCache.input) {
          sheetCache.input.getRange(CONFIG.CELLS.COMPANY_NAME).setValue(selectedCompany);
          ui.alert('宛名設定完了', `「${selectedCompany}」を宛先会社名欄に設定しました。`, ui.ButtonSet.OK);
        } else {
          ui.alert('選択完了', `選択された会社名：「${selectedCompany}」\n\n手動で宛先会社名欄にコピーしてください。`, ui.ButtonSet.OK);
        }
      } else {
        ui.alert('エラー', '無効な番号です。', ui.ButtonSet.OK);
      }
    }
    
  } catch (error) {
    console.error('宛名履歴表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `宛名履歴の表示中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システム設定確認
 * 必要なシートやフォルダの存在確認とシステム状態をチェック
 */
function checkSystemStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const issues = [];
    const info = [];
    
    // シートの存在確認
    const inputSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
    const templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
    const historySheet = spreadsheet.getSheetByName(CONFIG.SHEETS.HISTORY);
    const companyHistorySheet = spreadsheet.getSheetByName(CONFIG.SHEETS.COMPANY_HISTORY);
    
    if (!inputSheet) {
      issues.push('- 入力シートが存在しません');
    } else {
      info.push('✅ 入力シート: OK');
    }
    
    if (!templateSheet) {
      issues.push('- テンプレートシートが存在しません');
    } else {
      info.push('✅ テンプレートシート: OK');
    }
    
    if (!historySheet) {
      info.push('📋 送信履歴シート: 初回送信時に作成されます');
    } else {
      info.push('✅ 送信履歴シート: OK');
    }
    
    if (!companyHistorySheet) {
      info.push('📋 宛名履歴シート: 初回送信時に作成されます');
    } else {
      const companyCount = Math.max(0, companyHistorySheet.getLastRow() - 1);
      info.push(`✅ 宛名履歴シート: OK (${companyCount}件の宛名を記録済み)`);
    }
    
    // フォルダの存在確認
    const parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
    
    const estimateFolder = parentFolder.getFoldersByName(CONFIG.FOLDERS.ESTIMATES);
    const invoiceFolder = parentFolder.getFoldersByName(CONFIG.FOLDERS.INVOICES);
    const backupFolder = parentFolder.getFoldersByName(CONFIG.FOLDERS.BACKUP);
    
    if (!estimateFolder.hasNext()) {
      issues.push('- 見積書フォルダが存在しません');
    } else {
      info.push('✅ 見積書フォルダ: OK');
    }
    
    if (!invoiceFolder.hasNext()) {
      issues.push('- 請求書フォルダが存在しません');
    } else {
      info.push('✅ 請求書フォルダ: OK');
    }
    
    if (!backupFolder.hasNext()) {
      issues.push('- バックアップフォルダが存在しません');
    } else {
      info.push('✅ バックアップフォルダ: OK');
    }
    
    // 設定情報の表示
    info.push('');
    info.push('📧 メール送信者設定:');
    info.push(`   会社名: ${CONFIG.EMAIL.SENDER_COMPANY}`);
    info.push(`   部署: ${CONFIG.EMAIL.SENDER_DEPARTMENT}`);
    info.push(`   担当者: ${CONFIG.EMAIL.SENDER_NAME}`);
    
    // 結果表示
    let message = '📋 システム状態確認結果\n\n';
    
    if (issues.length > 0) {
      message += '⚠️ 以下の問題が見つかりました:\n';
      message += issues.join('\n') + '\n\n';
      message += '解決方法:\n';
      message += '- シートの問題: initialSetup()関数を実行\n';
      message += '- フォルダの問題: 手動でフォルダを作成\n\n';
    }
    
    message += '📊 システム情報:\n';
    message += info.join('\n');
    
    SpreadsheetApp.getUi().alert('システム状態確認', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('システム確認エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `システム確認中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 商品明細の表示行数を動的に調整
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} visibleRows 表示する行数（1-20）
 */
function adjustItemRowsVisibility(sheet, visibleRows = CONFIG.ITEMS_CONFIG.DEFAULT_VISIBLE_ROWS) {
  const maxRows = CONFIG.ITEMS_CONFIG.MAX_ROWS;
  const startRow = CONFIG.ITEMS_CONFIG.START_ROW;
  
  // 入力値の検証
  if (visibleRows < 1) visibleRows = 1;
  if (visibleRows > maxRows) visibleRows = maxRows;
  
  // 全ての明細行の背景色をリセット
  const allRowsRange = sheet.getRange(startRow, 1, maxRows, 4);
  allRowsRange.setBackground('#ffffff');
  
  // 使用する行数に応じて背景色を設定
  if (visibleRows > 0) {
    const visibleRange = sheet.getRange(startRow, 1, visibleRows, 4);
    visibleRange.setBackground('#f9f9f9'); // 薄いグレー
  }
  
  // 使用しない行の背景色を薄く設定
  if (visibleRows < maxRows) {
    const hiddenRange = sheet.getRange(startRow + visibleRows, 1, maxRows - visibleRows, 4);
    hiddenRange.setBackground('#f5f5f5'); // より薄いグレー
  }
}

/**
 * 商品明細行数を設定するメニュー関数
 */
function setItemRowCount() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '明細行数設定',
      `現在の最大行数: ${CONFIG.ITEMS_CONFIG.MAX_ROWS}行\n\n表示する明細行数を入力してください（1-${CONFIG.ITEMS_CONFIG.MAX_ROWS}）：`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const input = response.getResponseText().trim();
      const rowCount = parseInt(input);
      
      if (isNaN(rowCount) || rowCount < 1 || rowCount > CONFIG.ITEMS_CONFIG.MAX_ROWS) {
        ui.alert('エラー', `1から${CONFIG.ITEMS_CONFIG.MAX_ROWS}までの数値を入力してください。`, ui.ButtonSet.OK);
        return;
      }
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const inputSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
      const templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
      
      if (inputSheet) {
        adjustItemRowsVisibility(inputSheet, rowCount);
      }
      
      if (templateSheet) {
        adjustItemRowsVisibility(templateSheet, rowCount);
      }
      
      ui.alert('設定完了', `明細行数を${rowCount}行に設定しました。`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('明細行数設定エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `明細行数設定中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * テスト機能：設定の妥当性をチェック
 */
function testConfiguration() {
  try {
    const ui = SpreadsheetApp.getUi();
    let testResults = [];
    
    // 設定値の確認
    testResults.push(`✅ 商品明細最大行数: ${CONFIG.ITEMS_CONFIG.MAX_ROWS}`);
    testResults.push(`✅ 商品明細開始行: ${CONFIG.ITEMS_CONFIG.START_ROW}`);
    testResults.push(`✅ デフォルト表示行数: ${CONFIG.ITEMS_CONFIG.DEFAULT_VISIBLE_ROWS}`);
    
    // 計算された範囲の確認
    const calculatedRange = getItemsRangeString();
    testResults.push(`✅ 動的計算範囲: ${calculatedRange}`);
    
    // セル位置の確認
    testResults.push(`✅ 小計セル: ${CONFIG.CELLS.TOTAL_AMOUNT}`);
    testResults.push(`✅ 消費税セル: ${CONFIG.CELLS.TAX}`);
    testResults.push(`✅ 合計セル: ${CONFIG.CELLS.GRAND_TOTAL}`);
    
    // テンプレートセル位置の確認
    testResults.push(`✅ テンプレート小計: ${CONFIG.TEMPLATE_CELLS.TOTAL_AMOUNT}`);
    testResults.push(`✅ テンプレート消費税: ${CONFIG.TEMPLATE_CELLS.TAX}`);
    testResults.push(`✅ テンプレート合計: ${CONFIG.TEMPLATE_CELLS.GRAND_TOTAL}`);
    testResults.push(`✅ テンプレート備考: ${CONFIG.TEMPLATE_CELLS.REMARKS}`);
    
    // 範囲の妥当性チェック
    const endRow = CONFIG.ITEMS_CONFIG.START_ROW + CONFIG.ITEMS_CONFIG.MAX_ROWS - 1;
    const totalRow = parseInt(CONFIG.CELLS.TOTAL_AMOUNT.substring(1));
    
    if (totalRow > endRow) {
      testResults.push(`✅ 合計行位置は適切です (${totalRow} > ${endRow})`);
    } else {
      testResults.push(`❌ 合計行位置が明細行と重複しています (${totalRow} <= ${endRow})`);
    }
    
    const message = '🔧 設定テスト結果\n\n' + testResults.join('\n');
    ui.alert('設定テスト', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('設定テストエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `設定テスト中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * エクスポート対象シート選択エリアを設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 入力シート
 */
function setupSheetSelectionArea(sheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = spreadsheet.getSheets();
  
  // システムシート（入力、送信履歴、宛名履歴）は除外
  const excludeSheets = [CONFIG.SHEETS.INPUT, CONFIG.SHEETS.HISTORY, CONFIG.SHEETS.COMPANY_HISTORY];
  const exportableSheets = allSheets.filter(s => !excludeSheets.includes(s.getName()));
  
  // シート選択エリアをクリア
  const clearRange = sheet.getRange('F3:G10');
  clearRange.clearContent();
  clearRange.clearDataValidations();
  
  if (exportableSheets.length === 0) {
    // エクスポート可能なシートがない場合
    sheet.getRange('F3').setValue('エクスポート可能なシートがありません');
    sheet.getRange('F3').setFontColor('#ff0000');
    return;
  }
  
  // 各シートのチェックボックスを設定
  exportableSheets.forEach((sheetObj, index) => {
    if (index < 8) { // 最大8シートまで表示
      const row = 3 + index;
      
      // シート名を表示
      sheet.getRange(row, 6).setValue(sheetObj.getName());
      
      // チェックボックスを設定
      const checkbox = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .build();
      
      sheet.getRange(row, 7).setDataValidation(checkbox);
      
      // テンプレートシートはデフォルトでチェック
      if (sheetObj.getName() === CONFIG.SHEETS.TEMPLATE) {
        sheet.getRange(row, 7).setValue(true);
      } else {
        sheet.getRange(row, 7).setValue(false);
      }
    }
  });
  
  // 8シート以上ある場合の注意書き
  if (exportableSheets.length > 8) {
    sheet.getRange('F11').setValue(`注意: ${exportableSheets.length - 8}個のシートが表示されていません`);
    sheet.getRange('F11').setFontColor('#ff8800');
    sheet.getRange('F11').setFontSize(9);
  }
}

/**
 * エクスポート対象として選択されたシート名を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet 入力シート
 * @return {Array<string>} 選択されたシート名の配列
 */
function getSelectedSheetsForExport(inputSheet) {
  const selectedSheets = [];
  
  // エクスポート選択エリアをチェック
  for (let row = 3; row <= 10; row++) {
    const sheetName = inputSheet.getRange(row, 6).getValue();
    const isSelected = inputSheet.getRange(row, 7).getValue();
    
    if (sheetName && isSelected === true) {
      selectedSheets.push(sheetName);
    }
  }
  
  return selectedSheets;
}

/**
 * シート選択エリアを更新するメニュー関数
 */
function refreshSheetSelection() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = getSheetCache(spreadsheet);
    
    if (!sheetCache.input) {
      SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    setupSheetSelectionArea(sheetCache.input);
    SpreadsheetApp.getUi().alert('更新完了', 'シート選択エリアを更新しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('シート選択更新エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', `シート選択更新中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}