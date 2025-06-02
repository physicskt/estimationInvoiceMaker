/**
 * メイン処理：見積書・請求書の作成・送信・保存
 * スプレッドシート上のボタンから呼び出される
 */
function sendDocument() {
  try {
    // 処理開始のメッセージ
    SpreadsheetApp.getUi().alert('処理を開始します', '見積書・請求書の作成を開始します。しばらくお待ちください。', SpreadsheetApp.getUi().ButtonSet.OK);
    
    // 現在のスプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 入力データを取得
    const inputData = getInputData(spreadsheet);
    
    // 入力データの検証
    if (!validateInputData(inputData)) {
      return; // エラーメッセージは validateInputData 内で表示
    }
    
    // 確認ダイアログ
    const confirmResult = SpreadsheetApp.getUi().alert(
      '送信確認',
      `${inputData.documentType}を以下の宛先に送信します。\n\n宛先: ${inputData.companyName}\nメール: ${inputData.email}\n\n送信しますか？`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (confirmResult !== SpreadsheetApp.getUi().Button.YES) {
      SpreadsheetApp.getUi().alert('処理をキャンセルしました。');
      return;
    }
    
    // PDFを生成
    const pdfBlob = generatePDF(spreadsheet, inputData);
    
    // PDFをフォルダに保存
    const savedFile = savePDFToFolder(pdfBlob, inputData);
    
    // メール送信
    sendEmailWithPDF(pdfBlob, inputData);
    
    // 送信履歴を記録
    recordSendingHistory(spreadsheet, inputData, savedFile);
    
    // バックアップを作成
    createBackupDocument(inputData, savedFile);
    
    // 完了メッセージを表示
    SpreadsheetApp.getUi().alert(
      '送信完了', 
      `${inputData.documentType}を正常に送信しました。\n\n宛先: ${inputData.email}\nファイル名: ${savedFile.getName()}\n保存先: ${savedFile.getUrl()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', `処理中にエラーが発生しました:\n\n${error.message}\n\n管理者にお問い合わせください。`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * スプレッドシートから入力データを取得
 */
function getInputData(spreadsheet) {
  const inputSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
  
  const data = {
    documentType: inputSheet.getRange(CONFIG.CELLS.DOCUMENT_TYPE).getValue(),
    issueDate: inputSheet.getRange(CONFIG.CELLS.ISSUE_DATE).getValue(),
    companyName: inputSheet.getRange(CONFIG.CELLS.COMPANY_NAME).getValue(),
    contactName: inputSheet.getRange(CONFIG.CELLS.CONTACT_NAME).getValue(),
    address: inputSheet.getRange(CONFIG.CELLS.ADDRESS).getValue(),
    email: inputSheet.getRange(CONFIG.CELLS.EMAIL).getValue(),
    remarks: inputSheet.getRange(CONFIG.CELLS.REMARKS).getValue(),
    items: getItemsData(inputSheet),
    totalAmount: inputSheet.getRange(CONFIG.CELLS.TOTAL_AMOUNT).getValue(),
    tax: inputSheet.getRange(CONFIG.CELLS.TAX).getValue(),
    grandTotal: inputSheet.getRange(CONFIG.CELLS.GRAND_TOTAL).getValue()
  };
  
  return data;
}

/**
 * 商品明細データを取得
 */
function getItemsData(inputSheet) {
  const itemsRange = inputSheet.getRange(CONFIG.RANGES.ITEMS);
  const values = itemsRange.getValues();
  
  const items = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[0]) { // 品目名が入力されている行のみ
      items.push({
        name: row[0],
        quantity: row[1],
        unitPrice: row[2],
        subtotal: row[3]
      });
    }
  }
  
  return items;
}

/**
 * 入力データの検証
 */
function validateInputData(data) {
  const errors = [];
  
  // 必須項目のチェック
  if (!data.documentType) {
    errors.push('書類種別が入力されていません');
  } else if (data.documentType !== '見積書' && data.documentType !== '請求書') {
    errors.push('書類種別は「見積書」または「請求書」を入力してください');
  }
  
  if (!data.companyName) {
    errors.push('宛先会社名が入力されていません');
  }
  
  if (!data.email) {
    errors.push('メールアドレスが入力されていません');
  } else {
    // メールアドレスの形式チェック
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(data.email)) {
      errors.push('メールアドレスの形式が正しくありません');
    }
  }
  
  if (!data.issueDate) {
    errors.push('発行日が入力されていません');
  } else if (!(data.issueDate instanceof Date)) {
    errors.push('発行日の形式が正しくありません');
  }
  
  // 商品明細のチェック
  if (!data.items || data.items.length === 0) {
    errors.push('商品明細が入力されていません');
  } else {
    // 各明細の妥当性チェック
    for (let i = 0; i < data.items.length; i++) {
      const item = data.items[i];
      if (!item.name) {
        errors.push(`商品明細${i + 1}行目: 品目名が入力されていません`);
      }
      if (!item.quantity || item.quantity <= 0) {
        errors.push(`商品明細${i + 1}行目: 数量が正しく入力されていません`);
      }
      if (!item.unitPrice || item.unitPrice <= 0) {
        errors.push(`商品明細${i + 1}行目: 単価が正しく入力されていません`);
      }
    }
  }
  
  // 金額のチェック
  if (!data.grandTotal || data.grandTotal <= 0) {
    errors.push('合計金額が正しく計算されていません');
  }
  
  // エラーがある場合はアラートで表示
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert('入力エラー', errors.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  
  return true;
}

/**
 * PDFを生成
 */
function generatePDF(spreadsheet, inputData) {
  // テンプレートシートを取得
  const templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  // テンプレートシートにデータを反映
  updateTemplateSheet(templateSheet, inputData);
  
  // PDFとして出力
  const pdfBlob = DriveApp.getFileById(spreadsheet.getId()).getAs('application/pdf');
  
  // ファイル名を設定
  const fileName = generateFileName(inputData);
  pdfBlob.setName(fileName);
  
  return pdfBlob;
}

/**
 * テンプレートシートにデータを反映
 */
function updateTemplateSheet(templateSheet, inputData) {
  // 基本情報を設定
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.DOCUMENT_TYPE).setValue(inputData.documentType);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.ISSUE_DATE).setValue(Utilities.formatDate(inputData.issueDate, 'Asia/Tokyo', 'yyyy年MM月dd日'));
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.COMPANY_NAME).setValue(inputData.companyName);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.CONTACT_NAME).setValue(inputData.contactName);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.ADDRESS).setValue(inputData.address);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.REMARKS).setValue(inputData.remarks);
  
  // 商品明細を設定
  updateItemsInTemplate(templateSheet, inputData.items);
  
  // 合計金額を設定
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.TOTAL_AMOUNT).setValue(inputData.totalAmount);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.TAX).setValue(inputData.tax);
  templateSheet.getRange(CONFIG.TEMPLATE_CELLS.GRAND_TOTAL).setValue(inputData.grandTotal);
}

/**
 * テンプレートシートの商品明細を更新
 */
function updateItemsInTemplate(templateSheet, items) {
  const startRow = CONFIG.TEMPLATE_RANGES.ITEMS_START_ROW;
  
  // 既存の明細をクリア
  const clearRange = templateSheet.getRange(startRow, 1, CONFIG.TEMPLATE_RANGES.ITEMS_MAX_ROWS, 4);
  clearRange.clear();
  
  // 新しい明細を設定
  for (let i = 0; i < items.length && i < CONFIG.TEMPLATE_RANGES.ITEMS_MAX_ROWS; i++) {
    const item = items[i];
    const row = startRow + i;
    
    templateSheet.getRange(row, 1).setValue(item.name);
    templateSheet.getRange(row, 2).setValue(item.quantity);
    templateSheet.getRange(row, 3).setValue(item.unitPrice);
    templateSheet.getRange(row, 4).setValue(item.subtotal);
  }
}

/**
 * PDFファイル名を生成
 */
function generateFileName(inputData) {
  const date = Utilities.formatDate(inputData.issueDate, 'Asia/Tokyo', 'yyyyMMdd');
  return `${inputData.documentType}_${inputData.companyName}_${date}.pdf`;
}

/**
 * PDFをフォルダに保存
 */
function savePDFToFolder(pdfBlob, inputData) {
  const parentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  
  // 書類種別に応じてフォルダを取得または作成
  const folderName = inputData.documentType === '見積書' ? CONFIG.FOLDERS.ESTIMATES : CONFIG.FOLDERS.INVOICES;
  const targetFolder = getOrCreateFolder(parentFolder, folderName);
  
  // PDFを保存
  const savedFile = targetFolder.createFile(pdfBlob);
  
  return savedFile;
}

/**
 * メール送信
 */
function sendEmailWithPDF(pdfBlob, inputData) {
  const subject = `【${inputData.documentType}】${inputData.companyName}様宛`;
  
  const body = `${inputData.companyName} 御中

平素よりお世話になっております。
以下の通り、${inputData.documentType}をお送りします。

ご確認のほど、よろしくお願いいたします。

────────────────────────

${CONFIG.EMAIL.SENDER_COMPANY}
${CONFIG.EMAIL.SENDER_DEPARTMENT} ${CONFIG.EMAIL.SENDER_NAME}`;

  GmailApp.sendEmail(
    inputData.email,
    subject,
    body,
    {
      attachments: [pdfBlob]
    }
  );
}

/**
 * 送信履歴を記録
 */
function recordSendingHistory(spreadsheet, inputData, savedFile) {
  const historySheet = getOrCreateHistorySheet(spreadsheet);
  
  const lastRow = historySheet.getLastRow() + 1;
  const timestamp = new Date();
  
  historySheet.getRange(lastRow, 1).setValue(timestamp);
  historySheet.getRange(lastRow, 2).setValue(inputData.documentType);
  historySheet.getRange(lastRow, 3).setValue(inputData.companyName);
  historySheet.getRange(lastRow, 4).setValue(inputData.email);
  historySheet.getRange(lastRow, 5).setValue(savedFile.getName());
  historySheet.getRange(lastRow, 6).setValue(savedFile.getUrl());
}

/**
 * バックアップドキュメントを作成
 */
function createBackupDocument(inputData, savedFile) {
  const parentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  const backupFolder = getOrCreateFolder(parentFolder, CONFIG.FOLDERS.BACKUP);
  
  const docName = `${inputData.documentType}_バックアップ_${inputData.companyName}_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}`;
  const doc = DocumentApp.create(docName);
  
  // バックアップ内容を作成
  const body = doc.getBody();
  body.appendParagraph(`${inputData.documentType} 送信記録`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`送信日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm:ss')}`);
  body.appendParagraph(`宛先会社: ${inputData.companyName}`);
  body.appendParagraph(`担当者: ${inputData.contactName}`);
  body.appendParagraph(`メールアドレス: ${inputData.email}`);
  body.appendParagraph(`保存ファイル: ${savedFile.getName()}`);
  body.appendParagraph(`ファイルURL: ${savedFile.getUrl()}`);
  
  doc.saveAndClose();
  
  // バックアップフォルダに移動
  const docFile = DriveApp.getFileById(doc.getId());
  backupFolder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
}