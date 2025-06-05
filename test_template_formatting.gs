/**
 * テンプレートシートの書式継承のテスト
 * 手動実行用のテスト関数
 */
function testTemplateFormatting() {
  try {
    console.log('テンプレート書式継承のテストを開始します...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存のテンプレートシートをチェック
    let templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
    
    if (!templateSheet) {
      console.log('テンプレートシートが存在しないため、新規作成します');
      templateSheet = createTemplateSheet();
      console.log('テンプレートシート作成完了');
      return;
    }
    
    console.log('既存テンプレートシートから書式を保存中...');
    const savedFormatting = inheritTemplateFormatting(templateSheet);
    
    if (!savedFormatting) {
      console.log('書式の保存に失敗しました');
      return;
    }
    
    console.log('保存された書式範囲:', Object.keys(savedFormatting));
    
    // テスト用に既存のシートをバックアップ
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HHmmss');
    const backupName = `${CONFIG.SHEETS.TEMPLATE}_test_backup_${timestamp}`;
    templateSheet.setName(backupName);
    console.log(`既存シートを ${backupName} にバックアップしました`);
    
    // 新しいテンプレートシートを作成
    const newTemplateSheet = spreadsheet.insertSheet(CONFIG.SHEETS.TEMPLATE);
    setupTemplateSheetLayout(newTemplateSheet);
    console.log('新しいテンプレートシートを作成しました');
    
    // 保存した書式を適用
    applyTemplateFormatting(newTemplateSheet, savedFormatting);
    console.log('書式を適用しました');
    
    // G列以降に余計な書式が残っていないかチェック
    const extraRange = newTemplateSheet.getRange('G1:J10');
    const backgrounds = extraRange.getBackgrounds();
    const hasExtraFormatting = backgrounds.some(row => 
      row.some(cell => cell !== '#ffffff' && cell !== '#000000' && cell !== '')
    );
    
    if (hasExtraFormatting) {
      console.log('⚠️ G列以降に余計な書式が残っています');
    } else {
      console.log('✅ G列以降の余計な書式は正常にクリアされています');
    }
    
    console.log('テンプレート書式継承のテストが完了しました');
    console.log('元のシートを復元するには、バックアップシートの名前を変更してください');
    
  } catch (error) {
    console.error('テストエラー:', error);
  }
}

/**
 * バックアップシートから元のテンプレートシートを復元
 */
function restoreTemplateFromBackup() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = spreadsheet.getSheets();
    
    // バックアップシートを検索
    const backupSheets = allSheets.filter(sheet => 
      sheet.getName().includes('テンプレート_test_backup_') || 
      sheet.getName().includes('テンプレート_backup_')
    );
    
    if (backupSheets.length === 0) {
      console.log('バックアップシートが見つかりません');
      return;
    }
    
    // 最新のバックアップを使用
    const latestBackup = backupSheets[backupSheets.length - 1];
    console.log(`バックアップシート "${latestBackup.getName()}" から復元します`);
    
    // 現在のテンプレートシートを削除
    const currentTemplate = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
    if (currentTemplate) {
      spreadsheet.deleteSheet(currentTemplate);
    }
    
    // バックアップシートの名前をテンプレートに変更
    latestBackup.setName(CONFIG.SHEETS.TEMPLATE);
    console.log('テンプレートシートを復元しました');
    
  } catch (error) {
    console.error('復元エラー:', error);
  }
}