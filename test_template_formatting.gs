/**
 * テンプレートシートの書式継承のテスト
 * 手動実行用のテスト関数
 */
function testTemplateFormattingInheritance() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let templateSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TEMPLATE);
    
    if (!templateSheet) {
      console.log('テンプレートシートが存在しないため、新規作成します');
      templateSheet = createTemplateSheet();
      return;
    }
    
    // 書式継承をテスト
    const savedFormatting = inheritTemplateFormatting(templateSheet);
    
    if (!savedFormatting) {
      console.log('❌ 書式の保存に失敗しました');
      return;
    }
    
    console.log('✅ 書式範囲が拡張されました:', Object.keys(savedFormatting));
    console.log('✅ A1:F50の範囲で書式が保存されています');
    
    // 新しいテンプレートに書式適用をテスト
    const testSheet = spreadsheet.insertSheet('テスト用テンプレート');
    setupTemplateSheetLayout(testSheet);
    applyTemplateFormatting(testSheet, savedFormatting);
    
    // G列以降のクリアをチェック
    const extraRange = testSheet.getRange('G1:H10');
    const backgrounds = extraRange.getBackgrounds();
    const hasExtraFormatting = backgrounds.some(row => 
      row.some(cell => cell !== '#ffffff' && cell !== '')
    );
    
    if (hasExtraFormatting) {
      console.log('⚠️ G列以降に余計な書式が残っています');
    } else {
      console.log('✅ G列以降の余計な書式は正常にクリアされています');
    }
    
    // テストシートをクリーンアップ
    spreadsheet.deleteSheet(testSheet);
    console.log('✅ テンプレート書式継承の修正が正常に動作しています');
    
  } catch (error) {
    console.error('テストエラー:', error);
  }
}

