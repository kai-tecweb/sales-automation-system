/**
 * キーワード生成機能 - 修正版
 */

/**
 * メニュー用キーワード生成関数
 */
function generateKeywords() {
  try {
    console.log('🔤 キーワード生成を開始します...');
    
    // 制御パネルから設定を取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('制御パネル');
    if (!sheet) {
      throw new Error('制御パネルシートが見つかりません。システムを初期化してください。');
    }
    
    // 商材情報を取得
    const data = sheet.getRange('A2:B6').getValues();
    const productName = data[0][1] || '';
    const productDescription = data[1][1] || '';
    
    if (!productName || !productDescription) {
      SpreadsheetApp.getUi().alert('❌ エラー', '制御パネルで商材名と商材概要を入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 成功メッセージを表示（実際のAPI呼び出しは後で実装）
    const message = `✅ キーワード生成（デモ版）\n\n商材名: ${productName}\n商材概要: ${productDescription}\n\n実際のキーワード生成機能は次のフェーズで実装予定です。`;
    
    SpreadsheetApp.getUi().alert('キーワード生成', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log('✅ キーワード生成処理完了');
    
  } catch (error) {
    console.error('❌ キーワード生成エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `キーワード生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 企業検索（デモ版）
 */
function executeCompanySearch() {
  try {
    SpreadsheetApp.getUi().alert('企業検索', '企業検索機能（デモ版）\n\n次のフェーズで実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ 企業検索エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `企業検索でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 提案生成（デモ版）
 */
function generatePersonalizedProposals() {
  try {
    SpreadsheetApp.getUi().alert('提案生成', '提案生成機能（デモ版）\n\n次のフェーズで実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ 提案生成エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `提案生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 完全自動化実行（デモ版）
 */
function executeFullWorkflow() {
  try {
    SpreadsheetApp.getUi().alert('完全自動化', '完全自動化機能（デモ版）\n\n次のフェーズで実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ 完全自動化エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `完全自動化でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}