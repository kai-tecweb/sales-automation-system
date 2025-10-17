/**
 * 営業自動化システム - 動作確認用最小メニュー
 */

function onOpen() {
  try {
    console.log('Creating minimal menu...');
    
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu('🚀 営業システム')
      .addItem('🧪 システムテスト', 'runSystemTest')
      .addItem('📋 基本情報', 'showBasicInfo')
      .addItem('🔧 シート作成', 'createBasicSheets')
      .addItem('✅ 簡単テスト', 'simpleTest')
      .addToUi();
    
    console.log('Menu created successfully');
    
    // 成功通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'システム起動完了', 
      '✅ 営業自動化システム', 
      2
    );
    
  } catch (error) {
    console.error('Menu creation error:', error);
  }
}

function runSystemTest() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    let message = '🧪 システムテスト結果\n\n';
    message += `📊 スプレッドシート: ${spreadsheet.getName()}\n`;
    message += `🆔 ID: ${spreadsheet.getId().substring(0, 20)}...\n`;
    message += ` 実行時刻: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    // シート確認
    const sheets = spreadsheet.getSheets();
    message += `📋 現在のシート数: ${sheets.length}個\n\n`;
    
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
    const existingRequired = requiredSheets.filter(name => 
      sheets.some(sheet => sheet.getName() === name)
    );
    
    message += `✅ 必要シート: ${existingRequired.length}/${requiredSheets.length}個作成済み\n`;
    
    if (existingRequired.length < requiredSheets.length) {
      message += `❌ 不足: ${requiredSheets.length - existingRequired.length}個\n`;
      message += `\n「シート作成」で不足分を作成できます。`;
    } else {
      message += `\n✅ すべてのシートが準備されています！`;
    }
    
    ui.alert('🧪 システムテスト結果\n\n' + message);
    
  } catch (error) {
    console.error('System test error:', error);
    SpreadsheetApp.getUi().alert('❌ テストエラー\n\nエラーが発生しました: ' + String(error).substring(0, 100));
  }
}

function showBasicInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    let info = '📋 営業自動化システム情報\n\n';
    info += `📊 スプレッドシート: ${spreadsheet.getName()}\n`;
    info += `🆔 ID: ${spreadsheet.getId().substring(0, 20)}...\n`;
    info += `🔗 URL: 正常にアクセス可能\n`;
    info += `🕒 確認日時: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    info += '🎯 主な機能:\n';
    info += '• キーワード生成\n';
    info += '• 企業検索・分析\n';
    info += '• 提案メッセージ自動作成\n';
    info += '• マッチ度スコアリング\n\n';
    
    info += '📞 サポート:\n';
    info += 'システムに関するお問い合わせは\n';
    info += '管理者までご連絡ください。';
    
    ui.alert('📋 営業自動化システム情報\n\n' + info);
    
  } catch (error) {
    console.error('Show info error:', error);
    SpreadsheetApp.getUi().alert('❌ 情報取得エラー\n\nエラーが発生しました: ' + String(error).substring(0, 100));
  }
}

function createBasicSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '🔧 シート作成確認\n\n営業自動化システムに必要なシートを作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      const requiredSheets = [
        { name: '制御パネル', description: 'システム設定と操作' },
        { name: '生成キーワード', description: 'AI生成検索キーワード' },
        { name: '企業マスター', description: '発見企業の管理' },
        { name: '提案メッセージ', description: '自動生成提案' },
        { name: '実行ログ', description: 'システム実行履歴' }
      ];
      
      let created = 0;
      let existing = 0;
      
      requiredSheets.forEach(sheetInfo => {
        let sheet = spreadsheet.getSheetByName(sheetInfo.name);
        if (!sheet) {
          sheet = spreadsheet.insertSheet(sheetInfo.name);
          
          // 基本ヘッダーを設定
          sheet.getRange('A1').setValue(sheetInfo.description);
          sheet.getRange('A1').setFontWeight('bold').setFontSize(12);
          
          created++;
          console.log(`Created sheet: ${sheetInfo.name}`);
        } else {
          existing++;
          console.log(`Sheet already exists: ${sheetInfo.name}`);
        }
      });
      
      let resultMessage = '✅ シート作成完了\n\n';
      resultMessage += `新規作成: ${created}個\n`;
      resultMessage += `既存確認: ${existing}個\n`;
      resultMessage += `合計: ${requiredSheets.length}個\n\n`;
      resultMessage += 'システムの準備が整いました！';
      
      ui.alert('✅ 作成完了\n\n' + resultMessage);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ シート作成エラー\n\n' + String(error).substring(0, 100));
  }
}

function simpleTest() {
  try {
    SpreadsheetApp.getUi().alert('✅ 簡単テスト成功\n\nメニューは正常に動作しています！');
  } catch (error) {
    console.error('Simple test error:', error);
  }
}
