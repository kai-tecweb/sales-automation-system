/**
 * プロパティ管理システム
 * システムで使用するAPIキーとその他設定値の管理
 */

/**
 * APIキーの設定・取得・削除機能
 */
function setApiKeys() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // OpenAI APIキーの設定
    const openaiResult = ui.prompt(
      'OpenAI APIキー設定',
      'ChatGPT APIキーを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (openaiResult.getSelectedButton() === ui.Button.OK) {
      const openaiKey = openaiResult.getResponseText().trim();
      if (openaiKey) {
        PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', openaiKey);
        ui.alert('OpenAI APIキーを設定しました');
      }
    } else {
      return;
    }
    
    // Google Search APIキーの設定
    const googleResult = ui.prompt(
      'Google Search APIキー設定',
      'Google Custom Search APIキーを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (googleResult.getSelectedButton() === ui.Button.OK) {
      const googleKey = googleResult.getResponseText().trim();
      if (googleKey) {
        PropertiesService.getScriptProperties().setProperty('GOOGLE_SEARCH_API_KEY', googleKey);
        ui.alert('Google Search APIキーを設定しました');
      }
    } else {
      return;
    }
    
    // 検索エンジンIDの設定
    const engineResult = ui.prompt(
      'Google Search エンジンID設定',
      'Google Custom Search エンジンIDを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (engineResult.getSelectedButton() === ui.Button.OK) {
      const engineId = engineResult.getResponseText().trim();
      if (engineId) {
        PropertiesService.getScriptProperties().setProperty('GOOGLE_SEARCH_ENGINE_ID', engineId);
        
        // APIキー管理シートと制御パネルを更新
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          
          // デバッグ情報
          console.log('SHEET_NAMES利用可能か:', typeof SHEET_NAMES !== 'undefined');
          console.log('API_KEYS定数:', typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.API_KEYS : '未定義');
          
          // APIキー管理シートを更新（直接シート名を指定）
          let apiKeySheet = ss.getSheetByName('APIキー管理');
          console.log('APIキー管理シート取得結果:', apiKeySheet ? 'シートが見つかりました' : 'シートが見つかりません');
          
          if (apiKeySheet) {
            updateApiKeyManagementSheet(apiKeySheet);
            console.log('APIキー管理シート更新完了');
          } else {
            console.log('APIキー管理シートが見つかりません');
          }
          
          // 制御パネルのAPI設定状況を更新
          let controlSheet = ss.getSheetByName('制御パネル');
          if (controlSheet) {
            updateControlPanelApiStatus(controlSheet);
            console.log('制御パネル更新完了');
          }
        } catch (error) {
          console.log('シート更新エラー:', error);
          console.log('エラースタック:', error.stack);
        }
        
        ui.alert('すべてのAPIキーと設定を完了しました！\n制御パネルのAPIキー設定も更新されました。');
      }
    }
    
  } catch (error) {
    Logger.log('APIキー設定エラー: ' + error.toString());
    ui.alert('APIキー設定中にエラーが発生しました: ' + error.toString());
  }
}

/**
 * 設定されているAPIキーの確認
 */
function checkApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  
  const openaiKey = properties.getProperty('OPENAI_API_KEY');
  const googleKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
  const engineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
  
  let status = '【APIキー設定状況】\n\n';
  status += 'OpenAI APIキー: ' + (openaiKey ? '✅ 設定済み' : '❌ 未設定') + '\n';
  status += 'Google Search APIキー: ' + (googleKey ? '✅ 設定済み' : '❌ 未設定') + '\n';
  status += 'Google Search エンジンID: ' + (engineId ? '✅ 設定済み' : '❌ 未設定') + '\n\n';
  
  if (openaiKey && googleKey && engineId) {
    status += '🎉 すべてのAPIキーが設定されています！';
  } else {
    status += '⚠️ 一部のAPIキーが未設定です。「🔑 APIキー設定」から設定してください。';
  }
  
  ui.alert('APIキー設定状況', status, ui.ButtonSet.OK);
}

/**
 * APIキーの削除（デバッグ用）
 */
function clearApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'APIキー削除確認',
    '本当にすべてのAPIキーを削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('OPENAI_API_KEY');
    properties.deleteProperty('GOOGLE_SEARCH_API_KEY');
    properties.deleteProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    ui.alert('APIキー削除完了', '全てのAPIキーを削除しました。', ui.ButtonSet.OK);
  }
}

/**
 * APIキー管理シートの更新
 */
function updateApiKeyManagementSheet(sheet) {
  console.log('updateApiKeyManagementSheet関数開始');
  try {
    const properties = PropertiesService.getScriptProperties();
    
    const openaiKey = properties.getProperty('OPENAI_API_KEY');
    const googleKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
    const engineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    // 現在の日時
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    
    // APIキー管理シートのデータ構造
    const apiData = [
      ['API種別', '設定状況', 'キー（マスク表示）', '最終更新', '備考'],
      [
        'OpenAI ChatGPT API', 
        openaiKey ? '✅ 設定済み' : '❌ 未設定',
        openaiKey ? `${openaiKey.substring(0, 8)}...${openaiKey.substring(openaiKey.length - 4)}` : '未設定',
        openaiKey ? dateStr : '',
        openaiKey ? 'キーワード生成・提案生成で使用' : 'キーワード生成機能が利用不可'
      ],
      [
        'Google Custom Search API',
        googleKey ? '✅ 設定済み' : '❌ 未設定', 
        googleKey ? `${googleKey.substring(0, 8)}...${googleKey.substring(googleKey.length - 4)}` : '未設定',
        googleKey ? dateStr : '',
        googleKey ? '企業検索で使用' : '企業検索機能が利用不可'
      ],
      [
        'Google Search Engine ID',
        engineId ? '✅ 設定済み' : '❌ 未設定',
        engineId ? engineId : '未設定', 
        engineId ? dateStr : '',
        engineId ? 'カスタム検索エンジンID' : '企業検索機能が利用不可'
      ]
    ];
    
    // シートをクリアして新しいデータを設定
    sheet.clear();
    sheet.getRange(1, 1, apiData.length, 5).setValues(apiData);
    
    // ヘッダーのスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 200); // API種別
    sheet.setColumnWidth(2, 120); // 設定状況
    sheet.setColumnWidth(3, 200); // キー（マスク表示）
    sheet.setColumnWidth(4, 150); // 最終更新
    sheet.setColumnWidth(5, 250); // 備考
    
    // 設定状況に基づく行の色分け
    for (let i = 2; i <= apiData.length; i++) {
      const statusCell = sheet.getRange(i, 2);
      if (statusCell.getValue().includes('✅')) {
        sheet.getRange(i, 1, 1, 5).setBackground('#d4edda'); // 緑系
      } else {
        sheet.getRange(i, 1, 1, 5).setBackground('#f8d7da'); // 赤系
      }
    }
    
    // 更新情報を追加
    const summaryRow = apiData.length + 2;
    sheet.getRange(summaryRow, 1).setValue('【設定状況サマリー】').setFontWeight('bold');
    
    const allSet = openaiKey && googleKey && engineId;
    const partialSet = (openaiKey || googleKey || engineId) && !allSet;
    const noneSet = !openaiKey && !googleKey && !engineId;
    
    let statusMessage = '';
    if (allSet) {
      statusMessage = '🎉 すべてのAPIキーが設定されています！システム全機能が利用可能です。';
    } else if (partialSet) {
      statusMessage = '⚠️ 一部のAPIキーが未設定です。一部機能が制限されます。';
    } else {
      statusMessage = '❌ APIキーが設定されていません。「🔑 APIキー設定」から設定してください。';
    }
    
    sheet.getRange(summaryRow + 1, 1, 1, 5).merge();
    sheet.getRange(summaryRow + 1, 1).setValue(statusMessage);
    
    console.log('✅ APIキー管理シート更新完了');
    
    // 更新完了をUIに通知
    SpreadsheetApp.getActiveSpreadsheet().toast('APIキー管理シートを更新しました', '更新完了', 3);
    
  } catch (error) {
    console.error('❌ APIキー管理シート更新エラー:', error);
    console.error('エラー詳細:', error.toString());
    SpreadsheetApp.getUi().alert('シート更新エラー', `APIキー管理シートの更新中にエラーが発生しました:\n${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * APIキーの検証（内部使用）
 */
function validateApiKeys() {
  const properties = PropertiesService.getScriptProperties();
  
  const openaiKey = properties.getProperty('OPENAI_API_KEY');
  const googleKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
  const engineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
  
  return {
    openaiKey: !!openaiKey,
    googleKey: !!googleKey,
    engineId: !!engineId,
    allSet: !!(openaiKey && googleKey && engineId)
  };
}

/**
 * OpenAI API呼び出しのシンプルテスト
 */
function testOpenAIAPISimple() {
  try {
    console.log('=== シンプルOpenAI APIテスト開始 ===');
    
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('エラー', 'OpenAI APIキーが設定されていません', SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    
    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: 'Hello, just respond with "Test successful"'
        }
      ],
      max_tokens: 10
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    console.log('UrlFetchApp.fetch実行開始...');
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    
    console.log('UrlFetchApp.fetch実行完了');
    console.log('レスポンスコード:', response.getResponseCode());
    console.log('レスポンスの型:', typeof response);
    
    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      throw new Error(`API Error: ${errorText}`);
    }
    
    const responseText = response.getContentText();
    console.log('レスポンステキストの型:', typeof responseText);
    
    const data = JSON.parse(responseText);
    const result = data.choices[0].message.content;
    
    console.log('結果:', result);
    console.log('結果の型:', typeof result);
    
    SpreadsheetApp.getUi().alert('テスト成功', `OpenAI APIテスト成功！\n結果: ${result}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return result;
    
  } catch (error) {
    console.error('APIテストエラー:', error);
    SpreadsheetApp.getUi().alert('テストエラー', `OpenAI APIテストでエラーが発生しました:\n${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * システム初期化時のAPIキーチェック
 */
function checkApiKeysOnInit() {
  const validation = validateApiKeys();
  
  if (!validation.allSet) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'APIキー設定が必要です',
      'システムを使用するにはAPIキーの設定が必要です。\n\nメニューから「🔑 APIキー設定」を選択してください。',
      ui.ButtonSet.OK
    );
    return false;
  }
  
  return true;
}

/**
 * APIキー管理シート更新のテスト関数
 */
function testUpdateApiKeySheet() {
  try {
    console.log('=== APIキー管理シート更新テスト開始 ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apiKeySheet = ss.getSheetByName('APIキー管理');
    
    if (!apiKeySheet) {
      SpreadsheetApp.getUi().alert('エラー', 'APIキー管理シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    
    console.log('APIキー管理シートが見つかりました');
    updateApiKeyManagementSheet(apiKeySheet);
    
    SpreadsheetApp.getUi().alert('テスト完了', 'APIキー管理シートの更新テストが完了しました', SpreadsheetApp.getUi().ButtonSet.OK);
    return true;
    
  } catch (error) {
    console.error('テストエラー:', error);
    SpreadsheetApp.getUi().alert('テストエラー', `エラーが発生しました: ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}