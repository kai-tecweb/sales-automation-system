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
        ui.alert('すべてのAPIキーと設定を完了しました！');
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
    
    ui.alert('すべてのAPIキーを削除しました');
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