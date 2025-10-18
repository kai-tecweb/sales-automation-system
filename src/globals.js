/**
 * 共通定数とユーティリティ関数
 * 全ファイルで共有される定数や安全なシート操作関数
 */

// グローバル定数
const SHEET_NAMES = {
  CONTROL: '制御パネル',
  KEYWORDS: '生成キーワード',
  COMPANIES: '企業マスター',
  PROPOSALS: '提案メッセージ',
  LOGS: '実行ログ',
  API_KEYS: 'APIキー管理'
};

const API_KEYS = {
  GOOGLE_SEARCH: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY'),
  OPENAI: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  GOOGLE_SEARCH_ENGINE_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID')
};

/**
 * シート名定数の安全な取得
 * @returns {object} SHEET_NAMES オブジェクト
 */
function getSheetNames() {
  return {
    CONTROL: '制御パネル',
    KEYWORDS: '生成キーワード',
    COMPANIES: '企業マスター',
    PROPOSALS: '提案メッセージ',
    LOGS: '実行ログ'
  };
}

/**
 * API設定の安全な取得
 * @returns {object} API_KEYS オブジェクト
 */
function getApiKeys() {
  return {
    GOOGLE_SEARCH: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY'),
    OPENAI: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
    GOOGLE_SEARCH_ENGINE_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID')
  };
}

/**
 * 安全なシート取得
 * @param {string} sheetName シート名
 * @param {boolean} createIfNotExists 存在しない場合に作成するか
 * @returns {Sheet|null} シートオブジェクトまたはnull
 */
function getSafeSheet(sheetName, createIfNotExists = false) {
  try {
    console.log(`[getSafeSheet] Attempting to access: ${sheetName}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.error('[getSafeSheet] No active spreadsheet found');
      return null;
    }
    
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.warn(`[getSafeSheet] Sheet "${sheetName}" not found`);
      
      if (createIfNotExists) {
        try {
          console.log(`[getSafeSheet] Creating missing sheet: ${sheetName}`);
          sheet = spreadsheet.insertSheet(sheetName);
          console.log(`[getSafeSheet] Successfully created sheet: ${sheetName}`);
        } catch (createError) {
          console.error(`[getSafeSheet] Failed to create sheet "${sheetName}":`, createError);
          return null;
        }
      } else {
        // 利用可能なシートを表示
        const availableSheets = spreadsheet.getSheets().map(s => s.getName());
        console.error(`[getSafeSheet] Available sheets:`, availableSheets);
        return null;
      }
    }
    
    if (sheet) {
      console.log(`[getSafeSheet] Successfully accessed sheet: ${sheetName}`);
    }
    
    return sheet;
    
  } catch (error) {
    console.error(`[getSafeSheet] Error accessing sheet "${sheetName}":`, error);
    console.error(`[getSafeSheet] Error stack:`, error.stack);
    return null;
  }
}

/**
 * 必要なシートが全て存在するかチェック
 * @returns {object} チェック結果
 */
function checkRequiredSheets() {
  const results = {
    allExists: true,
    missingSheets: [],
    existingSheets: []
  };
  
  Object.values(SHEET_NAMES).forEach(sheetName => {
    const sheet = getSafeSheet(sheetName);
    if (sheet) {
      results.existingSheets.push(sheetName);
    } else {
      results.missingSheets.push(sheetName);
      results.allExists = false;
    }
  });
  
  return results;
}

/**
 * システムの初期化状況をチェック
 * @returns {object} 初期化結果
 */
function checkSystemStatus() {
  try {
    const sheetStatus = checkRequiredSheets();
    
    return {
      success: true,
      needsInitialization: !sheetStatus.allExists,
      sheetStatus: sheetStatus,
      message: sheetStatus.allExists ? 
        'システムは正常に初期化されています' : 
        `不足しているシート: ${sheetStatus.missingSheets.join(', ')}`
    };
  } catch (error) {
    return {
      success: false,
      needsInitialization: true,
      error: error.toString(),
      message: 'システム状況の確認に失敗しました'
    };
  }
}
