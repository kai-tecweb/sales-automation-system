/**
 * デバッグとエラー追跡用ユーティリティ
 */

/**
 * 関数実行の安全なラッパー
 * @param {Function} func 実行する関数
 * @param {string} funcName 関数名（デバッグ用）
 * @param {...any} args 関数の引数
 */
function safeExecute(func, funcName, ...args) {
  try {
    console.log(`[DEBUG] Starting execution: ${funcName}`);
    const result = func.apply(this, args);
    console.log(`[DEBUG] Successfully completed: ${funcName}`);
    return result;
  } catch (error) {
    console.error(`[ERROR] Failed in ${funcName}:`, error);
    console.error(`[ERROR] Stack trace:`, error.stack);
    
    // ユーザーに分かりやすいエラーメッセージを表示
    if (error.toString().includes('sheet is not defined')) {
      SpreadsheetApp.getUi().alert(
        '❌ シートエラー', 
        `関数「${funcName}」でシートが見つかりません。\n\n` +
        `解決方法:\n` +
        `1. メニューから「システム初期化」を実行\n` +
        `2. 必要なシートが作成されていることを確認\n` +
        `3. ページを更新してから再試行\n\n` +
        `エラー詳細: ${error.toString()}`
      );
    }
    
    throw error;
  }
}

/**
 * シート存在確認（詳細デバッグ付き）
 * @param {string} sheetName シート名
 */
function debugSheetAccess(sheetName) {
  try {
    console.log(`[DEBUG] Attempting to access sheet: ${sheetName}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log(`[DEBUG] Spreadsheet ID: ${spreadsheet.getId()}`);
    console.log(`[DEBUG] Spreadsheet Name: ${spreadsheet.getName()}`);
    
    const allSheets = spreadsheet.getSheets().map(s => s.getName());
    console.log(`[DEBUG] All existing sheets:`, allSheets);
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`[ERROR] Sheet "${sheetName}" not found`);
      console.error(`[ERROR] Available sheets:`, allSheets);
      return null;
    }
    
    console.log(`[DEBUG] Successfully accessed sheet: ${sheetName}`);
    return sheet;
    
  } catch (error) {
    console.error(`[ERROR] Failed to access sheet "${sheetName}":`, error);
    throw error;
  }
}

/**
 * 全機能の安全性チェック
 */
function runSystemDiagnostics() {
  const results = {
    timestamp: new Date().toISOString(),
    spreadsheetInfo: {},
    sheetStatus: {},
    functionStatus: {},
    constants: {},
    errors: []
  };
  
  try {
    // スプレッドシート情報
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    results.spreadsheetInfo = {
      id: spreadsheet.getId(),
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl(),
      totalSheets: spreadsheet.getSheets().length
    };
    
    // 定数確認
    results.constants = {
      SHEET_NAMES_defined: typeof SHEET_NAMES !== 'undefined',
      API_KEYS_defined: typeof API_KEYS !== 'undefined',
      SHEET_NAMES_values: typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES : 'undefined'
    };
    
    // シート状況確認
    if (typeof SHEET_NAMES !== 'undefined') {
      Object.entries(SHEET_NAMES).forEach(([key, sheetName]) => {
        try {
          const sheet = debugSheetAccess(sheetName);
          results.sheetStatus[key] = {
            name: sheetName,
            exists: !!sheet,
            lastRow: sheet ? sheet.getLastRow() : 0,
            lastColumn: sheet ? sheet.getLastColumn() : 0
          };
        } catch (error) {
          results.sheetStatus[key] = {
            name: sheetName,
            exists: false,
            error: error.toString()
          };
          results.errors.push(`Sheet ${sheetName}: ${error.toString()}`);
        }
      });
    }
    
    // 重要な関数の存在確認
    const criticalFunctions = [
      'getSafeSheet',
      'checkSystemStatus', 
      'initializeSheets',
      'safeExecute',
      'debugSheetAccess'
    ];
    
    criticalFunctions.forEach(funcName => {
      results.functionStatus[funcName] = typeof global[funcName] === 'function' || typeof this[funcName] === 'function';
    });
    
  } catch (error) {
    results.errors.push(`System diagnostics error: ${error.toString()}`);
  }
  
  // 結果を表示
  console.log('[DIAGNOSTICS] System diagnostic results:', results);
  
  // ユーザー向けサマリー
  let summary = '🔍 システム診断結果\n\n';
  summary += `📊 スプレッドシート: ${results.spreadsheetInfo.name}\n`;
  summary += `🆔 ID: ${results.spreadsheetInfo.id}\n`;
  summary += `📋 総シート数: ${results.spreadsheetInfo.totalSheets}\n\n`;
  
  if (Object.keys(results.sheetStatus).length > 0) {
    summary += '📝 シート状況:\n';
    Object.entries(results.sheetStatus).forEach(([key, status]) => {
      summary += `  ${status.exists ? '✅' : '❌'} ${status.name}\n`;
    });
    summary += '\n';
  }
  
  if (results.errors.length > 0) {
    summary += '⚠️ エラー:\n';
    results.errors.forEach(error => {
      summary += `  • ${error}\n`;
    });
  }
  
  SpreadsheetApp.getUi().alert('システム診断結果', summary);
  
  return results;
}
