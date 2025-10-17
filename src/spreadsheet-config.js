/**
 * スプレッドシート設定
 * 指定されたGoogle SheetsにバインドするためのID設定
 */

// バインド対象のスプレッドシートID
const SPREADSHEET_ID = '1Q4i_GjDX1x_hWbqvRJDrmqizZMOZ8q-FjX_pf7eGPlM';

/**
 * バインドされたスプレッドシートを取得
 */
function getBoundSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    console.error('スプレッドシートアクセスエラー:', error.toString());
    throw new Error(`指定されたスプレッドシート（ID: ${SPREADSHEET_ID}）にアクセスできません。アクセス権限を確認してください。`);
  }
}

/**
 * アクティブなスプレッドシートまたはバインドされたスプレッドシートを取得
 */
function getSafeSpreadsheet() {
  try {
    // まずバインドされたスプレッドシートを試行
    return getBoundSpreadsheet();
  } catch (error) {
    console.warn('バインドされたスプレッドシートへのアクセスに失敗:', error.toString());
    
    // フォールバックとしてアクティブなスプレッドシートを使用
    try {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSpreadsheet) {
        console.log('アクティブなスプレッドシートを使用します');
        return activeSpreadsheet;
      }
    } catch (activeError) {
      console.error('アクティブなスプレッドシートの取得にも失敗:', activeError.toString());
    }
    
    throw new Error('利用可能なスプレッドシートが見つかりません。スプレッドシートのアクセス権限を確認してください。');
  }
}

/**
 * スプレッドシート接続テスト
 */
function testSpreadsheetConnection() {
  try {
    const spreadsheet = getBoundSpreadsheet();
    const name = spreadsheet.getName();
    const url = spreadsheet.getUrl();
    
    console.log('スプレッドシート接続成功:');
    console.log('名前:', name);
    console.log('URL:', url);
    console.log('ID:', SPREADSHEET_ID);
    
    return {
      success: true,
      name: name,
      url: url,
      id: SPREADSHEET_ID
    };
  } catch (error) {
    console.error('スプレッドシート接続テスト失敗:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * バインド設定を指定されたスプレッドシートに適用
 */
function bindToSpecifiedSpreadsheet() {
  try {
    console.log(`スプレッドシート（ID: ${SPREADSHEET_ID}）への接続を確認しています...`);
    
    const result = testSpreadsheetConnection();
    if (!result.success) {
      throw new Error(result.error);
    }
    
    console.log('接続成功！システム初期化を開始します...');
    
    // システム初期化
    initializeSheets();
    
    console.log('バインディング完了！');
    console.log('スプレッドシート名:', result.name);
    console.log('スプレッドシートURL:', result.url);
    
    return result;
  } catch (error) {
    console.error('バインディングエラー:', error.toString());
    throw error;
  }
}
