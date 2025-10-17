/**
 * ユーティリティ機能
 * システム共通で使用される補助機能
 */

/**
 * データ品質チェック機能
 */
function validateCompanyData(companyData) {
  const errors = [];
  
  // 必須項目チェック
  const required = ['companyName', 'industry'];
  required.forEach(field => {
    if (!companyData[field] || companyData[field].toString().trim() === '') {
      errors.push(`必須項目「${field}」が不足しています`);
    }
  });
  
  // 会社名の妥当性チェック
  if (companyData.companyName) {
    const invalidPatterns = [
      /404/i, /not found/i, /error/i, /エラー/i, 
      /ページが見つかりません/i, /アクセス拒否/i
    ];
    
    if (invalidPatterns.some(pattern => pattern.test(companyData.companyName))) {
      errors.push('無効な会社名が検出されました');
    }
  }
  
  // 従業員数の妥当性チェック
  if (companyData.employees !== null && companyData.employees !== undefined) {
    if (isNaN(companyData.employees) || companyData.employees < 0 || companyData.employees > 1000000) {
      errors.push('従業員数が無効な値です');
    }
  }
  
  if (errors.length > 0) {
    throw new Error(`データ品質エラー: ${errors.join(', ')}`);
  }
  
  return true;
}

/**
 * テキストクリーニング機能
 */
function cleanText(text) {
  if (!text) return '';
  
  return text
    .replace(/\s+/g, ' ')           // 複数の空白を単一に
    .replace(/[\r\n]+/g, ' ')       // 改行を空白に
    .replace(/[^\w\s\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/g, ' ') // 特殊文字除去（日本語対応）
    .trim()                         // 前後の空白除去
    .substring(0, 500);            // 長さ制限
}

/**
 * 日本語テキスト正規化
 */
function normalizeJapaneseText(text) {
  if (!text) return '';
  
  return text
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })  // 全角英数字を半角に
    .replace(/[！-～]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })  // 全角記号を半角に
    .replace(/\u3000/g, ' ')        // 全角スペースを半角に
    .trim();
}

/**
 * URLの妥当性チェック
 */
function isValidUrl(url) {
  if (!url) return false;
  
  try {
    const urlObj = new URL(url);
    return urlObj.protocol === 'http:' || urlObj.protocol === 'https:';
  } catch (e) {
    return false;
  }
}

/**
 * メールアドレスの妥当性チェック
 */
function isValidEmail(email) {
  if (!email) return false;
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * 日付フォーマット
 */
function formatDate(date, format = 'YYYY-MM-DD HH:mm:ss') {
  if (!date) return '';
  
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const seconds = String(d.getSeconds()).padStart(2, '0');
  
  return format
    .replace('YYYY', year)
    .replace('MM', month)
    .replace('DD', day)
    .replace('HH', hours)
    .replace('mm', minutes)
    .replace('ss', seconds);
}

/**
 * 安全なJSON解析
 */
function safeJsonParse(jsonString, defaultValue = null) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    Logger.log(`JSON解析エラー: ${e.toString()}`);
    Logger.log(`問題のあるJSON: ${jsonString}`);
    return defaultValue;
  }
}

/**
 * 配列のシャッフル
 */
function shuffleArray(array) {
  const shuffled = [...array];
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
}

/**
 * 配列の重複除去
 */
function removeDuplicates(array, keyFunc = null) {
  if (!keyFunc) {
    return [...new Set(array)];
  }
  
  const seen = new Set();
  return array.filter(item => {
    const key = keyFunc(item);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

/**
 * オブジェクトの深いコピー
 */
function deepCopy(obj) {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }
  
  if (obj instanceof Date) {
    return new Date(obj.getTime());
  }
  
  if (obj instanceof Array) {
    return obj.map(item => deepCopy(item));
  }
  
  const copiedObj = {};
  Object.keys(obj).forEach(key => {
    copiedObj[key] = deepCopy(obj[key]);
  });
  
  return copiedObj;
}

/**
 * スプレッドシートの列インデックスを文字に変換
 */
function columnIndexToLetter(index) {
  let result = '';
  let num = index;
  
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  
  return result;
}

/**
 * スプレッドシートの列文字をインデックスに変換
 */
function columnLetterToIndex(letter) {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * レート制限対応のスリープ
 */
function smartSleep(baseMs = 1000, jitter = 0.2) {
  const jitterMs = baseMs * jitter * Math.random();
  const totalMs = baseMs + jitterMs;
  Utilities.sleep(totalMs);
}

/**
 * バッチ処理用のチャンク分割
 */
function chunkArray(array, chunkSize) {
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  return chunks;
}

/**
 * エラー情報の詳細ログ
 */
function logDetailedError(error, context = '') {
  const errorInfo = {
    message: error.toString(),
    name: error.name || 'Unknown',
    stack: error.stack || 'No stack trace',
    context: context,
    timestamp: new Date().toISOString()
  };
  
  Logger.log(`詳細エラー情報: ${JSON.stringify(errorInfo, null, 2)}`);
  
  // エラーログシートに記録（オプション）
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('エラーログ');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('エラーログ');
      errorSheet.getRange(1, 1, 1, 5).setValues([
        ['タイムスタンプ', 'エラー名', 'メッセージ', 'コンテキスト', 'スタックトレース']
      ]);
    }
    
    const lastRow = errorSheet.getLastRow();
    errorSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      errorInfo.timestamp,
      errorInfo.name,
      errorInfo.message,
      errorInfo.context,
      errorInfo.stack
    ]]);
    
  } catch (logError) {
    Logger.log(`エラーログ記録失敗: ${logError.toString()}`);
  }
}

/**
 * パフォーマンス測定
 */
function measurePerformance(func, name = 'Operation') {
  const startTime = new Date();
  let result;
  let error = null;
  
  try {
    result = func();
  } catch (e) {
    error = e;
  }
  
  const endTime = new Date();
  const duration = endTime - startTime;
  
  Logger.log(`パフォーマンス測定 [${name}]: ${duration}ms`);
  
  if (error) {
    throw error;
  }
  
  return result;
}

/**
 * システム状態の健全性チェック
 */
function performHealthCheck() {
  const healthStatus = {
    timestamp: new Date(),
    sheets: {},
    apiKeys: {},
    dataIntegrity: {},
    performance: {}
  };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シートの存在確認
    Object.values(SHEET_NAMES).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      healthStatus.sheets[sheetName] = {
        exists: !!sheet,
        rows: sheet ? sheet.getLastRow() : 0
      };
    });
    
    // APIキーの存在確認
    healthStatus.apiKeys = {
      openai: !!API_KEYS.OPENAI,
      googleSearch: !!API_KEYS.GOOGLE_SEARCH,
      googleSearchEngineId: !!API_KEYS.GOOGLE_SEARCH_ENGINE_ID
    };
    
    // データ整合性チェック
    healthStatus.dataIntegrity = checkDataIntegrity();
    
  } catch (error) {
    healthStatus.error = error.toString();
  }
  
  Logger.log(`システム健全性チェック結果: ${JSON.stringify(healthStatus, null, 2)}`);
  return healthStatus;
}

/**
 * データ整合性チェック
 */
function checkDataIntegrity() {
  const integrity = {
    companyProposalMatch: true,
    duplicateCompanies: 0,
    orphanedProposals: 0
  };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const companiesSheet = ss.getSheetByName(SHEET_NAMES.COMPANIES);
    const proposalsSheet = ss.getSheetByName(SHEET_NAMES.PROPOSALS);
    
    if (companiesSheet && proposalsSheet) {
      const companyIds = companiesSheet.getRange(2, 1, companiesSheet.getLastRow() - 1, 1)
        .getValues().flat().filter(id => id);
      const proposalCompanyIds = proposalsSheet.getRange(2, 1, proposalsSheet.getLastRow() - 1, 1)
        .getValues().flat().filter(id => id);
      
      // 孤立した提案のチェック
      integrity.orphanedProposals = proposalCompanyIds.filter(id => !companyIds.includes(id)).length;
      
      // 重複企業のチェック
      const companyNames = companiesSheet.getRange(2, 2, companiesSheet.getLastRow() - 1, 1)
        .getValues().flat().filter(name => name);
      integrity.duplicateCompanies = companyNames.length - new Set(companyNames).size;
    }
    
  } catch (error) {
    integrity.error = error.toString();
  }
  
  return integrity;
}
