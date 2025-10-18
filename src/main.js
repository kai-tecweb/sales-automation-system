/**
 * 営業自動化システム - メインファイル
 * 商材起点企業発掘・提案自動生成システム
 */

// ===========================================
// API設定と定数
// ===========================================

/**
 * OpenAI API設定
 */
const OPENAI_CONFIG = {
  baseURL: 'https://api.openai.com/v1',
  model: 'gpt-3.5-turbo',
  maxTokens: 2000,
  temperature: 0.7
};

/**
 * Google Custom Search API設定
 */
const GOOGLE_SEARCH_CONFIG = {
  baseURL: 'https://www.googleapis.com/customsearch/v1',
  defaultParams: {
    lr: 'lang_ja',
    cr: 'countryJP',
    safe: 'off',
    num: 10
  }
};

// ===========================================
// APIレート制限システム
// ===========================================

/**
 * OpenAI APIレート制限管理クラス
 */
class OpenAIRateLimiter {
  constructor() {
    this.lastRequestTime = 0;
    this.requestCount = 0;
    this.tokensUsed = 0;
    this.minutelyLimit = 3; // GPT-3.5-turbo: 3 requests/minute (無料プラン)
    this.dailyTokenLimit = 40000; // 1日のトークン制限
  }
  
  async waitIfNeeded() {
    const now = Date.now();
    const timeSinceLastRequest = now - this.lastRequestTime;
    
    // 20秒間隔を維持（3リクエスト/分制限対応）
    if (timeSinceLastRequest < 20000) {
      Logger.log(`OpenAI API制限: ${20000 - timeSinceLastRequest}ms待機中...`);
      await this.sleep(20000 - timeSinceLastRequest);
    }
    
    // 分間制限チェック
    if (this.requestCount >= this.minutelyLimit) {
      Logger.log('OpenAI API分間制限に達しました。60秒待機します...');
      await this.sleep(60000);
      this.requestCount = 0;
    }
    
    this.lastRequestTime = Date.now();
    this.requestCount++;
  }
  
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  updateTokenUsage(tokens) {
    this.tokensUsed += tokens;
    if (this.tokensUsed > this.dailyTokenLimit) {
      throw new Error('OpenAI API日次トークン制限に達しました');
    }
  }
}

/**
 * Google Search APIレート制限管理クラス
 */
class GoogleSearchRateLimiter {
  constructor() {
    this.lastRequestTime = 0;
    this.requestCount = 0;
    this.dailyLimit = 100; // Custom Search API: 100 queries/day (無料プラン)
  }
  
  async waitIfNeeded() {
    const now = Date.now();
    const timeSinceLastRequest = now - this.lastRequestTime;
    
    // 1.5秒間隔を維持（安全マージン）
    if (timeSinceLastRequest < 1500) {
      Logger.log(`Google Search API制限: ${1500 - timeSinceLastRequest}ms待機中...`);
      await this.sleep(1500 - timeSinceLastRequest);
    }
    
    // 日次制限チェック
    if (this.requestCount >= this.dailyLimit) {
      throw new Error('Google Search API日次制限に達しました');
    }
    
    this.lastRequestTime = Date.now();
    this.requestCount++;
  }
  
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * 統合レート制限マネージャー
 */
class APIRateLimitManager {
  constructor() {
    this.openAILimiter = new OpenAIRateLimiter();
    this.googleSearchLimiter = new GoogleSearchRateLimiter();
  }
  
  async waitForOpenAI() {
    await this.openAILimiter.waitIfNeeded();
  }
  
  async waitForGoogleSearch() {
    await this.googleSearchLimiter.waitIfNeeded();
  }
  
  updateOpenAITokens(tokens) {
    this.openAILimiter.updateTokenUsage(tokens);
  }
}

// グローバルレート制限マネージャーインスタンス
const rateLimitManager = new APIRateLimitManager();

// ===========================================
// API認証・ヘッダー取得
// ===========================================

/**
 * OpenAI APIヘッダー取得
 */
function getOpenAIHeaders() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI API Keyが設定されていません。PropertiesServiceで OPENAI_API_KEY を設定してください。');
  }
  return {
    'Authorization': `Bearer ${apiKey}`,
    'Content-Type': 'application/json'
  };
}

/**
 * Google Custom Search APIパラメータ取得
 */
function getGoogleSearchParams() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
  const engineId = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
  
  if (!apiKey || !engineId) {
    throw new Error('Google Search API設定が不完全です。GOOGLE_SEARCH_API_KEY と GOOGLE_SEARCH_ENGINE_ID を設定してください。');
  }
  
  return { apiKey, engineId };
}

// ===========================================
// エラーハンドリング
// ===========================================

/**
 * 指数バックオフ付きAPI呼び出し
 */
function apiCallWithRetry(apiCall, maxRetries = 3, baseDelay = 1000) {
  let attempt = 0;
  
  function makeRequest() {
    try {
      return apiCall();
    } catch (error) {
      attempt++;
      if (attempt >= maxRetries) {
        throw error;
      }
      
      const delay = baseDelay * Math.pow(2, attempt - 1);
      Logger.log(`API呼び出し失敗 (試行${attempt}/${maxRetries}): ${delay}ms後にリトライします`);
      Utilities.sleep(delay);
      return makeRequest();
    }
  }
  
  return makeRequest();
}

/**
 * OpenAI APIエラーハンドリング
 */
function handleOpenAIError(error, context) {
  const errorMap = {
    401: 'API キーが無効です',
    429: 'API使用量制限に達しました',
    500: 'OpenAIサーバーエラーです',
    503: 'OpenAIサービスが一時的に利用できません'
  };
  
  Logger.log(`OpenAI API Error in ${context}: ${error}`);
  
  if (error.code && errorMap[error.code]) {
    throw new Error(errorMap[error.code]);
  }
  
  throw new Error('OpenAI APIでエラーが発生しました');
}

/**
 * Google Search APIエラーハンドリング
 */
function handleGoogleSearchError(error, context) {
  const errorMap = {
    400: '検索クエリが無効です',
    403: 'API制限に達しました、または認証エラーです',
    500: 'Google Search APIサーバーエラーです'
  };
  
  Logger.log(`Google Search API Error in ${context}: ${error}`);
  
  if (error.code && errorMap[error.code]) {
    throw new Error(errorMap[error.code]);
  }
  
  throw new Error('Google Search APIでエラーが発生しました');
}

// ===========================================
// OpenAI ChatGPT API統合
// ===========================================

/**
 * OpenAI ChatGPT APIでキーワード生成
 */
async function generateKeywordsWithChatGPT(productInfo) {
  try {
    await rateLimitManager.waitForOpenAI();
    
    const requestData = {
      model: OPENAI_CONFIG.model,
      messages: [
        {
          role: 'system',
          content: `あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。

以下の4つのカテゴリで、各5個ずつ、合計20個のキーワードを生成してください：

1. 業界キーワード - 対象業界の具体的な名称
2. 課題キーワード - 企業が抱える問題・悩み  
3. 技術キーワード - 関連技術・システム・ツール
4. 成果キーワード - 期待される結果・効果

JSON形式で以下の構造で回答してください：
{
  "industryKeywords": ["キーワード1", "キーワード2", ...],
  "challengeKeywords": ["キーワード1", "キーワード2", ...],
  "technologyKeywords": ["キーワード1", "キーワード2", ...],
  "resultKeywords": ["キーワード1", "キーワード2", ...]
}`
        },
        {
          role: 'user',
          content: `商材名: ${productInfo.name}
商材概要: ${productInfo.description}
対象企業規模: ${productInfo.targetCompanySize || '中小企業〜大企業'}
価格帯: ${productInfo.priceRange || '要相談'}
主な機能: ${productInfo.features || ''}`
        }
      ],
      max_tokens: OPENAI_CONFIG.maxTokens,
      temperature: OPENAI_CONFIG.temperature
    };

    const response = apiCallWithRetry(() => {
      const apiResponse = UrlFetchApp.fetch(`${OPENAI_CONFIG.baseURL}/chat/completions`, {
        method: 'POST',
        headers: getOpenAIHeaders(),
        payload: JSON.stringify(requestData)
      });
      
      if (apiResponse.getResponseCode() !== 200) {
        throw new Error(`OpenAI API Error: ${apiResponse.getResponseCode()}`);
      }
      
      return JSON.parse(apiResponse.getContentText());
    });

    // トークン使用量更新
    if (response.usage && response.usage.total_tokens) {
      rateLimitManager.updateOpenAITokens(response.usage.total_tokens);
    }

    // レスポンス解析
    const content = response.choices[0].message.content;
    return JSON.parse(content);
    
  } catch (error) {
    handleOpenAIError(error, 'generateKeywordsWithChatGPT');
  }
}

/**
 * OpenAI ChatGPT APIで提案メッセージ生成
 */
async function generateProposalWithChatGPT(companyInfo, productInfo) {
  try {
    await rateLimitManager.waitForOpenAI();
    
    const requestData = {
      model: OPENAI_CONFIG.model,
      messages: [
        {
          role: 'system',
          content: `あなたは経験豊富な営業担当者です。企業情報と商材情報から、パーソナライズされた営業提案メッセージを作成してください。

要件：
- 企業の業界・規模に特化した内容
- 具体的な導入効果を示す
- 次のアクションを明確にする
- 日本のビジネスマナーに配慮
- 300-500文字程度

JSON形式で以下の構造で回答してください：
{
  "subject": "メール件名",
  "proposal": "提案メッセージ本文",
  "nextAction": "提案する次のステップ",
  "expectedBenefit": "期待される導入効果"
}`
        },
        {
          role: 'user',
          content: `【企業情報】
企業名: ${companyInfo.companyName}
業界: ${companyInfo.industry || '不明'}
企業規模: ${companyInfo.companySize || '不明'}
ウェブサイト: ${companyInfo.website}
企業概要: ${companyInfo.description || ''}

【商材情報】
商材名: ${productInfo.name}
商材概要: ${productInfo.description}
対象企業規模: ${productInfo.targetCompanySize || ''}
価格帯: ${productInfo.priceRange || ''}
主な機能: ${productInfo.features || ''}`
        }
      ],
      max_tokens: OPENAI_CONFIG.maxTokens,
      temperature: 0.8
    };

    const response = apiCallWithRetry(() => {
      const apiResponse = UrlFetchApp.fetch(`${OPENAI_CONFIG.baseURL}/chat/completions`, {
        method: 'POST',
        headers: getOpenAIHeaders(),
        payload: JSON.stringify(requestData)
      });
      
      if (apiResponse.getResponseCode() !== 200) {
        throw new Error(`OpenAI API Error: ${apiResponse.getResponseCode()}`);
      }
      
      return JSON.parse(apiResponse.getContentText());
    });

    // トークン使用量更新
    if (response.usage && response.usage.total_tokens) {
      rateLimitManager.updateOpenAITokens(response.usage.total_tokens);
    }

    // レスポンス解析
    const content = response.choices[0].message.content;
    return JSON.parse(content);
    
  } catch (error) {
    handleOpenAIError(error, 'generateProposalWithChatGPT');
  }
}

// ===========================================
// Google Custom Search API統合
// ===========================================

/**
 * Google Custom Search APIで企業検索
 */
async function searchCompaniesWithGoogle(keyword, startIndex = 1) {
  try {
    await rateLimitManager.waitForGoogleSearch();
    
    const { apiKey, engineId } = getGoogleSearchParams();
    const searchParams = {
      ...GOOGLE_SEARCH_CONFIG.defaultParams,
      key: apiKey,
      cx: engineId,
      q: keyword,
      start: startIndex
    };
    
    // URLパラメータ構築
    const url = `${GOOGLE_SEARCH_CONFIG.baseURL}?${Object.entries(searchParams)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&')}`;
    
    const response = apiCallWithRetry(() => {
      const apiResponse = UrlFetchApp.fetch(url, {
        method: 'GET'
      });
      
      if (apiResponse.getResponseCode() !== 200) {
        throw new Error(`Google Search API Error: ${apiResponse.getResponseCode()}`);
      }
      
      return JSON.parse(apiResponse.getContentText());
    });

    // 企業情報抽出
    const companies = [];
    if (response.items) {
      for (const item of response.items) {
        try {
          const companyInfo = extractCompanyInfo(item);
          if (companyInfo && isValidCompany(companyInfo)) {
            companies.push(companyInfo);
          }
        } catch (error) {
          Logger.log(`企業情報抽出エラー: ${error}`);
        }
      }
    }
    
    return companies;
    
  } catch (error) {
    handleGoogleSearchError(error, 'searchCompaniesWithGoogle');
  }
}

/**
 * 検索結果から企業情報を抽出
 */
function extractCompanyInfo(searchItem) {
  const url = searchItem.link;
  const domain = extractDomain(url);
  
  return {
    companyId: generateCompanyId(domain),
    companyName: extractCompanyName(searchItem.title),
    website: url,
    domain: domain,
    description: cleanText(searchItem.snippet),
    lastUpdated: new Date(),
    searchKeyword: '', // 後で設定
    source: 'Google Custom Search'
  };
}

/**
 * 有効な企業データかチェック
 */
function isValidCompany(companyInfo) {
  // 無効なパターンをフィルタリング
  const invalidPatterns = [
    'wikipedia', 'facebook.com', 'twitter.com', 'linkedin.com',
    'youtube.com', 'amazon.com', 'google.com', 'yahoo.com'
  ];
  
  const domain = companyInfo.domain.toLowerCase();
  return !invalidPatterns.some(pattern => domain.includes(pattern)) &&
         companyInfo.companyName &&
         companyInfo.companyName.length > 2;
}

/**
 * ドメイン抽出
 */
function extractDomain(url) {
  try {
    const urlObj = new URL(url);
    return urlObj.hostname.replace('www.', '');
  } catch (error) {
    return '';
  }
}

/**
 * 企業名抽出
 */
function extractCompanyName(title) {
  // タイトルから企業名を抽出するロジック
  return title.split('|')[0].split('-')[0].split('・')[0].trim();
}

/**
 * テキストクリーニング
 */
function cleanText(text) {
  return text ? text.replace(/\s+/g, ' ').trim() : '';
}

/**
 * 企業ID生成
 */
function generateCompanyId(domain) {
  return `comp_${domain.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}`;
}

// ===========================================
// システム健全性チェック
// ===========================================

/**
 * システム全体の健全性チェック
 */
function performHealthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    sheets: false,
    openaiAPI: false,
    googleSearchAPI: false,
    errors: []
  };
  
  try {
    // シート存在チェック
    results.sheets = checkSystemInitialization();
    Logger.log('✅ シートチェック: 正常');
  } catch (error) {
    results.errors.push(`シートエラー: ${error.toString()}`);
    Logger.log('❌ シートチェック: 異常');
  }
  
  try {
    // OpenAI API接続チェック
    results.openaiAPI = checkOpenAIConnection();
    Logger.log('✅ OpenAI APIチェック: 正常');
  } catch (error) {
    results.errors.push(`OpenAI APIエラー: ${error.toString()}`);
    Logger.log('❌ OpenAI APIチェック: 異常');
  }
  
  try {
    // Google Search API接続チェック
    results.googleSearchAPI = checkGoogleSearchConnection();
    Logger.log('✅ Google Search APIチェック: 正常');
  } catch (error) {
    results.errors.push(`Google Search APIエラー: ${error.toString()}`);
    Logger.log('❌ Google Search APIチェック: 異常');
  }
  
  // 結果をログシートに記録
  try {
    logHealthCheckResults(results);
  } catch (error) {
    Logger.log('ヘルスチェック結果の記録に失敗:', error);
  }
  
  return results;
}

/**
 * OpenAI API接続チェック
 */
function checkOpenAIConnection() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      throw new Error('OPENAI_API_KEY が設定されていません');
    }
    
    // 簡単なテストリクエスト
    const testRequest = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: 'Hello! Please respond with "OK" to confirm the connection.'
        }
      ],
      max_tokens: 10
    };
    
    const response = UrlFetchApp.fetch(`${OPENAI_CONFIG.baseURL}/chat/completions`, {
      method: 'POST',
      headers: getOpenAIHeaders(),
      payload: JSON.stringify(testRequest)
    });
    
    if (response.getResponseCode() === 200) {
      return true;
    } else {
      throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }
    
  } catch (error) {
    throw new Error(`OpenAI API接続失敗: ${error.toString()}`);
  }
}

/**
 * Google Search API接続チェック
 */
function checkGoogleSearchConnection() {
  try {
    const { apiKey, engineId } = getGoogleSearchParams();
    
    // 簡単なテスト検索
    const testUrl = `${GOOGLE_SEARCH_CONFIG.baseURL}?key=${apiKey}&cx=${engineId}&q=test&num=1`;
    
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET'
    });
    
    if (response.getResponseCode() === 200) {
      return true;
    } else {
      throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }
    
  } catch (error) {
    throw new Error(`Google Search API接続失敗: ${error.toString()}`);
  }
}

/**
 * ヘルスチェック結果をログシートに記録
 */
function logHealthCheckResults(results) {
  try {
    const sheet = getSafeSheet(SHEET_NAMES.LOGS);
    const status = results.errors.length === 0 ? '正常' : '異常';
    const errorSummary = results.errors.length > 0 ? results.errors.join('; ') : '';
    
    sheet.appendRow([
      results.timestamp,
      'システムヘルスチェック',
      status,
      `シート:${results.sheets ? '○' : '×'}, OpenAI:${results.openaiAPI ? '○' : '×'}, Google:${results.googleSearchAPI ? '○' : '×'}`,
      errorSummary
    ]);
    
  } catch (error) {
    Logger.log('ヘルスチェック結果記録エラー:', error);
  }
}

/**
 * API設定状況の確認
 */
function checkAPIConfiguration() {
  const config = {
    openaiKey: false,
    googleSearchKey: false,
    googleSearchEngineId: false
  };
  
  try {
    config.openaiKey = !!PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    config.googleSearchKey = !!PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
    config.googleSearchEngineId = !!PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
  } catch (error) {
    Logger.log('API設定確認エラー:', error);
  }
  
  return config;
}

/**
 * システム初期化状況の確認（スプレッドシートID指定版）
 */
function checkSystemInitializationById(spreadsheetId = null) {
  try {
    let spreadsheet;
    if (spreadsheetId) {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } else {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    if (!spreadsheet) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length > 0) {
      console.log('不足しているシート:', missingSheets);
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    console.log('✅ システム初期化済み - すべてのシートが利用可能です');
    return true;
  } catch (error) {
    console.error('❌ システム初期化チェックエラー:', error.toString());
    throw error;
  }
}

/**
 * システム初期化状況の確認
 */
function checkSystemInitialization() {
  try {
    const spreadsheet = getSafeSpreadsheet();
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length > 0) {
      console.log('不足しているシート:', missingSheets);
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    console.log('✅ システム初期化済み - すべてのシートが利用可能です');
    return true;
  } catch (error) {
    console.error('❌ システム初期化チェックエラー:', error.toString());
    throw error;
  }
}

/**
 * スプレッドシートの初期化
 */
function initializeSheets() {
  const ss = getSafeSpreadsheet();
  
  // 制御パネルシートの作成
  createControlPanel(ss);
  
  // プラン説明シートの作成
  createPlanInfoSheet(ss);
  
  // 各データシートの作成
  createKeywordsSheet(ss);
  createCompaniesSheet(ss);
  createProposalsSheet(ss);
  createLogsSheet(ss);
  
  Logger.log('システムの初期化が完了しました');
}



/**
 * 制御パネルシートの作成
 */
function createControlPanel(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONTROL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONTROL);
  }
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('営業自動化システム').setFontSize(16).setFontWeight('bold');
  
  // 入力項目設定
  const inputData = [
    ['商材名', ''],
    ['商材概要', ''],
    ['価格帯', '低価格'],
    ['対象企業規模', '中小企業'],
    ['優先地域', ''],
    ['', ''],
    ['実行ボタン', ''],
    ['', ''],
    ['', ''],
    ['検索企業数上限', 20],
    ['APIキー設定', '確認中...'],
    ['', ''],
    ['', ''],
    ['実行状況表示', '']
  ];
  
  sheet.getRange(2, 1, inputData.length, 2).setValues(inputData);
  
  // データ検証の設定
  const priceValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['低価格', '中価格', '高価格'])
    .build();
  sheet.getRange('B4').setDataValidation(priceValidation);
  
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['個人事業主', '中小企業', '大企業', 'すべて'])
    .build();
  sheet.getRange('B5').setDataValidation(sizeValidation);
  
  // 実行ボタンの設定
  createExecutionButtons(sheet);
  
  // ダッシュボード作成
  createDashboard(sheet);
  
  // API設定状況を更新
  updateControlPanelApiStatus(sheet);
}

/**
 * 実行ボタンの作成
 */
function createExecutionButtons(sheet) {
  // キーワード生成ボタン
  sheet.getRange('A8').setValue('キーワード生成');
  sheet.getRange('B8').setValue('企業検索');
  sheet.getRange('C8').setValue('全自動実行');
  
  // ボタンの背景色設定
  sheet.getRange('A8:C8').setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
}

/**
 * ダッシュボード作成
 */
function createDashboard(sheet) {
  // ダッシュボードエリア
  sheet.getRange('E1').setValue('システムダッシュボード').setFontSize(14).setFontWeight('bold');
  
  const dashboardData = [
    ['登録済み企業数', '=COUNTA(企業マスター!A:A)-1'],
    ['生成済みキーワード数', '=COUNTA(キーワード一覧!A:A)-1'],
    ['提案メッセージ数', '=COUNTA(提案メッセージ!A:A)-1'],
    ['システム稼働日数', '=TODAY()-DATE(2024,1,1)']
  ];
  
  sheet.getRange(2, 5, dashboardData.length, 2).setValues(dashboardData);
}

/**
 * 制御パネルのAPI設定状況を更新
 */
function updateControlPanelApiStatus(sheet) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    const openaiKey = properties.getProperty('OPENAI_API_KEY');
    const googleKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
    const engineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    let statusText = '❌ 未設定';
    if (openaiKey && googleKey && engineId) {
      statusText = '✅ 設定済み';
    } else if (openaiKey || googleKey || engineId) {
      statusText = '⚠️ 一部設定済み';
    }
    
    // APIキー設定行を更新（B12セル）
    sheet.getRange('B12').setValue(statusText);
    
  } catch (error) {
    console.error('制御パネルAPI設定状況更新エラー:', error);
    sheet.getRange('B12').setValue('❌ エラー');
  }
}

/**
 * キーワードシートの作成
 */
function createKeywordsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.KEYWORDS);
  }
  
  const headers = ['キーワード', 'カテゴリ', '優先度', '実行状況', '発見企業数', '実行日時', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e1f5fe');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 200);
}

/**
 * 企業シートの作成
 */
function createCompaniesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COMPANIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.COMPANIES);
  }
  
  const headers = [
    '企業名', 'ウェブサイト', '業界', '企業規模', '所在地', '説明',
    'マッチ度スコア', '発見キーワード', '最終更新日', '連絡状況', '備考'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e8');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
}

/**
 * 提案メッセージシートの作成
 */
function createProposalsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.PROPOSALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PROPOSALS);
  }
  
  const headers = [
    '企業名', 'メール件名', '提案メッセージ', '生成日時', '送信状況', '返信', '備考'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fff3e0');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
}

/**
 * ログシートの作成
 */
function createLogsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.LOGS);
  }
  
  const headers = ['日時', '処理', '結果', '詳細', 'エラー'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3e5f5');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 300);
}

/**
 * コントロールパネルの設定値を取得
 */
function getControlPanelSettings() {
  try {
    const sheet = getSafeSpreadsheet().getSheetByName(SHEET_NAMES.CONTROL);
    if (!sheet) {
      throw new Error('制御パネルシートが見つかりません');
    }
    
    const data = sheet.getRange('A2:B15').getValues();
    
    const settings = {
      productName: data[0][1] || '',
      productDescription: data[1][1] || '',
      priceRange: data[2][1] || '低価格',
      targetSize: data[3][1] || '中小企業',
      targetRegion: data[4][1] || '',
      searchResultsPerKeyword: data[9][1] || 20
    };
    
    console.log('制御パネル設定を取得しました:', settings);
    return settings;
  } catch (error) {
    console.error('制御パネル設定取得エラー:', error);
    throw error;
  }
}

/**
 * 結果をシートに保存
 */
function saveToSheet(sheetName, data, headers = null) {
  try {
    const sheet = getSafeSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }
    
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    if (data && data.length > 0) {
      const startRow = headers ? 2 : sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    }
    
    console.log(`${sheetName}に${data.length}件のデータを保存しました`);
  } catch (error) {
    console.error(`${sheetName}への保存エラー:`, error);
    throw error;
  }
}

/**
 * システムの健全性チェック
 */
function healthCheck() {
  const results = {
    timestamp: new Date(),
    spreadsheet: false,
    sheets: false,
    apis: false,
    errors: []
  };
  
  try {
    // スプレッドシートアクセス
    const ss = getSafeSpreadsheet();
    results.spreadsheet = true;
    
    // 必要シートの存在確認
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = ss.getSheets().map(sheet => sheet.getName());
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length === 0) {
      results.sheets = true;
    } else {
      results.errors.push(`不足シート: ${missingSheets.join(', ')}`);
    }
    
    // API設定確認
    const properties = PropertiesService.getScriptProperties();
    const hasGoogleApi = properties.getProperty('GOOGLE_SEARCH_API_KEY') && 
                        properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    const hasOpenAI = properties.getProperty('OPENAI_API_KEY');
    
    if (hasGoogleApi && hasOpenAI) {
      results.apis = true;
    } else {
      const missingApis = [];
      if (!hasGoogleApi) missingApis.push('Google Search API');
      if (!hasOpenAI) missingApis.push('OpenAI API');
      results.errors.push(`未設定API: ${missingApis.join(', ')}`);
    }
    
  } catch (error) {
    results.errors.push(error.toString());
  }
  
  return results;
}

/**
 * システム状況をユーザーに表示
 */
function showSystemStatus() {
  try {
    const health = healthCheck();
    
    if (health.spreadsheet && health.sheets && health.apis) {
      SpreadsheetApp.getUi().alert(
        '✅ システム正常',
        'システムは正常に動作しています。すべての機能をご利用いただけます。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      const message = `❌ システムに問題があります。\n\n解決方法:\n1. メニューバーから「拡張機能」→「Apps Script」を選択\n2. 関数一覧から「initializeSheets」を選択\n3. 実行ボタンをクリックしてシステムを初期化`;
      
      SpreadsheetApp.getUi().alert(
        'システム状況',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    console.log('システム状況:', health);
  } catch (error) {
    console.error('システム状況表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ エラー',
      `システム状況の確認中にエラーが発生しました:\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * カスタムメニューの作成
 */
function onOpen_DISABLED() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // システム状況を確認
    const health = healthCheck();
    
    if (!health.spreadsheet || !health.sheets) {
      // システム未初期化の場合は初期化メニューのみ表示
      ui.createMenu('🔧 営業自動化システム')
        .addItem('システム初期化', 'initializeSheets')
        .addItem('システム状況確認', 'showSystemStatus')
        .addToUi();
      return;
    }
    
    // 通常メニュー
    ui.createMenu('🎯 営業自動化システム')
      .addSubMenu(ui.createMenu('🔍 企業発掘')
        .addItem('キーワード生成', 'generateKeywords')
        .addItem('企業検索', 'searchCompanies')
        .addItem('全自動実行', 'executeFullWorkflow'))
      .addSubMenu(ui.createMenu('📝 提案生成')
        .addItem('AI提案生成', 'generateProposals')
        .addItem('提案一括生成', 'batchGenerateProposals'))
      .addSubMenu(ui.createMenu('⚙️ 設定')
        .addItem('API キー設定', 'showApiKeySettings')
        .addItem('システム設定', 'showSystemSettings'))
      .addSubMenu(ui.createMenu('ℹ️ ヘルプ')
        .addItem('使用方法', 'showUserGuide')
        .addItem('システム状況', 'showSystemStatus')
        .addItem('システム初期化', 'initializeSheets'))
      .addToUi();
      
  } catch (error) {
    console.error('メニュー作成エラー:', error);
    // エラー時は基本メニューのみ表示
    SpreadsheetApp.getUi()
      .createMenu('🔧 営業自動化システム (エラー)')
      .addItem('システム初期化', 'initializeSheets')
      .addToUi();
  }
}

/**
 * API キー設定ダイアログ
 */
function showApiKeySettings() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('api-key-settings')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'API キー設定');
}

/**
 * APIキーをスクリプトプロパティに保存
 */
function saveApiKeysToProperties(googleSearchKey, searchEngineId, openaiKey) {
  const properties = PropertiesService.getScriptProperties();
  
  if (googleSearchKey) {
    properties.setProperty('GOOGLE_SEARCH_API_KEY', googleSearchKey);
  }
  if (searchEngineId) {
    properties.setProperty('GOOGLE_SEARCH_ENGINE_ID', searchEngineId);
  }
  if (openaiKey) {
    properties.setProperty('OPENAI_API_KEY', openaiKey);
  }
  
  Logger.log('API キーがプロパティに保存されました');
}

/**
 * プラン説明シートの作成
 */
function createPlanInfoSheet(ss) {
  let sheet = ss.getSheetByName('プラン説明');
  if (!sheet) {
    sheet = ss.insertSheet('プラン説明', 1); // 2番目の位置に挿入
  } else {
    sheet.clear(); // 既存の場合はクリア
  }
  
  // タイトル設定
  sheet.getRange('A1').setValue('🎯 営業自動化システム - 課金プラン説明')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white');
  sheet.getRange('A1:G1').merge();
  
  // サブタイトル
  sheet.getRange('A2').setValue('各プランの機能比較表')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.getRange('A2:G2').merge();
  
  // ヘッダー行
  const headers = ['項目', 'ベーシック\n¥500/月', 'スタンダード\n¥1,500/月', 'プロフェッショナル\n¥5,500/月', 'エンタープライズ\n¥17,500/月'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // プラン比較データ
  const planData = [
    ['💰 月額料金', '¥500', '¥1,500', '¥5,500', '¥17,500'],
    ['🔢 アカウント数', '1アカウント', '1アカウント', '2アカウント', '5アカウント'],
    ['', '', '', '', ''],
    ['🔍 基本機能', '', '', '', ''],
    ['手動企業入力', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['企業検索上限', '10社/日', '50社/日', '100社/日', '500社/日'],
    ['基本レポート', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['', '', '', '', ''],
    ['🤖 AI機能', '', '', '', ''],
    ['キーワード自動生成', '❌ 利用不可', '✅ 利用可能\n(20個/回)', '✅ 利用可能\n(40個/回)', '✅ 利用可能\n(100個/回)'],
    ['AI提案文生成', '❌ 利用不可', '✅ 利用可能\n(50件/日)', '✅ 利用可能\n(100件/日)', '✅ 利用可能\n(500件/日)'],
    ['', '', '', '', ''],
    ['📊 高度な機能', '', '', '', ''],
    ['バッチ処理', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['カスタム分析', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['データエクスポート', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['高速処理', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['マルチアカウント', '❌ 利用不可', '❌ 利用不可', '✅ 2アカウント', '✅ 5アカウント'],
    ['カスタム連携', '❌ 利用不可', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能'],
    ['', '', '', '', ''],
    ['🛠️ サポート', '', '', '', ''],
    ['メールサポート', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['優先サポート', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['電話サポート', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['専任サポート', '❌ 利用不可', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能']
  ];
  
  sheet.getRange(5, 1, planData.length, planData[0].length).setValues(planData);
  
  // セルスタイル設定
  const dataRange = sheet.getRange(5, 1, planData.length, planData[0].length);
  
  // 項目列のスタイル
  sheet.getRange(5, 1, planData.length, 1)
    .setBackground('#f8f9fa')
    .setFontWeight('bold');
  
  // プラン列のスタイル
  for (let col = 2; col <= 5; col++) {
    const columnRange = sheet.getRange(5, col, planData.length, 1);
    columnRange.setHorizontalAlignment('center');
    
    // プラン別の背景色
    switch (col) {
      case 2: // ベーシック
        columnRange.setBackground('#fff3e0');
        break;
      case 3: // スタンダード
        columnRange.setBackground('#e8f5e8');
        break;
      case 4: // プロフェッショナル
        columnRange.setBackground('#e3f2fd');
        break;
      case 5: // エンタープライズ
        columnRange.setBackground('#fce4ec');
        break;
    }
  }
  
  // ボーダー設定
  dataRange.setBorder(true, true, true, true, true, true);
  
  // 特別な行のスタイル（カテゴリ見出し行）
  const categoryRows = [7, 10, 14, 22]; // 「🔍 基本機能」「🤖 AI機能」「📊 高度な機能」「🛠️ サポート」
  categoryRows.forEach(row => {
    sheet.getRange(row + 4, 1, 1, 5) // +4はヘッダー分のオフセット
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
  });
  
  // 現在のプラン表示エリア
  sheet.getRange('A' + (5 + planData.length + 2)).setValue('📋 現在のプラン情報')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white');
  sheet.getRange('A' + (5 + planData.length + 2) + ':C' + (5 + planData.length + 2)).merge();
  
  const currentPlanRow = 5 + planData.length + 3;
  sheet.getRange('A' + currentPlanRow).setValue('現在のプラン:');
  
  // getUserPlan関数があるかチェックして安全に呼び出し
  let currentPlan = 'ベーシック';
  try {
    if (typeof getUserPlan === 'function') {
      currentPlan = getUserPlan() || 'ベーシック';
    }
  } catch (error) {
    console.log('getUserPlan関数が見つかりません。デフォルト値を使用します。');
  }
  
  sheet.getRange('B' + currentPlanRow).setValue(currentPlan);
  
  sheet.getRange('A' + (currentPlanRow + 1)).setValue('月額料金:');
  sheet.getRange('B' + (currentPlanRow + 1)).setValue('=IF(B' + currentPlanRow + '="ベーシック","¥500",IF(B' + currentPlanRow + '="スタンダード","¥1,500",IF(B' + currentPlanRow + '="プロフェッショナル","¥5,500","¥17,500")))');
  
  sheet.getRange('A' + (currentPlanRow + 2)).setValue('今日の使用量:');
  sheet.getRange('B' + (currentPlanRow + 2)).setValue('企業検索: 0回');
  
  // プランアップグレード案内
  sheet.getRange('A' + (currentPlanRow + 4)).setValue('💡 プランアップグレードについて')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');
  sheet.getRange('A' + (currentPlanRow + 4) + ':C' + (currentPlanRow + 4)).merge();
  
  const upgradeInfo = [
    ['• AI機能をご利用になりたい場合は「スタンダード」以上をお選びください'],
    ['• より多くの企業検索をお求めの場合は上位プランをご検討ください'],
    ['• 高速処理・優先サポートは「プロフェッショナル」以上で利用可能です'],
    ['• 企業向け機能・専任サポートは「エンタープライズ」プランをご利用ください']
  ];
  
  sheet.getRange(currentPlanRow + 5, 1, upgradeInfo.length, 1).setValues(upgradeInfo);
  
  // API料金情報セクション
  const apiCostStartRow = currentPlanRow + 5 + upgradeInfo.length + 2;
  
  sheet.getRange('A' + apiCostStartRow).setValue('💳 API利用料金について（別途課金）')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#d32f2f')
    .setFontColor('white');
  sheet.getRange('A' + apiCostStartRow + ':E' + apiCostStartRow).merge();
  
  const apiCostInfo = [
    ['⚠️ 重要: 以下のAPI利用料金が別途発生します', '', '', '', ''],
    ['', '', '', '', ''],
    ['API名', '料金体系', '無料枠', '超過時料金', '備考'],
    ['OpenAI ChatGPT API', 'トークン従量課金', '$5.00相当/月', '$0.002/1Kトークン', 'AI機能利用時のみ'],
    ['Google Custom Search API', 'リクエスト従量課金', '100回/日', '$5.00/1000回', '企業検索時のみ'],
    ['', '', '', '', ''],
    ['📊 プラン別API料金目安（最大想定）', '', '', '', ''],
    ['プラン', 'システム利用料', 'API料金（最大）', '合計月額', 'アカウント数'],
    ['ベーシック', '¥500', '¥0', '¥500', '1アカウント'],
    ['スタンダード', '¥1,500', '¥2,000', '¥3,500', '1アカウント'],
    ['プロフェッショナル', '¥5,500', '¥6,000', '¥11,500', '2アカウント'],
    ['エンタープライズ', '¥17,500', '¥30,000', '¥47,500', '5アカウント'],
    ['', '', '', '', ''],
    ['💡 API料金節約のコツ', '', '', '', ''],
    ['• ベーシックプランではAI機能を使わないため、OpenAI料金は¥0です', '', '', '', ''],
    ['• 企業検索は1日100回まで無料なので、Google検索料金も抑えられます', '', '', '', ''],
    ['• マルチアカウントプランでは使用量を分散して効率的に利用できます', '', '', '', ''],
    ['• API料金はGoogle・OpenAIに直接支払い（当社には支払われません）', '', '', '', ''],
    ['• 実際の料金は使用量により変動し、上記は最大想定値です', '', '', '', '']
  ];
  
  sheet.getRange(apiCostStartRow + 1, 1, apiCostInfo.length, 5).setValues(apiCostInfo);
  
  // API料金表のスタイル設定
  // ヘッダー行のスタイル
  sheet.getRange(apiCostStartRow + 3, 1, 1, 5)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(apiCostStartRow + 7, 1, 1, 5)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // API料金データ行のスタイル
  sheet.getRange(apiCostStartRow + 4, 1, 2, 5)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(apiCostStartRow + 8, 1, 4, 4)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  // 注意事項のスタイル
  sheet.getRange(apiCostStartRow + 14, 1, 1, 5)
    .setBackground('#ff9800')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // 節約コツ部分のスタイル
  sheet.getRange(apiCostStartRow + 15, 1, 4, 1)
    .setBackground('#e8f5e8')
    .setFontSize(10);
  
  // 料金に関する重要な注意喚起
  const importantNoticeRow = apiCostStartRow + apiCostInfo.length + 2;
  
  sheet.getRange('A' + importantNoticeRow).setValue('🚨 API料金に関する重要な注意事項')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#d32f2f')
    .setFontColor('white');
  sheet.getRange('A' + importantNoticeRow + ':E' + importantNoticeRow).merge();
  
  const importantNotices = [
    ['1. API料金は当社プラン料金とは別に、各APIプロバイダーから直接請求されます'],
    ['2. API料金は使用量に応じて変動するため、事前に予算をご確認ください'],
    ['3. 無料枠を超過した場合、自動的に有料課金が開始されます'],
    ['4. API料金の詳細は各プロバイダーの公式サイトでご確認ください'],
    ['   • OpenAI API: https://openai.com/pricing'],
    ['   • Google Custom Search API: https://developers.google.com/custom-search/v1/overview']
  ];
  
  sheet.getRange(importantNoticeRow + 1, 1, importantNotices.length, 1).setValues(importantNotices);
  sheet.getRange(importantNoticeRow + 1, 1, importantNotices.length, 1)
    .setBackground('#ffebee')
    .setFontSize(10)
    .setWrap(true);
  
  // 列幅調整（API料金表対応）
  sheet.setColumnWidth(1, 250); // 項目列
  sheet.setColumnWidth(2, 180); // ベーシック
  sheet.setColumnWidth(3, 120); // スタンダード
  sheet.setColumnWidth(4, 150); // プロフェッショナル
  sheet.setColumnWidth(5, 150); // エンタープライズ
  
  // 行の高さ調整
  sheet.setRowHeight(1, 40); // タイトル
  sheet.setRowHeight(4, 40); // ヘッダー
  
  Logger.log('プラン説明シートが作成されました');
}
