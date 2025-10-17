# API仕様書 - 営業自動化システム v2.0

## 📋 API概要

営業自動化システムは以下の外部APIを統合して動作します：

1. **OpenAI ChatGPT API** - キーワード・提案生成
2. **Google Custom Search API** - 企業検索
3. **Google Sheets API** - データ管理
4. **Google Apps Script APIs** - システム統合

---

## 🤖 OpenAI ChatGPT API統合

### 基本設定
```javascript
const OPENAI_CONFIG = {
  baseURL: 'https://api.openai.com/v1',
  model: 'gpt-3.5-turbo',
  maxTokens: 2000,
  temperature: 0.7
};
```

### 認証
```javascript
function getOpenAIHeaders() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  return {
    'Authorization': `Bearer ${apiKey}`,
    'Content-Type': 'application/json'
  };
}
```

### キーワード生成API

#### リクエスト仕様
```javascript
function generateKeywordsRequest(productInfo) {
  return {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: `あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。

以下の4つのカテゴリで、各5個ずつ、合計20個のキーワードを生成してください：

1. 課題発見系：商材が解決する課題を抱えている企業を見つけるキーワード
2. 成長企業系：成長企業や新規事業に取り組む企業を見つけるキーワード  
3. 予算企業系：投資予算や導入予算を持つ企業を見つけるキーワード
4. タイミング系：導入タイミングが良い企業を見つけるキーワード

出力形式：
{
  "課題発見": ["キーワード1", "キーワード2", ...],
  "成長企業": ["キーワード1", "キーワード2", ...],
  "予算企業": ["キーワード1", "キーワード2", ...],
  "タイミング": ["キーワード1", "キーワード2", ...]
}`
      },
      {
        role: 'user',
        content: `商材名: ${productInfo.productName}
商材概要: ${productInfo.productDescription}
価格帯: ${productInfo.priceRange}
対象企業規模: ${productInfo.targetSize}
対象業界: ${productInfo.targetIndustry || '指定なし'}`
      }
    ],
    max_tokens: 2000,
    temperature: 0.7
  };
}
```

#### レスポンス処理
```javascript
function parseKeywordResponse(response) {
  try {
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      throw new Error(`OpenAI API Error: ${data.error.message}`);
    }
    
    const content = data.choices[0].message.content;
    const keywords = JSON.parse(content);
    
    // キーワード配列に変換
    const result = [];
    for (const [category, keywordList] of Object.entries(keywords)) {
      keywordList.forEach((keyword, index) => {
        result.push({
          keyword: keyword,
          category: category,
          priority: index + 1,
          status: '未実行',
          createdAt: new Date()
        });
      });
    }
    
    return result;
  } catch (error) {
    Logger.log(`キーワード解析エラー: ${error}`);
    throw new Error('キーワード生成レスポンスの解析に失敗しました');
  }
}
```

### 提案生成API

#### リクエスト仕様
```javascript
function generateProposalRequest(companyData, productInfo) {
  return {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: `あなたは経験豊富な営業担当者です。企業情報と商材情報を基に、効果的な営業提案メッセージを作成してください。

出力形式：
{
  "subject": "メール件名",
  "greeting": "挨拶文",
  "introduction": "自己紹介・会社紹介",
  "problemIdentification": "課題提起",
  "solutionProposal": "解決策提案",
  "benefits": "導入メリット",
  "nextStep": "次のアクション提案",
  "closing": "締めの挨拶"
}`
      },
      {
        role: 'user',
        content: `企業情報：
企業名: ${companyData.companyName}
業界: ${companyData.industry}
企業規模: ${companyData.companySize}
ウェブサイト: ${companyData.website}
企業概要: ${companyData.description}

商材情報：
商材名: ${productInfo.productName}
商材概要: ${productInfo.productDescription}
価格帯: ${productInfo.priceRange}
主な機能: ${productInfo.features || ''}

追加要件：
- 企業の業界に特化した内容にしてください
- 具体的な導入効果を示してください
- 次のアクションを明確にしてください`
      }
    ],
    max_tokens: 1500,
    temperature: 0.8
  };
}
```

#### エラーハンドリング
```javascript
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
```

---

## 🔍 Google Custom Search API統合

### 基本設定
```javascript
const GOOGLE_SEARCH_CONFIG = {
  baseURL: 'https://www.googleapis.com/customsearch/v1',
  defaultParams: {
    lr: 'lang_ja',
    cr: 'countryJP',
    safe: 'off',
    num: 10
  }
};
```

### 認証
```javascript
function getGoogleSearchParams() {
  return {
    key: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY'),
    cx: PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID')
  };
}
```

### 企業検索API

#### リクエスト仕様
```javascript
function buildSearchQuery(keyword, settings) {
  const baseQuery = keyword.keyword;
  const filters = [];
  
  // 企業サイト限定
  filters.push('site:co.jp OR site:com OR site:net');
  
  // 除外キーワード
  if (settings.excludeKeywords) {
    const excludes = settings.excludeKeywords.split(',').map(k => `-"${k.trim()}"`);
    filters.push(...excludes);
  }
  
  // 地域指定
  if (settings.targetRegion && settings.targetRegion !== '全国') {
    filters.push(`"${settings.targetRegion}"`);
  }
  
  return {
    q: `${baseQuery} ${filters.join(' ')}`,
    num: Math.min(settings.searchResultsPerKeyword || 10, 10),
    start: 1,
    ...GOOGLE_SEARCH_CONFIG.defaultParams
  };
}

function executeGoogleSearch(searchParams) {
  const authParams = getGoogleSearchParams();
  const allParams = { ...authParams, ...searchParams };
  
  const url = GOOGLE_SEARCH_CONFIG.baseURL + '?' + 
    Object.entries(allParams)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&');
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (SalesAutomationSystem/2.0)'
      }
    });
    
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Google Search API Error: ${error}`);
    throw new Error('企業検索APIでエラーが発生しました');
  }
}
```

#### レスポンス処理
```javascript
function processSearchResults(searchResponse, keyword) {
  if (!searchResponse.items) {
    return [];
  }
  
  const companies = [];
  
  for (const item of searchResponse.items) {
    try {
      const companyData = extractCompanyInfo(item);
      if (companyData && validateCompanyData(companyData)) {
        companies.push({
          ...companyData,
          discoveredBy: keyword.keyword,
          discoveredAt: new Date(),
          searchRank: companies.length + 1
        });
      }
    } catch (error) {
      Logger.log(`企業データ処理エラー: ${error}`);
    }
  }
  
  return companies;
}

function extractCompanyInfo(searchItem) {
  const url = searchItem.link;
  const domain = new URL(url).hostname;
  
  // 基本情報抽出
  const companyInfo = {
    companyId: generateCompanyId(domain),
    companyName: extractCompanyName(searchItem.title),
    website: url,
    domain: domain,
    description: cleanText(searchItem.snippet),
    lastUpdated: new Date()
  };
  
  // 詳細分析（ウェブサイトアクセス）
  try {
    const websiteData = analyzeWebsite(url);
    Object.assign(companyInfo, websiteData);
  } catch (error) {
    Logger.log(`ウェブサイト分析エラー (${url}): ${error}`);
  }
  
  return companyInfo;
}
```

#### API統合レート制限対応

##### OpenAI ChatGPT APIレート制限
```javascript
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
    
    // 日次トークン制限チェック
    if (this.tokensUsed >= this.dailyTokenLimit) {
      throw new Error('OpenAI APIの日次トークン制限に達しました');
    }
    
    this.lastRequestTime = Date.now();
    this.requestCount++;
  }
  
  updateTokenUsage(tokens) {
    this.tokensUsed += tokens;
    Logger.log(`OpenAI API使用トークン: ${tokens} (累計: ${this.tokensUsed}/${this.dailyTokenLimit})`);
  }
  
  sleep(ms) {
    return new Promise(resolve => {
      Utilities.sleep(ms);
      resolve();
    });
  }
}
```

##### Google Search APIレート制限
```javascript
class GoogleSearchRateLimiter {
  constructor() {
    this.lastRequestTime = 0;
    this.requestCount = 0;
    this.dailyLimit = 100; // カスタムサーチAPIの無料枠
    this.perSecondLimit = 10; // 1秒間の最大リクエスト数
  }
  
  async waitIfNeeded() {
    const now = Date.now();
    const timeSinceLastRequest = now - this.lastRequestTime;
    
    // 1.5秒間隔を維持（安全マージン含む）
    if (timeSinceLastRequest < 1500) {
      Logger.log(`Google Search API制限: ${1500 - timeSinceLastRequest}ms待機中...`);
      await this.sleep(1500 - timeSinceLastRequest);
    }
    
    // 日次制限チェック
    if (this.requestCount >= this.dailyLimit) {
      throw new Error('Google Search APIの日次制限に達しました');
    }
    
    this.lastRequestTime = Date.now();
    this.requestCount++;
  }
  
  sleep(ms) {
    return new Promise(resolve => {
      Utilities.sleep(ms);
      resolve();
    });
  }
}
```

##### 統合レート制限マネージャー
```javascript
class APIRateLimitManager {
  constructor() {
    this.openaiLimiter = new OpenAIRateLimiter();
    this.googleSearchLimiter = new GoogleSearchRateLimiter();
    this.lastApiCall = null;
    this.minimumIntervalBetweenDifferentApis = 2000; // 異なるAPI間の最小間隔
  }
  
  async waitForOpenAI() {
    await this.waitBetweenDifferentApis('openai');
    await this.openaiLimiter.waitIfNeeded();
    this.lastApiCall = { api: 'openai', timestamp: Date.now() };
  }
  
  async waitForGoogleSearch() {
    await this.waitBetweenDifferentApis('google-search');
    await this.googleSearchLimiter.waitIfNeeded();
    this.lastApiCall = { api: 'google-search', timestamp: Date.now() };
  }
  
  async waitBetweenDifferentApis(currentApi) {
    if (!this.lastApiCall || this.lastApiCall.api === currentApi) {
      return;
    }
    
    const timeSinceLastCall = Date.now() - this.lastApiCall.timestamp;
    if (timeSinceLastCall < this.minimumIntervalBetweenDifferentApis) {
      const waitTime = this.minimumIntervalBetweenDifferentApis - timeSinceLastCall;
      Logger.log(`異なるAPI間の待機: ${waitTime}ms`);
      await this.sleep(waitTime);
    }
  }
  
  updateOpenAITokens(tokens) {
    this.openaiLimiter.updateTokenUsage(tokens);
  }
  
  sleep(ms) {
    return new Promise(resolve => {
      Utilities.sleep(ms);
      resolve();
    });
  }
  
  getRemainingLimits() {
    return {
      openai: {
        requestsRemaining: this.openaiLimiter.minutelyLimit - this.openaiLimiter.requestCount,
        tokensRemaining: this.openaiLimiter.dailyTokenLimit - this.openaiLimiter.tokensUsed
      },
      googleSearch: {
        requestsRemaining: this.googleSearchLimiter.dailyLimit - this.googleSearchLimiter.requestCount
      }
    };
  }
}
```

---

## 📊 Google Sheets API統合

### シート操作API

#### 安全なシート取得
```javascript
function getSafeSheet(sheetName, createIfNotExists = false) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet && createIfNotExists) {
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet(sheet, sheetName);
  }
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません`);
  }
  
  return sheet;
}
```

#### バッチ書き込み
```javascript
function batchWriteToSheet(sheetName, data, startRow = 2) {
  const sheet = getSafeSheet(sheetName);
  
  if (data.length === 0) return;
  
  const range = sheet.getRange(
    startRow, 
    1, 
    data.length, 
    data[0].length
  );
  
  range.setValues(data);
  SpreadsheetApp.flush(); // 即座に反映
}
```

#### データ検索・フィルタ
```javascript
function findCompaniesWithScore(minScore = 70, maxResults = 10) {
  const sheet = getSafeSheet(SHEET_NAMES.COMPANIES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const scoreColumnIndex = headers.indexOf('マッチ度スコア');
  if (scoreColumnIndex === -1) {
    throw new Error('マッチ度スコア列が見つかりません');
  }
  
  const results = [];
  for (let i = 1; i < data.length && results.length < maxResults; i++) {
    const row = data[i];
    const score = row[scoreColumnIndex];
    
    if (typeof score === 'number' && score >= minScore) {
      const companyData = {};
      headers.forEach((header, index) => {
        companyData[header] = row[index];
      });
      results.push(companyData);
    }
  }
  
  return results.sort((a, b) => b['マッチ度スコア'] - a['マッチ度スコア']);
}
```

---

## 🔧 システム統合API

### プロパティサービス管理
```javascript
class ConfigManager {
  static setConfig(key, value, encrypted = true) {
    const finalValue = encrypted ? this.encrypt(value) : value;
    PropertiesService.getScriptProperties().setProperty(key, finalValue);
  }
  
  static getConfig(key, encrypted = true) {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value && encrypted ? this.decrypt(value) : value;
  }
  
  static validateAPIKeys() {
    const requiredKeys = [
      'OPENAI_API_KEY',
      'GOOGLE_SEARCH_API_KEY', 
      'GOOGLE_SEARCH_ENGINE_ID'
    ];
    
    const missing = requiredKeys.filter(key => !this.getConfig(key));
    
    if (missing.length > 0) {
      throw new Error(`必要なAPI設定が不足しています: ${missing.join(', ')}`);
    }
    
    return true;
  }
}
```

### ユーザーインターフェース統合
```javascript
class UIManager {
  static showProgress(message, percentage = null) {
    const displayMessage = percentage !== null 
      ? `${message} (${percentage}%)`
      : message;
      
    SpreadsheetApp.getActiveSpreadsheet().toast(
      displayMessage,
      '処理中...',
      3
    );
  }
  
  static showSuccess(message, details = '') {
    SpreadsheetApp.getUi().alert(
      '✅ 処理完了',
      `${message}\n\n${details}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  static showError(error, context = '') {
    const errorMessage = context 
      ? `${context}\n\nエラー: ${error.toString()}`
      : error.toString();
      
    SpreadsheetApp.getUi().alert(
      '❌ エラー発生',
      errorMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  static confirmAction(title, message) {
    const response = SpreadsheetApp.getUi().alert(
      title,
      message,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    return response === SpreadsheetApp.getUi().Button.YES;
  }
}
```

---

## 📈 API監視・ログ

### API使用量追跡
```javascript
class APIUsageTracker {
  static trackUsage(apiName, endpoint, tokens = null) {
    const usage = {
      timestamp: new Date(),
      api: apiName,
      endpoint: endpoint,
      tokens: tokens,
      sessionId: Session.getTemporaryActiveUserKey()
    };
    
    this.logToSheet(usage);
    this.updateDailyQuota(apiName, tokens);
  }
  
  static logToSheet(usage) {
    try {
      const sheet = getSafeSheet('API使用ログ', true);
      sheet.appendRow([
        usage.timestamp,
        usage.api,
        usage.endpoint,
        usage.tokens || '',
        usage.sessionId
      ]);
    } catch (error) {
      Logger.log(`API使用量ログエラー: ${error}`);
    }
  }
  
  static getDailyUsage(apiName) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const sheet = getSafeSheet('API使用ログ');
    const data = sheet.getDataRange().getValues();
    
    return data.slice(1).filter(row => {
      const logDate = new Date(row[0]);
      return logDate >= today && row[1] === apiName;
    }).length;
  }
}
```

### パフォーマンス監視
```javascript
class PerformanceMonitor {
  static measureExecution(funcName, func) {
    const startTime = new Date();
    
    try {
      const result = func();
      const endTime = new Date();
      const duration = endTime - startTime;
      
      this.logPerformance(funcName, duration, 'SUCCESS');
      return result;
    } catch (error) {
      const endTime = new Date();
      const duration = endTime - startTime;
      
      this.logPerformance(funcName, duration, 'ERROR', error);
      throw error;
    }
  }
  
  static logPerformance(funcName, duration, status, error = null) {
    const performanceLog = {
      timestamp: new Date(),
      function: funcName,
      duration: duration,
      status: status,
      error: error ? error.toString() : null
    };
    
    Logger.log(`Performance: ${JSON.stringify(performanceLog)}`);
  }
}
```

---

## 🛡️ API セキュリティ

### 入力検証
```javascript
class APIValidator {
  static validateOpenAIRequest(request) {
    if (!request.messages || !Array.isArray(request.messages)) {
      throw new Error('Invalid messages format');
    }
    
    if (request.max_tokens && (request.max_tokens < 1 || request.max_tokens > 4000)) {
      throw new Error('Invalid max_tokens value');
    }
    
    if (request.temperature && (request.temperature < 0 || request.temperature > 2)) {
      throw new Error('Invalid temperature value');
    }
    
    return true;
  }
  
  static validateSearchQuery(query) {
    if (!query || typeof query !== 'string') {
      throw new Error('Invalid search query');
    }
    
    if (query.length > 500) {
      throw new Error('Search query too long');
    }
    
    // 危険な文字列をチェック
    const dangerousPatterns = [
      /<script/i,
      /javascript:/i,
      /on\w+=/i
    ];
    
    if (dangerousPatterns.some(pattern => pattern.test(query))) {
      throw new Error('Query contains dangerous content');
    }
    
    return true;
  }
}
```

### エラー処理統合
```javascript
class APIErrorHandler {
  static handleError(error, apiName, operation) {
    const errorInfo = {
      api: apiName,
      operation: operation,
      error: error.toString(),
      timestamp: new Date(),
      stack: error.stack
    };
    
    // ログ記録
    Logger.log(`API Error: ${JSON.stringify(errorInfo)}`);
    
    // エラー種別による処理分岐
    if (error.message.includes('quota') || error.message.includes('limit')) {
      return this.handleQuotaError(apiName);
    }
    
    if (error.message.includes('network') || error.message.includes('timeout')) {
      return this.handleNetworkError(operation);
    }
    
    if (error.message.includes('auth') || error.message.includes('key')) {
      return this.handleAuthError(apiName);
    }
    
    // 一般的なエラー
    throw new Error(`${apiName} API でエラーが発生しました: ${error.message}`);
  }
  
  static handleQuotaError(apiName) {
    throw new Error(`${apiName} APIの使用量制限に達しました。しばらく時間をおいて再実行してください。`);
  }
  
  static handleNetworkError(operation) {
    throw new Error(`ネットワークエラーが発生しました。インターネット接続を確認して再実行してください。`);
  }
  
  static handleAuthError(apiName) {
    throw new Error(`${apiName} APIの認証に失敗しました。API キーの設定を確認してください。`);
  }
}
```

---

## 📝 API使用例

### 完全な実行フロー例（レート制限対応版）
```javascript
async function executeFullAPIWorkflow() {
  const rateLimitManager = new APIRateLimitManager();
  
  try {
    // 1. 設定検証
    ConfigManager.validateAPIKeys();
    UIManager.showProgress('システム設定を確認中...');
    
    // 2. キーワード生成（OpenAI API）
    UIManager.showProgress('キーワードを生成中...');
    await rateLimitManager.waitForOpenAI();
    
    const settings = getControlPanelSettings();
    const keywords = await generateKeywordsWithChatGPT(settings);
    
    // 使用トークン数を記録
    rateLimitManager.updateOpenAITokens(500); // 推定トークン数
    
    UIManager.showProgress(`キーワード生成完了: ${keywords.length}個`);
    
    // 3. 企業検索（Google Search API）
    const companies = [];
    const totalKeywords = Math.min(keywords.length, 20); // 制限内で実行
    
    for (let i = 0; i < totalKeywords; i++) {
      UIManager.showProgress(`企業検索中... (${i+1}/${totalKeywords})`);
      
      // Google Search APIレート制限待機
      await rateLimitManager.waitForGoogleSearch();
      
      const searchResults = await executeGoogleSearch(keywords[i], settings);
      companies.push(...searchResults);
      
      // 進捗表示
      const progress = Math.round((i + 1) / totalKeywords * 30); // 30%まで
      UIManager.showProgress(`企業検索進行中... ${progress}%`);
    }
    
    // 4. スコア計算（ローカル処理）
    UIManager.showProgress('マッチ度を計算中... 60%');
    companies.forEach(company => {
      company.matchScore = calculateMatchScore(company, settings);
    });
    
    // 5. 高スコア企業の提案生成（OpenAI API）
    const highScoreCompanies = companies
      .filter(c => c.matchScore >= 70)
      .slice(0, 10); // 最大10件まで
      
    const proposals = [];
    
    for (let i = 0; i < highScoreCompanies.length; i++) {
      const progress = 60 + Math.round((i + 1) / highScoreCompanies.length * 30);
      UIManager.showProgress(`AI提案生成中... ${progress}% (${i+1}/${highScoreCompanies.length})`);
      
      // OpenAI APIレート制限待機
      await rateLimitManager.waitForOpenAI();
      
      const proposal = await generateProposalForCompany(highScoreCompanies[i], settings);
      proposals.push(proposal);
      
      // 使用トークン数を記録
      rateLimitManager.updateOpenAITokens(800); // 推定トークン数
    }
    
    // 6. 結果保存
    UIManager.showProgress('結果を保存中... 95%');
    batchWriteToSheet('企業マスター', companies);
    batchWriteToSheet('提案メッセージ', proposals);
    
    // 7. 残り制限を表示
    const remainingLimits = rateLimitManager.getRemainingLimits();
    const limitInfo = 
      `OpenAI残り: ${remainingLimits.openai.requestsRemaining}リクエスト, ${remainingLimits.openai.tokensRemaining}トークン\n` +
      `Google検索残り: ${remainingLimits.googleSearch.requestsRemaining}リクエスト`;
    
    // 8. 完了通知
    UIManager.showSuccess(
      '全自動実行が完了しました',
      `キーワード: ${keywords.length}個\n企業: ${companies.length}社\n提案: ${proposals.length}件\n\n${limitInfo}`
    );
    
  } catch (error) {
    if (error.message.includes('制限')) {
      UIManager.showError(error, 'API使用量制限により処理が中断されました');
    } else {
      APIErrorHandler.handleError(error, 'System', 'Full Workflow');
      UIManager.showError(error, '全自動実行中にエラーが発生しました');
    }
  }
}

// より安全な個別API実行例
async function safeOpenAICall(prompt, rateLimitManager) {
  try {
    await rateLimitManager.waitForOpenAI();
    
    const response = await UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: getOpenAIHeaders(),
      payload: JSON.stringify({
        model: 'gpt-3.5-turbo',
        messages: prompt,
        max_tokens: 1000
      })
    });
    
    const data = JSON.parse(response.getContentText());
    
    // 実際の使用トークン数を記録
    if (data.usage && data.usage.total_tokens) {
      rateLimitManager.updateOpenAITokens(data.usage.total_tokens);
    }
    
    return data;
  } catch (error) {
    Logger.log(`OpenAI API呼び出しエラー: ${error}`);
    throw error;
  }
}

async function safeGoogleSearchCall(query, rateLimitManager) {
  try {
    await rateLimitManager.waitForGoogleSearch();
    
    const searchParams = buildSearchQuery(query, getControlPanelSettings());
    const response = await executeGoogleSearch(searchParams);
    
    return response;
  } catch (error) {
    Logger.log(`Google Search API呼び出しエラー: ${error}`);
    throw error;
  }
}
```

---

## 💰 課金プラン・利用制限体系

### 一般ユーザー課金プラン

#### プラン一覧
| プラン名 | 月額料金 | キーワード生成 | 企業検索上限 | AI提案生成 | アカウント数 | 主な制限 |
|---------|---------|-------------|------------|-----------|------------|----------|
| **ベーシック** | ¥500 | ❌ 利用不可 | 10社/日 | ❌ 利用不可 | 1アカウント | 手動入力・基本テンプレートのみ |
| **スタンダード** | ¥1,500 | ✅ 利用可能 | 50社/日 | ✅ 利用可能 | 1アカウント | 全機能利用可能 |
| **プロフェッショナル** | ¥5,500 | ✅ 利用可能 | 100社/日 | ✅ 利用可能 | 2アカウント | 高速処理・優先サポート |
| **エンタープライズ** | ¥17,500 | ✅ 利用可能 | 500社/日 | ✅ 利用可能 | 5アカウント | 最大性能・専任サポート |

#### プラン詳細仕様

⚠️ **重要**: 以下のプラン料金に加えて、**各API利用料金が別途発生します**

##### 📊 API料金体系（別途課金）

| API名 | 料金体系 | 無料枠 | 超過時料金 | 月額目安 |
|-------|---------|-------|-----------|----------|
| **OpenAI ChatGPT API** | トークン従量課金 | $5.00相当/月 | $0.002/1Kトークン | ¥500-3,000 |
| **Google Custom Search API** | リクエスト従量課金 | 100回/日 | $5.00/1000回 | ¥0-1,000 |

##### 💰 プラン別API料金目安

| プラン | システム利用料 | API料金（最大想定） | **合計月額目安** |
|--------|---------------|------------------|-----------------|
| **ベーシック** | ¥500 | ¥0 | **¥500** |
| **スタンダード** | ¥1,500 | ¥2,000 | **¥3,500** |
| **プロフェッショナル** | ¥5,500 | ¥6,000 | **¥11,500** |
| **エンタープライズ** | ¥17,500 | ¥30,000 | **¥47,500** |

*API料金はGoogleとOpenAIに直接支払い（当社には支払われません）。料金は最大想定値です。*

##### 🥉 ベーシックプラン（¥500/月）
```javascript
const BASIC_PLAN_LIMITS = {
  monthlyFee: 500,
  keywordGeneration: false,        // キーワード生成機能無効
  proposalGeneration: false,       // AI提案生成機能無効
  maxCompaniesPerDay: 10,          // 1日10社まで（手動入力のみ）
  maxProposalsPerDay: 0,           // AI提案生成不可
  features: {
    manualCompanyEntry: true,      // 手動企業入力のみ
    basicTemplateProposal: true,   // 基本テンプレート提案のみ
    basicReports: true,            // 基本レポート
    emailSupport: true             // メールサポート
  },
  restrictions: {
    chatgptApiAccess: false,       // ChatGPT API利用不可
    aiProposalGeneration: false,   // AI提案生成不可
    googleSearchApi: true,         // Google検索は制限付き（手動企業確認用）
    advancedAnalytics: false       // 高度な分析機能無効
  }
};
```

##### 🥈 スタンダードプラン（¥1,500/月）
```javascript
const STANDARD_PLAN_LIMITS = {
  monthlyFee: 1500,
  keywordGeneration: true,         // キーワード生成機能有効
  maxCompaniesPerDay: 50,          // 1日50社まで
  maxProposalsPerDay: 50,          // 1日50提案まで
  maxKeywordsPerGeneration: 20,    // 1回20キーワードまで
  features: {
    fullWorkflow: true,            // 全自動ワークフロー
    advancedProposals: true,       // 高度な提案生成
    comprehensiveReports: true,    // 総合レポート
    prioritySupport: true          // 優先サポート
  },
  restrictions: {
    chatgptApiAccess: true,        // ChatGPT API利用可能
    googleSearchApi: true,         // Google検索フル利用
    advancedAnalytics: true        // 高度な分析機能有効
  }
};
```

##### 🥇 プロフェッショナルプラン（¥5,500/月）
```javascript
const PROFESSIONAL_PLAN_LIMITS = {
  monthlyFee: 5500,
  keywordGeneration: true,         // キーワード生成機能有効
  maxCompaniesPerDay: 100,         // 1日100社まで
  maxProposalsPerDay: 100,         // 1日100提案まで
  maxKeywordsPerGeneration: 40,    // 1回40キーワードまで
  maxAccounts: 2,                  // 2アカウントまで
  features: {
    batchProcessing: true,         // バッチ処理
    customAnalytics: true,         // カスタム分析
    exportFeatures: true,          // データエクスポート
    phoneSupport: true,            // 電話サポート
    fastProcessing: true,          // 高速処理
    multiAccountAccess: true       // マルチアカウントアクセス
  },
  apiLimits: {
    chatgptRequestsPerDay: 200,    // ChatGPT API 1日200回（アカウント合計）
    googleSearchPerDay: 300        // Google検索 1日300回（アカウント合計）
  }
};
```

##### 💎 エンタープライズプラン（¥17,500/月）
```javascript
const ENTERPRISE_PLAN_LIMITS = {
  monthlyFee: 17500,
  keywordGeneration: true,         // キーワード生成機能有効
  maxCompaniesPerDay: 500,         // 1日500社まで
  maxProposalsPerDay: 500,         // 1日500提案まで
  maxKeywordsPerGeneration: 100,   // 1回100キーワードまで
  maxAccounts: 5,                  // 5アカウントまで
  features: {
    unlimitedWorkflows: true,      // 無制限ワークフロー
    aiOptimization: true,          // AI最適化
    multiUserAccess: true,         // マルチユーザー（5アカウント）
    dedicatedSupport: true,        // 専任サポート
    customIntegrations: true,      // カスタム連携
    advancedSecurity: true,        // 高度なセキュリティ
    enterpriseFeatures: true       // エンタープライズ機能
  },
  apiLimits: {
    chatgptRequestsPerDay: 1000,   // ChatGPT API 1日1000回（アカウント合計）
    googleSearchPerDay: 1500       // Google検索 1日1500回（アカウント合計）
  }
};
```

### 制限管理システム

#### 日次制限チェック
```javascript
function checkDailyLimits(userPlan, requestType) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const properties = PropertiesService.getUserProperties();
  
  // 今日の使用量取得
  const todayUsage = JSON.parse(properties.getProperty(`usage_${today}`) || '{}');
  
  const planLimits = getPlanLimits(userPlan);
  
  switch (requestType) {
    case 'company_search':
      if ((todayUsage.companies || 0) >= planLimits.maxCompaniesPerDay) {
        throw new Error(`${userPlan}プランの1日企業検索上限（${planLimits.maxCompaniesPerDay}社）に達しました`);
      }
      break;
      
    case 'keyword_generation':
      if (!planLimits.keywordGeneration) {
        throw new Error(`${userPlan}プランではキーワード生成機能は利用できません`);
      }
      if ((todayUsage.keywords || 0) >= planLimits.maxKeywordsPerGeneration) {
        throw new Error(`${userPlan}プランの1日キーワード生成上限に達しました`);
      }
      break;
      
    case 'proposal_generation':
      if (!planLimits.proposalGeneration) {
        throw new Error(`${userPlan}プランではAI提案生成機能は利用できません`);
      }
      if ((todayUsage.proposals || 0) >= planLimits.maxProposalsPerDay) {
        throw new Error(`${userPlan}プランの1日提案生成上限（${planLimits.maxProposalsPerDay}件）に達しました`);
      }
      break;
  }
}
```

#### 使用量カウント
```javascript
function incrementUsage(requestType) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const properties = PropertiesService.getUserProperties();
  
  const todayUsage = JSON.parse(properties.getProperty(`usage_${today}`) || '{}');
  
  switch (requestType) {
    case 'company_search':
      todayUsage.companies = (todayUsage.companies || 0) + 1;
      break;
    case 'keyword_generation':
      todayUsage.keywords = (todayUsage.keywords || 0) + 1;
      break;
    case 'proposal_generation':
      todayUsage.proposals = (todayUsage.proposals || 0) + 1;
      break;
  }
  
  properties.setProperty(`usage_${today}`, JSON.stringify(todayUsage));
}
```

#### プラン取得・管理
```javascript
function getUserPlan() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('user_plan') || 'BASIC';
}

function getPlanLimits(planName) {
  const plans = {
    'BASIC': BASIC_PLAN_LIMITS,
    'STANDARD': STANDARD_PLAN_LIMITS,
    'PROFESSIONAL': PROFESSIONAL_PLAN_LIMITS,
    'ENTERPRISE': ENTERPRISE_PLAN_LIMITS
  };
  
  return plans[planName] || plans['BASIC'];
}

function upgradePlan(newPlan) {
  const validPlans = ['BASIC', 'STANDARD', 'PROFESSIONAL', 'ENTERPRISE'];
  
  if (!validPlans.includes(newPlan)) {
    throw new Error('無効なプラン名です');
  }
  
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('user_plan', newPlan);
  properties.setProperty('plan_upgrade_date', new Date().toISOString());
  
  // 使用量リセット（プランアップグレード時）
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  properties.deleteProperty(`usage_${today}`);
}
```

### エラーメッセージ・UI

#### 制限到達時の表示
```javascript
function showPlanLimitError(limitType, currentPlan) {
  const ui = SpreadsheetApp.getUi();
  
  let message = '🚫 プラン制限に達しました\n\n';
  
  switch (limitType) {
    case 'keyword_generation_disabled':
      message += '❌ キーワード生成機能は利用できません\n';
      message += `現在のプラン: ${currentPlan}\n\n`;
      message += '💡 スタンダードプラン（¥1,500/月）以上でご利用いただけます\n';
      message += '📈 プランアップグレードをご検討ください';
      break;
      
    case 'daily_company_limit':
      const limits = getPlanLimits(currentPlan);
      message += `📊 1日の企業検索上限（${limits.maxCompaniesPerDay}社）に達しました\n`;
      message += `現在のプラン: ${currentPlan}\n\n`;
      message += '⏰ 明日の00:00にリセットされます\n';
      message += '📈 より多くの検索をご希望の場合はプランアップグレードをご検討ください';
      break;
  }
  
  ui.alert('プラン制限', message, ui.ButtonSet.OK);
}
```

---

## 💰 API利用料金詳細説明

### 🚨 重要な注意事項

当システムの利用には、システム利用料金（月額¥500-¥17,500）に加えて、**各APIプロバイダーへの利用料金が別途発生**します。

### 📊 API料金体系

#### OpenAI ChatGPT API
```
基本情報:
- プロバイダー: OpenAI
- 課金方式: トークン従量課金
- 無料枠: $5.00相当/月（新規アカウント）
- 料金: $0.002 per 1K tokens (GPT-3.5-turbo)
- 公式サイト: https://openai.com/pricing
```

**使用例と料金目安:**
- キーワード生成1回: 約500-800トークン（¥1.4-2.2）
- AI提案生成1回: 約800-1,200トークン（¥2.2-3.3）
- 月50回AI機能利用: 約¥150-300
- 月200回AI機能利用: 約¥600-1,200

#### Google Custom Search API
```
基本情報:
- プロバイダー: Google
- 課金方式: リクエスト従量課金
- 無料枠: 100回/日（約3,000回/月）
- 料金: $5.00 per 1,000 queries
- 公式サイト: https://developers.google.com/custom-search/v1/overview
```

**使用例と料金目安:**
- 企業検索1回: 1リクエスト
- 月100回検索（無料枠内）: ¥0
- 月3,000回検索（無料枠上限）: ¥0
- 月5,000回検索: 約¥1,100（超過分2,000回）

### 💳 プラン別総額料金シミュレーション

#### ベーシックプラン利用者
```
システム利用料: ¥500/月
OpenAI API: ¥0（AI機能未使用）
Google Search API: ¥0（無料枠内想定）
---
月額合計: ¥500
```

#### スタンダードプラン利用者（軽量使用）
```
システム利用料: ¥1,500/月
OpenAI API: ¥300（月30回AI利用）
Google Search API: ¥0（無料枠内）
---
月額合計: ¥1,800
```

#### スタンダードプラン利用者（標準使用）
```
システム利用料: ¥1,500/月
OpenAI API: ¥800（月80回AI利用）
Google Search API: ¥500（月4,000回検索）
---
月額合計: ¥2,800
```

#### プロフェッショナルプラン利用者（2アカウント）
```
システム利用料: ¥5,500/月
OpenAI API: ¥3,500（月200回AI利用×2アカウント）
Google Search API: ¥2,500（月3,000回検索×2アカウント）
---
月額合計: ¥11,500
```

#### エンタープライズプラン利用者（5アカウント）
```
システム利用料: ¥17,500/月
OpenAI API: ¥17,500（月1,000回AI利用×5アカウント）
Google Search API: ¥12,500（月15,000回検索×5アカウント）
---
月額合計: ¥47,500
```

### ⚙️ API料金管理機能

#### 使用量監視
```javascript
function trackAPIUsage() {
  // OpenAI使用量追跡
  const openaiUsage = getOpenAIUsageToday();
  const estimatedCost = (openaiUsage.tokens / 1000) * 0.002 * 140; // USD to JPY
  
  // Google Search使用量追跡
  const searchUsage = getGoogleSearchUsageToday();
  const searchCost = Math.max(0, (searchUsage - 100) / 1000 * 5 * 140);
  
  Logger.log(`今日のAPI料金: OpenAI=¥${estimatedCost}, Google=¥${searchCost}`);
}
```

#### 料金アラート機能
```javascript
function checkAPIBudget() {
  const monthlyBudget = getUserAPIBudget(); // ユーザー設定予算
  const currentUsage = getCurrentMonthAPIUsage();
  
  if (currentUsage > monthlyBudget * 0.8) {
    showBudgetAlert('API料金が予算の80%に達しました');
  }
  
  if (currentUsage > monthlyBudget) {
    showBudgetAlert('API料金が予算を超過しました', true);
  }
}
```

### � プラン別API使用量詳細（最大想定）

#### ベーシックプラン（1アカウント）
- Google検索: 10回/日（無料枠内）
- OpenAI API: 利用なし
- **月額API料金: ¥0**

#### スタンダードプラン（1アカウント）
- Google検索: 50回/日（月1,500回想定）
- OpenAI API: キーワード生成+提案生成（月50回想定）
- **月額API料金: 最大¥2,000**

#### プロフェッショナルプラン（2アカウント）
- Google検索: 100回/日×2（月3,000回想定）
- OpenAI API: キーワード生成+提案生成（月200回想定）
- **月額API料金: 最大¥6,000**

#### エンタープライズプラン（5アカウント）
- Google検索: 500回/日×5（月15,000回想定）
- OpenAI API: キーワード生成+提案生成（月1,000回想定）
- **月額API料金: 最大¥30,000**

### �💡 API料金節約のベストプラクティス

1. **ベーシックプランの活用**
   - AI機能が不要な場合はベーシックプランでAPI料金¥0

2. **効率的なキーワード設計**
   - 1回のキーワード生成で多くの検索語を取得
   - 重複検索を避ける

3. **検索回数の最適化**
   - 1日100回の無料枠を効率的に活用
   - 質の高いキーワードに絞って検索

4. **バッチ処理の活用**
   - 複数企業の提案を一度に生成
   - API呼び出し回数を削減

5. **アカウント間での効率的な使用量分散**
   - マルチアカウントプランでは使用量を分散
   - チーム全体での計画的な利用

6. **定期的な使用量確認**
   - 月末の予期しない高額請求を防止
   - 予算管理機能の活用

### 🔒 API料金に関する免責事項

- API料金は各プロバイダーの料金体系変更により変動する可能性があります
- 当システムの料金予測は目安であり、実際の請求額と異なる場合があります
- API料金の支払いは各プロバイダーとの直接契約となります
- 当社システム利用料金とAPI利用料金は別途管理・請求されます

---

## 🔄 マルチアカウント運用システム

### 🚨 なぜ複数アカウントが必要なのか

#### 1日の作業量制限の現実
```
【1アカウントでの1日限界】
- Google Search API無料枠: 100回/日
- OpenAI APIレート制限: 3リクエスト/分（最大4,320回/日理論値）
- 実際の処理時間: レート制限により大幅に時間がかかる

【具体例：スタンダードプラン（50社/日）の場合】
- キーワード生成: 1-2回（20-40キーワード）
- 企業検索: 20-40回（Google Search API使用）
- AI提案生成: 50回（OpenAI API使用、約17分必要）
- 合計処理時間: 約30-60分

【問題点】
✅ 処理自体は可能だが、時間がかかりすぎる
❌ 大量処理時にボトルネックとなる
❌ ビジネス効率が低下する
```

#### プロフェッショナル・エンタープライズプランでの必要性
```
【プロフェッショナルプラン：100社/日】
- 1アカウントでは2-3時間必要
- 2アカウント並行なら1-1.5時間に短縮

【エンタープライズプラン：500社/日】
- 1アカウントでは10-15時間必要（現実的でない）
- 5アカウント並行なら2-3時間に短縮
```

### 🏗️ マルチアカウント実装アーキテクチャ

#### システム構成
```javascript
/**
 * マルチアカウント管理システム
 */
class MultiAccountManager {
  constructor() {
    this.accounts = [];
    this.currentAccountIndex = 0;
    this.accountUsage = new Map();
    this.rateLimiters = new Map();
  }
  
  /**
   * アカウント初期化
   */
  initializeAccounts(accountConfigs) {
    this.accounts = accountConfigs.map((config, index) => ({
      id: config.id || `account_${index + 1}`,
      email: config.email,
      apiKeys: {
        openai: config.openaiKey,
        googleSearch: config.googleSearchKey,
        googleSearchEngineId: config.googleSearchEngineId
      },
      spreadsheetId: config.spreadsheetId,
      status: 'ready',
      lastUsed: null
    }));
    
    // 各アカウントのレート制限管理を初期化
    this.accounts.forEach(account => {
      this.rateLimiters.set(account.id, new APIRateLimitManager());
      this.accountUsage.set(account.id, {
        dailySearchCount: 0,
        dailyOpenAICount: 0,
        lastResetDate: new Date().toDateString()
      });
    });
  }
  
  /**
   * 最適なアカウント選択
   */
  selectOptimalAccount(taskType) {
    const availableAccounts = this.accounts.filter(account => {
      const usage = this.accountUsage.get(account.id);
      
      // 日付変更チェック
      if (usage.lastResetDate !== new Date().toDateString()) {
        usage.dailySearchCount = 0;
        usage.dailyOpenAICount = 0;
        usage.lastResetDate = new Date().toDateString();
      }
      
      // タスクタイプ別の利用可能性チェック
      switch (taskType) {
        case 'search':
          return usage.dailySearchCount < 90; // 安全マージン10件
        case 'openai':
          return usage.dailyOpenAICount < 200; // 1日の推奨上限
        default:
          return true;
      }
    });
    
    if (availableAccounts.length === 0) {
      throw new Error(`${taskType}タスクを実行可能なアカウントがありません`);
    }
    
    // 使用量の少ないアカウントを優先選択
    return availableAccounts.sort((a, b) => {
      const usageA = this.accountUsage.get(a.id);
      const usageB = this.accountUsage.get(b.id);
      const totalUsageA = usageA.dailySearchCount + usageA.dailyOpenAICount;
      const totalUsageB = usageB.dailySearchCount + usageB.dailyOpenAICount;
      return totalUsageA - totalUsageB;
    })[0];
  }
  
  /**
   * 使用量更新
   */
  updateUsage(accountId, taskType) {
    const usage = this.accountUsage.get(accountId);
    switch (taskType) {
      case 'search':
        usage.dailySearchCount++;
        break;
      case 'openai':
        usage.dailyOpenAICount++;
        break;
    }
  }
  
  /**
   * アカウント状態取得
   */
  getAccountStatus() {
    return this.accounts.map(account => ({
      id: account.id,
      email: account.email,
      usage: this.accountUsage.get(account.id),
      status: account.status
    }));
  }
}
```

### 🔧 実装方法

#### 方法1: 複数Google Apps Scriptプロジェクト
```javascript
/**
 * メインコントローラー（マスターアカウント）
 */
class MasterAccountController {
  constructor() {
    this.workerAccounts = [
      {
        id: 'worker_1',
        scriptId: 'YOUR_WORKER_SCRIPT_ID_1',
        email: 'worker1@yourcompany.com'
      },
      {
        id: 'worker_2', 
        scriptId: 'YOUR_WORKER_SCRIPT_ID_2',
        email: 'worker2@yourcompany.com'
      }
      // エンタープライズプランなら最大5アカウント
    ];
  }
  
  /**
   * 並行処理実行
   */
  async executeParallelProcessing(tasks) {
    const chunks = this.distributeTask(tasks, this.workerAccounts.length);
    const promises = [];
    
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const worker = this.workerAccounts[i];
      
      promises.push(this.executeWorkerTask(worker, chunk));
    }
    
    try {
      const results = await Promise.all(promises);
      return this.mergeResults(results);
    } catch (error) {
      Logger.log(`並行処理エラー: ${error}`);
      throw error;
    }
  }
  
  /**
   * ワーカーアカウントでタスク実行
   */
  async executeWorkerTask(worker, taskChunk) {
    // Google Apps Script Execution APIを使用
    const response = await UrlFetchApp.fetch(
      `https://script.googleapis.com/v1/scripts/${worker.scriptId}:run`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.getAccessToken()}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          function: 'processTaskChunk',
          parameters: [taskChunk]
        })
      }
    );
    
    return JSON.parse(response.getContentText());
  }
  
  /**
   * タスク分散
   */
  distributeTask(tasks, workerCount) {
    const chunks = [];
    const chunkSize = Math.ceil(tasks.length / workerCount);
    
    for (let i = 0; i < tasks.length; i += chunkSize) {
      chunks.push(tasks.slice(i, i + chunkSize));
    }
    
    return chunks;
  }
}
```

#### 方法2: スプレッドシート共有による協調動作
```javascript
/**
 * 共有スプレッドシート経由での協調システム
 */
class SharedSpreadsheetCoordinator {
  constructor(masterSpreadsheetId) {
    this.masterSpreadsheetId = masterSpreadsheetId;
    this.taskQueueSheet = 'タスクキュー';
    this.resultSheet = '処理結果';
  }
  
  /**
   * タスクキューに追加
   */
  addTasksToQueue(tasks) {
    const sheet = SpreadsheetApp.openById(this.masterSpreadsheetId)
      .getSheetByName(this.taskQueueSheet);
    
    const taskData = tasks.map(task => [
      task.id,
      task.type,
      JSON.stringify(task.data),
      '待機中',
      new Date(),
      '', // 担当アカウントID
      ''  // 完了時刻
    ]);
    
    sheet.getRange(sheet.getLastRow() + 1, 1, taskData.length, 7)
      .setValues(taskData);
  }
  
  /**
   * タスク取得（各ワーカーが実行）
   */
  claimTask(workerAccountId) {
    const sheet = SpreadsheetApp.openById(this.masterSpreadsheetId)
      .getSheetByName(this.taskQueueSheet);
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === '待機中') { // ステータス列
        // タスクを自分に割り当て
        sheet.getRange(i + 1, 4).setValue('処理中');
        sheet.getRange(i + 1, 6).setValue(workerAccountId);
        sheet.getRange(i + 1, 7).setValue(new Date());
        
        return {
          row: i + 1,
          id: data[i][0],
          type: data[i][1],
          data: JSON.parse(data[i][2])
        };
      }
    }
    
    return null; // 利用可能なタスクなし
  }
  
  /**
   * タスク完了報告
   */
  completeTask(taskRow, result) {
    const sheet = SpreadsheetApp.openById(this.masterSpreadsheetId)
      .getSheetByName(this.taskQueueSheet);
    
    sheet.getRange(taskRow, 4).setValue('完了');
    
    // 結果を別シートに保存
    const resultSheet = SpreadsheetApp.openById(this.masterSpreadsheetId)
      .getSheetByName(this.resultSheet);
    
    resultSheet.appendRow([
      new Date(),
      result.taskId,
      result.type,
      JSON.stringify(result.data)
    ]);
  }
}
```

### ⚙️ 運用設定例

#### プロフェッショナルプラン（2アカウント）設定
```javascript
const PROFESSIONAL_MULTI_ACCOUNT_CONFIG = {
  accounts: [
    {
      id: 'primary_account',
      email: 'primary@company.com',
      role: 'coordinator', // タスク分散・結果集約
      apiKeys: {
        openai: 'OPENAI_KEY_1',
        googleSearch: 'GOOGLE_SEARCH_KEY_1',
        googleSearchEngineId: 'SEARCH_ENGINE_ID_1'
      },
      spreadsheetId: 'MASTER_SPREADSHEET_ID'
    },
    {
      id: 'worker_account',
      email: 'worker@company.com', 
      role: 'worker', // 実作業専用
      apiKeys: {
        openai: 'OPENAI_KEY_2',
        googleSearch: 'GOOGLE_SEARCH_KEY_2',
        googleSearchEngineId: 'SEARCH_ENGINE_ID_2'
      },
      spreadsheetId: 'WORKER_SPREADSHEET_ID'
    }
  ],
  workload_distribution: {
    daily_target: 100, // 企業検索数
    per_account_limit: 50,
    parallel_processing: true,
    coordination_method: 'shared_spreadsheet'
  }
};
```

#### エンタープライズプラン（5アカウント）設定
```javascript
const ENTERPRISE_MULTI_ACCOUNT_CONFIG = {
  accounts: [
    // マスターアカウント（1つ）
    {
      id: 'master_account',
      role: 'master',
      responsibilities: ['task_distribution', 'result_aggregation', 'monitoring']
    },
    // ワーカーアカウント（4つ）
    ...Array.from({length: 4}, (_, i) => ({
      id: `worker_${i + 1}`,
      role: 'worker',
      responsibilities: ['task_execution'],
      daily_quota: {
        searches: 100,
        openai_requests: 200
      }
    }))
  ],
  processing_strategy: {
    load_balancing: 'round_robin',
    failover: 'automatic',
    monitoring_interval: 300000 // 5分間隔
  }
};
```

### 📊 パフォーマンス比較

#### 処理時間比較表
```
【100社の企業検索・提案生成の場合】

単一アカウント:
- 企業検索: 100回 × 1.5秒間隔 = 2.5分
- AI提案生成: 100回 × 20秒間隔 = 33.3分
- 合計: 約36分

2アカウント並行:
- 企業検索: 50回ずつ並行 = 1.25分
- AI提案生成: 50回ずつ並行 = 16.7分
- 合計: 約18分（50%短縮）

5アカウント並行:
- 企業検索: 20回ずつ並行 = 0.5分
- AI提案生成: 20回ずつ並行 = 6.7分
- 合計: 約7.2分（80%短縮）
```

### 🛠️ セットアップ手順

#### 1. 複数Googleアカウント準備
```
1. Googleアカウント作成（プラン数分）
2. 各アカウントでGoogle Apps Script有効化
3. 各アカウントでAPI設定
   - OpenAI APIキー取得・設定
   - Google Custom Search API有効化
   - 検索エンジンID取得
```

#### 2. スプレッドシート共有設定
```
1. マスタースプレッドシート作成
2. 全アカウントに編集権限付与
3. タスクキュー・結果集約シート作成
4. 各ワーカーアカウントに個別スプレッドシート作成
```

#### 3. スクリプト配布・設定
```javascript
/**
 * 各アカウントでの初期設定
 */
function setupWorkerAccount() {
  // 1. API キー設定
  PropertiesService.getScriptProperties().setProperties({
    'OPENAI_API_KEY': 'YOUR_OPENAI_KEY',
    'GOOGLE_SEARCH_API_KEY': 'YOUR_GOOGLE_SEARCH_KEY',
    'GOOGLE_SEARCH_ENGINE_ID': 'YOUR_SEARCH_ENGINE_ID',
    'MASTER_SPREADSHEET_ID': 'SHARED_SPREADSHEET_ID',
    'WORKER_ACCOUNT_ID': 'UNIQUE_WORKER_ID'
  });
  
  // 2. トリガー設定（定期実行）
  ScriptApp.newTrigger('processTaskQueue')
    .timeBased()
    .everyMinutes(5)
    .create();
}
```

### 🚦 制限事項・注意点

#### Google Apps Scriptの制限
```
- 実行時間制限: 6分/実行
- 日次実行制限: 6時間/日
- 同時実行制限: アカウント別に管理
- API呼び出し制限: 各APIの制限に従う
```

#### 運用上の注意点
```
⚠️ アカウント管理
- 各アカウントの使用量監視が必要
- APIキーの適切な管理・ローテーション
- アカウント停止リスクの分散

⚠️ データ整合性
- 重複処理の防止機構が必要
- 結果マージ時の整合性チェック
- エラー処理・リトライ機構の実装
```

### 💡 推奨運用パターン

#### 小規模運用（プロフェッショナルプラン）
- メインアカウント: UI操作・結果確認
- ワーカーアカウント: バックグラウンド処理
- 協調方法: 共有スプレッドシート

#### 大規模運用（エンタープライズプラン）  
- マスターアカウント: 統括管理・監視
- ワーカーアカウント群: 専用処理
- 協調方法: Google Apps Script Execution API
- 監視: リアルタイムダッシュボード

---

*最終更新: 2025年10月17日*  
*API仕様書 v2.0*
