# 🔌 API仕様書

**営業自動化システム**  
**バージョン**: 1.0.0  
**作成日**: 2025年10月17日  

---

## 📑 目次

1. [概要](#概要)
2. [OpenAI ChatGPT API](#openai-chatgpt-api)
3. [Google Custom Search API](#google-custom-search-api)
4. [エラーハンドリング](#エラーハンドリング)
5. [レート制限対応](#レート制限対応)
6. [セキュリティ](#セキュリティ)
7. [テスト仕様](#テスト仕様)

---

## 🎯 概要

営業自動化システムが使用する外部APIの詳細仕様を定義します。主に以下の2つのAPIを使用してシステムの核となる機能を提供します。

### 使用API一覧
1. **OpenAI ChatGPT API** - テキスト生成・分析
2. **Google Custom Search API** - ウェブ検索

---

## 🤖 OpenAI ChatGPT API

### 基本情報
- **エンドポイント**: `https://api.openai.com/v1/chat/completions`
- **認証方式**: Bearer Token
- **使用モデル**: `gpt-3.5-turbo`
- **リクエスト形式**: JSON (POST)

### 認証
```javascript
const headers = {
  'Authorization': `Bearer ${OPENAI_API_KEY}`,
  'Content-Type': 'application/json'
};
```

### 共通リクエストパラメータ
```javascript
{
  "model": "gpt-3.5-turbo",
  "messages": [
    {
      "role": "system", 
      "content": "システムプロンプト"
    },
    {
      "role": "user",
      "content": "ユーザープロンプト"
    }
  ],
  "max_tokens": 1500,
  "temperature": 0.7,
  "top_p": 1,
  "frequency_penalty": 0,
  "presence_penalty": 0
}
```

### 使用例1: キーワード生成

#### リクエスト
```javascript
{
  "model": "gpt-3.5-turbo",
  "messages": [
    {
      "role": "system",
      "content": "あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。"
    },
    {
      "role": "user", 
      "content": "商材名: 営業支援ツール\n概要: 営業プロセスを効率化するSaaS\n価格帯: 中価格\n対象: 中小企業\n\n以下の4カテゴリで検索キーワードを生成してください：\n1. painPointHunting: 課題企業発見用\n2. growthTargeting: 成長企業発見用\n3. budgetTargeting: 予算企業発見用\n4. timingCapture: 導入タイミング企業発見用\n\nJSON形式で回答してください。"
    }
  ],
  "max_tokens": 2000,
  "temperature": 0.7
}
```

#### レスポンス
```javascript
{
  "id": "chatcmpl-xxx",
  "object": "chat.completion",
  "created": 1699123456,
  "model": "gpt-3.5-turbo",
  "choices": [
    {
      "index": 0,
      "message": {
        "role": "assistant",
        "content": "{\n  \"strategicKeywords\": {\n    \"painPointHunting\": {\n      \"keywords\": [\"営業効率化 課題\", \"売上向上 悩み\", \"顧客管理 問題\"],\n      \"strategy\": \"営業プロセスに課題を抱える企業を発見\"\n    },\n    \"growthTargeting\": {\n      \"keywords\": [\"急成長 企業\", \"売上拡大 中小企業\", \"事業拡張\"],\n      \"strategy\": \"成長段階で営業強化が必要な企業をターゲット\"\n    }\n  }\n}"
      },
      "finish_reason": "stop"
    }
  ],
  "usage": {
    "prompt_tokens": 180,
    "completion_tokens": 250,
    "total_tokens": 430
  }
}
```

### 使用例2: 企業情報分析

#### リクエスト
```javascript
{
  "model": "gpt-3.5-turbo",
  "messages": [
    {
      "role": "system",
      "content": "あなたは企業情報分析の専門家です。Webサイトから正確な企業情報を抽出してください。"
    },
    {
      "role": "user",
      "content": "WebサイトHTML：[HTMLコンテンツ]\nURL：https://example.com\n\n以下のJSON形式で企業情報を抽出してください：\n{\n  \"companyName\": \"会社名\",\n  \"industry\": \"業界\",\n  \"employees\": \"従業員数\",\n  \"location\": \"所在地\",\n  \"hasContactForm\": true/false,\n  \"businessDescription\": \"事業内容\",\n  \"isValidCompany\": true/false\n}"
    }
  ],
  "max_tokens": 800,
  "temperature": 0.3
}
```

### 使用例3: 提案メッセージ生成

#### リクエスト
```javascript
{
  "model": "gpt-3.5-turbo", 
  "messages": [
    {
      "role": "system",
      "content": "あなたは営業のプロフェッショナルです。企業情報と商材情報を基に、効果的で個別最適化された提案メッセージを作成してください。"
    },
    {
      "role": "user",
      "content": "【商材情報】\n商材名: 営業支援ツール\n概要: SFA機能付きCRM\n価格帯: 中価格\n\n【企業情報】\n会社名: ABC株式会社\n業界: IT\n規模: 中小企業\n事業内容: Webシステム開発\n\n課題訴求型(A)と成功事例型(B)の2パターンの提案を作成し、JSON形式で回答してください。"
    }
  ],
  "max_tokens": 1500,
  "temperature": 0.7
}
```

### トークン管理
| 機能 | 推定トークン数 | max_tokens設定 |
|------|---------------|----------------|
| キーワード生成 | 1500-2000 | 2000 |
| 企業情報分析 | 600-800 | 800 |
| 提案メッセージ生成 | 1200-1500 | 1500 |

### エラーレスポンス
```javascript
{
  "error": {
    "message": "Rate limit reached for requests",
    "type": "rate_limit_error",
    "param": null,
    "code": "rate_limit_exceeded"
  }
}
```

---

## 🔍 Google Custom Search API

### 基本情報
- **エンドポイント**: `https://www.googleapis.com/customsearch/v1`
- **認証方式**: API Key
- **リクエスト形式**: GET
- **制限**: 100回/日（無料版）

### 認証
```javascript
const url = `https://www.googleapis.com/customsearch/v1?key=${GOOGLE_SEARCH_API_KEY}&cx=${SEARCH_ENGINE_ID}&q=${query}`;
```

### リクエストパラメータ
| パラメータ | 必須 | 説明 | 例 |
|-----------|------|------|-----|
| `key` | ○ | Google API Key | `AIza...` |
| `cx` | ○ | 検索エンジンID | `017576662512468239146:omuauf_lfve` |
| `q` | ○ | 検索クエリ | `営業効率化 課題 site:co.jp` |
| `num` | △ | 結果数（1-10） | `3` |
| `start` | △ | 開始位置 | `1` |
| `lr` | △ | 言語制限 | `lang_ja` |
| `safe` | △ | セーフサーチ | `active` |

### 使用例: 企業検索

#### リクエストURL
```
https://www.googleapis.com/customsearch/v1?key=YOUR_API_KEY&cx=YOUR_SEARCH_ENGINE_ID&q=営業効率化%20課題%20site:co.jp%20OR%20site:com&num=3&lr=lang_ja
```

#### レスポンス
```javascript
{
  "kind": "customsearch#search",
  "url": {
    "type": "application/json",
    "template": "https://www.googleapis.com/customsearch/v1?q={searchTerms}&num={count?}&start={startIndex?}&lr={language?}&safe={safe?}&cx={cx?}&sort={sort?}&filter={filter?}&gl={gl?}&cr={cr?}&googlehost={googleHost?}&c2coff={disableCnTwTranslation?}&hq={hq?}&hl={hl?}&siteSearch={siteSearch?}&siteSearchFilter={siteSearchFilter?}&exactTerms={exactTerms?}&excludeTerms={excludeTerms?}&linkSite={linkSite?}&orTerms={orTerms?}&relatedSite={relatedSite?}&dateRestrict={dateRestrict?}&lowRange={lowRange?}&highRange={highRange?}&searchType={searchType}&fileType={fileType?}&rights={rights?}&imgSize={imgSize?}&imgType={imgType?}&imgColorType={imgColorType?}&imgDominantColor={imgDominantColor?}&alt=json"
  },
  "queries": {
    "request": [
      {
        "title": "Google Custom Search - 営業効率化 課題 site:co.jp OR site:com",
        "totalResults": "1240",
        "searchTerms": "営業効率化 課題 site:co.jp OR site:com",
        "count": 3,
        "startIndex": 1,
        "inputEncoding": "utf8",
        "outputEncoding": "utf8",
        "safe": "off",
        "cx": "017576662512468239146:omuauf_lfve"
      }
    ]
  },
  "searchInformation": {
    "searchTime": 0.234567,
    "formattedSearchTime": "0.23",
    "totalResults": "1240",
    "formattedTotalResults": "1,240"
  },
  "items": [
    {
      "kind": "customsearch#result",
      "title": "営業効率化の課題と解決策 | ABC株式会社",
      "htmlTitle": "営業効率化の課題と解決策 | ABC株式会社",
      "link": "https://www.abc-corp.co.jp/solutions/sales-efficiency/",
      "displayLink": "www.abc-corp.co.jp",
      "snippet": "当社では営業プロセスの効率化に関する課題を抱えており、顧客管理システムの導入を検討しています。売上向上と業務効率化を両立するソリューションを...",
      "htmlSnippet": "当社では営業プロセスの効率化に関する課題を抱えており、顧客管理システムの導入を検討しています。売上向上と業務効率化を両立するソリューションを...",
      "formattedUrl": "https://www.abc-corp.co.jp/solutions/sales-efficiency/",
      "htmlFormattedUrl": "https://www.abc-corp.co.jp/solutions/sales-efficiency/"
    }
  ]
}
```

### 検索クエリの最適化

#### 基本検索パターン
```javascript
const buildSearchQuery = (keyword, settings) => {
  let query = keyword;
  
  // 地域指定
  if (settings.preferredRegion) {
    query += ` ${settings.preferredRegion}`;
  }
  
  // サイト絞り込み
  query += ' site:co.jp OR site:com OR site:net OR site:org';
  
  // 除外キーワード
  query += ' -求人 -採用 -リクルート';
  
  return encodeURIComponent(query);
};
```

#### 業界特化検索
```javascript
const industrySpecificQueries = {
  'IT': 'site:co.jp OR site:com (IT OR システム OR ソフトウェア)',
  '製造業': 'site:co.jp (製造 OR 工場 OR メーカー)',
  'サービス業': 'site:co.jp (サービス OR コンサル OR 支援)'
};
```

### エラーレスポンス
```javascript
{
  "error": {
    "code": 400,
    "message": "Bad Request",
    "errors": [
      {
        "message": "Bad Request",
        "domain": "global",
        "reason": "badRequest"
      }
    ]
  }
}
```

---

## 🛡️ エラーハンドリング

### エラー分類と対応

#### OpenAI API エラー
| エラーコード | 説明 | 対応方法 |
|-------------|------|----------|
| 400 | Bad Request | リクエスト内容を確認 |
| 401 | Unauthorized | APIキーを確認 |
| 429 | Rate Limit | 指数バックオフで再試行 |
| 500 | Server Error | 1分後に再試行 |

#### Google Search API エラー
| エラーコード | 説明 | 対応方法 |
|-------------|------|----------|
| 400 | Bad Request | クエリパラメータを確認 |
| 403 | Quota Exceeded | 日次制限に達した |
| 404 | Not Found | エンドポイントを確認 |

### 統一エラーハンドリング
```javascript
function handleApiError(error, apiType) {
  const errorInfo = {
    timestamp: new Date().toISOString(),
    apiType: apiType,
    statusCode: error.getResponseCode ? error.getResponseCode() : null,
    message: error.toString(),
    retryable: isRetryableError(error)
  };
  
  Logger.log(`API Error [${apiType}]:`, JSON.stringify(errorInfo));
  
  if (errorInfo.retryable) {
    return { shouldRetry: true, waitTime: calculateWaitTime(errorInfo) };
  }
  
  throw new Error(`${apiType} API Error: ${errorInfo.message}`);
}
```

---

## ⏱️ レート制限対応

### 制限値一覧
| API | 制限値 | 期間 | 課金プラン |
|-----|--------|------|-----------|
| OpenAI ChatGPT | 3 RPM | 分 | 無料 |
| OpenAI ChatGPT | 3,500 RPM | 分 | Pay-as-you-go |
| Google Search | 100回 | 日 | 無料 |
| Google Search | 10,000回 | 日 | 有料 |

### 指数バックオフ実装
```javascript
async function apiCallWithRetry(apiFunction, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await apiFunction();
    } catch (error) {
      if (isRateLimitError(error)) {
        const waitTime = Math.pow(2, i) * 1000 + Math.random() * 1000;
        Logger.log(`Rate limit hit. Waiting ${waitTime}ms before retry ${i + 1}/${maxRetries}`);
        
        if (i < maxRetries - 1) {
          Utilities.sleep(waitTime);
          continue;
        }
      }
      throw error;
    }
  }
}

function isRateLimitError(error) {
  const errorStr = error.toString().toLowerCase();
  return errorStr.includes('rate') || 
         errorStr.includes('limit') || 
         errorStr.includes('quota') ||
         (error.getResponseCode && error.getResponseCode() === 429);
}
```

### 使用量追跡
```javascript
class ApiUsageTracker {
  constructor() {
    this.usage = {
      openai: { requests: 0, tokens: 0 },
      google: { requests: 0 }
    };
  }
  
  trackOpenAI(tokens) {
    this.usage.openai.requests++;
    this.usage.openai.tokens += tokens;
  }
  
  trackGoogle() {
    this.usage.google.requests++;
  }
  
  getDailyUsage() {
    return this.usage;
  }
  
  resetDaily() {
    this.usage = {
      openai: { requests: 0, tokens: 0 },
      google: { requests: 0 }
    };
  }
}
```

---

## 🔒 セキュリティ

### APIキー管理
```javascript
// 安全な保存
function setApiKeys(openaiKey, googleKey, searchEngineId) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'OPENAI_API_KEY': openaiKey,
    'GOOGLE_SEARCH_API_KEY': googleKey,
    'GOOGLE_SEARCH_ENGINE_ID': searchEngineId
  });
}

// 安全な取得
function getApiKeys() {
  const properties = PropertiesService.getScriptProperties();
  return {
    openai: properties.getProperty('OPENAI_API_KEY'),
    google: properties.getProperty('GOOGLE_SEARCH_API_KEY'),
    searchEngineId: properties.getProperty('GOOGLE_SEARCH_ENGINE_ID')
  };
}
```

### ログセキュリティ
```javascript
function sanitizeLogData(data) {
  const sanitized = JSON.stringify(data, (key, value) => {
    // APIキーをマスク
    if (typeof key === 'string' && key.toLowerCase().includes('key')) {
      return value ? `${value.substring(0, 4)}...${value.substring(value.length - 4)}` : value;
    }
    return value;
  });
  
  return sanitized;
}
```

---

## 🧪 テスト仕様

### API接続テスト
```javascript
function testApiConnections() {
  const results = {
    openai: false,
    google: false,
    timestamp: new Date()
  };
  
  // OpenAI API テスト
  try {
    const testResponse = callOpenAIAPI({
      model: 'gpt-3.5-turbo',
      messages: [{ role: 'user', content: 'Hello' }],
      max_tokens: 10
    });
    results.openai = !!testResponse;
  } catch (error) {
    Logger.log('OpenAI API test failed:', error.toString());
  }
  
  // Google Search API テスト
  try {
    const testUrl = `https://www.googleapis.com/customsearch/v1?key=${getApiKeys().google}&cx=${getApiKeys().searchEngineId}&q=test`;
    const response = UrlFetchApp.fetch(testUrl);
    results.google = response.getResponseCode() === 200;
  } catch (error) {
    Logger.log('Google Search API test failed:', error.toString());
  }
  
  return results;
}
```

### 機能別テスト
```javascript
function testKeywordGeneration() {
  const testInput = {
    productName: "テスト商材",
    productDescription: "テスト用の商材説明",
    priceRange: "中価格",
    targetSize: "中小企業"
  };
  
  try {
    const keywords = generateStrategicKeywords(testInput);
    return {
      success: keywords.length > 0,
      keywordCount: keywords.length,
      categories: [...new Set(keywords.map(k => k.category))]
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
```

### パフォーマンステスト
```javascript
function benchmarkApiCalls() {
  const results = [];
  
  // OpenAI API ベンチマーク
  const startTime = new Date();
  try {
    callOpenAIAPI({
      model: 'gpt-3.5-turbo',
      messages: [{ role: 'user', content: 'Generate 5 keywords for sales automation' }],
      max_tokens: 100
    });
    results.push({
      api: 'OpenAI',
      duration: new Date() - startTime,
      success: true
    });
  } catch (error) {
    results.push({
      api: 'OpenAI',
      duration: new Date() - startTime,
      success: false,
      error: error.toString()
    });
  }
  
  return results;
}
```

---

**このAPI仕様書は、営業自動化システムのAPI統合に関する完全な定義です。開発・運用時の参照資料として活用してください。**
