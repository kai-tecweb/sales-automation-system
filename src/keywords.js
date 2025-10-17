/**
 * キーワード生成機能
 * 商材情報から戦略的な企業検索キーワードを自動生成
 */

/**
 * 戦略的キーワードの生成
 */
function executeKeywordGeneration() {
  try {
    // システム初期化確認
    checkSystemInitialization();
    
    updateExecutionStatus('キーワード生成を開始します...');
    
    const settings = getControlPanelSettings();
    
    // 入力値チェック
    if (!settings.productName || !settings.productDescription) {
      throw new Error('商材名と商材概要は必須です');
    }
    
    // ChatGPT APIでキーワード生成
    const keywords = generateKeywordsWithChatGPT(settings);
    
    // 結果をスプレッドシートに保存
    saveKeywordsToSheet(keywords);
    
    updateExecutionStatus(`キーワード生成完了: ${keywords.length}個のキーワードを生成しました`);
    logExecution('キーワード生成', 'SUCCESS', keywords.length, 0);
    
    return keywords;
    
  } catch (error) {
    handleSystemError('キーワード生成', error);
  }
}

/**
 * ChatGPT APIを使用してキーワードを生成
 */
function generateKeywordsWithChatGPT(settings) {
  const prompt = createKeywordGenerationPrompt(settings);
  
  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: 'あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: 2000,
    temperature: 0.7
  };
  
  const response = callOpenAIAPI(payload);
  return parseKeywordResponse(response);
}

/**
 * キーワード生成用プロンプトの作成
 */
function createKeywordGenerationPrompt(settings) {
  return `
商材情報：
- 商材名: ${settings.productName}
- 概要: ${settings.productDescription}
- 価格帯: ${settings.priceRange}
- 対象企業規模: ${settings.targetSize}
- 優先地域: ${settings.preferredRegion || '指定なし'}

あなたは営業戦略コンサルタントです。この商材が刺さりそうな企業を見つけるための検索キーワードを戦略的に生成してください。

以下の4つのカテゴリに分けて、それぞれ10-15個のキーワードを生成してください：

1. painPointHunting: 課題を抱える企業を発見するためのキーワード
2. growthTargeting: 成長している企業を発見するためのキーワード  
3. budgetTargeting: 予算のある企業を発見するためのキーワード
4. timingCapture: 導入タイミングが良い企業を発見するためのキーワード

出力は以下のJSON形式でお願いします：

{
  "strategicKeywords": {
    "painPointHunting": {
      "keywords": ["キーワード1", "キーワード2", ...],
      "strategy": "この戦略の説明"
    },
    "growthTargeting": {
      "keywords": ["キーワード1", "キーワード2", ...],
      "strategy": "この戦略の説明"
    },
    "budgetTargeting": {
      "keywords": ["キーワード1", "キーワード2", ...],
      "strategy": "この戦略の説明"
    },
    "timingCapture": {
      "keywords": ["キーワード1", "キーワード2", ...],
      "strategy": "この戦略の説明"
    }
  }
}

キーワードは具体的で検索に効果的なものにしてください。企業名ではなく、状況や課題、成長段階を表現するキーワードを重視してください。
`;
}

/**
 * OpenAI API呼び出し
 */
function callOpenAIAPI(payload) {
  const apiKey = API_KEYS.OPENAI;
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません');
  }
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  return apiCallWithRetry(() => {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API Error: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.choices[0].message.content;
  });
}

/**
 * キーワードレスポンスのパース
 */
function parseKeywordResponse(response) {
  try {
    // JSON部分を抽出（markdown形式の場合に対応）
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('有効なJSONレスポンスが見つかりません');
    }
    
    const data = JSON.parse(jsonMatch[0]);
    const strategicKeywords = data.strategicKeywords;
    
    const allKeywords = [];
    
    // 各カテゴリのキーワードを配列に変換
    Object.keys(strategicKeywords).forEach(category => {
      const categoryData = strategicKeywords[category];
      const priority = getPriorityByCategory(category);
      
      categoryData.keywords.forEach(keyword => {
        allKeywords.push({
          keyword: keyword,
          category: category,
          priority: priority,
          strategy: categoryData.strategy,
          executed: false,
          hitCount: 0,
          lastExecuted: ''
        });
      });
    });
    
    return allKeywords;
    
  } catch (error) {
    Logger.log('キーワードパースエラー:', error);
    Logger.log('レスポンス:', response);
    throw new Error(`キーワードパースエラー: ${error.toString()}`);
  }
}

/**
 * カテゴリ別優先度の決定
 */
function getPriorityByCategory(category) {
  const priorityMap = {
    'painPointHunting': '高',
    'timingCapture': '高',
    'growthTargeting': '中',
    'budgetTargeting': '中'
  };
  
  return priorityMap[category] || '低';
}

/**
 * キーワードをスプレッドシートに保存
 */
function saveKeywordsToSheet(keywords) {
  const sheet = getSafeSheet(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    console.error('KEYWORDS sheet not found - cannot save keywords');
    throw new Error('キーワードシートが見つかりません。システム初期化を実行してください。');
  }
  
  // 既存データのクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  
  // 新しいデータの挿入
  const dataToInsert = keywords.map(kw => [
    kw.keyword,
    kw.category,
    kw.priority,
    kw.strategy,
    kw.executed,
    kw.hitCount,
    kw.lastExecuted
  ]);
  
  if (dataToInsert.length > 0) {
    sheet.getRange(2, 1, dataToInsert.length, 7).setValues(dataToInsert);
  }
  
  Logger.log(`${keywords.length}個のキーワードを保存しました`);
}

/**
 * API呼び出しのリトライ機能
 */
function apiCallWithRetry(apiFunction, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return apiFunction();
    } catch (error) {
      const errorStr = error.toString();
      
      // API制限エラーの場合は指数バックオフで再試行
      if (errorStr.includes('quota') || errorStr.includes('limit') || errorStr.includes('rate_limit')) {
        const waitTime = Math.pow(2, i) * 1000; // 1秒, 2秒, 4秒
        Logger.log(`API制限検出。${waitTime}ms待機後に再試行します...`);
        Utilities.sleep(waitTime);
        continue;
      }
      
      // その他のエラーは即座にスロー
      throw error;
    }
  }
  
  throw new Error('API制限により処理を中止しました。時間をおいて再実行してください。');
}

/**
 * 実行ログの記録
 */
function logExecution(type, status, successCount, errorCount, errorDetails = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
  const lastRow = sheet.getLastRow();
  const executionId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;
  
  const logData = [
    executionId,
    new Date(),
    type,
    status,
    successCount,
    errorCount,
    errorDetails,
    '', // 処理時間は後で更新
    1 // API使用量
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, logData.length).setValues([logData]);
}
