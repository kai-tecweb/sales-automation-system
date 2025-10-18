/**
 * キーワード生成機能 - OpenAI API統合版
 */

/**
 * メニュー用キーワード生成関数
 */
function generateKeywords() {
  try {
    console.log('🔤 キーワード生成を開始します...');
    updateExecutionStatus('キーワード生成を開始します...');
    
    // 制御パネルから設定を取得
    const settings = getControlPanelSettings();
    
    // 入力値チェック
    if (!settings.productName || !settings.productDescription) {
      SpreadsheetApp.getUi().alert('❌ エラー', '制御パネルで商材名と商材概要を入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // APIキーの確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('❌ エラー', 'OpenAI APIキーが設定されていません。\n\nメニュー > ⚙️ システム管理 > ⚙️ API設定管理 から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // ChatGPT APIでキーワード生成
    updateExecutionStatus('AI分析中...');
    const keywords = generateKeywordsWithChatGPT(settings);
    
    // 結果をスプレッドシートに保存
    updateExecutionStatus('結果を保存中...');
    saveKeywordsToSheet(keywords);
    
    const successMessage = `✅ キーワード生成完了！\n\n${keywords.length}個の戦略的キーワードを生成しました。\n「生成キーワード」シートで確認できます。`;
    
    updateExecutionStatus(`キーワード生成完了: ${keywords.length}個のキーワードを生成しました`);
    logExecution('キーワード生成', 'SUCCESS', keywords.length, 0);
    
    SpreadsheetApp.getUi().alert('キーワード生成完了', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log('✅ キーワード生成処理完了');
    
  } catch (error) {
    console.error('❌ キーワード生成エラー:', error);
    updateExecutionStatus(`エラー: ${error.message}`);
    logExecution('キーワード生成', 'ERROR', 0, 1);
    SpreadsheetApp.getUi().alert('❌ エラー', `キーワード生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 制御パネル設定を取得
 */
function getControlPanelSettings() {
  const sheet = getSafeSheet(SHEET_NAMES.CONTROL);
  if (!sheet) {
    throw new Error('制御パネルシートが見つかりません。システムを初期化してください。');
  }
  
  // 制御パネルからデータを取得
  const data = sheet.getRange('A2:B6').getValues();
  
  return {
    productName: data[0][1] || '',
    productDescription: data[1][1] || '',
    priceRange: data[2][1] || '',
    targetSize: data[3][1] || '',
    preferredRegion: data[4][1] || ''
  };
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
        content: 'あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。JSON形式で回答してください。'
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

この商材が刺さりそうな企業を見つけるための検索キーワードを戦略的に生成してください。

以下の4つのカテゴリに分けて、それぞれ8-12個のキーワードを生成してください：

1. painPointHunting: 課題を抱える企業を発見するためのキーワード
2. growthTargeting: 成長している企業を発見するためのキーワード  
3. budgetTargeting: 予算のある企業を発見するためのキーワード
4. timingCapture: 導入タイミングが良い企業を発見するためのキーワード

出力は以下のJSON形式でお願いします：
{
  "strategicKeywords": {
    "painPointHunting": {
      "strategy": "課題発見戦略の説明",
      "keywords": ["キーワード1", "キーワード2", ...]
    },
    "growthTargeting": {
      "strategy": "成長企業発見戦略の説明", 
      "keywords": ["キーワード1", "キーワード2", ...]
    },
    "budgetTargeting": {
      "strategy": "予算企業発見戦略の説明",
      "keywords": ["キーワード1", "キーワード2", ...]
    },
    "timingCapture": {
      "strategy": "タイミング発見戦略の説明",
      "keywords": ["キーワード1", "キーワード2", ...]
    }
  }
}`;
}

/**
 * OpenAI APIを呼び出し
 */
function callOpenAIAPI(payload) {
  return apiCallWithRetry(() => {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
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
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    
    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      throw new Error(`OpenAI API Error (${response.getResponseCode()}): ${errorText}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.choices[0].message.content;
  });
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
        console.log(`API制限検出。${waitTime}ms待機後に再試行します...`);
        Utilities.sleep(waitTime);
        continue;
      }
      
      // その他のエラーは即座にスロー
      throw error;
    }
  }
  
  throw new Error('API制限により処理を中止しました。時間をおいて再実行してください。');
}

// 企業検索機能はcompanies.jsファイルで実装されています

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
    console.error('キーワードパースエラー:', error);
    console.error('レスポンス:', response);
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
  
  console.log(`${keywords.length}個のキーワードを保存しました`);
}

/**
 * 提案生成（デモ版）
 */
function generatePersonalizedProposals() {
  try {
    SpreadsheetApp.getUi().alert('提案生成', '提案生成機能（デモ版）\n\n次のフェーズで実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ 提案生成エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `提案生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 完全自動化実行（デモ版）
 */
function executeFullWorkflow() {
  try {
    SpreadsheetApp.getUi().alert('完全自動化', '完全自動化機能（デモ版）\n\n次のフェーズで実装予定です。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ 完全自動化エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `完全自動化でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}