/**
 * 企業検索・分析機能
 * Google Custom Search APIで企業を検索し、詳細情報を分析
 */

/**
 * 制御パネル設定を取得（企業検索用）
 */
function getControlPanelSettingsForCompanies() {
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
    preferredRegion: data[4][1] || '',
    maxCompanies: 20 // 1回の実行で検索する最大企業数
  };
}

/**
 * 企業検索の実行
 */
function executeCompanySearch() {
  const startTime = new Date();
  
  try {
    console.log('🏢 企業検索を開始します...');
    updateExecutionStatus('企業検索を開始します...');
    
    // システム初期化確認
    const systemStatus = checkSystemStatus();
    console.log('システム状況:', systemStatus);
    
    if (systemStatus && systemStatus.needsInitialization) {
      throw new Error(`システムが初期化されていません: ${systemStatus.message}`);
    }
    
    const settings = getControlPanelSettingsForCompanies();
    const keywords = getUnprocessedKeywords();
    
    // APIキーの確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
    const searchEngineId = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    if (!apiKey || !searchEngineId) {
      SpreadsheetApp.getUi().alert(
        '❌ エラー', 
        'Google Search APIキーまたは検索エンジンIDが設定されていません。\n\nメニュー > ⚙️ システム管理 > ⚙️ API設定管理 から設定してください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    if (keywords.length === 0) {
      SpreadsheetApp.getUi().alert(
        '❌ エラー', 
        '未処理のキーワードがありません。\nまずキーワード生成を実行してください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    let totalCompanies = 0;
    let totalErrors = 0;
    
    // 企業数上限まで検索を実行
    for (const keyword of keywords) {
      if (totalCompanies >= settings.maxCompanies) {
        break;
      }
      
      try {
        updateExecutionStatus(`キーワード「${keyword.keyword}」で検索中...`);
        
        const companies = searchCompaniesByKeyword(keyword, settings);
        totalCompanies += companies.length;
        
        // キーワードの実行状況を更新
        updateKeywordStatus(keyword.keyword, companies.length);
        
        // レート制限対策
        Utilities.sleep(1000);
        
      } catch (error) {
        Logger.log(`キーワード「${keyword.keyword}」の検索でエラー: ${error.toString()}`);
        totalErrors++;
      }
    }
    
    const processingTime = Math.round((new Date() - startTime) / 1000);
    const message = `企業検索完了: ${totalCompanies}件の企業を発見しました（処理時間: ${processingTime}秒）`;
    
    updateExecutionStatus(message);
    logExecution('企業検索', `${keywords.length}キーワード`, totalCompanies, totalErrors, '', processingTime);
    
    return totalCompanies;
    
  } catch (error) {
    console.error('❌ 企業検索エラー:', error);
    updateExecutionStatus(`エラー: ${error.message}`);
    
    // ユーザーフレンドリーなエラーメッセージ
    let userMessage = error.message;
    if (error.message.includes('Cannot read properties')) {
      userMessage = 'システム初期化に問題があります。メニュー > システム管理 > システム初期化 を実行してください。';
    } else if (error.message.includes('API')) {
      userMessage = 'API設定に問題があります。メニュー > システム管理 > API設定管理 で設定を確認してください。';
    }
    
    SpreadsheetApp.getUi().alert('❌ エラー', `企業検索でエラーが発生しました:\n\n${userMessage}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return 0;
  }
}

/**
 * 未処理キーワードの取得
 */
function getUnprocessedKeywords() {
  const sheet = getSafeSheet(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    console.error('KEYWORDS sheet not found');
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return data
    .filter(row => !row[4]) // 実行済みフラグがfalseのもの
    .filter(row => row[2] === '高') // 優先度が高のもの
    .map(row => ({
      keyword: row[0],
      category: row[1],
      priority: row[2],
      strategy: row[3],
      executed: row[4],
      hitCount: row[5],
      lastExecuted: row[6]
    }))
    .slice(0, 10); // 一度に処理するキーワード数を制限
}

/**
 * キーワードによる企業検索
 */
function searchCompaniesByKeyword(keywordObj, settings) {
  const searchResults = performGoogleSearch(keywordObj.keyword, settings);
  const companies = [];
  
  for (const result of searchResults) {
    try {
      const companyData = analyzeCompanyWebsite(result, keywordObj);
      
      if (companyData && isValidCompany(companyData)) {
        // 重複チェック
        if (!isDuplicateCompany(companyData)) {
          const matchScore = calculateMatchScore(companyData, settings);
          companyData.matchScore = matchScore;
          
          saveCompanyToSheet(companyData);
          companies.push(companyData);
        }
      }
      
    } catch (error) {
      Logger.log(`企業分析エラー (${result.link}): ${error.toString()}`);
    }
  }
  
  return companies;
}

/**
 * Google Custom Search APIでの検索
 */
function performGoogleSearch(keyword, settings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
  const searchEngineId = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
  
  if (!apiKey || !searchEngineId) {
    throw new Error('Google Search APIキーまたは検索エンジンIDが設定されていません');
  }
  
  // 検索クエリの構築
  let query = keyword;
  if (settings.preferredRegion) {
    query += ` ${settings.preferredRegion}`;
  }
  
  // 企業サイトに絞り込むフィルター
  query += ' site:co.jp OR site:com OR site:net OR site:org';
  
  const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${searchEngineId}&q=${encodeURIComponent(query)}&num=3`;
  
  return apiCallWithRetry(() => {
    const response = UrlFetchApp.fetch(url);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Google Search API Error: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.items) {
      return [];
    }
    
    return data.items.map(item => ({
      title: item.title,
      link: item.link,
      snippet: item.snippet
    }));
  });
}

/**
 * 企業ウェブサイトの分析
 */
function analyzeCompanyWebsite(searchResult, keyword) {
  try {
    // ウェブサイトのHTMLを取得
    const html = fetchWebsiteContent(searchResult.link);
    
    // ChatGPTで企業情報を抽出
    const companyInfo = extractCompanyInfoWithChatGPT(html, searchResult);
    
    if (companyInfo) {
      companyInfo.discoveryKeyword = keyword.keyword;
      companyInfo.registrationDate = new Date();
      companyInfo.officialUrl = searchResult.link;
    }
    
    return companyInfo;
    
  } catch (error) {
    Logger.log(`ウェブサイト分析エラー (${searchResult.link}): ${error.toString()}`);
    return null;
  }
}

/**
 * ウェブサイトコンテンツの取得
 */
function fetchWebsiteContent(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    });
    
    if (response.getResponseCode() === 200) {
      // HTMLコンテンツを制限（ChatGPTのトークン制限対応）
      const content = response.getContentText();
      return content.substring(0, 3000); // 最初の3000文字のみ
    }
    
    return '';
    
  } catch (error) {
    Logger.log(`ウェブサイト取得エラー (${url}): ${error.toString()}`);
    return '';
  }
}

/**
 * ChatGPTで企業情報を抽出
 */
function extractCompanyInfoWithChatGPT(html, searchResult) {
  const prompt = `
WebサイトHTML（抜粋）：
${html}

URL：${searchResult.link}
タイトル：${searchResult.title}

このWebサイトから企業情報を抽出し、以下のJSON形式で回答してください：

{
  "companyName": "正式会社名（不明な場合はタイトルから推測）",
  "industry": "業界（IT/製造業/サービス業/医療/教育/金融/不動産/その他）",
  "employees": "従業員数（数値、不明なら null）",
  "location": "本社所在地（都道府県まで、不明なら null）",
  "hasContactForm": true/false,
  "hasPhoneContact": true/false,
  "isPublicCompany": true/false/null,
  "businessDescription": "主な事業内容（200文字以内）",
  "companySize": "大企業/中小企業/スタートアップ/個人事業主",
  "webFeatures": "サイトから読み取れる特徴や強み（100文字以内）",
  "isValidCompany": true/false
}

注意：
- 会社名が明確でない場合やスパムサイト、個人ブログの場合は "isValidCompany": false
- 正式な企業サイトの場合のみ "isValidCompany": true
- 従業員数は「10名」「50人」などがあれば数値に変換
- 上場企業かどうかは「東証」「JASDAQ」などのキーワードから判断
`;

  try {
    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: 'あなたは企業情報分析の専門家です。Webサイトから正確な企業情報を抽出してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 800,
      temperature: 0.3
    };
    
    const response = callOpenAIAPIForCompanies(payload);
    return parseCompanyInfoResponse(response);
    
  } catch (error) {
    console.error(`企業情報抽出エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * OpenAI API呼び出し（企業検索用）
 */
function callOpenAIAPIForCompanies(payload) {
  return apiCallWithRetryForCompanies(() => {
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
 * API呼び出しのリトライ機能（企業検索用）
 */
function apiCallWithRetryForCompanies(apiFunction, maxRetries = 3) {
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

/**
 * 企業情報レスポンスのパース
 */
function parseCompanyInfoResponse(response) {
  try {
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('有効なJSONレスポンスが見つかりません');
    }
    
    const data = JSON.parse(jsonMatch[0]);
    
    // データの検証
    if (!data.isValidCompany || !data.companyName) {
      return null;
    }
    
    return {
      companyName: data.companyName,
      industry: data.industry || 'その他',
      employees: data.employees,
      location: data.location,
      contactMethod: determineContactMethod(data.hasContactForm, data.hasPhoneContact),
      isPublicCompany: data.isPublicCompany,
      businessDescription: data.businessDescription || '',
      companySize: data.companySize || '中小企業',
      webFeatures: data.webFeatures || ''
    };
    
  } catch (error) {
    Logger.log('企業情報パースエラー:', error);
    Logger.log('レスポンス:', response);
    return null;
  }
}

/**
 * 問い合わせ方法の決定
 */
function determineContactMethod(hasForm, hasPhone) {
  if (hasForm && hasPhone) {
    return 'フォーム・電話';
  } else if (hasForm) {
    return 'フォーム';
  } else if (hasPhone) {
    return '電話';
  } else {
    return 'その他';
  }
}

/**
 * 企業データの有効性チェック
 */
function isValidCompany(companyData) {
  return companyData &&
         companyData.companyName &&
         companyData.companyName.length >= 2 &&
         !companyData.companyName.includes('404') &&
         !companyData.companyName.includes('Not Found');
}

/**
 * 重複企業のチェック
 */
function isDuplicateCompany(companyData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANIES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return false;
  }
  
  const existingCompanies = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  return existingCompanies.some(row => 
    row[0] && row[0].toString().toLowerCase() === companyData.companyName.toLowerCase()
  );
}

/**
 * 企業をスプレッドシートに保存
 */
function saveCompanyToSheet(companyData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANIES);
  const lastRow = sheet.getLastRow();
  const companyId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()) + 1 : 1;
  
  const rowData = [
    companyId,
    companyData.companyName,
    companyData.officialUrl,
    companyData.industry,
    companyData.employees,
    companyData.location,
    companyData.contactMethod,
    companyData.isPublicCompany ? '上場' : '非上場',
    companyData.businessDescription,
    companyData.companySize,
    companyData.matchScore,
    companyData.discoveryKeyword,
    companyData.registrationDate
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  Logger.log(`企業「${companyData.companyName}」を保存しました（スコア: ${companyData.matchScore}）`);
}

/**
 * キーワードの実行状況を更新
 */
function updateKeywordStatus(keyword, hitCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.KEYWORDS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === keyword) {
      const rowNum = i + 2;
      sheet.getRange(rowNum, 5).setValue(true); // 実行済み
      sheet.getRange(rowNum, 6).setValue(hitCount); // ヒット件数
      sheet.getRange(rowNum, 7).setValue(new Date()); // 最終実行日
      break;
    }
  }
}
