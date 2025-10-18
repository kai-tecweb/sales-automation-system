/**
 * API統合テスト機能
 * OpenAI ChatGPT APIとGoogle Search APIの接続と機能テスト
 */

/**
 * API統合テストの実行
 */
function testAPIIntegration() {
  try {
    console.log('🔌 API統合テスト開始...');
    
    updateExecutionStatus('API統合テストを実行中...');
    
    const testResults = {
      openaiAPI: testOpenAIAPI(),
      googleSearchAPI: testGoogleSearchAPI(),
      apiConfiguration: testAPIConfiguration(),
      timestamp: new Date()
    };
    
    // テスト結果の評価
    const allAPIsWorking = testResults.openaiAPI.status === 'OK' && 
                          testResults.googleSearchAPI.status === 'OK' && 
                          testResults.apiConfiguration.status === 'OK';
    
    const resultMessage = allAPIsWorking 
      ? 'すべてのAPI接続が正常に動作しています' 
      : '一部のAPIに問題があります';
    
    updateExecutionStatus(`API統合テスト完了: ${resultMessage}`);
    
    // テスト結果をログに記録
    logExecution('API統合テスト', allAPIsWorking ? 'SUCCESS' : 'WARNING', 
                allAPIsWorking ? 3 : 0, allAPIsWorking ? 0 : 1);
    
    // 結果を表示
    displayAPITestResults(testResults);
    
    return testResults;
    
  } catch (error) {
    console.error('❌ API統合テストエラー:', error);
    updateExecutionStatus(`API統合テストエラー: ${error.message}`);
    handleSystemError('API統合テスト', error);
  }
}

/**
 * OpenAI API接続テスト
 */
function testOpenAIAPI() {
  try {
    console.log('🤖 OpenAI API接続テスト開始...');
    
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!apiKey) {
      return {
        status: 'ERROR',
        message: 'OpenAI APIキーが設定されていません',
        details: 'API設定からOpenAI APIキーを設定してください'
      };
    }
    
    // 軽量なテストリクエスト
    const testPrompt = 'こんにちは';
    
    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: testPrompt
        }
      ],
      max_tokens: 10, // テスト用に少ない制限
      temperature: 0.1
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    console.log('🔄 OpenAI APIリクエスト送信中...');
    
    // タイムアウト制限付きでリクエスト実行
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const responseData = JSON.parse(response.getContentText());
      
      return {
        status: 'OK',
        message: 'OpenAI API接続正常',
        details: {
          model: responseData.model,
          usage: responseData.usage,
          response: responseData.choices[0]?.message?.content || '応答なし'
        }
      };
    } else {
      const errorText = response.getContentText();
      return {
        status: 'ERROR',
        message: `OpenAI API接続エラー (${responseCode})`,
        details: errorText
      };
    }
    
  } catch (error) {
    console.error('❌ OpenAI APIテストエラー:', error);
    return {
      status: 'ERROR',
      message: 'OpenAI API接続テスト失敗',
      details: error.message
    };
  }
}

/**
 * Google Search API接続テスト
 */
function testGoogleSearchAPI() {
  try {
    console.log('🔍 Google Search API接続テスト開始...');
    
    const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
    const searchEngineId = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    if (!apiKey || !searchEngineId) {
      return {
        status: 'ERROR',
        message: 'Google Search API設定が不完全です',
        details: {
          hasApiKey: !!apiKey,
          hasSearchEngineId: !!searchEngineId
        }
      };
    }
    
    // 軽量なテストクエリ
    const testQuery = 'test';
    const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${searchEngineId}&q=${encodeURIComponent(testQuery)}&num=1`;
    
    console.log('🔄 Google Search APIリクエスト送信中...');
    
    const response = UrlFetchApp.fetch(url);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const responseData = JSON.parse(response.getContentText());
      
      return {
        status: 'OK',
        message: 'Google Search API接続正常',
        details: {
          totalResults: responseData.searchInformation?.totalResults || 0,
          searchTime: responseData.searchInformation?.searchTime || 0,
          itemCount: responseData.items?.length || 0
        }
      };
    } else {
      const errorText = response.getContentText();
      return {
        status: 'ERROR',
        message: `Google Search API接続エラー (${responseCode})`,
        details: errorText
      };
    }
    
  } catch (error) {
    console.error('❌ Google Search APIテストエラー:', error);
    return {
      status: 'ERROR',
      message: 'Google Search API接続テスト失敗',
      details: error.message
    };
  }
}

/**
 * API設定の総合チェック
 */
function testAPIConfiguration() {
  try {
    console.log('⚙️ API設定チェック開始...');
    
    const properties = PropertiesService.getScriptProperties();
    
    const config = {
      openaiKey: properties.getProperty('OPENAI_API_KEY'),
      googleSearchKey: properties.getProperty('GOOGLE_SEARCH_API_KEY'),
      searchEngineId: properties.getProperty('GOOGLE_SEARCH_ENGINE_ID')
    };
    
    const issues = [];
    
    // 各設定の確認
    if (!config.openaiKey) {
      issues.push('OpenAI APIキーが未設定');
    } else if (config.openaiKey.length < 20) {
      issues.push('OpenAI APIキーが短すぎます');
    }
    
    if (!config.googleSearchKey) {
      issues.push('Google Search APIキーが未設定');
    } else if (config.googleSearchKey.length < 20) {
      issues.push('Google Search APIキーが短すぎます');
    }
    
    if (!config.searchEngineId) {
      issues.push('カスタム検索エンジンIDが未設定');
    } else if (!config.searchEngineId.includes(':')) {
      issues.push('カスタム検索エンジンIDの形式が正しくありません');
    }
    
    return {
      status: issues.length === 0 ? 'OK' : 'WARNING',
      message: issues.length === 0 ? 'API設定は正常です' : `設定に問題があります: ${issues.length}件`,
      details: {
        issues: issues,
        hasOpenAI: !!config.openaiKey,
        hasGoogleSearch: !!config.googleSearchKey,
        hasSearchEngine: !!config.searchEngineId
      }
    };
    
  } catch (error) {
    console.error('❌ API設定チェックエラー:', error);
    return {
      status: 'ERROR',
      message: 'API設定チェック失敗',
      details: error.message
    };
  }
}

/**
 * APIテスト結果の表示
 */
function displayAPITestResults(testResults) {
  try {
    const sheet = getSafeSheet(SHEET_NAMES.CONTROL);
    if (!sheet) return;
    
    // API テスト結果表示エリア
    const startRow = 12;
    
    // 既存の結果をクリア
    sheet.getRange(startRow, 8, 15, 3).clearContent();
    
    let row = startRow;
    
    // ヘッダー
    sheet.getRange(row, 8).setValue('🔌 API統合テスト結果');
    sheet.getRange(row, 8).setFontWeight('bold');
    sheet.getRange(row, 8).setBackground('#4285f4');
    sheet.getRange(row, 8).setFontColor('white');
    row += 2;
    
    // OpenAI API結果
    sheet.getRange(row, 8).setValue('OpenAI API');
    sheet.getRange(row, 9).setValue(testResults.openaiAPI.status);
    sheet.getRange(row, 10).setValue(testResults.openaiAPI.message);
    setStatusColor(sheet.getRange(row, 9), testResults.openaiAPI.status);
    row++;
    
    // Google Search API結果
    sheet.getRange(row, 8).setValue('Google Search API');
    sheet.getRange(row, 9).setValue(testResults.googleSearchAPI.status);
    sheet.getRange(row, 10).setValue(testResults.googleSearchAPI.message);
    setStatusColor(sheet.getRange(row, 9), testResults.googleSearchAPI.status);
    row++;
    
    // API設定結果
    sheet.getRange(row, 8).setValue('API設定');
    sheet.getRange(row, 9).setValue(testResults.apiConfiguration.status);
    sheet.getRange(row, 10).setValue(testResults.apiConfiguration.message);
    setStatusColor(sheet.getRange(row, 9), testResults.apiConfiguration.status);
    row += 2;
    
    // タイムスタンプ
    sheet.getRange(row, 8).setValue('テスト実行時刻:');
    sheet.getRange(row, 9).setValue(testResults.timestamp);
    
    console.log('✅ APIテスト結果を制御パネルに表示しました');
    
  } catch (error) {
    console.error('❌ APIテスト結果表示エラー:', error);
  }
}

/**
 * ステータスに応じた色設定
 */
function setStatusColor(range, status) {
  switch (status) {
    case 'OK':
      range.setBackground('#d4edda');
      range.setFontColor('#155724');
      break;
    case 'WARNING':
      range.setBackground('#fff3cd');
      range.setFontColor('#856404');
      break;
    case 'ERROR':
      range.setBackground('#f8d7da');
      range.setFontColor('#721c24');
      break;
  }
}

/**
 * API クォータ使用量チェック
 */
function checkAPIQuotaUsage() {
  try {
    console.log('📊 API使用量チェック開始...');
    
    const today = new Date();
    const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    
    const sheet = getSafeSheet(SHEET_NAMES.LOGS);
    if (!sheet) {
      return {
        status: 'WARNING',
        message: 'ログシートが見つからないため、使用量を確認できません'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    let openaiCalls = 0;
    let googleSearchCalls = 0;
    
    // 今日のAPI呼び出し数をカウント
    for (let i = 1; i < data.length; i++) {
      const logDate = new Date(data[i][0]);
      const action = data[i][1];
      
      if (logDate >= startOfDay) {
        if (action.includes('キーワード') || action.includes('提案')) {
          openaiCalls++;
        }
        if (action.includes('企業検索')) {
          googleSearchCalls++;
        }
      }
    }
    
    return {
      status: 'OK',
      message: '使用量チェック完了',
      usage: {
        date: today.toDateString(),
        openaiCalls: openaiCalls,
        googleSearchCalls: googleSearchCalls,
        estimatedCost: (openaiCalls * 0.002) + (googleSearchCalls * 0.005) // 概算コスト
      }
    };
    
  } catch (error) {
    console.error('❌ API使用量チェックエラー:', error);
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * API制限チェック機能
 */
function checkAPILimits() {
  try {
    const quotaCheck = checkAPIQuotaUsage();
    
    const limits = {
      openaiDaily: 100,      // 1日のOpenAI呼び出し制限
      googleSearchDaily: 50  // 1日のGoogle Search呼び出し制限
    };
    
    const warnings = [];
    
    if (quotaCheck.usage) {
      if (quotaCheck.usage.openaiCalls >= limits.openaiDaily * 0.8) {
        warnings.push('OpenAI API使用量が制限の80%に達しました');
      }
      
      if (quotaCheck.usage.googleSearchCalls >= limits.googleSearchDaily * 0.8) {
        warnings.push('Google Search API使用量が制限の80%に達しました');
      }
    }
    
    return {
      status: warnings.length === 0 ? 'OK' : 'WARNING',
      message: warnings.length === 0 ? 'API使用量は正常範囲内です' : `注意: ${warnings.length}件の警告`,
      warnings: warnings,
      usage: quotaCheck.usage,
      limits: limits
    };
    
  } catch (error) {
    console.error('❌ API制限チェックエラー:', error);
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}