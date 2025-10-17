/**
 * 提案メッセージ生成機能
 * 企業情報に基づいた個別最適化された提案を生成
 */

/**
 * メニューから呼び出される詳細ログ付き提案生成の実行
 */
function executeProposalGenerationEnhanced() {
  const startTime = new Date();
  
  try {
    // システム初期化確認
    checkSystemInitialization();
    
    updateExecutionStatus('詳細ログ付き提案メッセージ生成を開始します...');
    
    const settings = getControlPanelSettings();
    const highScoreCompanies = getHighScoreCompanies(70, 10); // スコア70以上、最大10社
    
    if (highScoreCompanies.length === 0) {
      throw new Error('提案対象となる高スコア企業がありません。まず企業検索を実行してください。');
    }
    
    let successCount = 0;
    let errorCount = 0;
    const errorDetails = [];
    
    for (const company of highScoreCompanies) {
      try {
        updateExecutionStatus(`「${company.companyName}」の提案メッセージを生成中...`);
        
        const result = generateProposalForCompanyWithDetails(company, settings);
        
        if (result.success) {
          saveProposalToSheet(company.companyId, result.proposals);
          successCount++;
          
          // 成功ログの記録
          logDetailedProposalResult(company.companyId, company.companyName, 'SUCCESS', null, result.responseInfo);
        } else {
          errorCount++;
          errorDetails.push({
            companyId: company.companyId,
            companyName: company.companyName,
            error: result.error
          });
          
          // エラーログの記録
          logDetailedProposalResult(company.companyId, company.companyName, 'ERROR', result.error, result.responseInfo);
        }
        
        // API制限対策
        Utilities.sleep(2000);
        
      } catch (error) {
        errorCount++;
        const errorInfo = {
          companyId: company.companyId,
          companyName: company.companyName,
          error: {
            type: 'SYSTEM_ERROR',
            message: error.toString(),
            timestamp: new Date()
          }
        };
        
        errorDetails.push(errorInfo);
        logDetailedProposalResult(company.companyId, company.companyName, 'ERROR', errorInfo.error, null);
        
        Logger.log(`提案生成システムエラー (${company.companyName}): ${error.toString()}`);
      }
    }
    
    const processingTime = Math.round((new Date() - startTime) / 1000);
    
    // 詳細結果の表示
    displayDetailedProposalResults(successCount, errorCount, errorDetails, processingTime);
    
    const message = `詳細ログ付き提案メッセージ生成完了: ${successCount}件成功、${errorCount}件失敗（処理時間: ${processingTime}秒）`;
    updateExecutionStatus(message);
    logExecution('詳細提案生成', `${highScoreCompanies.length}社対象`, successCount, errorCount, '', processingTime);
    
    // 結果の通知
    SpreadsheetApp.getUi().alert(
      '詳細ログ付き提案生成完了',
      `${message}\n\n詳細なエラー情報は「提案生成詳細ログ」シートをご確認ください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return {
      successCount: successCount,
      errorCount: errorCount,
      errorDetails: errorDetails
    };
    
  } catch (error) {
    const errorMessage = `詳細ログ付き提案メッセージ生成エラー: ${error.toString()}`;
    updateExecutionStatus(errorMessage);
    logExecution('詳細提案生成', 'ERROR', 0, 1, error.toString());
    
    SpreadsheetApp.getUi().alert(
      'エラー',
      errorMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    throw error;
  }
}

/**
 * 詳細エラー情報付き個別企業向け提案の生成
 */
function generateProposalForCompanyWithDetails(company, settings) {
  const requestStartTime = new Date();
  
  try {
    const prompt = createProposalPrompt(company, settings);
    
    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: 'あなたは営業のプロフェッショナルです。企業情報と商材情報を基に、効果的で個別最適化された提案メッセージを作成してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 1500,
      temperature: 0.7
    };
    
    // API呼び出しの詳細監視
    const apiResult = callOpenAIAPIWithEnhancedLogging(payload);
    
    if (!apiResult.success) {
      return {
        success: false,
        error: apiResult.error,
        responseInfo: apiResult.responseInfo
      };
    }
    
    // レスポンス解析の詳細監視
    const parseResult = parseProposalResponseWithLogging(apiResult.response);
    
    if (!parseResult.success) {
      return {
        success: false,
        error: parseResult.error,
        responseInfo: {
          ...apiResult.responseInfo,
          rawResponse: apiResult.response,
          parseAttempt: parseResult.parseAttempt
        }
      };
    }
    
    const processingTime = new Date() - requestStartTime;
    
    return {
      success: true,
      proposals: parseResult.proposals,
      responseInfo: {
        ...apiResult.responseInfo,
        processingTime: processingTime,
        responseLength: apiResult.response.length
      }
    };
    
  } catch (error) {
    return {
      success: false,
      error: {
        type: 'UNKNOWN_ERROR',
        message: error.toString(),
        timestamp: new Date(),
        stack: error.stack
      },
      responseInfo: null
    };
  }
}

/**
 * 提案生成エラーログの表示
 */
function showProposalErrorLog() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('提案生成詳細ログ');
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert(
        '提案生成ログなし',
        '詳細ログはまだありません。「詳細ログ付き提案生成」を実行してからご確認ください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // ログシートをアクティブにする
    spreadsheet.setActiveSheet(logSheet);
    
    SpreadsheetApp.getUi().alert(
      '提案生成詳細ログ',
      '詳細なログ情報を「提案生成詳細ログ」シートで確認できます。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      `ログ表示エラー: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 従来の提案メッセージ生成（互換性維持）
 */
function executeProposalGeneration() {
  const startTime = new Date();
  
  try {
    // システム初期化確認
    checkSystemInitialization();
    
    updateExecutionStatus('提案メッセージ生成を開始します...');
    
    const settings = getControlPanelSettings();
    const highScoreCompanies = getHighScoreCompanies(70, 10); // スコア70以上、最大10社
    
    if (highScoreCompanies.length === 0) {
      throw new Error('提案対象となる高スコア企業がありません。まず企業検索を実行してください。');
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    for (const company of highScoreCompanies) {
      try {
        updateExecutionStatus(`「${company.companyName}」の提案メッセージを生成中...`);
        
        const proposals = generateProposalForCompany(company, settings);
        
        if (proposals) {
          saveProposalToSheet(company.companyId, proposals);
          successCount++;
        }
        
        // API制限対策
        Utilities.sleep(2000);
        
      } catch (error) {
        Logger.log(`提案生成エラー (${company.companyName}): ${error.toString()}`);
        errorCount++;
      }
    }
    
    const processingTime = Math.round((new Date() - startTime) / 1000);
    const message = `提案メッセージ生成完了: ${successCount}件の提案を生成しました（処理時間: ${processingTime}秒）`;
    
    updateExecutionStatus(message);
    logExecution('提案生成', `${highScoreCompanies.length}社対象`, successCount, errorCount, '', processingTime);
    
    return successCount;
    
  } catch (error) {
    const errorMessage = `提案メッセージ生成エラー: ${error.toString()}`;
    updateExecutionStatus(errorMessage);
    logExecution('提案生成', 'ERROR', 0, 1, error.toString());
    throw error;
  }
}

/**
 * 個別企業向け提案の生成
 */
function generateProposalForCompany(company, settings) {
  const prompt = createProposalPrompt(company, settings);
  
  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: 'あなたは営業のプロフェッショナルです。企業情報と商材情報を基に、効果的で個別最適化された提案メッセージを作成してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: 1500,
    temperature: 0.7
  };
  
  try {
    const response = callOpenAIAPI(payload);
    return parseProposalResponse(response);
  } catch (error) {
    Logger.log(`提案生成APIエラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 提案生成用プロンプトの作成（改良版）
 */
function createProposalPrompt(company, settings) {
  return `
【商材情報】
商材名: ${settings.productName}
概要: ${settings.productDescription}
価格帯: ${settings.priceRange}
対象企業規模: ${settings.targetSize}

【企業情報】
会社名: ${company.companyName}
業界: ${company.industry}
規模: ${company.companySize}
従業員数: ${company.employees || '不明'}
所在地: ${company.location || '不明'}
事業内容: ${company.businessDescription}
問い合わせ方法: ${company.contactMethod}
マッチ度スコア: ${company.matchScore}点
発見キーワード: ${company.discoveryKeyword}

重要: 必ず以下のJSON形式で回答してください。他の説明文は一切含めず、JSONのみを返してください：

{
  "patternA": {
    "subject": "課題訴求型の件名（30文字以内）",
    "body": "課題を前面に出した本文（300文字程度、具体的で説得力のある内容）",
    "approach": "このパターンの戦略的根拠"
  },
  "patternB": {
    "subject": "成功事例型の件名（30文字以内）", 
    "body": "実績や成功事例を前面に出した本文（300文字程度、信頼性を重視）",
    "approach": "このパターンの戦略的根拠"
  },
  "contactForm": "問い合わせフォーム用短文（150文字程度、簡潔で魅力的）",
  "recommendedPattern": "A",
  "painPoint": "この企業の想定される課題（100文字程度）",
  "valueProposition": "商材が提供できる具体的価値（100文字程度）",
  "timing": "最適なアプローチタイミング（50文字程度）",
  "personalizedElements": "この企業特有の個別要素（業界特性、規模感など）"
}

注意事項:
- contactFormは必ず文字列で入力してください
- JSONの構文エラーがないよう注意してください
- 全てのフィールドを必ず含めてください
- 件名は開封率を高める工夫を、本文は具体的なメリットと次のアクションを明確に示してください
`;
}

/**
 * 提案レスポンスのパース
 */
function parseProposalResponse(response) {
  try {
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('有効なJSONレスポンスが見つかりません');
    }
    
    const data = JSON.parse(jsonMatch[0]);
    
    // 必須フィールドの検証
    const requiredFields = ['patternA', 'patternB', 'contactForm'];
    for (const field of requiredFields) {
      if (!data[field] || typeof data[field] !== 'object') {
        throw new Error(`必須フィールド「${field}」が不正です`);
      }
    }
    
    return {
      patternA: {
        subject: data.patternA.subject || '',
        body: data.patternA.body || '',
        approach: data.patternA.approach || ''
      },
      patternB: {
        subject: data.patternB.subject || '',
        body: data.patternB.body || '',
        approach: data.patternB.approach || ''
      },
      contactForm: data.contactForm || '',
      recommendedPattern: data.recommendedPattern || 'A',
      painPoint: data.painPoint || '',
      valueProposition: data.valueProposition || '',
      timing: data.timing || '',
      personalizedElements: data.personalizedElements || '',
      generatedDate: new Date()
    };
    
  } catch (error) {
    Logger.log('提案パースエラー:', error);
    Logger.log('レスポンス:', response);
    throw new Error(`提案パースエラー: ${error.toString()}`);
  }
}

/**
 * 提案をスプレッドシートに保存
 */
function saveProposalToSheet(companyId, proposals) {
  const sheet = getSafeSheet(SHEET_NAMES.PROPOSALS);
  if (!sheet) {
    console.error('PROPOSALS sheet not found - cannot save proposal');
    throw new Error('提案シートが見つかりません。システム初期化を実行してください。');
  }
  
  // 既存の提案をチェック（重複回避）
  if (isProposalExists(companyId)) {
    Logger.log(`企業ID ${companyId} の提案は既に存在します。スキップします。`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  
  const rowData = [
    companyId,
    proposals.patternA.subject,
    proposals.patternA.body,
    proposals.patternB.subject,
    proposals.patternB.body,
    proposals.contactForm,
    proposals.recommendedPattern,
    proposals.painPoint,
    proposals.valueProposition,
    proposals.timing,
    proposals.generatedDate
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  Logger.log(`企業ID ${companyId} の提案メッセージを保存しました`);
}

/**
 * 提案の存在チェック
 */
function isProposalExists(companyId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PROPOSALS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return false;
  }
  
  const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return existingIds.includes(companyId);
}

/**
 * 企業別提案の取得
 */
function getProposalByCompanyId(companyId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PROPOSALS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return null;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const proposalRow = data.find(row => row[0] === companyId);
  
  if (!proposalRow) {
    return null;
  }
  
  return {
    companyId: proposalRow[0],
    patternA: {
      subject: proposalRow[1],
      body: proposalRow[2]
    },
    patternB: {
      subject: proposalRow[3],
      body: proposalRow[4]
    },
    contactForm: proposalRow[5],
    recommendedPattern: proposalRow[6],
    painPoint: proposalRow[7],
    valueProposition: proposalRow[8],
    timing: proposalRow[9],
    generatedDate: proposalRow[10]
  };
}

/**
 * 全提案の取得（統計用）
 */
function getAllProposals() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PROPOSALS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  return data.map(row => ({
    companyId: row[0],
    patternA: {
      subject: row[1],
      body: row[2]
    },
    patternB: {
      subject: row[3],
      body: row[4]
    },
    contactForm: row[5],
    recommendedPattern: row[6],
    painPoint: row[7],
    valueProposition: row[8],
    timing: row[9],
    generatedDate: row[10]
  }));
}

/**
 * 提案メッセージの品質チェック
 */
function validateProposalQuality(proposals) {
  const issues = [];
  
  // 件名の長さチェック
  if (proposals.patternA.subject.length > 30) {
    issues.push('パターンAの件名が長すぎます（30文字以内推奨）');
  }
  
  if (proposals.patternB.subject.length > 30) {
    issues.push('パターンBの件名が長すぎます（30文字以内推奨）');
  }
  
  // 本文の長さチェック
  if (proposals.patternA.body.length < 100 || proposals.patternA.body.length > 400) {
    issues.push('パターンAの本文の長さが適切ではありません（100-400文字推奨）');
  }
  
  if (proposals.patternB.body.length < 100 || proposals.patternB.body.length > 400) {
    issues.push('パターンBの本文の長さが適切ではありません（100-400文字推奨）');
  }
  
  // フォーム用メッセージの長さチェック
  if (proposals.contactForm.length > 200) {
    issues.push('フォーム用メッセージが長すぎます（200文字以内推奨）');
  }
  
  // 必須要素の存在チェック
  const hasCallToAction = 
    proposals.patternA.body.includes('お問い合わせ') || 
    proposals.patternA.body.includes('ご相談') || 
    proposals.patternA.body.includes('ご連絡');
  
  if (!hasCallToAction) {
    issues.push('明確なコールトゥアクションが不足しています');
  }
  
  return {
    isValid: issues.length === 0,
    issues: issues,
    score: Math.max(0, 100 - issues.length * 20)
  };
}

/**
 * 提案メッセージのプレビュー生成
 */
function generateProposalPreview(companyId) {
  const company = getCompanyById(companyId);
  const proposal = getProposalByCompanyId(companyId);
  
  if (!company || !proposal) {
    return null;
  }
  
  return {
    company: {
      name: company.companyName,
      industry: company.industry,
      matchScore: company.matchScore
    },
    preview: {
      recommendedPattern: proposal.recommendedPattern,
      selectedProposal: proposal.recommendedPattern === 'A' ? proposal.patternA : proposal.patternB,
      contactForm: proposal.contactForm,
      timing: proposal.timing,
      painPoint: proposal.painPoint,
      valueProposition: proposal.valueProposition
    }
  };
}

/**
 * 企業情報の取得（ID指定）
 */
function getCompanyById(companyId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANIES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return null;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const companyRow = data.find(row => row[0] === companyId);
  
  if (!companyRow) {
    return null;
  }
  
  return {
    companyId: companyRow[0],
    companyName: companyRow[1],
    officialUrl: companyRow[2],
    industry: companyRow[3],
    employees: companyRow[4],
    location: companyRow[5],
    contactMethod: companyRow[6],
    isPublicCompany: companyRow[7],
    businessDescription: companyRow[8],
    companySize: companyRow[9],
    matchScore: companyRow[10],
    discoveryKeyword: companyRow[11],
    registrationDate: companyRow[12]
  };
}

/**
 * 詳細エラー情報付きOpenAI API呼び出し
 */
function callOpenAIAPIWithEnhancedLogging(payload) {
  const requestStartTime = new Date();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    return {
      success: false,
      error: {
        type: 'API_KEY_MISSING',
        message: 'OpenAI APIキーが設定されていません',
        timestamp: new Date()
      },
      responseInfo: null
    };
  }
  
  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseTime = new Date() - requestStartTime;
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    const responseInfo = {
      statusCode: statusCode,
      responseTime: responseTime,
      requestTokens: Math.ceil(JSON.stringify(payload).length / 3),
      timestamp: new Date()
    };
    
    if (statusCode !== 200) {
      let errorType = 'API_ERROR';
      let errorMessage = `HTTPエラー ${statusCode}`;
      
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error) {
          errorMessage = errorData.error.message || errorMessage;
          
          if (statusCode === 429) {
            errorType = 'QUOTA_EXCEEDED';
          } else if (statusCode === 401) {
            errorType = 'AUTHENTICATION_ERROR';
          } else if (statusCode === 408 || statusCode === 504) {
            errorType = 'TIMEOUT_ERROR';
          }
        }
      } catch (parseError) {
        errorMessage += ` - レスポンス解析失敗: ${responseText.substring(0, 100)}`;
      }
      
      return {
        success: false,
        error: {
          type: errorType,
          message: errorMessage,
          statusCode: statusCode,
          timestamp: new Date(),
          rawResponse: responseText.substring(0, 300)
        },
        responseInfo: responseInfo
      };
    }
    
    try {
      const data = JSON.parse(responseText);
      const responseTokens = data.usage ? data.usage.total_tokens : 0;
      
      responseInfo.responseTokens = responseTokens;
      responseInfo.model = data.model;
      
      return {
        success: true,
        response: data.choices[0].message.content,
        responseInfo: responseInfo
      };
      
    } catch (parseError) {
      return {
        success: false,
        error: {
          type: 'RESPONSE_PARSE_ERROR',
          message: `レスポンス解析エラー: ${parseError.toString()}`,
          timestamp: new Date(),
          rawResponse: responseText.substring(0, 300)
        },
        responseInfo: responseInfo
      };
    }
    
  } catch (error) {
    return {
      success: false,
      error: {
        type: 'CONNECTION_ERROR',
        message: `API接続エラー: ${error.toString()}`,
        timestamp: new Date()
      },
      responseInfo: {
        requestTime: new Date() - requestStartTime,
        timestamp: new Date()
      }
    };
  }
}

/**
 * 詳細ログ付きレスポンス解析（改良版）
 */
function parseProposalResponseWithLogging(response) {
  try {
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return {
        success: false,
        error: {
          type: 'JSON_NOT_FOUND',
          message: '有効なJSONレスポンスが見つかりません',
          timestamp: new Date(),
          responsePreview: response.substring(0, 200)
        },
        parseAttempt: 'JSON_MATCH_FAILED'
      };
    }
    
    let data;
    try {
      data = JSON.parse(jsonMatch[0]);
    } catch (jsonError) {
      return {
        success: false,
        error: {
          type: 'JSON_PARSE_ERROR',
          message: `JSON解析エラー: ${jsonError.toString()}`,
          timestamp: new Date(),
          jsonString: jsonMatch[0].substring(0, 200)
        },
        parseAttempt: 'JSON_PARSE_FAILED'
      };
    }
    
    // デバッグ情報の詳細ログ
    Logger.log('解析対象JSON構造:', JSON.stringify(data, null, 2));
    
    // 柔軟なフィールド検証（contactFormを文字列として許可）
    const requiredFields = ['patternA', 'patternB'];
    for (const field of requiredFields) {
      if (!data[field] || typeof data[field] !== 'object') {
        return {
          success: false,
          error: {
            type: 'FIELD_VALIDATION_ERROR',
            message: `必須オブジェクトフィールド「${field}」が不正です - 型: ${typeof data[field]}, 値: ${JSON.stringify(data[field])}`,
            timestamp: new Date(),
            availableFields: Object.keys(data),
            actualData: JSON.stringify(data).substring(0, 300)
          },
          parseAttempt: 'FIELD_VALIDATION_FAILED'
        };
      }
    }
    
    // contactFormは文字列または配列でも許可
    if (!data['contactForm']) {
      return {
        success: false,
        error: {
          type: 'FIELD_VALIDATION_ERROR',
          message: `必須フィールド「contactForm」が存在しません`,
          timestamp: new Date(),
          availableFields: Object.keys(data),
          actualData: JSON.stringify(data).substring(0, 300)
        },
        parseAttempt: 'CONTACT_FORM_MISSING'
      };
    }
    
    // contactFormの型に応じた処理
    let contactFormValue = '';
    if (typeof data.contactForm === 'string') {
      contactFormValue = data.contactForm;
    } else if (Array.isArray(data.contactForm)) {
      contactFormValue = data.contactForm.join(' ');
    } else if (typeof data.contactForm === 'object') {
      contactFormValue = data.contactForm.message || data.contactForm.text || JSON.stringify(data.contactForm);
    } else {
      contactFormValue = String(data.contactForm || '');
    }
    
    const proposals = {
      patternA: {
        subject: data.patternA.subject || '',
        body: data.patternA.body || '',
        approach: data.patternA.approach || ''
      },
      patternB: {
        subject: data.patternB.subject || '',
        body: data.patternB.body || '',
        approach: data.patternB.approach || ''
      },
      contactForm: contactFormValue,
      recommendedPattern: data.recommendedPattern || 'A',
      painPoint: data.painPoint || '',
      valueProposition: data.valueProposition || '',
      timing: data.timing || '',
      personalizedElements: data.personalizedElements || '',
      generatedDate: new Date()
    };
    
    // 最終的な品質チェック
    if (!proposals.patternA.subject || !proposals.patternA.body || 
        !proposals.patternB.subject || !proposals.patternB.body) {
      return {
        success: false,
        error: {
          type: 'CONTENT_VALIDATION_ERROR',
          message: 'パターンAまたはパターンBの件名・本文が不完全です',
          timestamp: new Date(),
          proposals: proposals
        },
        parseAttempt: 'CONTENT_VALIDATION_FAILED'
      };
    }
    
    return {
      success: true,
      proposals: proposals,
      parseAttempt: 'SUCCESS'
    };
    
  } catch (error) {
    return {
      success: false,
      error: {
        type: 'UNEXPECTED_ERROR',
        message: `予期しないパースエラー: ${error.toString()}`,
        timestamp: new Date(),
        stack: error.stack
      },
      parseAttempt: 'UNEXPECTED_ERROR'
    };
  }
}

/**
 * 詳細ログの記録
 */
function logDetailedProposalResult(companyId, companyName, status, error, responseInfo) {
  try {
    const sheet = getOrCreateSheetForLogging('提案生成詳細ログ');
    
    // ヘッダーの設定
    if (sheet.getLastRow() === 0) {
      const headers = [
        '企業ID', '企業名', '実行時刻', 'ステータス', 'エラータイプ', 'エラーメッセージ', 
        'HTTPステータス', 'レスポンス時間(ms)', 'リクエストトークン', 'レスポンストークン', 
        'モデル', '詳細情報'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    const timestamp = new Date();
    const rowData = [
      companyId,
      companyName,
      timestamp,
      status,
      error ? error.type || '' : '',
      error ? error.message || '' : '',
      responseInfo ? responseInfo.statusCode || '' : '',
      responseInfo ? responseInfo.responseTime || '' : '',
      responseInfo ? responseInfo.requestTokens || '' : '',
      responseInfo ? responseInfo.responseTokens || '' : '',
      responseInfo ? responseInfo.model || '' : '',
      error ? JSON.stringify(error).substring(0, 100) : (responseInfo ? `処理時間:${responseInfo.processingTime}ms` : '')
    ];
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
    
  } catch (logError) {
    Logger.log(`詳細ログ記録エラー: ${logError.toString()}`);
  }
}

/**
 * 詳細結果の表示
 */
function displayDetailedProposalResults(successCount, errorCount, errorDetails, processingTime) {
  if (errorCount === 0) {
    return;
  }
  
  try {
    const sheet = getOrCreateSheetForLogging('提案エラー詳細');
    
    // 前回の結果をクリア
    sheet.clear();
    
    // ヘッダー情報
    sheet.getRange(1, 1).setValue('🚨 詳細提案生成エラーレポート');
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    
    sheet.getRange(2, 1).setValue(`実行時刻: ${new Date()}`);
    sheet.getRange(3, 1).setValue(`成功: ${successCount}件, 失敗: ${errorCount}件, 処理時間: ${processingTime}秒`);
    
    // エラー詳細テーブル
    const headers = ['企業名', 'エラータイプ', 'エラーメッセージ', '発生時刻', 'HTTPステータス', '詳細情報'];
    sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(5, 1, 1, headers.length).setFontWeight('bold');
    
    let row = 6;
    for (const errorDetail of errorDetails) {
      const error = errorDetail.error;
      const rowData = [
        errorDetail.companyName,
        error.type || 'Unknown',
        error.message || '',
        error.timestamp || '',
        error.statusCode || '',
        error.rawResponse ? error.rawResponse.substring(0, 50) : ''
      ];
      
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    }
    
    // エラー統計
    const errorTypes = {};
    errorDetails.forEach(detail => {
      const type = detail.error.type || 'Unknown';
      errorTypes[type] = (errorTypes[type] || 0) + 1;
    });
    
    sheet.getRange(row + 1, 1).setValue('📊 エラー統計:');
    sheet.getRange(row + 1, 1).setFontWeight('bold');
    
    let statsRow = row + 2;
    for (const [type, count] of Object.entries(errorTypes)) {
      sheet.getRange(statsRow, 1).setValue(`${type}: ${count}件`);
      statsRow++;
    }
    
  } catch (displayError) {
    Logger.log(`詳細結果表示エラー: ${displayError.toString()}`);
  }
}

/**
 * ログ用シートの取得または作成
 */
function getOrCreateSheetForLogging(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}
