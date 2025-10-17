/**
 * 提案メッセージ生成機能（詳細エラーログ対応版）
 * 企業情報に基づいた個別最適化された提案を生成
 */

/**
 * 詳細ログ用のエラー情報構造
 */
const PROPOSAL_ERROR_TYPES = {
  API_CONNECTION: 'API接続エラー',
  API_QUOTA: 'API使用量制限',
  API_TIMEOUT: 'APIタイムアウト',
  RESPONSE_PARSE: 'レスポンス解析エラー',
  DATA_VALIDATION: 'データ検証エラー',
  SYSTEM_ERROR: 'システムエラー',
  UNKNOWN: '不明なエラー'
};

/**
 * 強化された提案メッセージ生成
 */
function executeProposalGenerationEnhanced() {
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
    const errorDetails = [];
    
    for (const company of highScoreCompanies) {
      try {
        updateExecutionStatus(`「${company.companyName}」の提案メッセージを生成中...`);
        
        const result = generateProposalForCompanyEnhanced(company, settings);
        
        if (result.success) {
          saveProposalToSheet(company.companyId, result.proposals);
          successCount++;
          
          // 成功ログの記録
          logProposalResult(company.companyId, company.companyName, 'SUCCESS', null, result.responseInfo);
        } else {
          errorCount++;
          errorDetails.push({
            companyId: company.companyId,
            companyName: company.companyName,
            error: result.error
          });
          
          // エラーログの記録
          logProposalResult(company.companyId, company.companyName, 'ERROR', result.error, result.responseInfo);
        }
        
        // API制限対策
        Utilities.sleep(2000);
        
      } catch (error) {
        errorCount++;
        const errorInfo = {
          companyId: company.companyId,
          companyName: company.companyName,
          error: {
            type: PROPOSAL_ERROR_TYPES.SYSTEM_ERROR,
            message: error.toString(),
            timestamp: new Date()
          }
        };
        
        errorDetails.push(errorInfo);
        logProposalResult(company.companyId, company.companyName, 'ERROR', errorInfo.error, null);
        
        Logger.log(`提案生成システムエラー (${company.companyName}): ${error.toString()}`);
      }
    }
    
    const processingTime = Math.round((new Date() - startTime) / 1000);
    
    // 詳細結果の表示
    displayProposalResults(successCount, errorCount, errorDetails, processingTime);
    
    const message = `提案メッセージ生成完了: ${successCount}件成功、${errorCount}件失敗（処理時間: ${processingTime}秒）`;
    updateExecutionStatus(message);
    logExecution('提案生成', `${highScoreCompanies.length}社対象`, successCount, errorCount, '', processingTime);
    
    return {
      successCount: successCount,
      errorCount: errorCount,
      errorDetails: errorDetails
    };
    
  } catch (error) {
    const errorMessage = `提案メッセージ生成エラー: ${error.toString()}`;
    updateExecutionStatus(errorMessage);
    logExecution('提案生成', 'ERROR', 0, 1, error.toString());
    throw error;
  }
}

/**
 * 詳細エラー情報付き個別企業向け提案の生成
 */
function generateProposalForCompanyEnhanced(company, settings) {
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
    const apiResult = callOpenAIAPIWithDetails(payload);
    
    if (!apiResult.success) {
      return {
        success: false,
        error: apiResult.error,
        responseInfo: apiResult.responseInfo
      };
    }
    
    // レスポンス解析の詳細監視
    const parseResult = parseProposalResponseEnhanced(apiResult.response);
    
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
        type: PROPOSAL_ERROR_TYPES.UNKNOWN,
        message: error.toString(),
        timestamp: new Date(),
        stack: error.stack
      },
      responseInfo: null
    };
  }
}

/**
 * 詳細情報付きOpenAI API呼び出し
 */
function callOpenAIAPIWithDetails(payload) {
  const requestStartTime = new Date();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    return {
      success: false,
      error: {
        type: PROPOSAL_ERROR_TYPES.API_CONNECTION,
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
      requestTokens: estimateTokenCount(payload.messages),
      timestamp: new Date()
    };
    
    if (statusCode !== 200) {
      let errorType = PROPOSAL_ERROR_TYPES.API_CONNECTION;
      let errorMessage = `HTTPエラー ${statusCode}`;
      
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error) {
          errorMessage = errorData.error.message || errorMessage;
          
          // エラータイプの詳細分類
          if (statusCode === 429) {
            errorType = PROPOSAL_ERROR_TYPES.API_QUOTA;
          } else if (statusCode === 408 || statusCode === 504) {
            errorType = PROPOSAL_ERROR_TYPES.API_TIMEOUT;
          }
        }
      } catch (parseError) {
        // エラーレスポンスのパースに失敗した場合
        errorMessage += ` - レスポンス解析失敗: ${responseText.substring(0, 200)}`;
      }
      
      return {
        success: false,
        error: {
          type: errorType,
          message: errorMessage,
          statusCode: statusCode,
          timestamp: new Date(),
          rawResponse: responseText.substring(0, 500)
        },
        responseInfo: responseInfo
      };
    }
    
    // 成功レスポンスの処理
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
          type: PROPOSAL_ERROR_TYPES.RESPONSE_PARSE,
          message: `レスポンス解析エラー: ${parseError.toString()}`,
          timestamp: new Date(),
          rawResponse: responseText.substring(0, 500)
        },
        responseInfo: responseInfo
      };
    }
    
  } catch (error) {
    return {
      success: false,
      error: {
        type: PROPOSAL_ERROR_TYPES.API_CONNECTION,
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
 * 詳細エラー情報付き提案レスポンス解析
 */
function parseProposalResponseEnhanced(response) {
  try {
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return {
        success: false,
        error: {
          type: PROPOSAL_ERROR_TYPES.RESPONSE_PARSE,
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
          type: PROPOSAL_ERROR_TYPES.RESPONSE_PARSE,
          message: `JSON解析エラー: ${jsonError.toString()}`,
          timestamp: new Date(),
          jsonString: jsonMatch[0].substring(0, 300)
        },
        parseAttempt: 'JSON_PARSE_FAILED'
      };
    }
    
    // 必須フィールドの詳細検証
    const requiredFields = ['patternA', 'patternB', 'contactForm'];
    for (const field of requiredFields) {
      if (!data[field]) {
        return {
          success: false,
          error: {
            type: PROPOSAL_ERROR_TYPES.DATA_VALIDATION,
            message: `必須フィールド「${field}」が存在しません`,
            timestamp: new Date(),
            availableFields: Object.keys(data)
          },
          parseAttempt: 'FIELD_VALIDATION_FAILED'
        };
      }
      
      if (typeof data[field] !== 'object') {
        return {
          success: false,
          error: {
            type: PROPOSAL_ERROR_TYPES.DATA_VALIDATION,
            message: `フィールド「${field}」が正しいオブジェクト形式ではありません`,
            timestamp: new Date(),
            fieldType: typeof data[field],
            fieldValue: String(data[field]).substring(0, 100)
          },
          parseAttempt: 'FIELD_TYPE_VALIDATION_FAILED'
        };
      }
    }
    
    // サブフィールドの検証
    const patternFields = ['subject', 'body'];
    for (const pattern of ['patternA', 'patternB']) {
      for (const subField of patternFields) {
        if (!data[pattern][subField] || typeof data[pattern][subField] !== 'string') {
          return {
            success: false,
            error: {
              type: PROPOSAL_ERROR_TYPES.DATA_VALIDATION,
              message: `${pattern}.${subField}が正しい文字列ではありません`,
              timestamp: new Date(),
              fieldValue: data[pattern][subField]
            },
            parseAttempt: 'SUBFIELD_VALIDATION_FAILED'
          };
        }
      }
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
      contactForm: data.contactForm || '',
      recommendedPattern: data.recommendedPattern || 'A',
      painPoint: data.painPoint || '',
      valueProposition: data.valueProposition || '',
      timing: data.timing || '',
      personalizedElements: data.personalizedElements || '',
      generatedDate: new Date()
    };
    
    return {
      success: true,
      proposals: proposals,
      parseAttempt: 'SUCCESS'
    };
    
  } catch (error) {
    return {
      success: false,
      error: {
        type: PROPOSAL_ERROR_TYPES.RESPONSE_PARSE,
        message: `予期しないパースエラー: ${error.toString()}`,
        timestamp: new Date(),
        stack: error.stack
      },
      parseAttempt: 'UNEXPECTED_ERROR'
    };
  }
}

/**
 * 提案生成結果の詳細ログ記録
 */
function logProposalResult(companyId, companyName, status, error, responseInfo) {
  try {
    const sheet = getOrCreateSheet('提案生成ログ');
    
    // ヘッダーの設定（初回のみ）
    if (sheet.getLastRow() === 0) {
      const headers = [
        '企業ID', '企業名', '実行時刻', 'ステータス', 'エラータイプ', 'エラーメッセージ', 
        'HTTPステータス', 'レスポンス時間(ms)', 'リクエストトークン', 'レスポンストークン', 
        '処理時間(ms)', 'モデル', 'その他詳細'
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
      responseInfo ? responseInfo.processingTime || '' : '',
      responseInfo ? responseInfo.model || '' : '',
      error ? JSON.stringify(error).substring(0, 200) : ''
    ];
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
    
    // 最新100行のみ保持（パフォーマンス対策）
    if (lastRow > 101) {
      sheet.deleteRows(2, lastRow - 100);
    }
    
  } catch (logError) {
    Logger.log(`ログ記録エラー: ${logError.toString()}`);
  }
}

/**
 * 提案生成結果の詳細表示
 */
function displayProposalResults(successCount, errorCount, errorDetails, processingTime) {
  if (errorCount === 0) {
    return; // エラーがない場合は詳細表示不要
  }
  
  try {
    const sheet = getOrCreateSheet('提案エラー詳細');
    
    // 前回の結果をクリア
    sheet.clear();
    
    // ヘッダー情報
    sheet.getRange(1, 1).setValue('🚨 提案生成エラー詳細レポート');
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    
    sheet.getRange(2, 1).setValue(`実行時刻: ${new Date()}`);
    sheet.getRange(3, 1).setValue(`成功: ${successCount}件, 失敗: ${errorCount}件, 処理時間: ${processingTime}秒`);
    
    // エラー詳細テーブル
    const headers = ['企業名', 'エラータイプ', 'エラーメッセージ', '発生時刻', '詳細情報'];
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
        error.statusCode ? `HTTP:${error.statusCode}` : (error.rawResponse ? error.rawResponse.substring(0, 50) : '')
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
    
    // 推奨対策
    sheet.getRange(statsRow + 1, 1).setValue('💡 推奨対策:');
    sheet.getRange(statsRow + 1, 1).setFontWeight('bold');
    
    statsRow += 2;
    if (errorTypes[PROPOSAL_ERROR_TYPES.API_QUOTA]) {
      sheet.getRange(statsRow, 1).setValue('• OpenAI APIの使用量制限: 課金設定を確認してください');
      statsRow++;
    }
    if (errorTypes[PROPOSAL_ERROR_TYPES.API_CONNECTION]) {
      sheet.getRange(statsRow, 1).setValue('• API接続エラー: APIキーの設定を確認してください');
      statsRow++;
    }
    if (errorTypes[PROPOSAL_ERROR_TYPES.RESPONSE_PARSE]) {
      sheet.getRange(statsRow, 1).setValue('• レスポンス解析エラー: APIモデルの設定を確認してください');
      statsRow++;
    }
    
  } catch (displayError) {
    Logger.log(`結果表示エラー: ${displayError.toString()}`);
  }
}

/**
 * トークン数の概算
 */
function estimateTokenCount(messages) {
  try {
    const text = messages.map(msg => msg.content).join(' ');
    return Math.ceil(text.length / 3); // 概算: 1トークン ≈ 3文字
  } catch (error) {
    return 0;
  }
}

/**
 * シート取得または作成
 */
function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}
