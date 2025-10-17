/**
 * 詳細レスポンス分析とデバッグ機能
 */

/**
 * 最新の提案生成レスポンスを分析
 */
function analyzeLastProposalResponse() {
  try {
    // 最新のログから最後のAPIレスポンスを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('提案生成詳細ログ');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('ログなし', '提案生成ログがありません。まず詳細ログ付き提案生成を実行してください。');
      return;
    }
    
    // 最後のエラー企業でテスト実行
    const lastRow = sheet.getLastRow();
    const lastLogRow = sheet.getRange(lastRow, 1, 1, 12).getValues()[0];
    const companyId = lastLogRow[0];
    const companyName = lastLogRow[1];
    
    if (!companyId) {
      SpreadsheetApp.getUi().alert('データエラー', '企業IDが見つかりません。');
      return;
    }
    
    // 企業情報の取得
    const company = getCompanyById(companyId);
    if (!company) {
      SpreadsheetApp.getUi().alert('企業データエラー', `企業ID ${companyId} のデータが見つかりません。`);
      return;
    }
    
    // 設定の取得
    const settings = getControlPanelSettings();
    
    // テスト実行とレスポンス分析
    updateExecutionStatus(`「${companyName}」のレスポンス分析を開始...`);
    
    const result = testProposalGenerationWithAnalysis(company, settings);
    
    // 分析結果の表示
    displayResponseAnalysis(result, companyName);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('分析エラー', `エラー: ${error.toString()}`);
  }
}

/**
 * 詳細分析付きテスト実行
 */
function testProposalGenerationWithAnalysis(company, settings) {
  const prompt = createProposalPrompt(company, settings);
  
  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: 'あなたは営業のプロフェッショナルです。企業情報と商材情報を基に、効果的で個別最適化された提案メッセージを作成してください。レスポンスは必ずJSON形式のみで返してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: 1500,
    temperature: 0.7
  };
  
  // API呼び出し
  const apiResult = callOpenAIAPIWithEnhancedLogging(payload);
  
  if (!apiResult.success) {
    return {
      type: 'API_ERROR',
      apiResult: apiResult,
      rawResponse: null,
      parseResult: null
    };
  }
  
  // レスポンス内容の保存
  const rawResponse = apiResult.response;
  
  // 解析の実行
  const parseResult = parseProposalResponseWithLogging(rawResponse);
  
  return {
    type: 'COMPLETE_ANALYSIS',
    apiResult: apiResult,
    rawResponse: rawResponse,
    parseResult: parseResult,
    prompt: prompt.substring(0, 300) + '...'
  };
}

/**
 * レスポンス分析結果の表示
 */
function displayResponseAnalysis(result, companyName) {
  try {
    const sheet = getOrCreateSheetForLogging('レスポンス分析結果');
    
    // 前回の結果をクリア
    sheet.clear();
    
    // ヘッダー
    sheet.getRange(1, 1).setValue('🔬 提案生成レスポンス詳細分析');
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    
    sheet.getRange(2, 1).setValue(`対象企業: ${companyName}`);
    sheet.getRange(3, 1).setValue(`分析時刻: ${new Date()}`);
    
    let row = 5;
    
    // API結果の表示
    sheet.getRange(row, 1).setValue('📡 API呼び出し結果:');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    if (result.apiResult) {
      sheet.getRange(row, 1).setValue(`ステータス: ${result.apiResult.success ? '✅ 成功' : '❌ 失敗'}`);
      row++;
      
      if (result.apiResult.responseInfo) {
        const info = result.apiResult.responseInfo;
        sheet.getRange(row, 1).setValue(`HTTPステータス: ${info.statusCode || 'N/A'}`);
        row++;
        sheet.getRange(row, 1).setValue(`レスポンス時間: ${info.responseTime || 'N/A'}ms`);
        row++;
        sheet.getRange(row, 1).setValue(`使用トークン: ${info.responseTokens || 'N/A'}`);
        row++;
      }
      
      if (!result.apiResult.success && result.apiResult.error) {
        sheet.getRange(row, 1).setValue(`エラー: ${result.apiResult.error.message}`);
        row++;
      }
    }
    
    row++;
    
    // レスポンス内容の表示
    if (result.rawResponse) {
      sheet.getRange(row, 1).setValue('📄 生のレスポンス内容:');
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      // レスポンスを1000文字ずつに分割して表示
      const responseChunks = splitTextIntoChunks(result.rawResponse, 1000);
      for (let i = 0; i < responseChunks.length && i < 5; i++) { // 最大5チャンク
        sheet.getRange(row, 1).setValue(`チャンク${i + 1}:`);
        sheet.getRange(row, 2).setValue(responseChunks[i]);
        row++;
      }
      
      row++;
    }
    
    // 解析結果の表示
    if (result.parseResult) {
      sheet.getRange(row, 1).setValue('🔍 解析結果:');
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      sheet.getRange(row, 1).setValue(`解析状況: ${result.parseResult.success ? '✅ 成功' : '❌ 失敗'}`);
      row++;
      
      if (!result.parseResult.success && result.parseResult.error) {
        sheet.getRange(row, 1).setValue(`エラータイプ: ${result.parseResult.error.type}`);
        row++;
        sheet.getRange(row, 1).setValue(`エラーメッセージ: ${result.parseResult.error.message}`);
        row++;
        
        if (result.parseResult.error.availableFields) {
          sheet.getRange(row, 1).setValue(`利用可能フィールド: ${result.parseResult.error.availableFields.join(', ')}`);
          row++;
        }
        
        if (result.parseResult.error.actualData) {
          sheet.getRange(row, 1).setValue('実際のデータ:');
          sheet.getRange(row, 2).setValue(result.parseResult.error.actualData);
          row++;
        }
      }
      
      if (result.parseResult.success && result.parseResult.proposals) {
        sheet.getRange(row, 1).setValue('✅ 正常に解析されたデータ:');
        row++;
        const proposals = result.parseResult.proposals;
        sheet.getRange(row, 1).setValue(`パターンA件名: ${proposals.patternA.subject}`);
        row++;
        sheet.getRange(row, 1).setValue(`パターンB件名: ${proposals.patternB.subject}`);
        row++;
        sheet.getRange(row, 1).setValue(`問い合わせフォーム: ${proposals.contactForm}`);
        row++;
      }
    }
    
    // 分析シートをアクティブにする
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    
    SpreadsheetApp.getUi().alert(
      'レスポンス分析完了',
      `「${companyName}」のレスポンス分析が完了しました。詳細は「レスポンス分析結果」シートをご確認ください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`分析結果表示エラー: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('表示エラー', `分析結果の表示に失敗しました: ${error.toString()}`);
  }
}

/**
 * テキストを指定文字数で分割
 */
function splitTextIntoChunks(text, chunkSize) {
  const chunks = [];
  for (let i = 0; i < text.length; i += chunkSize) {
    chunks.push(text.substring(i, i + chunkSize));
  }
  return chunks;
}
