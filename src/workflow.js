/**
 * 全自動実行制御機能
 * システム全体のワークフローを統合管理
 */

/**
 * 全自動ワークフローの実行
 */
function executeFullWorkflow() {
  const startTime = new Date();
  let workflowLog = [];
  
  try {
    updateExecutionStatus('全自動実行を開始します...');
    workflowLog.push('全自動実行開始');
    
    const settings = getControlPanelSettings();
    
    // 入力値の検証
    validateWorkflowInputs(settings);
    
    // ステップ1: キーワード生成
    workflowLog.push('ステップ1: キーワード生成開始');
    updateExecutionStatus('ステップ1/4: 戦略的キーワードを生成中...');
    
    const keywords = generateStrategicKeywords();
    workflowLog.push(`キーワード生成完了: ${keywords.length}個`);
    
    // 少し待機（API制限対策）
    Utilities.sleep(2000);
    
    // ステップ2: 企業検索
    workflowLog.push('ステップ2: 企業検索開始');
    updateExecutionStatus('ステップ2/4: 企業を検索・分析中...');
    
    const companiesFound = executeCompanySearch();
    workflowLog.push(`企業検索完了: ${companiesFound}件の企業を発見`);
    
    // 少し待機
    Utilities.sleep(2000);
    
    // ステップ3: 提案メッセージ生成
    workflowLog.push('ステップ3: 提案メッセージ生成開始');
    updateExecutionStatus('ステップ3/4: 提案メッセージを生成中...');
    
    const proposalsGenerated = generatePersonalizedProposals();
    workflowLog.push(`提案生成完了: ${proposalsGenerated}件の提案を作成`);
    
    // ステップ4: 結果サマリー生成
    workflowLog.push('ステップ4: 結果サマリー生成開始');
    updateExecutionStatus('ステップ4/4: 結果をまとめています...');
    
    const summary = generateWorkflowSummary();
    workflowLog.push('結果サマリー生成完了');
    
    // 実行完了
    const processingTime = Math.round((new Date() - startTime) / 1000);
    const completionMessage = `全自動実行完了！ 処理時間: ${processingTime}秒`;
    
    updateExecutionStatus(completionMessage);
    workflowLog.push(completionMessage);
    
    // 実行ログの記録
    logWorkflowExecution(workflowLog, summary, processingTime);
    
    // 完了通知の表示
    showWorkflowCompletion(summary, processingTime);
    
    return summary;
    
  } catch (error) {
    const errorMessage = `全自動実行エラー: ${error.toString()}`;
    updateExecutionStatus(errorMessage);
    workflowLog.push(errorMessage);
    
    logExecution('全自動実行', 'ERROR', 0, 1, error.toString());
    
    // エラー通知の表示
    showWorkflowError(error, workflowLog);
    
    throw error;
  }
}

/**
 * ワークフロー入力値の検証
 */
function validateWorkflowInputs(settings) {
  const errors = [];
  
  if (!settings.productName || settings.productName.trim() === '') {
    errors.push('商材名が入力されていません');
  }
  
  if (!settings.productDescription || settings.productDescription.trim() === '') {
    errors.push('商材概要が入力されていません');
  }
  
  if (!settings.priceRange || settings.priceRange === '') {
    errors.push('価格帯が選択されていません');
  }
  
  if (!settings.targetSize || settings.targetSize === '') {
    errors.push('対象企業規模が選択されていません');
  }
  
  if (settings.maxCompanies < 1 || settings.maxCompanies > 100) {
    errors.push('検索企業数上限は1-100の範囲で設定してください');
  }
  
  // APIキーの存在確認
  if (!API_KEYS.OPENAI) {
    errors.push('OpenAI APIキーが設定されていません');
  }
  
  if (!API_KEYS.GOOGLE_SEARCH || !API_KEYS.GOOGLE_SEARCH_ENGINE_ID) {
    errors.push('Google Search APIキーまたは検索エンジンIDが設定されていません');
  }
  
  if (errors.length > 0) {
    throw new Error(`入力エラー: ${errors.join(', ')}`);
  }
}

/**
 * ワークフロー結果サマリーの生成
 */
function generateWorkflowSummary() {
  const keywordStats = getKeywordStatistics();
  const companyStats = getMatchingStatistics();
  const proposalStats = getProposalStatistics();
  
  return {
    keywords: keywordStats,
    companies: companyStats,
    proposals: proposalStats,
    recommendations: generateRecommendations(keywordStats, companyStats, proposalStats),
    nextActions: generateNextActions(companyStats, proposalStats)
  };
}

/**
 * キーワード統計の取得
 */
function getKeywordStatistics() {
  try {
    const SHEET_NAMES = getSheetNames ? getSheetNames() : {
      CONTROL: '制御パネル',
      KEYWORDS: '生成キーワード',
      COMPANIES: '企業マスター',
      PROPOSALS: '提案メッセージ',
      LOGS: '実行ログ'
    };
    
    const sheet = getSafeSheet(SHEET_NAMES.KEYWORDS);
    if (!sheet) {
      console.error('KEYWORDS sheet not found in getKeywordStatistics');
      return {
        total: 0,
        executed: 0,
        highPriority: 0,
        totalHits: 0
      };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        total: 0,
        executed: 0,
        highPriority: 0,
        totalHits: 0
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    return {
      total: data.length,
      executed: data.filter(row => row[4]).length,
      highPriority: data.filter(row => row[2] === '高').length,
      totalHits: data.reduce((sum, row) => sum + (row[5] || 0), 0),
      categories: getCategoryBreakdown(data)
    };
    
  } catch (error) {
    console.error('Error in getKeywordStatistics:', error);
    return {
      total: 0,
      executed: 0,
      highPriority: 0,
      totalHits: 0
    };
  }
}

/**
 * カテゴリ別内訳の取得
 */
function getCategoryBreakdown(keywordData) {
  const categories = {};
  
  keywordData.forEach(row => {
    const category = row[1];
    if (!categories[category]) {
      categories[category] = { count: 0, executed: 0, hits: 0 };
    }
    categories[category].count++;
    if (row[4]) categories[category].executed++;
    categories[category].hits += row[5] || 0;
  });
  
  return categories;
}

/**
 * 提案統計の取得
 */
function getProposalStatistics() {
  const proposals = getAllProposals();
  
  if (proposals.length === 0) {
    return {
      total: 0,
      patternARecommended: 0,
      patternBRecommended: 0,
      averageBodyLength: 0
    };
  }
  
  const patternACount = proposals.filter(p => p.recommendedPattern === 'A').length;
  const patternBCount = proposals.filter(p => p.recommendedPattern === 'B').length;
  
  const totalBodyLength = proposals.reduce((sum, p) => {
    return sum + (p.patternA.body.length + p.patternB.body.length) / 2;
  }, 0);
  
  return {
    total: proposals.length,
    patternARecommended: patternACount,
    patternBRecommended: patternBCount,
    averageBodyLength: Math.round(totalBodyLength / proposals.length),
    recentGenerated: proposals.filter(p => {
      const genDate = new Date(p.generatedDate);
      const today = new Date();
      return (today - genDate) < 24 * 60 * 60 * 1000; // 24時間以内
    }).length
  };
}

/**
 * 推奨事項の生成
 */
function generateRecommendations(keywordStats, companyStats, proposalStats) {
  const recommendations = [];
  
  // キーワード関連の推奨
  if (keywordStats.executed / keywordStats.total < 0.5) {
    recommendations.push('未実行のキーワードが多数あります。追加の企業検索を実行することをお勧めします。');
  }
  
  // 企業関連の推奨
  if (companyStats.averageScore < 60) {
    recommendations.push('マッチ度スコアが低めです。キーワード戦略の見直しをお勧めします。');
  }
  
  if (companyStats.highScoreCount < 5) {
    recommendations.push('高スコア企業が少数です。検索キーワードを追加するか、条件を調整してください。');
  }
  
  // 提案関連の推奨
  if (proposalStats.total < companyStats.totalCompanies * 0.3) {
    recommendations.push('提案生成率が低いです。より多くの企業に対して提案を生成することをお勧めします。');
  }
  
  // 業界バランスの推奨
  const topIndustry = companyStats.topIndustries[0];
  if (topIndustry && topIndustry.count / companyStats.totalCompanies > 0.7) {
    recommendations.push(`${topIndustry.industry}業界に偏っています。他の業界もターゲットに含めることを検討してください。`);
  }
  
  return recommendations;
}

/**
 * 次のアクションの生成
 */
function generateNextActions(companyStats, proposalStats) {
  const actions = [];
  
  // 高スコア企業への提案
  if (companyStats.highScoreCount > proposalStats.total) {
    const pendingCount = companyStats.highScoreCount - proposalStats.total;
    actions.push(`${pendingCount}件の高スコア企業への提案生成が残っています。`);
  }
  
  // アプローチ実行の推奨
  if (proposalStats.total > 0) {
    actions.push('生成された提案メッセージを使用して、実際のアプローチを開始してください。');
  }
  
  // 追加検索の推奨
  if (companyStats.totalCompanies < 20) {
    actions.push('より多くの企業を発見するため、追加のキーワード検索を実行してください。');
  }
  
  // 結果分析の推奨
  actions.push('アプローチ結果を記録し、成功率の高いパターンを分析してください。');
  
  return actions;
}

/**
 * ワークフロー実行ログの記録
 */
function logWorkflowExecution(workflowLog, summary, processingTime) {
  try {
    const SHEET_NAMES = getSheetNames ? getSheetNames() : {
      CONTROL: '制御パネル',
      KEYWORDS: '生成キーワード',
      COMPANIES: '企業マスター',
      PROPOSALS: '提案メッセージ',
      LOGS: '実行ログ'
    };
    
    const sheet = getSafeSheet(SHEET_NAMES.LOGS);
    if (!sheet) {
      console.error('LOGS sheet not found - cannot log workflow execution');
      return;
    }
    
    const lastRow = sheet.getLastRow();
  const executionId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;
  
  const logData = [
    executionId,
    new Date(),
    '全自動実行',
    '全ワークフロー',
    summary.companies.totalCompanies,
    0, // エラー件数
    workflowLog.join(' | '),
    processingTime,
    summary.keywords.total + summary.companies.totalCompanies + summary.proposals.total // API使用量概算
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, logData.length).setValues([logData]);
  
  } catch (error) {
    console.error('Error in logWorkflowExecution:', error);
  }
}

/**
 * ワークフロー完了通知の表示
 */
function showWorkflowCompletion(summary, processingTime) {
  const message = `
営業自動化システム実行完了！

📊 実行結果サマリー:
・生成キーワード数: ${summary.keywords.total}個
・発見企業数: ${summary.companies.totalCompanies}件
・平均マッチ度: ${summary.companies.averageScore}点
・生成提案数: ${summary.proposals.total}件
・処理時間: ${processingTime}秒

🎯 高スコア企業数: ${summary.companies.highScoreCount}件

次のステップ: 生成された提案メッセージを確認し、アプローチを開始してください。
  `;
  
  updateExecutionStatus(message.replace(/\n/g, ' '));
  Logger.log(message);
}

/**
 * ワークフローエラー通知の表示
 */
function showWorkflowError(error, workflowLog) {
  const message = `
営業自動化システムでエラーが発生しました。

エラー内容: ${error.toString()}

実行ログ:
${workflowLog.join('\n')}

対処方法:
1. APIキーの設定を確認してください
2. インターネット接続を確認してください
3. 入力値を確認してください
4. しばらく時間をおいて再実行してください
  `;
  
  updateExecutionStatus(`エラー発生: ${error.toString()}`);
  Logger.log(message);
}

/**
 * 段階的実行（デバッグ用）
 */
function executeWorkflowStep(stepNumber) {
  try {
    const settings = getControlPanelSettings();
    
    switch (stepNumber) {
      case 1:
        updateExecutionStatus('キーワード生成のみ実行中...');
        return generateStrategicKeywords();
        
      case 2:
        updateExecutionStatus('企業検索のみ実行中...');
        return executeCompanySearch();
        
      case 3:
        updateExecutionStatus('提案生成のみ実行中...');
        return generatePersonalizedProposals();
        
      case 4:
        updateExecutionStatus('サマリー生成のみ実行中...');
        return generateWorkflowSummary();
        
      default:
        throw new Error('無効なステップ番号です（1-4を指定してください）');
    }
    
  } catch (error) {
    updateExecutionStatus(`ステップ${stepNumber}でエラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * ワークフロー状態のリセット
 */
function resetWorkflowState() {
  const sheets = [SHEET_NAMES.KEYWORDS, SHEET_NAMES.COMPANIES, SHEET_NAMES.PROPOSALS];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
    }
  });
  
  updateExecutionStatus('ワークフロー状態をリセットしました');
  Logger.log('全データがクリアされました');
}
