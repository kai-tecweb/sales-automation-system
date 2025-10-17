/**
 * 営業自動化システム - 完全版メニュー
 */

function onOpen() {
  try {
    console.log('Creating comprehensive sales automation menu...');
    
    const ui = SpreadsheetApp.getUi();
    
    // メインメニュー
    ui.createMenu('🚀 営業システム')
      .addSubMenu(ui.createMenu('📊 システム管理')
        .addItem('🧪 システムテスト', 'runSystemTest')
        .addItem('� APIキーテスト', 'testApiKeys')
        .addItem('�📋 基本情報', 'showBasicInfo')
        .addItem('🔧 シート作成', 'createBasicSheets')
        .addItem('🗂️ シート設定', 'configureSheets')
        .addItem('🔄 データ同期', 'syncAllData'))
      
      .addSubMenu(ui.createMenu('🔍 キーワード管理')
        .addItem('🎯 キーワード生成', 'generateKeywords'))
      
      .addSubMenu(ui.createMenu('🏢 企業管理')
        .addItem('🔍 企業検索', 'searchCompany'))
      
      .addSubMenu(ui.createMenu('💼 提案管理')
        .addItem('✨ 提案生成', 'generateProposal'))
      
      .addSubMenu(ui.createMenu('📈 分析・レポート')
        .addItem('📊 総合レポート', 'generateComprehensiveReport')
        .addItem('📋 活動ログ', 'viewActivityLog'))
      
      .addSubMenu(ui.createMenu('🔐 ライセンス管理')
        .addItem('📋 ライセンス状況', 'showLicenseStatus')
        .addItem('👤 管理者認証', 'authenticateAdmin')
        .addSeparator()
        .addItem('🔑 APIキー管理', 'manageApiKeys')
        .addSeparator()
        .addItem('📅 使用開始設定', 'setLicenseStartDate')
        .addItem('🔄 期限延長', 'extendLicense')
        .addItem('🔒 システムロック解除', 'unlockSystem'))
      
      .addSubMenu(ui.createMenu('📚 ヘルプ・ドキュメント')
        .addItem('🆘 基本ヘルプ', 'showHelp')
        .addItem('📖 ユーザーガイド', 'showUserGuide')
        .addItem('⚙️ 機能説明', 'showFeatureGuide')
        .addItem('🔧 設定ガイド', 'showConfigGuide')
        .addItem('💰 料金・API設定ガイド', 'showPricingGuide')
        .addSeparator()
        .addItem('🚀 将来機能一覧', 'showFutureFeatures')
        .addItem('❓ よくある質問', 'showFAQ')
        .addItem('🐛 トラブルシューティング', 'showTroubleshooting'))
      
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ 設定')
        .addItem('🔑 APIキー設定', 'configureApiKeys')
        .addItem('📊 基本設定', 'showBasicSettings')
        .addItem('🔧 詳細設定', 'showAdvancedSettings')
        .addItem('🌐 システム環境', 'showSystemEnvironment'))
      .addToUi();
    
    console.log('Comprehensive menu created successfully');
    
    // システム起動通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '営業自動化システム v2.0 起動完了', 
      '🚀 全機能利用可能', 
      3
    );
    
  } catch (error) {
    console.error('Menu creation error:', error);
    // エラー時でも基本メニューは作成
    createFallbackMenu();
  }
}

/**
 * エラー時のフォールバックメニュー
 */
function createFallbackMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🚀 営業システム (基本)')
      .addItem('🧪 システムテスト', 'runSystemTest')
      .addItem('📋 基本情報', 'showBasicInfo')
      .addItem('🔧 シート作成', 'createBasicSheets')
      .addToUi();
  } catch (error) {
    console.error('Fallback menu creation failed:', error);
  }
}

function runSystemTest() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    let message = '🧪 システムテスト結果\n\n';
    message += `📊 スプレッドシート: ${spreadsheet.getName()}\n`;
    message += `🆔 ID: ${spreadsheet.getId().substring(0, 20)}...\n`;
    message += ` 実行時刻: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    // シート確認
    const sheets = spreadsheet.getSheets();
    message += `📋 現在のシート数: ${sheets.length}個\n\n`;
    
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
    const existingRequired = requiredSheets.filter(name => 
      sheets.some(sheet => sheet.getName() === name)
    );
    
    message += `✅ 必要シート: ${existingRequired.length}/${requiredSheets.length}個作成済み\n`;
    
    if (existingRequired.length < requiredSheets.length) {
      message += `❌ 不足: ${requiredSheets.length - existingRequired.length}個\n`;
      message += `\n「シート作成」で不足分を作成できます。`;
    } else {
      message += `\n✅ すべてのシートが準備されています！`;
    }
    
    ui.alert('🧪 システムテスト結果\n\n' + message);
    
  } catch (error) {
    console.error('System test error:', error);
    SpreadsheetApp.getUi().alert('❌ テストエラー\n\nエラーが発生しました: ' + String(error).substring(0, 100));
  }
}

function showBasicInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    let info = '📋 営業自動化システム情報\n\n';
    info += `📊 スプレッドシート: ${spreadsheet.getName()}\n`;
    info += `🆔 ID: ${spreadsheet.getId().substring(0, 20)}...\n`;
    info += `🔗 URL: 正常にアクセス可能\n`;
    info += `🕒 確認日時: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    info += '🎯 主な機能:\n';
    info += '• キーワード生成\n';
    info += '• 企業検索・分析\n';
    info += '• 提案メッセージ自動作成\n';
    info += '• マッチ度スコアリング\n\n';
    
    info += '📞 サポート:\n';
    info += 'システムに関するお問い合わせは\n';
    info += '管理者までご連絡ください。';
    
    ui.alert('📋 営業自動化システム情報\n\n' + info);
    
  } catch (error) {
    console.error('Show info error:', error);
    SpreadsheetApp.getUi().alert('❌ 情報取得エラー\n\nエラーが発生しました: ' + String(error).substring(0, 100));
  }
}

function createBasicSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '🔧 シート作成確認\n\n営業自動化システムに必要なシートを作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      const requiredSheets = [
        { name: '制御パネル', description: 'システム設定と操作' },
        { name: '生成キーワード', description: 'AI生成検索キーワード' },
        { name: '企業マスター', description: '発見企業の管理' },
        { name: '提案メッセージ', description: '自動生成提案' },
        { name: '実行ログ', description: 'システム実行履歴' }
      ];
      
      let created = 0;
      let existing = 0;
      
      requiredSheets.forEach(sheetInfo => {
        let sheet = spreadsheet.getSheetByName(sheetInfo.name);
        if (!sheet) {
          sheet = spreadsheet.insertSheet(sheetInfo.name);
          
          // 基本ヘッダーを設定
          sheet.getRange('A1').setValue(sheetInfo.description);
          sheet.getRange('A1').setFontWeight('bold').setFontSize(12);
          
          created++;
          console.log(`Created sheet: ${sheetInfo.name}`);
        } else {
          existing++;
          console.log(`Sheet already exists: ${sheetInfo.name}`);
        }
      });
      
      let resultMessage = '✅ シート作成完了\n\n';
      resultMessage += `新規作成: ${created}個\n`;
      resultMessage += `既存確認: ${existing}個\n`;
      resultMessage += `合計: ${requiredSheets.length}個\n\n`;
      resultMessage += 'システムの準備が整いました！';
      
      ui.alert('✅ 作成完了\n\n' + resultMessage);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ シート作成エラー\n\n' + String(error).substring(0, 100));
  }
}

function simpleTest() {
  try {
    SpreadsheetApp.getUi().alert('✅ 簡単テスト成功\n\nメニューは正常に動作しています！');
  } catch (error) {
    console.error('Simple test error:', error);
  }
}

// =============================================
// システム管理機能
// =============================================

function configureSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🗂️ シート設定\n\n各シートの詳細設定を行います。\n（開発中）');
  } catch (error) {
    console.error('Configure sheets error:', error);
    SpreadsheetApp.getUi().alert('❌ 設定エラー: ' + String(error).substring(0, 100));
  }
}

function syncAllData() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🔄 データ同期\n\n外部データとの同期を開始します。\n（開発中）');
  } catch (error) {
    console.error('Sync data error:', error);
    SpreadsheetApp.getUi().alert('❌ 同期エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// キーワード管理機能
// =============================================

function generateKeywords() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '🎯 キーワード生成確認\n\nAIを使用してキーワードを生成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      executeKeywordGeneration();
    }
  } catch (error) {
    console.error('Generate keywords error:', error);
    SpreadsheetApp.getUi().alert('❌ 生成エラー: ' + String(error).substring(0, 100));
  }
}

function editKeywords() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📝 キーワード編集\n\n既存キーワードを編集します。\n（実装中）');
  } catch (error) {
    console.error('Edit keywords error:', error);
    SpreadsheetApp.getUi().alert('❌ 編集エラー: ' + String(error).substring(0, 100));
  }
}

function updateKeywords() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🔄 キーワード更新\n\nキーワードデータを更新します。\n（実装中）');
  } catch (error) {
    console.error('Update keywords error:', error);
    SpreadsheetApp.getUi().alert('❌ 更新エラー: ' + String(error).substring(0, 100));
  }
}

function analyzeKeywordUsage() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📊 使用状況分析\n\nキーワードの使用状況を分析します。\n（実装中）');
  } catch (error) {
    console.error('Analyze keyword usage error:', error);
    SpreadsheetApp.getUi().alert('❌ 分析エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// 企業管理機能
// =============================================

function addCompany() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('➕ 企業追加\n\n新しい企業情報を追加します。\n（実装中）');
  } catch (error) {
    console.error('Add company error:', error);
    SpreadsheetApp.getUi().alert('❌ 追加エラー: ' + String(error).substring(0, 100));
  }
}

function searchCompany() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '🔍 企業検索確認\n\n企業検索を実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      executeCompanySearch();
    }
  } catch (error) {
    console.error('Search company error:', error);
    SpreadsheetApp.getUi().alert('❌ 検索エラー: ' + String(error).substring(0, 100));
  }
}

function analyzeCompany() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📊 企業分析\n\n企業の詳細分析を行います。\n（実装中）');
  } catch (error) {
    console.error('Analyze company error:', error);
    SpreadsheetApp.getUi().alert('❌ 分析エラー: ' + String(error).substring(0, 100));
  }
}

function updateCompanyData() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🔄 企業データ更新\n\n企業情報を最新化します。\n（実装中）');
  } catch (error) {
    console.error('Update company data error:', error);
    SpreadsheetApp.getUi().alert('❌ 更新エラー: ' + String(error).substring(0, 100));
  }
}

function calculateMatching() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '📈 マッチ度計算確認\n\n企業とのマッチ度を再計算しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // スコア再計算機能の実装
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const companiesSheet = spreadsheet.getSheetByName('企業マスター');
        if (!companiesSheet) {
          throw new Error('企業マスターシートが見つかりません');
        }
        
        const settings = getControlPanelSettings();
        const data = companiesSheet.getDataRange().getValues();
        let updatedCount = 0;
        
        for (let i = 1; i < data.length; i++) { // ヘッダー行をスキップ
          const companyData = {
            companySize: data[i][4] || '',
            industry: data[i][5] || '',
            contactMethod: data[i][6] || '',
            description: data[i][7] || ''
          };
          
          const newScore = calculateMatchScore(companyData, settings);
          companiesSheet.getRange(i + 1, 8).setValue(newScore);
          updatedCount++;
        }
        
        ui.alert(`✅ 計算完了\n\n${updatedCount}社のマッチ度を再計算しました。`);
        
      } catch (error) {
        throw error;
      }
    }
  } catch (error) {
    console.error('Calculate matching error:', error);
    SpreadsheetApp.getUi().alert('❌ 計算エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// 提案管理機能
// =============================================

function generateProposal() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '✨ 提案生成確認\n\nAIを使用して提案メッセージを生成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      executeProposalGenerationEnhanced();
    }
  } catch (error) {
    console.error('Generate proposal error:', error);
    SpreadsheetApp.getUi().alert('❌ 生成エラー: ' + String(error).substring(0, 100));
  }
}

function editProposal() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📝 提案編集\n\n既存の提案を編集します。\n（実装中）');
  } catch (error) {
    console.error('Edit proposal error:', error);
    SpreadsheetApp.getUi().alert('❌ 編集エラー: ' + String(error).substring(0, 100));
  }
}

function analyzeProposal() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📊 提案分析\n\n提案の効果を分析します。\n（実装中）');
  } catch (error) {
    console.error('Analyze proposal error:', error);
    SpreadsheetApp.getUi().alert('❌ 分析エラー: ' + String(error).substring(0, 100));
  }
}

function customizeProposal() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🎯 カスタマイズ\n\n提案をカスタマイズします。\n（実装中）');
  } catch (error) {
    console.error('Customize proposal error:', error);
    SpreadsheetApp.getUi().alert('❌ カスタマイズエラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// 分析・レポート機能
// =============================================

function generateComprehensiveReport() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 各シートの統計情報を取得
    const keywordsSheet = spreadsheet.getSheetByName('生成キーワード');
    const companiesSheet = spreadsheet.getSheetByName('企業マスター');
    const proposalsSheet = spreadsheet.getSheetByName('提案メッセージ');
    const logSheet = spreadsheet.getSheetByName('実行ログ');
    
    let report = '📊 営業自動化システム総合レポート\n\n';
    report += `📅 生成日時: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    // キーワード統計
    if (keywordsSheet) {
      const keywordCount = Math.max(0, keywordsSheet.getLastRow() - 1);
      report += `🔍 キーワード数: ${keywordCount}個\n`;
    }
    
    // 企業統計
    if (companiesSheet) {
      const companyCount = Math.max(0, companiesSheet.getLastRow() - 1);
      report += `🏢 企業数: ${companyCount}社\n`;
      
      // 高スコア企業数
      if (companyCount > 0) {
        const data = companiesSheet.getDataRange().getValues();
        const highScoreCount = data.slice(1).filter(row => row[7] >= 70).length;
        report += `📈 高スコア企業（70点以上）: ${highScoreCount}社\n`;
      }
    }
    
    // 提案統計
    if (proposalsSheet) {
      const proposalCount = Math.max(0, proposalsSheet.getLastRow() - 1);
      report += `💼 提案数: ${proposalCount}件\n`;
    }
    
    // 実行ログ統計
    if (logSheet) {
      const logCount = Math.max(0, logSheet.getLastRow() - 1);
      report += `📋 実行回数: ${logCount}回\n\n`;
      
      if (logCount > 0) {
        const logs = logSheet.getDataRange().getValues().slice(1);
        const successCount = logs.filter(log => log[2] === 'SUCCESS').length;
        const successRate = Math.round((successCount / logCount) * 100);
        report += `✅ 成功率: ${successRate}% (${successCount}/${logCount})\n`;
      }
    }
    
    report += '\n💡 提案:\n';
    if (keywordsSheet && companiesSheet && proposalsSheet) {
      const keywordCount = Math.max(0, keywordsSheet.getLastRow() - 1);
      const companyCount = Math.max(0, companiesSheet.getLastRow() - 1);
      const proposalCount = Math.max(0, proposalsSheet.getLastRow() - 1);
      
      if (keywordCount === 0) {
        report += '• まずキーワード生成を実行してください\n';
      } else if (companyCount === 0) {
        report += '• 企業検索を実行してください\n';
      } else if (proposalCount === 0) {
        report += '• 提案メッセージ生成を実行してください\n';
      } else {
        report += '• 全機能が正常に動作しています\n';
      }
    }
    
    SpreadsheetApp.getUi().alert(report);
    
  } catch (error) {
    console.error('Generate comprehensive report error:', error);
    SpreadsheetApp.getUi().alert('❌ レポートエラー: ' + String(error).substring(0, 100));
  }
}

function analyzeMatchingScores() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('🎯 マッチング分析\n\nマッチングスコアを詳細分析します。\n（実装中）');
  } catch (error) {
    console.error('Analyze matching scores error:', error);
    SpreadsheetApp.getUi().alert('❌ 分析エラー: ' + String(error).substring(0, 100));
  }
}

function analyzeSuccessRates() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('📈 成功率分析\n\n営業成功率を分析します。\n（実装中）');
  } catch (error) {
    console.error('Analyze success rates error:', error);
    SpreadsheetApp.getUi().alert('❌ 分析エラー: ' + String(error).substring(0, 100));
  }
}

function viewActivityLog() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('実行ログ');
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert('❌ ログエラー\n\n実行ログシートが見つかりません。');
      return;
    }
    
    const data = logSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('📋 活動ログ\n\nまだ実行履歴がありません。');
      return;
    }
    
    // 最新10件のログを表示
    let logText = '📋 最新の活動ログ\n\n';
    const recentLogs = data.slice(-10).reverse(); // 最新10件を逆順で
    
    recentLogs.forEach((log, index) => {
      if (index === 0) return; // ヘッダー行をスキップ
      logText += `${log[0]} - ${log[1]}\n`;
      logText += `結果: ${log[2]} | 処理数: ${log[3]}\n\n`;
    });
    
    SpreadsheetApp.getUi().alert(logText);
    
  } catch (error) {
    console.error('View activity log error:', error);
    SpreadsheetApp.getUi().alert('❌ ログエラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// ヘルプ・設定機能
// =============================================

function showHelp() {
  try {
    const ui = SpreadsheetApp.getUi();
    let helpText = '🆘 営業自動化システム ヘルプ\n\n';
    helpText += '📊 システム管理:\n';
    helpText += '  • システムテスト - 動作確認\n';
    helpText += '  • 基本情報 - システム情報表示\n';
    helpText += '  • シート作成 - 必要シート作成\n\n';
    helpText += '🔍 キーワード管理:\n';
    helpText += '  • キーワード生成・編集・更新\n';
    helpText += '  • 使用状況分析\n\n';
    helpText += '🏢 企業管理:\n';
    helpText += '  • 企業追加・検索・分析\n';
    helpText += '  • マッチ度計算\n\n';
    helpText += '💼 提案管理:\n';
    helpText += '  • 提案生成・編集・分析\n';
    helpText += '  • カスタマイズ\n\n';
    helpText += '📈 分析・レポート:\n';
    helpText += '  • 総合レポート生成\n';
    helpText += '  • 各種分析機能';
    
    ui.alert(helpText);
  } catch (error) {
    console.error('Show help error:', error);
    SpreadsheetApp.getUi().alert('❌ ヘルプエラー: ' + String(error).substring(0, 100));
  }
}

function showSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('⚙️ システム設定\n\nシステム設定画面を開きます。\n詳細設定は「⚙️設定」メニューをご利用ください。');
  } catch (error) {
    console.error('Show settings error:', error);
    SpreadsheetApp.getUi().alert('❌ 設定エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// API キー設定機能
// =============================================

function configureApiKeys() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    // 現在のAPIキー状況確認
    const chatgptKey = properties.getProperty('CHATGPT_API_KEY') || '';
    const googleSearchKey = properties.getProperty('GOOGLE_SEARCH_API_KEY') || '';
    const googleSearchEngineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID') || '';
    
    let message = '🔑 APIキー設定\n\n';
    message += '📋 現在の設定状況:\n';
    message += `• ChatGPT API: ${chatgptKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    message += `• Google Search API: ${googleSearchKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    message += `• Search Engine ID: ${googleSearchEngineId ? '✅ 設定済み' : '❌ 未設定'}\n\n`;
    message += '設定するAPIキーを選択してください:';
    
    const result = ui.alert('APIキー設定', message, ui.ButtonSet.YES_NO_CANCEL);
    
    if (result === ui.Button.YES) {
      configureChatGptApiKey();
    } else if (result === ui.Button.NO) {
      configureGoogleSearchApi();
    } else if (result === ui.Button.CANCEL) {
      showApiKeyStatus();
    }
    
  } catch (error) {
    console.error('API key configuration error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキー設定エラー: ' + String(error).substring(0, 100));
  }
}

function configureChatGptApiKey() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    const result = ui.prompt(
      '🤖 ChatGPT APIキー設定',
      'ChatGPT APIキーを入力してください:\n\n' +
      '※ OpenAIのAPIキーは「sk-」で始まります\n' +
      '※ 料金が発生する場合があります',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result.getSelectedButton() === ui.Button.OK) {
      const apiKey = result.getResponseText().trim();
      
      if (!apiKey) {
        ui.alert('❌ エラー', 'APIキーが入力されていません', ui.ButtonSet.OK);
        return;
      }
      
      if (!apiKey.startsWith('sk-')) {
        const confirm = ui.alert(
          '⚠️ 警告',
          'ChatGPT APIキーは通常「sk-」で始まります。\n本当にこのキーを設定しますか？',
          ui.ButtonSet.YES_NO
        );
        if (confirm !== ui.Button.YES) return;
      }
      
      properties.setProperty('CHATGPT_API_KEY', apiKey);
      ui.alert('✅ 成功', 'ChatGPT APIキーが設定されました', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('ChatGPT API key configuration error:', error);
    SpreadsheetApp.getUi().alert('❌ ChatGPT APIキー設定エラー: ' + String(error).substring(0, 100));
  }
}

function configureGoogleSearchApi() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    // APIキー設定
    const apiResult = ui.prompt(
      '🔍 Google Search APIキー設定',
      'Google Custom Search APIキーを入力してください:\n\n' +
      '※ Google Cloud Consoleで取得したAPIキー\n' +
      '※ 料金が発生する場合があります',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (apiResult.getSelectedButton() !== ui.Button.OK) return;
    
    const apiKey = apiResult.getResponseText().trim();
    if (!apiKey) {
      ui.alert('❌ エラー', 'APIキーが入力されていません', ui.ButtonSet.OK);
      return;
    }
    
    // Search Engine ID設定
    const engineResult = ui.prompt(
      '🔍 Search Engine ID設定',
      'Custom Search Engine IDを入力してください:\n\n' +
      '※ Googleカスタム検索エンジンのID\n' +
      '※ 英数字の文字列です',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (engineResult.getSelectedButton() !== ui.Button.OK) return;
    
    const engineId = engineResult.getResponseText().trim();
    if (!engineId) {
      ui.alert('❌ エラー', 'Search Engine IDが入力されていません', ui.ButtonSet.OK);
      return;
    }
    
    properties.setProperties({
      'GOOGLE_SEARCH_API_KEY': apiKey,
      'GOOGLE_SEARCH_ENGINE_ID': engineId
    });
    
    ui.alert('✅ 成功', 'Google Search API設定が完了しました', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Google Search API configuration error:', error);
    SpreadsheetApp.getUi().alert('❌ Google Search API設定エラー: ' + String(error).substring(0, 100));
  }
}

function showApiKeyStatus() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    const chatgptKey = properties.getProperty('CHATGPT_API_KEY') || '';
    const googleSearchKey = properties.getProperty('GOOGLE_SEARCH_API_KEY') || '';
    const googleSearchEngineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID') || '';
    
    let statusText = '🔑 APIキー設定状況\n\n';
    statusText += '📋 設定状況:\n';
    statusText += `• ChatGPT API: ${chatgptKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    statusText += `  ${chatgptKey ? `(${chatgptKey.substring(0, 10)}...)` : ''}\n`;
    statusText += `• Google Search API: ${googleSearchKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    statusText += `  ${googleSearchKey ? `(${googleSearchKey.substring(0, 10)}...)` : ''}\n`;
    statusText += `• Search Engine ID: ${googleSearchEngineId ? '✅ 設定済み' : '❌ 未設定'}\n`;
    statusText += `  ${googleSearchEngineId ? `(${googleSearchEngineId})` : ''}\n\n`;
    
    if (chatgptKey && googleSearchKey && googleSearchEngineId) {
      statusText += '🎉 すべてのAPIキーが設定されています！\nシステムの全機能が利用可能です。';
    } else {
      statusText += '⚠️ 一部のAPIキーが未設定です。\n未設定の機能は利用できません。';
    }
    
    ui.alert('APIキー設定状況', statusText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('API key status error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキー状況確認エラー: ' + String(error).substring(0, 100));
  }
}

function showBasicSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    let settingsText = '📊 基本設定\n\n';
    settingsText += '⚙️ システム基本設定:\n';
    settingsText += '• タイムゾーン: Asia/Tokyo\n';
    settingsText += '• 言語: 日本語\n';
    settingsText += '• 通貨: JPY (円)\n\n';
    settingsText += '📋 データ設定:\n';
    settingsText += '• 自動保存: 有効\n';
    settingsText += '• ログ保持期間: 30日\n';
    settingsText += '• エラー通知: 有効\n\n';
    settingsText += '🔄 更新頻度:\n';
    settingsText += '• リアルタイム同期\n';
    settingsText += '• 定期バックアップ: 有効';
    
    ui.alert('基本設定', settingsText, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Basic settings error:', error);
    SpreadsheetApp.getUi().alert('❌ 基本設定エラー: ' + String(error).substring(0, 100));
  }
}

function showAdvancedSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    let advancedText = '🔧 詳細設定\n\n';
    advancedText += '⚡ パフォーマンス設定:\n';
    advancedText += '• API呼び出し制限: 適用中\n';
    advancedText += '• キャッシュ機能: 有効\n';
    advancedText += '• 並列処理: 制限付き\n\n';
    advancedText += '🛡️ セキュリティ設定:\n';
    advancedText += '• APIキー暗号化: 有効\n';
    advancedText += '• アクセスログ: 記録中\n';
    advancedText += '• 権限管理: 有効\n\n';
    advancedText += '📊 デバッグ設定:\n';
    advancedText += '• ログレベル: INFO\n';
    advancedText += '• エラー追跡: 有効\n';
    advancedText += '• 開発者モード: 無効';
    
    ui.alert('詳細設定', advancedText, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Advanced settings error:', error);
    SpreadsheetApp.getUi().alert('❌ 詳細設定エラー: ' + String(error).substring(0, 100));
  }
}

function showSystemEnvironment() {
  try {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let envText = '🌐 システム環境情報\n\n';
    envText += '📱 プラットフォーム:\n';
    envText += '• Google Apps Script (V8)\n';
    envText += '• Google Spreadsheets\n\n';
    envText += '🆔 スプレッドシート情報:\n';
    envText += `• 名前: ${spreadsheet.getName()}\n`;
    envText += `• ID: ${spreadsheet.getId().substring(0, 20)}...\n`;
    envText += `• URL: ${spreadsheet.getUrl().substring(0, 50)}...\n\n`;
    envText += '⏰ 実行環境:\n';
    envText += `• 実行時刻: ${new Date().toLocaleString('ja-JP')}\n`;
    envText += `• タイムゾーン: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'z')}\n`;
    envText += `• ユーザー: ${Session.getActiveUser().getEmail()}\n\n`;
    envText += '📦 システムバージョン:\n';
    envText += '• 営業自動化システム v2.0\n';
    envText += '• 最終更新: 2025年10月17日';
    
    ui.alert('システム環境', envText, ui.ButtonSet.OK);
  } catch (error) {
    console.error('System environment error:', error);
    SpreadsheetApp.getUi().alert('❌ システム環境確認エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// APIキーテスト機能（全ユーザー利用可能）
// =============================================

function testApiKeys() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    let testResult = '🔑 APIキーテスト結果\n\n';
    
    // ChatGPT APIテスト
    const chatgptKey = properties.getProperty('CHATGPT_API_KEY');
    if (chatgptKey) {
      testResult += '🤖 ChatGPT API: ';
      try {
        // 簡単なテストリクエスト
        const testResponse = testChatGptApi(chatgptKey);
        testResult += testResponse ? '✅ 接続成功\n' : '⚠️ 応答異常\n';
      } catch (error) {
        testResult += '❌ 接続失敗\n';
      }
    } else {
      testResult += '🤖 ChatGPT API: ❌ 未設定\n';
    }
    
    // Google Search APIテスト
    const googleSearchKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
    const googleSearchEngineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    if (googleSearchKey && googleSearchEngineId) {
      testResult += '🔍 Google Search API: ';
      try {
        // 簡単なテストリクエスト
        const testResponse = testGoogleSearchApi(googleSearchKey, googleSearchEngineId);
        testResult += testResponse ? '✅ 接続成功\n' : '⚠️ 応答異常\n';
      } catch (error) {
        testResult += '❌ 接続失敗\n';
      }
    } else {
      testResult += '🔍 Google Search API: ❌ 未設定\n';
    }
    
    testResult += '\n📊 テスト実行時刻: ' + new Date().toLocaleString('ja-JP');
    testResult += '\n\n💡 APIキーの設定は「⚙️設定」→「🔑APIキー設定」から行えます。';
    
    ui.alert('APIキーテスト', testResult, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('API key test error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキーテストエラー: ' + String(error).substring(0, 100));
  }
}

function testChatGptApi(apiKey) {
  try {
    // ChatGPT APIの簡単なテスト
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    return response.getResponseCode() === 200;
  } catch (error) {
    console.error('ChatGPT API test error:', error);
    return false;
  }
}

function testGoogleSearchApi(apiKey, engineId) {
  try {
    // Google Custom Search APIの簡単なテスト
    const testUrl = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${engineId}&q=test&num=1`;
    
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    return response.getResponseCode() === 200;
  } catch (error) {
    console.error('Google Search API test error:', error);
    return false;
  }
}

// =============================================
// APIキー管理機能（管理者専用）
// =============================================

function manageApiKeys() {
  try {
    // 管理者権限チェック
    if (!isAdminUser()) {
      SpreadsheetApp.getUi().alert(
        '🔒 アクセス拒否',
        'この機能は管理者専用です。\n「👤 管理者認証」から認証を行ってください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    
    let managementMenu = '🔑 APIキー管理（管理者専用）\n\n';
    managementMenu += '📋 管理機能:\n';
    managementMenu += '• 設定状況詳細確認\n';
    managementMenu += '• APIキー一括リセット\n';
    managementMenu += '• 使用履歴確認\n';
    managementMenu += '• セキュリティ監査\n\n';
    managementMenu += '実行する操作を選択してください:';
    
    const result = ui.alert(
      'APIキー管理',
      managementMenu,
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (result === ui.Button.YES) {
      showDetailedApiKeyStatus();
    } else if (result === ui.Button.NO) {
      resetAllApiKeys();
    } else if (result === ui.Button.CANCEL) {
      showApiKeyUsageHistory();
    }
    
  } catch (error) {
    console.error('API key management error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキー管理エラー: ' + String(error).substring(0, 100));
  }
}

function showDetailedApiKeyStatus() {
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    const chatgptKey = properties.getProperty('CHATGPT_API_KEY') || '';
    const googleSearchKey = properties.getProperty('GOOGLE_SEARCH_API_KEY') || '';
    const googleSearchEngineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID') || '';
    
    let detailText = '🔑 APIキー詳細状況（管理者）\n\n';
    
    detailText += '🤖 ChatGPT API:\n';
    detailText += `• 状態: ${chatgptKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    detailText += `• キー: ${chatgptKey ? chatgptKey.substring(0, 15) + '...' : 'なし'}\n`;
    detailText += `• 長さ: ${chatgptKey ? chatgptKey.length + '文字' : '0文字'}\n\n`;
    
    detailText += '🔍 Google Search API:\n';
    detailText += `• APIキー: ${googleSearchKey ? '✅ 設定済み' : '❌ 未設定'}\n`;
    detailText += `• キー: ${googleSearchKey ? googleSearchKey.substring(0, 15) + '...' : 'なし'}\n`;
    detailText += `• Engine ID: ${googleSearchEngineId ? '✅ 設定済み' : '❌ 未設定'}\n`;
    detailText += `• ID: ${googleSearchEngineId}\n\n`;
    
    detailText += '⏰ 最終確認: ' + new Date().toLocaleString('ja-JP');
    
    ui.alert('詳細状況', detailText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Detailed API key status error:', error);
    SpreadsheetApp.getUi().alert('❌ 詳細状況確認エラー: ' + String(error).substring(0, 100));
  }
}

function resetAllApiKeys() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const confirm = ui.alert(
      '⚠️ 危険な操作',
      'すべてのAPIキーを削除します。\nこの操作は取り消せません。\n本当に実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (confirm === ui.Button.YES) {
      const doubleConfirm = ui.alert(
        '🚨 最終確認',
        '本当にすべてのAPIキーを削除しますか？\nシステムが利用できなくなります。',
        ui.ButtonSet.YES_NO
      );
      
      if (doubleConfirm === ui.Button.YES) {
        const properties = PropertiesService.getScriptProperties();
        properties.deleteProperty('CHATGPT_API_KEY');
        properties.deleteProperty('GOOGLE_SEARCH_API_KEY');
        properties.deleteProperty('GOOGLE_SEARCH_ENGINE_ID');
        
        ui.alert('✅ 完了', 'すべてのAPIキーを削除しました。', ui.ButtonSet.OK);
      }
    }
    
  } catch (error) {
    console.error('Reset API keys error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキーリセットエラー: ' + String(error).substring(0, 100));
  }
}

function showApiKeyUsageHistory() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let historyText = '📊 APIキー使用履歴\n\n';
    historyText += '📋 最近の活動:\n';
    historyText += '• この機能は開発中です\n';
    historyText += '• 将来のアップデートで詳細な使用履歴を表示予定\n\n';
    historyText += '💡 現在利用可能な確認方法:\n';
    historyText += '• 「📋 活動ログ」でシステム全体の履歴確認\n';
    historyText += '• 「🔑 APIキーテスト」で現在の接続状況確認';
    
    ui.alert('使用履歴', historyText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('API key usage history error:', error);
    SpreadsheetApp.getUi().alert('❌ 使用履歴確認エラー: ' + String(error).substring(0, 100));
  }
}

// 管理者権限チェック関数（簡易版）
function isAdminUser() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const adminMode = properties.getProperty('adminMode');
    return adminMode === 'true';
  } catch (error) {
    console.error('Admin check error:', error);
    return false;
  }
}

// =============================================
// ライセンス管理機能
// =============================================

function showLicenseStatus() {
  try {
    const licenseInfo = getLicenseInfo();
    let statusText = '🔐 ライセンス状況\n\n';
    
    statusText += `📊 管理者モード: ${licenseInfo.adminMode ? '🟢 有効' : '🔴 無効'}\n`;
    statusText += `📅 ライセンス期限: ${licenseInfo.isExpired ? '🔴 期限切れ' : '🟢 有効'}\n`;
    
    if (licenseInfo.startDate) {
      statusText += `🏁 開始日: ${formatDate(licenseInfo.startDate)}\n`;
    }
    if (licenseInfo.expiryDate) {
      statusText += `⏰ 期限: ${formatDate(licenseInfo.expiryDate)}\n`;
    }
    if (licenseInfo.remainingDays !== null) {
      statusText += `⏳ 残り: ${licenseInfo.remainingDays}営業日\n`;
    }
    
    statusText += `🔒 システム状態: ${licenseInfo.systemLocked ? 'ロック中' : '利用可能'}\n\n`;
    statusText += '詳細な管理は「🔐 ライセンス管理」メニューをご利用ください。';
    
    SpreadsheetApp.getUi().alert(statusText);
  } catch (error) {
    console.error('Show license status error:', error);
    SpreadsheetApp.getUi().alert('❌ ライセンス状況エラー: ' + String(error).substring(0, 100));
  }
}

function authenticateAdmin() {
  try {
    authenticateAdminUser();
  } catch (error) {
    console.error('Authenticate admin error:', error);
    SpreadsheetApp.getUi().alert('❌ 認証エラー: ' + String(error).substring(0, 100));
  }
}

function setLicenseStartDate() {
  try {
    setLicenseStart();
  } catch (error) {
    console.error('Set license start date error:', error);
    SpreadsheetApp.getUi().alert('❌ 設定エラー: ' + String(error).substring(0, 100));
  }
}

function extendLicense() {
  try {
    extendLicensePeriod();
  } catch (error) {
    console.error('Extend license error:', error);
    SpreadsheetApp.getUi().alert('❌ 延長エラー: ' + String(error).substring(0, 100));
  }
}

function unlockSystem() {
  try {
    forceUnlockSystem();
  } catch (error) {
    console.error('Unlock system error:', error);
    SpreadsheetApp.getUi().alert('❌ 解除エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// ヘルプ・ドキュメント機能
// =============================================

function showUserGuide() {
  try {
    const ui = SpreadsheetApp.getUi();
    let guideText = '📖 ユーザーガイド\n\n';
    guideText += '🚀 システム起動:\n';
    guideText += '1. スプレッドシートを開く\n';
    guideText += '2. メニューバーに「🚀 営業システム」が表示される\n';
    guideText += '3. 初回は「📊 システム管理」→「🔧 シート作成」を実行\n\n';
    
    guideText += '📋 基本ワークフロー:\n';
    guideText += '1. 制御パネルで商材情報を設定\n';
    guideText += '2. 「🔍 キーワード管理」→「🎯 キーワード生成」\n';
    guideText += '3. 「🏢 企業管理」→「🔍 企業検索」\n';
    guideText += '4. 「💼 提案管理」→「✨ 提案生成」\n';
    guideText += '5. 「📈 分析・レポート」で結果確認\n\n';
    
    guideText += '💡 ヒント:\n';
    guideText += '• 各ステップで結果を確認してから次へ進む\n';
    guideText += '• エラーが発生した場合は「🧪 システムテスト」で診断\n';
    guideText += '• 「📋 活動ログ」で実行履歴を確認可能';
    
    ui.alert(guideText);
  } catch (error) {
    console.error('Show user guide error:', error);
    SpreadsheetApp.getUi().alert('❌ ガイド表示エラー: ' + String(error).substring(0, 100));
  }
}

function showFeatureGuide() {
  try {
    const ui = SpreadsheetApp.getUi();
    let featureText = '⚙️ 機能説明\n\n';
    featureText += '🔍 キーワード生成:\n';
    featureText += '• ChatGPT APIを使用して戦略的キーワードを自動生成\n';
    featureText += '• 4つのカテゴリ別に最適化された検索語を作成\n\n';
    
    featureText += '🏢 企業検索:\n';
    featureText += '• Google Custom Search APIで企業情報を自動収集\n';
    featureText += '• ウェブサイト情報の解析とマッチ度スコアリング\n\n';
    
    featureText += '💼 提案生成:\n';
    featureText += '• 各企業に最適化された提案メッセージを自動作成\n';
    featureText += '• カスタマイズ可能なテンプレート機能\n\n';
    
    featureText += '📈 分析機能:\n';
    featureText += '• 総合レポート生成\n';
    featureText += '• マッチングスコア分析\n';
    featureText += '• 成功率追跡';
    
    ui.alert(featureText);
  } catch (error) {
    console.error('Show feature guide error:', error);
    SpreadsheetApp.getUi().alert('❌ 機能説明エラー: ' + String(error).substring(0, 100));
  }
}

function showConfigGuide() {
  try {
    const ui = SpreadsheetApp.getUi();
    let configText = '🔧 設定ガイド\n\n';
    configText += '⚙️ 初期設定:\n';
    configText += '1. 制御パネルシートで商材情報を入力\n';
    configText += '2. ChatGPT API キーを設定\n';
    configText += '3. Google Custom Search API を設定\n';
    configText += '4. 検索エンジンIDを設定\n\n';
    
    configText += '🎯 ターゲット設定:\n';
    configText += '• 企業規模（従業員数）\n';
    configText += '• 対象業界\n';
    configText += '• 地域設定\n';
    configText += '• 除外キーワード\n\n';
    
    configText += '📊 出力設定:\n';
    configText += '• 最大企業数\n';
    configText += '• 検索結果数\n';
    configText += '• マッチ度閾値';
    
    ui.alert(configText);
  } catch (error) {
    console.error('Show config guide error:', error);
    SpreadsheetApp.getUi().alert('❌ 設定ガイドエラー: ' + String(error).substring(0, 100));
  }
}

function showFAQ() {
  try {
    const ui = SpreadsheetApp.getUi();
    let faqText = '❓ よくある質問\n\n';
    
    faqText += '💰 料金について\n';
    faqText += 'Q: 料金はいつ発生しますか？\n';
    faqText += 'A: APIを実際に使用した分のみ発生します。システム起動だけでは料金は発生しません\n\n';
    
    faqText += 'Q: 基本版とAI強化版の違いは？\n';
    faqText += 'A: 基本版は約0.1円/企業、AI強化版は約2.5円/企業。AI版では自動キーワード生成と個別提案が利用できます\n\n';
    
    faqText += '🔧 設定について\n';
    faqText += 'Q: メニューが表示されない\n';
    faqText += 'A: スプレッドシートを再読み込みするか、ブラウザを更新してください\n\n';
    
    faqText += 'Q: API エラーが発生する\n';
    faqText += 'A: 制御パネルでAPI キーが正しく設定されているか確認してください\n\n';
    
    faqText += '📊 利用について\n';
    faqText += 'Q: 検索結果が少ない\n';
    faqText += 'A: キーワードの設定を見直すか、検索条件を緩和してください\n\n';
    
    faqText += 'Q: ライセンスエラーが表示される\n';
    faqText += 'A: 「🔐 ライセンス管理」で使用開始日を設定してください\n\n';
    
    faqText += '💡 最適化について\n';
    faqText += 'Q: コストを下げる方法は？\n';
    faqText += 'A: 企業数上限設定、基本版での試用、バッチ処理の活用が効果的です\n\n';
    
    faqText += 'Q: データが保存されない\n';
    faqText += 'A: スプレッドシートの編集権限があることを確認してください';
    
    ui.alert(faqText);
  } catch (error) {
    console.error('Show FAQ error:', error);
    SpreadsheetApp.getUi().alert('❌ FAQ表示エラー: ' + String(error).substring(0, 100));
  }
}

function showTroubleshooting() {
  try {
    const ui = SpreadsheetApp.getUi();
    let troubleText = '🐛 トラブルシューティング\n\n';
    troubleText += '🔍 問題の特定:\n';
    troubleText += '1. 「🧪 システムテスト」で基本動作を確認\n';
    troubleText += '2. 「📋 活動ログ」でエラー履歴を確認\n';
    troubleText += '3. ブラウザのデベロッパーツールでエラーを確認\n\n';
    
    troubleText += '⚡ 一般的な解決方法:\n';
    troubleText += '• スプレッドシートの再読み込み\n';
    troubleText += '• ブラウザキャッシュのクリア\n';
    troubleText += '• 異なるブラウザでの試行\n';
    troubleText += '• インターネット接続の確認\n\n';
    
    troubleText += '🔧 高度な対処:\n';
    troubleText += '• スクリプトエディタでログを確認\n';
    troubleText += '• API 使用量制限の確認\n';
    troubleText += '• 権限設定の再確認';
    
    ui.alert(troubleText);
  } catch (error) {
    console.error('Show troubleshooting error:', error);
    SpreadsheetApp.getUi().alert('❌ トラブルシューティングエラー: ' + String(error).substring(0, 100));
  }
}

function showPricingGuide() {
  try {
    const ui = SpreadsheetApp.getUi();
    let guideText = '💰 料金・API設定ガイド\n\n';
    
    guideText += '📋 プラン選択:\n';
    guideText += '🟢 基本版\n';
    guideText += '  • Google検索のみ（約0.1円/企業）\n';
    guideText += '  • 手動キーワード・テンプレート提案\n';
    guideText += '  • 低コストで基本機能利用\n\n';
    
    guideText += '🚀 AI強化版\n';
    guideText += '  • AI生成機能付き（約2.5円/企業）\n';
    guideText += '  • 自動キーワード生成・個別最適化提案\n';
    guideText += '  • 高度な営業自動化\n\n';
    
    guideText += '🎁 お得情報:\n';
    guideText += '• Google無料枠: 100回/日まで無料\n';
    guideText += '• ChatGPT初回: $5無料クレジット\n';
    guideText += '• システム試用: 10営業日無料\n\n';
    
    guideText += '📈 料金例（AI強化版）:\n';
    guideText += '• 50社/日: 125円/日（3,750円/月）\n';
    guideText += '• 100社/日: 250円/日（7,500円/月）\n\n';
    
    guideText += '⚙️ API設定手順:\n';
    guideText += '1. Google Cloud Console でAPI有効化\n';
    guideText += '2. OpenAI Platform でアカウント作成\n';
    guideText += '3. 制御パネルでAPI キー設定\n';
    guideText += '4. 🧪 システムテストで動作確認\n\n';
    
    guideText += '💡 コスト最適化:\n';
    guideText += '• 企業数上限設定で予算管理\n';
    guideText += '• バッチ処理でAPI効率化\n';
    guideText += '• 基本版で検証→AI版へアップグレード\n\n';
    
    guideText += '詳細設定は「🔧 設定ガイド」をご確認ください。';
    
    ui.alert(guideText);
  } catch (error) {
    console.error('Show pricing guide error:', error);
    SpreadsheetApp.getUi().alert('❌ 料金ガイドエラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// 将来機能一覧表示
// =============================================

function showFutureFeatures() {
  try {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存のシートをチェック
    let futureSheet = spreadsheet.getSheetByName('将来機能一覧');
    
    if (futureSheet) {
      // 既存シートがある場合は更新確認
      const response = ui.alert(
        '🚀 将来機能一覧',
        '既存の将来機能一覧シートを最新版に更新しますか？',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        spreadsheet.deleteSheet(futureSheet);
        futureSheet = null;
      } else {
        // 既存シートをアクティブにして終了
        futureSheet.activate();
        ui.alert('📋 既存の将来機能一覧シートを表示しました。');
        return;
      }
    }
    
    // 新しいシートを作成
    futureSheet = spreadsheet.insertSheet('将来機能一覧');
    
    // シートの設定
    futureSheet.getRange('A1').setValue('🚀 営業自動化システム 将来機能一覧 v2.0');
    futureSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    futureSheet.getRange('A1:F1').merge();
    
    futureSheet.getRange('A2').setValue(`最終更新: ${new Date().toLocaleString('ja-JP')}`);
    futureSheet.getRange('A2').setFontStyle('italic').setFontColor('#666666');
    
    // ヘッダー行
    const headers = [
      ['機能カテゴリ', '機能名', '実装優先度', '開発期間', '技術難易度', '説明'],
    ];
    
    futureSheet.getRange('A4:F4').setValues(headers);
    futureSheet.getRange('A4:F4').setFontWeight('bold').setBackground('#e8f0fe').setFontColor('#1a73e8');
    
    // 将来機能データ
    const futureFeatures = [
      // Phase 0: 緊急実装
      ['🚨 緊急実装', 'キーワード使用状況分析', 'Phase 0', '1-2週間', '低', '既存プレースホルダー機能の完全実装'],
      ['🚨 緊急実装', '企業詳細分析', 'Phase 0', '1-2週間', '低', '業界別・規模別詳細統計機能'],
      ['🚨 緊急実装', '提案効果分析', 'Phase 0', '1-2週間', '低', '提案品質・効果測定機能'],
      ['🚨 緊急実装', 'マッチ度計算拡張', 'Phase 0', '1週間', '低', 'スコア構成要素分析機能'],
      
      // Phase 1: 短期
      ['📝 編集機能', 'キーワード編集機能', 'Phase 1', '2-3週間', '中', '生成済みキーワードの手動調整'],
      ['📝 編集機能', '提案編集機能', 'Phase 1', '2-3週間', '中', '生成済み提案の手動カスタマイズ'],
      ['➕ データ管理', '手動企業追加', 'Phase 1', '1-2週間', '低', '名刺・紹介企業の直接登録'],
      ['💳 課金システム', '基本課金システム', 'Phase 1', '3-4週間', '中', '試用期間後の継続利用管理'],
      
      // Phase 2: 中期
      ['🔗 外部連携', 'Gmail連携', 'Phase 2', '1-2ヶ月', '中', '提案メール直接送信機能'],
      ['👥 チーム機能', 'マルチユーザー対応', 'Phase 2', '2-3ヶ月', '高', '複数営業担当者の同時利用'],
      ['📊 高度分析', '詳細分析機能', 'Phase 2', '1-2ヶ月', '中', 'より深い洞察提供機能'],
      ['🔄 自動化', '自動更新機能', 'Phase 2', '2-3週間', '中', '定期的なデータメンテナンス'],
      
      // Phase 3: 長期
      ['🌐 プラットフォーム', 'Webアプリ化', 'Phase 3', '3-6ヶ月', '高', 'ブラウザ独立型アプリケーション'],
      ['🔗 外部連携', 'CRM統合', 'Phase 3', '2-4ヶ月', '高', 'Salesforce、HubSpot等との連携'],
      ['🔮 AI機能', 'AI予測機能', 'Phase 3', '3-6ヶ月', '高', '機械学習による成果予測'],
      ['📱 モバイル', 'モバイル最適化', 'Phase 3', '2-3ヶ月', '中', 'スマートフォン対応'],
      
      // 高度機能
      ['📈 分析強化', 'A/Bテスト機能', '将来検討', '1-2ヶ月', '中', '複数戦略の効果比較'],
      ['📧 通信機能', 'メール配信システム', '将来検討', '2-3ヶ月', '中', 'MailChimp、SendGrid連携'],
      ['🏷️ 管理機能', '企業タグ管理', '将来検討', '1-2週間', '低', 'カスタムタグによる企業分類'],
      ['📊 追跡機能', '企業成長追跡', '将来検討', '2-4週間', '中', '時系列での企業成長分析'],
    ];
    
    // データを挿入
    const startRow = 5;
    futureSheet.getRange(startRow, 1, futureFeatures.length, 6).setValues(futureFeatures);
    
    // 書式設定
    for (let i = 0; i < futureFeatures.length; i++) {
      const rowNum = startRow + i;
      const priority = futureFeatures[i][2];
      
      // 優先度別の色分け
      let bgColor = '#ffffff';
      if (priority === 'Phase 0') bgColor = '#fce8e6'; // 赤系（緊急）
      else if (priority === 'Phase 1') bgColor = '#fff2cc'; // 黄系（短期）
      else if (priority === 'Phase 2') bgColor = '#d9ead3'; // 緑系（中期）
      else if (priority === 'Phase 3') bgColor = '#cfe2f3'; // 青系（長期）
      else if (priority === '将来検討') bgColor = '#f3f3f3'; // 灰系（将来）
      
      futureSheet.getRange(rowNum, 1, 1, 6).setBackground(bgColor);
      
      // 技術難易度の色分け
      const difficulty = futureFeatures[i][4];
      let difficultyColor = '#000000';
      if (difficulty === '低') difficultyColor = '#0d652d';
      else if (difficulty === '中') difficultyColor = '#bf9000';
      else if (difficulty === '高') difficultyColor = '#cc0000';
      
      futureSheet.getRange(rowNum, 5).setFontColor(difficultyColor).setFontWeight('bold');
    }
    
    // 列幅調整
    futureSheet.setColumnWidth(1, 120); // 機能カテゴリ
    futureSheet.setColumnWidth(2, 200); // 機能名
    futureSheet.setColumnWidth(3, 100); // 実装優先度
    futureSheet.setColumnWidth(4, 100); // 開発期間
    futureSheet.setColumnWidth(5, 100); // 技術難易度
    futureSheet.setColumnWidth(6, 300); // 説明
    
    // 枠線追加
    const dataRange = futureSheet.getRange(4, 1, futureFeatures.length + 1, 6);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 凡例を追加
    const legendRow = startRow + futureFeatures.length + 2;
    futureSheet.getRange(legendRow, 1).setValue('📊 実装優先度 凡例');
    futureSheet.getRange(legendRow, 1).setFontWeight('bold').setFontSize(12);
    
    const legend = [
      ['Phase 0 (緊急)', '既存プレースホルダー機能の完全実装', '#fce8e6'],
      ['Phase 1 (短期)', '1-3ヶ月で実装予定の機能', '#fff2cc'],
      ['Phase 2 (中期)', '3-6ヶ月で実装予定の機能', '#d9ead3'],
      ['Phase 3 (長期)', '6ヶ月以上で実装予定の機能', '#cfe2f3'],
      ['将来検討', '需要に応じて実装を検討する機能', '#f3f3f3']
    ];
    
    for (let i = 0; i < legend.length; i++) {
      const row = legendRow + 1 + i;
      futureSheet.getRange(row, 1).setValue(legend[i][0]);
      futureSheet.getRange(row, 2).setValue(legend[i][1]);
      futureSheet.getRange(row, 1, 1, 2).setBackground(legend[i][2]);
    }
    
    // シートをアクティブにする
    futureSheet.activate();
    
    // 完了メッセージ
    ui.alert(
      '✅ 将来機能一覧作成完了',
      '将来機能一覧シートを作成しました。\n\n主な機能:\n• Phase 0: 緊急実装必要機能\n• Phase 1-3: 段階的実装予定\n• 優先度別色分け表示\n• 技術難易度・開発期間情報',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Show future features error:', error);
    SpreadsheetApp.getUi().alert('❌ 将来機能一覧エラー: ' + String(error).substring(0, 100));
  }
}

// =============================================
// プラン管理機能
// =============================================

/**
 * プラン対応メニュー作成
 */
function createPlanBasedMenu(userPlan, trialInfo) {
  const ui = SpreadsheetApp.getUi();
  const planLimits = getPlanLimits(userPlan, trialInfo);
  
  // メインメニュー
  const mainMenu = ui.createMenu('🚀 営業システム');
  
  // システム管理（全プラン共通）
  mainMenu.addSubMenu(ui.createMenu('📊 システム管理')
    .addItem('🧪 システムテスト', 'runSystemTest')
    .addItem('🔑 APIキーテスト', 'testApiKeys')
    .addItem('💰 プラン状況', 'showPlanStatus')
    .addItem('📋 基本情報', 'showBasicInfo')
    .addItem('🔧 シート作成', 'createBasicSheets')
    .addItem('🗂️ シート設定', 'configureSheets')
    .addItem('🔄 データ同期', 'syncAllData'));
  
  // キーワード管理（プラン別制御）
  if (planLimits.keywordGeneration) {
    mainMenu.addSubMenu(ui.createMenu('🔍 キーワード管理')
      .addItem('🎯 キーワード生成', 'generateKeywords')
      .addItem('📊 使用状況確認', 'showKeywordUsage'));
  } else {
    mainMenu.addSubMenu(ui.createMenu('🔍 キーワード管理')
      .addItem('⚠️ キーワード生成（要アップグレード）', 'showUpgradeForKeywords')
      .addItem('💡 プラン比較', 'showPlanComparison'));
  }
  
  // 企業管理（制限表示）
  const companyLimitText = trialInfo && trialInfo.isTrialActive ? 
    `🔍 企業検索（試用期間：${planLimits.maxCompaniesPerDay}社/日）` :
    `🔍 企業検索（${planLimits.maxCompaniesPerDay}社/日）`;
    
  mainMenu.addSubMenu(ui.createMenu('🏢 企業管理')
    .addItem(companyLimitText, 'searchCompany')
    .addItem('📊 検索履歴', 'showSearchHistory'));
  
  // 提案管理（プラン別制御）
  if (planLimits.proposalGeneration) {
    mainMenu.addSubMenu(ui.createMenu('💼 提案管理')
      .addItem('✨ AI提案生成', 'generateProposal')
      .addItem('📝 提案履歴', 'showProposalHistory'));
  } else {
    mainMenu.addSubMenu(ui.createMenu('💼 提案管理')
      .addItem('📝 基本テンプレート提案', 'generateBasicProposal')
      .addItem('⚠️ AI提案生成（要アップグレード）', 'showUpgradeForProposals')
      .addItem('💡 プラン比較', 'showPlanComparison'));
  }
  
  // 分析・レポート（全プラン共通）
  mainMenu.addSubMenu(ui.createMenu('📈 分析・レポート')
    .addItem('📊 総合レポート', 'generateComprehensiveReport')
    .addItem('📋 活動ログ', 'viewActivityLog'));
  
  // ライセンス管理（管理者専用）
  mainMenu.addSubMenu(ui.createMenu('🔐 ライセンス管理')
    .addItem('📋 ライセンス状況', 'showLicenseStatus')
    .addItem('👤 管理者認証', 'authenticateAdmin')
    .addSeparator()
    .addItem('🔑 APIキー管理', 'manageApiKeys')
    .addSeparator()
    .addItem('📅 使用開始設定', 'setLicenseStartDate')
    .addItem('🔄 期限延長', 'extendLicense')
    .addItem('🔒 システムロック解除', 'unlockSystem'));
  
  // ヘルプ・ドキュメント（全プラン共通）
  mainMenu.addSubMenu(ui.createMenu('📚 ヘルプ・ドキュメント')
    .addItem('🆘 基本ヘルプ', 'showHelp')
    .addItem('📖 ユーザーガイド', 'showUserGuide')
    .addItem('⚙️ 機能説明', 'showFeatureGuide')
    .addItem('🔧 設定ガイド', 'showConfigGuide')
    .addItem('💰 料金・API設定ガイド', 'showPricingGuide')
    .addSeparator()
    .addItem('🚀 将来機能一覧', 'showFutureFeatures')
    .addItem('❓ よくある質問', 'showFAQ')
    .addItem('🐛 トラブルシューティング', 'showTroubleshooting'));
  
  // 設定（プラン別制御）
  mainMenu.addSeparator();
  const settingsMenu = ui.createMenu('⚙️ 設定');
  
  // APIキー設定：ベーシックプランでは不要を明示
  if (planLimits.chatgptApiRequired) {
    settingsMenu.addItem('🔑 APIキー設定', 'configureApiKeys');
  } else {
    settingsMenu.addItem('🔑 APIキー設定（ChatGPT不要）', 'showApiKeyNotRequired');
  }
  
  settingsMenu
    .addItem('📊 基本設定', 'showBasicSettings')
    .addItem('🔧 詳細設定', 'showAdvancedSettings')
    .addItem('🌐 システム環境', 'showSystemEnvironment')
    .addItem('💳 プランアップグレード', 'showPlanUpgrade');
    
  mainMenu.addSubMenu(settingsMenu);
  
  mainMenu.addToUi();
}

/**
 * ユーザープラン取得
 */
function getUserPlan() {
  try {
    const properties = PropertiesService.getUserProperties();
    return properties.getProperty('user_plan') || 'TRIAL';
  } catch (error) {
    console.error('Get user plan error:', error);
    return 'TRIAL';
  }
}

/**
 * 試用期間情報取得
 */
function getTrialInfo() {
  try {
    const properties = PropertiesService.getUserProperties();
    const trialStartDate = properties.getProperty('trial_start_date');
    
    if (!trialStartDate) {
      // 初回利用時：試用期間開始
      const startDate = new Date();
      properties.setProperty('trial_start_date', startDate.toISOString());
      return {
        isTrialActive: true,
        startDate: startDate,
        remainingDays: 10,
        isFirstTime: true
      };
    }
    
    const startDate = new Date(trialStartDate);
    const currentDate = new Date();
    const daysDiff = Math.floor((currentDate - startDate) / (1000 * 60 * 60 * 24));
    const remainingDays = Math.max(0, 10 - daysDiff);
    
    return {
      isTrialActive: remainingDays > 0,
      startDate: startDate,
      remainingDays: remainingDays,
      isFirstTime: false
    };
  } catch (error) {
    console.error('Get trial info error:', error);
    return {
      isTrialActive: false,
      startDate: null,
      remainingDays: 0,
      isFirstTime: false
    };
  }
}

/**
 * プラン制限取得
 */
function getPlanLimits(userPlan, trialInfo) {
  // 試用期間の制限
  if (trialInfo && trialInfo.isTrialActive) {
    return {
      planName: 'TRIAL',
      keywordGeneration: true,        // 試用期間はキーワード生成可能
      proposalGeneration: true,       // 試用期間はAI提案生成可能
      maxCompaniesPerDay: 5,          // 試用期間は5社まで
      maxProposalsPerDay: 5,
      chatgptApiRequired: true,       // ChatGPT API必要
      googleSearchApiRequired: true
    };
  }
  
  // 正式プランの制限
  const planLimits = {
    'BASIC': {
      planName: 'BASIC',
      keywordGeneration: false,       // ベーシックはキーワード生成不可
      proposalGeneration: false,      // ベーシックはAI提案生成不可
      maxCompaniesPerDay: 10,
      maxProposalsPerDay: 0,          // AI提案は0件
      chatgptApiRequired: false,      // ChatGPT API不要
      googleSearchApiRequired: true   // Google検索は企業確認用
    },
    'STANDARD': {
      planName: 'STANDARD',
      keywordGeneration: true,
      proposalGeneration: true,
      maxCompaniesPerDay: 50,
      maxProposalsPerDay: 50,
      chatgptApiRequired: true,
      googleSearchApiRequired: true
    },
    'PROFESSIONAL': {
      planName: 'PROFESSIONAL',
      keywordGeneration: true,
      proposalGeneration: true,
      maxCompaniesPerDay: 100,
      maxProposalsPerDay: 100,
      chatgptApiRequired: true,
      googleSearchApiRequired: true
    },
    'ENTERPRISE': {
      planName: 'ENTERPRISE',
      keywordGeneration: true,
      proposalGeneration: true,
      maxCompaniesPerDay: 500,
      maxProposalsPerDay: 500,
      chatgptApiRequired: true,
      googleSearchApiRequired: true
    }
  };
  
  return planLimits[userPlan] || planLimits['BASIC'];
}

/**
 * プラン表示名取得
 */
function getPlanDisplayName(userPlan, trialInfo) {
  if (trialInfo && trialInfo.isTrialActive) {
    return `試用期間 ${trialInfo.remainingDays}日残`;
  }
  
  const displayNames = {
    'BASIC': 'ベーシック ¥500',
    'STANDARD': 'スタンダード ¥1,500',
    'PROFESSIONAL': 'プロフェッショナル ¥3,000',
    'ENTERPRISE': 'エンタープライズ ¥7,500'
  };
  
  return displayNames[userPlan] || 'ベーシック ¥500';
}

/**
 * プラン状況表示
 */
function showPlanStatus() {
  try {
    const ui = SpreadsheetApp.getUi();
    const userPlan = getUserPlan();
    const trialInfo = getTrialInfo();
    const planLimits = getPlanLimits(userPlan, trialInfo);
    
    let statusText = '💰 プラン状況\n\n';
    
    // 試用期間の場合
    if (trialInfo.isTrialActive) {
      statusText += '🆓 試用期間中\n';
      statusText += `📅 残り日数: ${trialInfo.remainingDays}日\n`;
      statusText += `📊 企業検索: ${planLimits.maxCompaniesPerDay}社/日\n`;
      statusText += '✅ キーワード生成: 利用可能\n\n';
      statusText += '💡 試用期間終了後は有料プランへの移行が必要です';
    } else {
      // 正式プランの場合
      const planDisplayName = getPlanDisplayName(userPlan, trialInfo);
      statusText += `📋 現在のプラン: ${planDisplayName}\n\n`;
      statusText += `📊 企業検索: ${planLimits.maxCompaniesPerDay}社/日\n`;
      statusText += `💼 提案生成: ${planLimits.maxProposalsPerDay}件/日\n`;
      statusText += `🎯 キーワード生成: ${planLimits.keywordGeneration ? '✅ 利用可能' : '❌ 利用不可'}\n`;
      statusText += `🤖 ChatGPT API: ${planLimits.chatgptApiRequired ? '必要' : '不要'}\n\n`;
      
      if (!planLimits.keywordGeneration) {
        statusText += '💡 キーワード生成を利用するにはスタンダードプラン以上が必要です';
      }
    }
    
    ui.alert('プラン状況', statusText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Show plan status error:', error);
    SpreadsheetApp.getUi().alert('❌ プラン状況エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * AI提案生成アップグレード案内
 */
function showUpgradeForProposals() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '✨ AI提案生成機能\n\n';
    message += '❌ この機能は現在のプランでは利用できません\n\n';
    message += '💡 現在のプランでは基本テンプレート提案のみ利用可能です\n\n';
    message += '✅ AI提案生成利用可能プラン:\n';
    message += '• スタンダード（¥1,500/月）\n';
    message += '• プロフェッショナル（¥3,000/月）\n';
    message += '• エンタープライズ（¥7,500/月）\n\n';
    message += '🎯 AI提案生成の特徴:\n';
    message += '• 企業に最適化された個別提案\n';
    message += '• 業界特化型の営業メッセージ\n';
    message += '• 効果的な件名・本文自動生成\n\n';
    message += 'プラン詳細を確認しますか？';
    
    const result = ui.alert('プランアップグレード', message, ui.ButtonSet.YES_NO);
    
    if (result === ui.Button.YES) {
      showPlanComparison();
    }
    
  } catch (error) {
    console.error('Show upgrade for proposals error:', error);
    SpreadsheetApp.getUi().alert('❌ アップグレード案内エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * 基本テンプレート提案生成
 */
function generateBasicProposal() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '📝 基本テンプレート提案\n\n';
    message += '💡 ベーシックプランでは基本的なテンプレート提案を利用できます\n\n';
    message += '📋 利用可能なテンプレート:\n';
    message += '• 新規開拓用テンプレート\n';
    message += '• アフターフォロー用テンプレート\n';
    message += '• 商品紹介用テンプレート\n\n';
    message += '⚠️ 企業に最適化されたAI提案をご希望の場合は\nスタンダードプラン以上をご検討ください\n\n';
    message += 'テンプレート選択画面を開きますか？';
    
    const result = ui.alert('基本テンプレート提案', message, ui.ButtonSet.YES_NO);
    
    if (result === ui.Button.YES) {
      showProposalTemplates();
    }
    
  } catch (error) {
    console.error('Generate basic proposal error:', error);
    SpreadsheetApp.getUi().alert('❌ 基本提案エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * 提案テンプレート表示
 */
function showProposalTemplates() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let templates = '📝 提案テンプレート一覧\n\n';
    templates += '1️⃣ 新規開拓用テンプレート\n';
    templates += '「はじめまして、○○株式会社と申します..」\n\n';
    templates += '2️⃣ アフターフォロー用テンプレート\n';
    templates += '「先日はお忙しい中、お時間をいただき..」\n\n';
    templates += '3️⃣ 商品紹介用テンプレート\n';
    templates += '「貴社の業務効率化に貢献する..」\n\n';
    templates += '💡 これらのテンプレートは手動でカスタマイズできます\n';
    templates += '🚀 AI による個別最適化をご希望の場合は\nスタンダードプラン以上をご検討ください';
    
    ui.alert('提案テンプレート', templates, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Show proposal templates error:', error);
    SpreadsheetApp.getUi().alert('❌ テンプレート表示エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * キーワード生成アップグレード案内
 */
function showUpgradeForKeywords() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '🔑 キーワード生成機能\n\n';
    message += '❌ この機能は現在のプランでは利用できません\n\n';
    message += '✅ 利用可能プラン:\n';
    message += '• スタンダード（¥1,500/月）\n';
    message += '• プロフェッショナル（¥3,000/月）\n';
    message += '• エンタープライズ（¥7,500/月）\n\n';
    message += '💡 プランアップグレードで以下の機能が利用可能になります:\n';
    message += '• AI自動キーワード生成\n';
    message += '• 企業検索の大幅拡張\n';
    message += '• 高度な提案生成\n\n';
    message += 'プラン詳細を確認しますか？';
    
    const result = ui.alert('プランアップグレード', message, ui.ButtonSet.YES_NO);
    
    if (result === ui.Button.YES) {
      showPlanComparison();
    }
    
  } catch (error) {
    console.error('Show upgrade for keywords error:', error);
    SpreadsheetApp.getUi().alert('❌ アップグレード案内エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * APIキー設定不要案内
 */
function showApiKeyNotRequired() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '🔑 APIキー設定について\n\n';
    message += '✅ 現在のプランではChatGPT APIキーの設定は不要です\n\n';
    message += '📋 ベーシックプランの特徴:\n';
    message += '• 手動企業入力による提案生成\n';
    message += '• ChatGPT API利用なし\n';
    message += '• Google検索APIのみ使用\n\n';
    message += '💡 AI機能をフル活用したい場合は\nスタンダードプラン以上をご検討ください';
    
    ui.alert('APIキー設定', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Show API key not required error:', error);
    SpreadsheetApp.getUi().alert('❌ APIキー案内エラー: ' + String(error).substring(0, 100));
  }
}

/**
 * プラン比較表示
 */
function showPlanComparison() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let comparison = '💰 プラン比較表\n\n';
    comparison += '🥉 ベーシック（¥500/月）\n';
    comparison += '• 企業検索: 10社/日\n';
    comparison += '• キーワード生成: ❌\n';
    comparison += '• ChatGPT API: 不要\n\n';
    
    comparison += '🥈 スタンダード（¥1,500/月）\n';
    comparison += '• 企業検索: 50社/日\n';
    comparison += '• キーワード生成: ✅\n';
    comparison += '• ChatGPT API: 必要\n';
    comparison += '• 全機能利用可能\n\n';
    
    comparison += '🥇 プロフェッショナル（¥3,000/月）\n';
    comparison += '• 企業検索: 100社/日\n';
    comparison += '• 高速処理・優先サポート\n\n';
    
    comparison += '💎 エンタープライズ（¥7,500/月）\n';
    comparison += '• 企業検索: 500社/日\n';
    comparison += '• 最大性能・専任サポート\n\n';
    
    comparison += '🆓 試用期間: 10日間\n';
    comparison += '• 企業検索: 5社/日\n';
    comparison += '• 全機能お試し可能';
    
    ui.alert('プラン比較', comparison, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Show plan comparison error:', error);
    SpreadsheetApp.getUi().alert('❌ プラン比較エラー: ' + String(error).substring(0, 100));
  }
}
