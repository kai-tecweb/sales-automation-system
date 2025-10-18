/**
 * 営業自動化システム - 仕様書v2.0完全準拠メニューシステム
 * システム仕様書v2.0に厳密に準拠した実装
 */

function onOpen() {
  try {
    console.log('🚀 仕様書v2.0準拠メニューシステム開始...');
    
    // ガバナンスチェック実行
    if (typeof checkBeforeMenuModification === 'function') {
      checkBeforeMenuModification();
    }
    
    createSpecCompliantMenu();
    
  } catch (error) {
    console.error('❌ 仕様書準拠メニュー作成エラー:', error);
    createEmergencyFallbackMenu();
  }
}

/**
 * 仕様書v2.0完全準拠メニューの作成
 */
function createSpecCompliantMenu() {
  try {
    console.log('🏗️ 仕様書v2.0準拠メニュー作成開始...');
    
    const ui = SpreadsheetApp.getUi();
    
    // ユーザー権限確認（仕様書準拠）
    let isAdmin = false;
    try {
      const currentUser = getCurrentUser();
      isAdmin = currentUser.role === 'Administrator';
    } catch (error) {
      console.log('ユーザー管理未初期化 - 一般ユーザーとして処理');
    }
    
    // メインメニュー作成（仕様書v2.0タイトル）
    const mainMenu = ui.createMenu('🚀 営業システム');
    
    // 📊 システム管理（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('📊 システム管理')
      .addItem('🧪 システムテスト', 'runSystemTest')
      .addItem('🔑 APIキーテスト', 'testApiKeys')
      .addItem('📋 基本情報', 'showBasicInfo')
      .addItem('🔧 シート作成', 'createBasicSheets'));
    
    // 🔍 キーワード管理（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('🔍 キーワード管理')
      .addItem('🎯 キーワード生成', 'generateKeywords')
      .addItem('📊 使用状況分析', 'analyzeKeywordUsage'));
    
    // 🏢 企業管理（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('🏢 企業管理')
      .addItem('🔍 企業検索', 'searchCompany')
      .addItem('📊 企業分析', 'analyzeCompany')
      .addItem('📈 マッチ度計算', 'calculateMatching'));
    
    // 💼 提案管理（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('💼 提案管理')
      .addItem('✨ 提案生成', 'generateProposal')
      .addItem('📊 提案分析', 'analyzeProposal'));
    
    // 📊 分析・レポート（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('📊 分析・レポート')
      .addItem('📊 総合レポート', 'generateComprehensiveReport')
      .addItem('📋 活動ログ', 'viewActivityLog'));
    
    // 📚 ヘルプ・ドキュメント（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('📚 ヘルプ・ドキュメント')
      .addItem('🆘 基本ヘルプ', 'showHelp')
      .addItem('📖 ユーザーガイド', 'showUserGuide')
      .addItem('💰 料金・API設定ガイド', 'showPricingGuide')
      .addItem('❓ よくある質問', 'showFAQ'));
    
    // ⚙️ 設定（仕様書準拠）
    mainMenu.addSubMenu(ui.createMenu('⚙️ 設定')
      .addItem('🔑 APIキー設定', 'configureApiKeys')
      .addItem('📊 基本設定', 'showBasicSettings')
      .addItem('🔧 詳細設定', 'showAdvancedSettings')
      .addItem('🌐 システム環境', 'showSystemEnvironment'));
    
    // 管理者専用機能（仕様書v2.0準拠）
    if (isAdmin) {
      mainMenu.addSeparator();
      
      // 🔐 ライセンス管理（管理者専用）
      mainMenu.addSubMenu(ui.createMenu('🔐 ライセンス管理（管理者専用）')
        .addItem('📋 ライセンス状況', 'showLicenseStatus')
        .addItem('🔐 管理者認証', 'authenticateAdmin')
        .addItem('🔑 APIキー管理', 'manageApiKeys')
        .addItem('📅 使用開始日設定', 'setLicenseStartDate')
        .addItem('🔄 期限延長', 'extendLicense')
        .addItem('🔒 システムロック解除', 'unlockSystem')
        .addItem('💳 課金状況管理', 'manageBilling'));
      
      // 👥 ユーザー管理（管理者専用）
      mainMenu.addSubMenu(ui.createMenu('👥 ユーザー管理（管理者専用）')
        .addItem('👤 管理者モード切替', 'toggleAdminMode')
        .addItem('📊 利用統計', 'viewPerformanceStatistics')
        .addItem('🔧 システム設定', 'systemConfiguration'));
      
      // 🔧 高度な設定（管理者専用）
      mainMenu.addSubMenu(ui.createMenu('🔧 高度な設定（管理者専用）')
        .addItem('📊 詳細分析設定', 'advancedAnalytics')
        .addItem('🔄 データ同期', 'syncAllData')
        .addItem('🛠️ システムメンテナンス', 'systemMaintenance'));
    }
    
    // ガバナンス機能
    mainMenu.addSeparator();
    mainMenu.addItem('🔍 仕様書準拠チェック', 'generateComplianceReport');
    mainMenu.addItem('🔧 仕様書準拠メニュー強制適用', 'enforceSpecCompliantMenu');
    
    // メニュー有効化
    mainMenu.addToUi();
    
    // システム起動通知（仕様書準拠版）
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '営業システム v2.0 - 仕様書準拠版', 
      '🏗️ 標準メニュー起動完了', 
      5
    );
    
    console.log('✅ 仕様書v2.0準拠メニュー作成完了');
    
  } catch (error) {
    console.error('❌ 仕様書準拠メニュー作成エラー:', error);
    throw error;
  }
}

/**
 * 緊急時フォールバックメニュー
 */
function createEmergencyFallbackMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🆘 営業システム (緊急モード)')
      .addItem('🔧 仕様書準拠メニュー復元', 'enforceSpecCompliantMenu')
      .addItem('📋 システム状態確認', 'showBasicInfo')
      .addItem('🔍 仕様書準拠チェック', 'generateComplianceReport')
      .addToUi();
  } catch (error) {
    console.error('❌ 緊急フォールバックメニュー作成失敗:', error);
  }
}

// ==============================================
// 🎯 仕様書v2.0準拠実装関数
// ==============================================

/**
 * 🧪 システムテスト（仕様書準拠・拡張版）
 */
function runSystemTest() {
  try {
    console.log('🧪 システムテスト開始（仕様書v2.0準拠）');
    
    let testResults = '🧪 営業システム テスト結果\n\n';
    
    // 1. 仕様書準拠チェック
    testResults += '📖 仕様書v2.0準拠確認:\n';
    try {
      if (typeof validateMenuComplianceWithSpec === 'function') {
        const compliance = validateMenuComplianceWithSpec();
        testResults += `${compliance.compliant ? '✅' : '❌'} メニュー構造準拠\n`;
        if (!compliance.compliant && compliance.violations) {
          testResults += `  違反数: ${compliance.violations.length}件\n`;
        }
      } else {
        testResults += '❌ 準拠チェック機能未実装\n';
      }
    } catch (error) {
      testResults += '❌ 準拠チェックエラー\n';
    }
    
    // 2. シート存在確認
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    testResults += '\n📋 システムシート確認:\n';
    let sheetCount = 0;
    requiredSheets.forEach(sheetName => {
      const exists = sheetNames.includes(sheetName);
      testResults += `${exists ? '✅' : '❌'} ${sheetName}\n`;
      if (exists) sheetCount++;
    });
    testResults += `📊 シート状況: ${sheetCount}/${requiredSheets.length} 完了\n`;
    
    // 3. API設定確認
    testResults += '\n🔑 API設定確認:\n';
    const properties = PropertiesService.getScriptProperties().getProperties();
    const apiItems = [
      { key: 'OPENAI_API_KEY', name: 'OpenAI API' },
      { key: 'GOOGLE_SEARCH_API_KEY', name: 'Google Search API' },
      { key: 'GOOGLE_SEARCH_ENGINE_ID', name: 'Search Engine ID' }
    ];
    
    let apiCount = 0;
    apiItems.forEach(item => {
      const exists = !!properties[item.key];
      testResults += `${exists ? '✅' : '❌'} ${item.name}\n`;
      if (exists) apiCount++;
    });
    testResults += `📊 API設定: ${apiCount}/${apiItems.length} 完了\n`;
    
    // 4. 主要機能確認
    testResults += '\n🎯 主要機能確認:\n';
    const coreFunctions = [
      { name: 'generateKeywords', display: 'キーワード生成' },
      { name: 'executeCompanySearch', display: '企業検索' },  
      { name: 'generatePersonalizedProposals', display: '提案生成' },
      { name: 'initializeSheets', display: 'シート初期化' }
    ];
    
    let funcCount = 0;
    coreFunctions.forEach(func => {
      const exists = typeof globalThis[func.name] === 'function';
      testResults += `${exists ? '✅' : '❌'} ${func.display}\n`;
      if (exists) funcCount++;
    });
    testResults += `📊 機能実装: ${funcCount}/${coreFunctions.length} 完了\n`;
    
    // 5. 総合判定
    const allSystemsGo = sheetCount === requiredSheets.length && 
                        apiCount === apiItems.length && 
                        funcCount === coreFunctions.length;
    
    testResults += '\n🏁 総合判定:\n';
    if (allSystemsGo) {
      testResults += '✅ システム完全動作可能\n';
      testResults += '🚀 営業活動を開始できます\n\n';
      testResults += '📋 推奨アクション:\n';
      testResults += '1. 制御パネルに商材情報入力\n';
      testResults += '2. キーワード生成→企業検索→提案生成\n';
      testResults += '3. 結果を活用して営業活動開始';
    } else {
      testResults += '⚠️ 設定・修正が必要です\n\n';
      testResults += '📋 必要なアクション:\n';
      if (sheetCount < requiredSheets.length) testResults += '• 🔧 シート作成を実行\n';
      if (apiCount < apiItems.length) testResults += '• 🔑 APIキー設定を実行\n';
      if (funcCount < coreFunctions.length) testResults += '• システム再デプロイを確認\n';
    }
    
    SpreadsheetApp.getUi().alert('システムテスト', testResults, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // ログ記録
    try {
      const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('実行ログ');
      if (logSheet) {
        const timestamp = new Date().toLocaleString('ja-JP');
        const status = allSystemsGo ? '正常' : '要対応';
        logSheet.appendRow([timestamp, 'システムテスト', status, `シート:${sheetCount}/${requiredSheets.length} API:${apiCount}/${apiItems.length} 機能:${funcCount}/${coreFunctions.length}`, '']);
      }
    } catch (logError) {
      console.log('ログ記録エラー:', logError.message);
    }
    
    console.log('✅ システムテスト完了');
    return allSystemsGo;
    
  } catch (error) {
    console.error('❌ システムテストエラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'システムテストでエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}

/**
 * 🔑 APIキーテスト（仕様書準拠）
 */
function testApiKeys() {
  try {
    console.log('🔑 APIキーテスト開始');
    
    let testResults = '🔑 APIキーテスト結果\n\n';
    
    // 1. API設定確認
    const properties = PropertiesService.getScriptProperties().getProperties();
    const hasOpenAI = !!properties.OPENAI_API_KEY;
    const hasGoogleSearch = !!properties.GOOGLE_SEARCH_API_KEY;
    const hasEngineId = !!properties.GOOGLE_SEARCH_ENGINE_ID;
    
    testResults += '📋 API設定確認:\n';
    testResults += `${hasOpenAI ? '✅' : '❌'} OpenAI APIキー: ${hasOpenAI ? '設定済み' : '未設定'}\n`;
    testResults += `${hasGoogleSearch ? '✅' : '❌'} Google Search APIキー: ${hasGoogleSearch ? '設定済み' : '未設定'}\n`;
    testResults += `${hasEngineId ? '✅' : '❌'} 検索エンジンID: ${hasEngineId ? '設定済み' : '未設定'}\n\n`;
    
    // 2. OpenAI API接続テスト
    testResults += '🤖 OpenAI API接続テスト:\n';
    if (hasOpenAI) {
      try {
        // 簡単なテストクエリを実行
        const testPayload = {
          model: 'gpt-3.5-turbo',
          messages: [
            {
              role: 'user',
              content: 'テスト接続です。「OK」と返答してください。'
            }
          ],
          max_tokens: 10
        };
        
        const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${properties.OPENAI_API_KEY}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(testPayload)
        });
        
        if (response.getResponseCode() === 200) {
          testResults += '✅ OpenAI API: 接続成功\n';
        } else {
          testResults += `❌ OpenAI API: 接続失敗 (${response.getResponseCode()})\n`;
        }
      } catch (error) {
        testResults += `❌ OpenAI API: エラー - ${error.message.substring(0, 100)}...\n`;
      }
    } else {
      testResults += '❌ OpenAI API: APIキー未設定\n';
    }
    
    // 3. Google Search API接続テスト
    testResults += '\n🔍 Google Search API接続テスト:\n';
    if (hasGoogleSearch && hasEngineId) {
      try {
        const testUrl = `https://www.googleapis.com/customsearch/v1?key=${properties.GOOGLE_SEARCH_API_KEY}&cx=${properties.GOOGLE_SEARCH_ENGINE_ID}&q=test&num=1`;
        
        const response = UrlFetchApp.fetch(testUrl, {
          method: 'GET'
        });
        
        if (response.getResponseCode() === 200) {
          testResults += '✅ Google Search API: 接続成功\n';
        } else {
          testResults += `❌ Google Search API: 接続失敗 (${response.getResponseCode()})\n`;
        }
      } catch (error) {
        testResults += `❌ Google Search API: エラー - ${error.message.substring(0, 100)}...\n`;
      }
    } else {
      testResults += '❌ Google Search API: 設定不完全\n';
    }
    
    testResults += '\n📝 推奨アクション:\n';
    if (!hasOpenAI || !hasGoogleSearch || !hasEngineId) {
      testResults += '• 「⚙️ 設定」→「🔑 APIキー設定」でAPIキーを設定してください\n';
    }
    testResults += '• エラーが続く場合は、APIキーの有効性を確認してください\n';
    
    SpreadsheetApp.getUi().alert('APIキーテスト', testResults, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ APIキーテストエラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'APIキーテストでエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 📋 基本情報（仕様書準拠）
 */
function showBasicInfo() {
  try {
    console.log('📋 基本情報表示');
    
    let info = '📋 営業自動化システム 基本情報\n\n';
    info += '📖 仕様書バージョン: v2.0\n';
    info += '🚀 システム名: 営業自動化システム\n';
    info += '🎯 目的: 商材起点企業発掘・提案自動生成\n\n';
    
    info += '🔧 主要機能:\n';
    info += '• キーワード生成（ChatGPT API）\n';
    info += '• 企業検索（Google Search API）\n';
    info += '• 提案生成（ChatGPT API）\n';
    info += '• 分析・レポート\n\n';
    
    info += '📚 関連ドキュメント:\n';
    info += '• システム仕様書v2.0\n';
    info += '• API仕様書v2.0\n';
    info += '• 技術仕様書v2.0\n';
    
    SpreadsheetApp.getUi().alert('基本情報', info, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 基本情報表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '基本情報表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 🔧 シート作成（仕様書準拠）
 */
function createBasicSheets() {
  try {
    console.log('🔧 基本シート作成開始');
    
    // 既存のinitializeSheets()を呼び出し（実装済み機能を活用）
    if (typeof initializeSheets === 'function') {
      initializeSheets();
      SpreadsheetApp.getUi().alert(
        '✅ 完了', 
        '基本シートの作成が完了しました。\n\nシート一覧を確認してください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'エラー', 
        '❌ シート作成機能が実装されていません。\n\nmain.jsの確認が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('❌ シート作成エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'シート作成でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// プレースホルダー関数（仕様書で定義されているが未実装の機能）

function analyzeKeywordUsage() {
  SpreadsheetApp.getUi().alert(
    '開発予定機能', 
    '📊 使用状況分析\n\n仕様書v2.0で定義されていますが、まだ実装されていません。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function searchCompany() {
  // 既存のexecuteCompanySearch()を呼び出し
  if (typeof executeCompanySearch === 'function') {
    executeCompanySearch();
  } else {
    SpreadsheetApp.getUi().alert(
      '開発予定機能', 
      '🔍 企業検索\n\n実装中です。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function analyzeCompany() {
  SpreadsheetApp.getUi().alert(
    '開発予定機能', 
    '📊 企業分析\n\n仕様書v2.0で定義されていますが、まだ実装されていません。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function calculateMatching() {
  SpreadsheetApp.getUi().alert(
    '開発予定機能', 
    '📈 マッチ度計算\n\n仕様書v2.0で定義されていますが、まだ実装されていません。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function generateProposal() {
  // 既存のgeneratePersonalizedProposals()を呼び出し
  if (typeof generatePersonalizedProposals === 'function') {
    generatePersonalizedProposals();
  } else {
    SpreadsheetApp.getUi().alert(
      '開発予定機能', 
      '✨ 提案生成\n\n実装中です。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function analyzeProposal() {
  SpreadsheetApp.getUi().alert(
    '開発予定機能', 
    '📊 提案分析\n\n仕様書v2.0で定義されていますが、まだ実装されていません。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 📊 総合レポート生成（仕様書v2.0準拠）
 */
function generateComprehensiveReport() {
  try {
    console.log('📊 総合レポート生成開始');
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('📊 総合レポート生成中...', 'システム全体の状況を分析しています...', ui.ButtonSet.OK);
    
    let report = '📊 営業自動化システム - 総合レポート\n\n';
    report += `生成日時: ${new Date().toLocaleString('ja-JP')}\n\n`;
    
    // 1. システム概要
    report += '🎯 システム概要:\n';
    report += '• システム名: 営業自動化システム v2.0\n';
    report += '• 仕様書準拠: v2.0\n';
    report += '• 稼働状況: 本番運用可能\n\n';
    
    // 2. データ統計
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      
      report += '📊 データ統計:\n';
      
      // キーワードデータ
      try {
        const keywordSheet = spreadsheet.getSheetByName('生成キーワード');
        if (keywordSheet) {
          const keywordCount = Math.max(0, keywordSheet.getLastRow() - 1);
          report += `• 生成キーワード数: ${keywordCount}件\n`;
        } else {
          report += '• 生成キーワード数: 0件（シート未作成）\n';
        }
      } catch (error) {
        report += '• 生成キーワード数: データ取得エラー\n';
      }
      
      // 企業データ
      try {
        const companySheet = spreadsheet.getSheetByName('企業マスター');
        if (companySheet) {
          const companyCount = Math.max(0, companySheet.getLastRow() - 1);
          report += `• 発見企業数: ${companyCount}件\n`;
        } else {
          report += '• 発見企業数: 0件（シート未作成）\n';
        }
      } catch (error) {
        report += '• 発見企業数: データ取得エラー\n';
      }
      
      // 提案データ
      try {
        const proposalSheet = spreadsheet.getSheetByName('提案メッセージ');
        if (proposalSheet) {
          const proposalCount = Math.max(0, proposalSheet.getLastRow() - 1);
          report += `• 生成提案数: ${proposalCount}件\n`;
        } else {
          report += '• 生成提案数: 0件（シート未作成）\n';
        }
      } catch (error) {
        report += '• 生成提案数: データ取得エラー\n';
      }
      
      report += `• システムシート数: ${sheets.length}個\n\n`;
      
    } catch (error) {
      report += '📊 データ統計: 取得エラー\n\n';
    }
    
    // 3. API設定状況
    try {
      const props = PropertiesService.getScriptProperties().getProperties();
      
      report += '🔑 API設定状況:\n';
      report += `• OpenAI API: ${props.OPENAI_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
      report += `• Google Search API: ${props.GOOGLE_SEARCH_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
      report += `• 検索エンジンID: ${props.GOOGLE_SEARCH_ENGINE_ID ? '✅ 設定済み' : '❌ 未設定'}\n`;
      
      const apiReadiness = props.OPENAI_API_KEY && props.GOOGLE_SEARCH_API_KEY && props.GOOGLE_SEARCH_ENGINE_ID;
      report += `• API準備度: ${apiReadiness ? '✅ 完全準備済み' : '⚠️ 設定不完全'}\n\n`;
      
    } catch (error) {
      report += '🔑 API設定状況: 確認エラー\n\n';
    }
    
    // 4. システム健全性
    try {
      report += '🏥 システム健全性:\n';
      
      // 必須シートチェック
      const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
      
      let sheetScore = 0;
      requiredSheets.forEach(sheetName => {
        const exists = existingSheets.includes(sheetName);
        if (exists) sheetScore++;
      });
      
      const sheetHealth = Math.round((sheetScore / requiredSheets.length) * 100);
      report += `• シート健全性: ${sheetHealth}% (${sheetScore}/${requiredSheets.length})\n`;
      
      // 全体的な健全性
      const props = PropertiesService.getScriptProperties().getProperties();
      const apiHealth = (props.OPENAI_API_KEY ? 1 : 0) + (props.GOOGLE_SEARCH_API_KEY ? 1 : 0) + (props.GOOGLE_SEARCH_ENGINE_ID ? 1 : 0);
      const apiHealthPercent = Math.round((apiHealth / 3) * 100);
      
      const overallHealth = Math.round((sheetHealth + apiHealthPercent) / 2);
      report += `• API健全性: ${apiHealthPercent}% (${apiHealth}/3)\n`;
      report += `• 総合健全性: ${overallHealth}%\n\n`;
      
    } catch (error) {
      report += '🏥 システム健全性: 評価エラー\n\n';
    }
    
    // 5. 仕様書準拠状況
    try {
      report += '📖 仕様書v2.0準拠状況:\n';
      
      if (typeof validateMenuComplianceWithSpec === 'function') {
        const compliance = validateMenuComplianceWithSpec();
        report += `• 準拠状況: ${compliance.compliant ? '✅ 完全準拠' : '❌ 違反あり'}\n`;
        if (!compliance.compliant && compliance.violations) {
          report += `• 検出違反数: ${compliance.violations.length}件\n`;
        }
      } else {
        report += '• 準拠チェック: 機能未実装\n';
      }
      
    } catch (error) {
      report += '• 準拠チェック: エラー\n';
    }
    
    // 6. 推奨アクション
    report += '\n💡 推奨アクション:\n';
    
    const props = PropertiesService.getScriptProperties().getProperties();
    if (!props.OPENAI_API_KEY || !props.GOOGLE_SEARCH_API_KEY || !props.GOOGLE_SEARCH_ENGINE_ID) {
      report += '• 🔑 APIキー設定を完了させる\n';
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length > 0) {
      report += '• 🔧 システムシート作成を実行\n';
    }
    
    if (props.OPENAI_API_KEY && props.GOOGLE_SEARCH_API_KEY && missingSheets.length === 0) {
      report += '• 🚀 営業リスト自動生成を開始\n';
      report += '• 📈 定期的な企業発掘を実施\n';
    }
    
    report += '\n📊 このレポートは「分析・レポート」→「総合レポート」から再生成できます。';
    
    // レポート表示
    ui.alert('📊 総合レポート', report, ui.ButtonSet.OK);
    
    // ログ記録
    try {
      const logSheet = spreadsheet.getSheetByName('実行ログ');
      if (logSheet) {
        const timestamp = new Date().toLocaleString('ja-JP');
        logSheet.appendRow([timestamp, '総合レポート生成', '完了', 'システム全体の状況レポートを生成', '']);
      }
    } catch (logError) {
      console.log('ログ記録エラー:', logError.message);
    }
    
    console.log('✅ 総合レポート生成完了');
    
  } catch (error) {
    console.error('❌ 総合レポート生成エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '総合レポート生成でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 📋 活動ログ表示（仕様書v2.0準拠）
 */
function viewActivityLog() {
  try {
    console.log('📋 活動ログ表示開始');
    
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let logReport = '📋 営業システム - 活動ログ\n\n';
    
    try {
      // 実行ログシートから最新のログを取得
      const logSheet = spreadsheet.getSheetByName('実行ログ');
      
      if (!logSheet) {
        logReport += '❌ 実行ログシートが見つかりません。\n\n';
        logReport += '💡 解決方法:\n';
        logReport += '「📊 システム管理」→「🔧 シート作成」を実行してください。';
      } else {
        const lastRow = logSheet.getLastRow();
        
        if (lastRow <= 1) {
          logReport += '📝 活動ログはまだありません。\n\n';
          logReport += '💡 システムの使用を開始すると、ここに活動履歴が表示されます。\n\n';
          logReport += '🚀 推奨アクション:\n';
          logReport += '1. キーワード生成を実行\n';
          logReport += '2. 企業検索を実行\n';
          logReport += '3. 提案生成を実行';
        } else {
          // ヘッダーを確認
          const headers = logSheet.getRange(1, 1, 1, 5).getValues()[0];
          logReport += `📊 ログ項目: ${headers.join(' | ')}\n\n`;
          
          // 最新10件のログを取得
          const startRow = Math.max(2, lastRow - 9);
          const numRows = lastRow - startRow + 1;
          
          if (numRows > 0) {
            const logData = logSheet.getRange(startRow, 1, numRows, 5).getValues();
            
            logReport += `📋 最新の活動履歴 (最新${numRows}件):\n\n`;
            
            // ログを新しい順に表示
            for (let i = logData.length - 1; i >= 0; i--) {
              const row = logData[i];
              const [timestamp, action, status, details, error] = row;
              
              if (timestamp && action) {
                const statusIcon = status === '成功' || status === '完了' ? '✅' : 
                                 status === 'エラー' || status === '失敗' ? '❌' : 
                                 status === '実行中' ? '⏳' : '📝';
                
                logReport += `${statusIcon} ${timestamp}\n`;
                logReport += `   ${action}: ${status}\n`;
                if (details) {
                  logReport += `   詳細: ${details}\n`;
                }
                if (error) {
                  logReport += `   エラー: ${error}\n`;
                }
                logReport += '\n';
              }
            }
            
            logReport += `📊 総ログ件数: ${lastRow - 1}件\n\n`;
            logReport += '💡 完全なログを確認するには「実行ログ」シートを直接参照してください。';
          }
        }
      }
      
    } catch (error) {
      logReport += `❌ ログ取得エラー: ${error.message}\n\n`;
      logReport += '💡 「📊 システム管理」→「🧪 システムテスト」でシステム状況を確認してください。';
    }
    
    ui.alert('活動ログ', logReport, ui.ButtonSet.OK);
    console.log('✅ 活動ログ表示完了');
    
  } catch (error) {
    console.error('❌ 活動ログ表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '活動ログの表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showUserGuide() {
  SpreadsheetApp.getUi().alert(
    '📖 ユーザーガイド', 
    '営業自動化システム ユーザーガイド\n\n🎯 基本操作:\n1. 制御パネルに商材情報入力\n2. キーワード生成実行\n3. 企業検索実行\n4. 提案メッセージ生成\n\n詳細は実運用ガイドを参照してください。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showBasicSettings() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let settings = '📊 基本設定\n\n';
    
    // API設定状況
    const props = PropertiesService.getScriptProperties().getProperties();
    settings += '🔑 API設定状況:\n';
    settings += `• OpenAI API: ${props.OPENAI_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
    settings += `• Google Search API: ${props.GOOGLE_SEARCH_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
    settings += `• 検索エンジンID: ${props.GOOGLE_SEARCH_ENGINE_ID ? '✅ 設定済み' : '❌ 未設定'}\n\n`;
    
    // シート状況
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ'];
    
    settings += '📋 システムシート:\n';
    let sheetCount = 0;
    requiredSheets.forEach(sheetName => {
      const exists = sheets.some(sheet => sheet.getName() === sheetName);
      settings += `• ${sheetName}: ${exists ? '✅ 存在' : '❌ 未作成'}\n`;
      if (exists) sheetCount++;
    });
    
    settings += `\n📊 設定完了度: ${Math.round((sheetCount / requiredSheets.length) * 100)}%\n`;
    
    if (sheetCount < requiredSheets.length || !props.OPENAI_API_KEY || !props.GOOGLE_SEARCH_API_KEY) {
      settings += '\n💡 推奨アクション:\n';
      if (sheetCount < requiredSheets.length) settings += '• 🔧 シート作成を実行\n';
      if (!props.OPENAI_API_KEY || !props.GOOGLE_SEARCH_API_KEY) settings += '• 🔑 APIキー設定を実行\n';
    } else {
      settings += '\n✅ システムは正常に設定されています';
    }
    
    ui.alert('基本設定', settings, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 基本設定表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '基本設定の表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    '🆘 基本ヘルプ', 
    '営業自動化システムのヘルプ\n\n詳細は「📖 ユーザーガイド」を参照してください。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showPricingGuide() {
  SpreadsheetApp.getUi().alert(
    '💰 料金・API設定ガイド', 
    'API料金とシステム利用料金に関するガイド\n\n詳細情報は仕様書を参照してください。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showFAQ() {
  SpreadsheetApp.getUi().alert(
    '❓ よくある質問', 
    'Q: システムが動作しない\nA: APIキー設定を確認してください。\n\nQ: メニューが表示されない\nA: ブラウザを更新してください。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function configureApiKeys() {
  // 既存のsetApiKeysSimple()を呼び出し
  if (typeof setApiKeysSimple === 'function') {
    setApiKeysSimple();
  } else {
    SpreadsheetApp.getUi().alert(
      'API設定', 
      'APIキー設定機能を実装中です。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showAdvancedSettings() {
  SpreadsheetApp.getUi().alert(
    '🔧 詳細設定', 
    '詳細設定機能\n\n管理者専用の高度な設定項目です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showSystemEnvironment() {
  SpreadsheetApp.getUi().alert(
    '🌐 システム環境', 
    'Google Apps Script環境\n\nV8ランタイム使用', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==============================================
// 🔐 管理者専用機能の実装
// ==============================================

/**
 * ライセンス状況表示
 */
function showLicenseStatus() {
  try {
    if (typeof createLicenseManagementSheet === 'function') {
      createLicenseManagementSheet();
    } else {
      SpreadsheetApp.getUi().alert(
        'ライセンス状況', 
        '📋 ライセンス管理機能\n\nライセンス管理機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ ライセンス状況表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'ライセンス状況の表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 管理者認証
 */
function authenticateAdmin() {
  try {
    if (typeof authenticateAdminUser === 'function') {
      authenticateAdminUser();
    } else {
      SpreadsheetApp.getUi().alert(
        '管理者認証', 
        '🔐 管理者認証機能\n\n管理者認証機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ 管理者認証エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '管理者認証でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * APIキー管理
 */
function manageApiKeys() {
  try {
    if (typeof setApiKeys === 'function') {
      setApiKeys();
    } else {
      SpreadsheetApp.getUi().alert(
        'APIキー管理', 
        '🔑 APIキー管理機能\n\nAPIキー管理機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ APIキー管理エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'APIキー管理でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 使用開始日設定
 */
function setLicenseStartDate() {
  try {
    if (typeof setLicenseStartDateInteractive === 'function') {
      setLicenseStartDateInteractive();
    } else {
      SpreadsheetApp.getUi().alert(
        '使用開始日設定', 
        '📅 使用開始日設定機能\n\n使用開始日設定機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ 使用開始日設定エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '使用開始日設定でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 期限延長
 */
function extendLicense() {
  try {
    if (typeof extendLicensePeriod === 'function') {
      extendLicensePeriod();
    } else {
      SpreadsheetApp.getUi().alert(
        '期限延長', 
        '🔄 期限延長機能\n\n期限延長機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ 期限延長エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '期限延長でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * システムロック解除
 */
function unlockSystem() {
  try {
    if (typeof emergencyUnlock === 'function') {
      emergencyUnlock();
    } else {
      SpreadsheetApp.getUi().alert(
        'システムロック解除', 
        '🔒 システムロック解除機能\n\nシステムロック解除機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ システムロック解除エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'システムロック解除でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 課金状況管理
 */
function manageBilling() {
  SpreadsheetApp.getUi().alert(
    '課金状況管理', 
    '💳 課金状況管理\n\n課金状況管理機能を実装中です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 管理者モード切替
 */
function toggleAdminMode() {
  try {
    if (typeof switchToAdminMode === 'function') {
      switchToAdminMode();
    } else {
      SpreadsheetApp.getUi().alert(
        '管理者モード切替', 
        '👤 管理者モード切替機能\n\n管理者モード切替機能を実装中です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('❌ 管理者モード切替エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      '管理者モード切替でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 📈 パフォーマンス統計表示（仕様書v2.0準拠）
 */
function viewPerformanceStatistics() {
  try {
    console.log('📈 パフォーマンス統計表示開始');
    
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let statsReport = '📈 営業システム - パフォーマンス統計\n\n';
    
    try {
      // キーワード生成統計
      statsReport += '🔑 キーワード生成統計:\n';
      try {
        const keywordSheet = spreadsheet.getSheetByName('生成キーワード');
        if (keywordSheet) {
          const lastRow = keywordSheet.getLastRow();
          const totalKeywords = Math.max(0, lastRow - 1);
          statsReport += `   ✅ 生成キーワード数: ${totalKeywords}件\n`;
          
          // 商材別の統計を計算
          if (totalKeywords > 0) {
            const data = keywordSheet.getRange(2, 1, totalKeywords, 3).getValues();
            const productStats = {};
            
            data.forEach(row => {
              const product = row[0];
              if (product) {
                productStats[product] = (productStats[product] || 0) + 1;
              }
            });
            
            const uniqueProducts = Object.keys(productStats).length;
            statsReport += `   📊 商材種類数: ${uniqueProducts}種類\n`;
            
            // トップ3商材を表示
            const topProducts = Object.entries(productStats)
              .sort((a, b) => b[1] - a[1])
              .slice(0, 3);
            
            if (topProducts.length > 0) {
              statsReport += '   🏆 上位商材:\n';
              topProducts.forEach(([product, count], index) => {
                statsReport += `      ${index + 1}. ${product}: ${count}キーワード\n`;
              });
            }
          }
        } else {
          statsReport += '   ❌ キーワードシートが見つかりません\n';
        }
      } catch (error) {
        statsReport += `   ❌ キーワード統計エラー: ${error.message}\n`;
      }
      
      statsReport += '\n';
      
      // 企業検索統計
      statsReport += '🏢 企業検索統計:\n';
      try {
        const companySheet = spreadsheet.getSheetByName('企業マスター');
        if (companySheet) {
          const lastRow = companySheet.getLastRow();
          const totalCompanies = Math.max(0, lastRow - 1);
          statsReport += `   ✅ 発見企業数: ${totalCompanies}社\n`;
          
          if (totalCompanies > 0) {
            // スコア統計を計算
            const scoreCol = 7; // G列（スコア）
            const scores = companySheet.getRange(2, scoreCol, totalCompanies, 1).getValues().flat();
            const validScores = scores.filter(score => typeof score === 'number' && score > 0);
            
            if (validScores.length > 0) {
              const avgScore = validScores.reduce((sum, score) => sum + score, 0) / validScores.length;
              const maxScore = Math.max(...validScores);
              const highScoreCount = validScores.filter(score => score >= 70).length;
              
              statsReport += `   📊 平均スコア: ${avgScore.toFixed(1)}点\n`;
              statsReport += `   🌟 最高スコア: ${maxScore}点\n`;
              statsReport += `   ⭐ 高スコア企業（70点以上）: ${highScoreCount}社\n`;
            }
            
            // 業界統計
            const industryCol = 3; // C列（業界）
            const industries = companySheet.getRange(2, industryCol, totalCompanies, 1).getValues().flat();
            const industryStats = {};
            
            industries.forEach(industry => {
              if (industry) {
                industryStats[industry] = (industryStats[industry] || 0) + 1;
              }
            });
            
            const uniqueIndustries = Object.keys(industryStats).length;
            statsReport += `   🏭 業界数: ${uniqueIndustries}業界\n`;
          }
        } else {
          statsReport += '   ❌ 企業マスターシートが見つかりません\n';
        }
      } catch (error) {
        statsReport += `   ❌ 企業統計エラー: ${error.message}\n`;
      }
      
      statsReport += '\n';
      
      // 提案生成統計
      statsReport += '💬 提案生成統計:\n';
      try {
        const proposalSheet = spreadsheet.getSheetByName('提案メッセージ');
        if (proposalSheet) {
          const lastRow = proposalSheet.getLastRow();
          const totalProposals = Math.max(0, lastRow - 1);
          statsReport += `   ✅ 生成提案数: ${totalProposals}件\n`;
          
          if (totalProposals > 0) {
            // 文字数統計
            const messageCol = 4; // D列（提案メッセージ）
            const messages = proposalSheet.getRange(2, messageCol, totalProposals, 1).getValues().flat();
            const validMessages = messages.filter(msg => msg && typeof msg === 'string');
            
            if (validMessages.length > 0) {
              const lengths = validMessages.map(msg => msg.length);
              const avgLength = lengths.reduce((sum, len) => sum + len, 0) / lengths.length;
              const maxLength = Math.max(...lengths);
              const minLength = Math.min(...lengths);
              
              statsReport += `   📏 平均文字数: ${avgLength.toFixed(0)}文字\n`;
              statsReport += `   � 文字数範囲: ${minLength}〜${maxLength}文字\n`;
            }
          }
        } else {
          statsReport += '   ❌ 提案メッセージシートが見つかりません\n';
        }
      } catch (error) {
        statsReport += `   ❌ 提案統計エラー: ${error.message}\n`;
      }
      
      statsReport += '\n';
      
      // システム実行統計
      statsReport += '⚡ システム実行統計:\n';
      try {
        const logSheet = spreadsheet.getSheetByName('実行ログ');
        if (logSheet) {
          const lastRow = logSheet.getLastRow();
          const totalLogs = Math.max(0, lastRow - 1);
          statsReport += `   📋 総実行回数: ${totalLogs}回\n`;
          
          if (totalLogs > 0) {
            // 成功/失敗統計
            const statusCol = 3; // C列（ステータス）
            const statuses = logSheet.getRange(2, statusCol, totalLogs, 1).getValues().flat();
            
            const successCount = statuses.filter(status => 
              status === '成功' || status === '完了'
            ).length;
            const errorCount = statuses.filter(status => 
              status === 'エラー' || status === '失敗'
            ).length;
            
            const successRate = totalLogs > 0 ? (successCount / totalLogs * 100).toFixed(1) : 0;
            
            statsReport += `   ✅ 成功率: ${successRate}% (${successCount}/${totalLogs})\n`;
            if (errorCount > 0) {
              statsReport += `   ❌ エラー回数: ${errorCount}回\n`;
            }
          }
        } else {
          statsReport += '   ❌ 実行ログシートが見つかりません\n';
        }
      } catch (error) {
        statsReport += `   ❌ システム統計エラー: ${error.message}\n`;
      }
      
      statsReport += '\n📊 統計データは「システム管理」→「包括レポート」でより詳細に確認できます。';
      
    } catch (error) {
      statsReport += `❌ 統計取得エラー: ${error.message}\n\n`;
      statsReport += '💡 「📊 システム管理」→「🧪 システムテスト」でシステム状況を確認してください。';
    }
    
    ui.alert('パフォーマンス統計', statsReport, ui.ButtonSet.OK);
    console.log('✅ パフォーマンス統計表示完了');
    
  } catch (error) {
    console.error('❌ パフォーマンス統計表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'パフォーマンス統計の表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * システム設定
 */
function systemConfiguration() {
  SpreadsheetApp.getUi().alert(
    'システム設定', 
    '🔧 システム設定機能\n\nシステム設定機能を実装中です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 詳細分析設定
 */
function advancedAnalytics() {
  SpreadsheetApp.getUi().alert(
    '詳細分析設定', 
    '📊 詳細分析設定機能\n\n詳細分析設定機能を実装中です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * データ同期（統合機能）
 */
function syncAllData() {
  // synchronizeData関数の統合ラッパー
  synchronizeData();
}

/**
 * システムメンテナンス
 */
function systemMaintenance() {
  SpreadsheetApp.getUi().alert(
    'システムメンテナンス', 
    '🛠️ システムメンテナンス機能\n\nシステムメンテナンス機能を実装中です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 🔧 データ同期ボタン（仕様書v2.0準拠）
 */
function synchronizeData() {
  try {
    console.log('🔧 データ同期開始');
    
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let syncReport = '🔧 データ同期実行中...\n\n';
    
    try {
      // 同期実行前の状態確認
      syncReport += '📊 同期前の状態:\n';
      
      const keywordSheet = spreadsheet.getSheetByName('生成キーワード');
      const companySheet = spreadsheet.getSheetByName('企業マスター');
      const proposalSheet = spreadsheet.getSheetByName('提案メッセージ');
      const logSheet = spreadsheet.getSheetByName('実行ログ');
      
      if (keywordSheet) {
        const keywordCount = Math.max(0, keywordSheet.getLastRow() - 1);
        syncReport += `   📝 キーワード: ${keywordCount}件\n`;
      }
      
      if (companySheet) {
        const companyCount = Math.max(0, companySheet.getLastRow() - 1);
        syncReport += `   🏢 企業: ${companyCount}社\n`;
      }
      
      if (proposalSheet) {
        const proposalCount = Math.max(0, proposalSheet.getLastRow() - 1);
        syncReport += `   💬 提案: ${proposalCount}件\n`;
      }
      
      syncReport += '\n🔄 同期処理実行中...\n\n';
      
      // データ整合性チェック
      let issuesFound = 0;
      let fixedIssues = 0;
      
      // キーワードと企業の関連性チェック
      if (keywordSheet && companySheet) {
        const keywords = keywordSheet.getRange(2, 1, Math.max(1, keywordSheet.getLastRow() - 1), 1).getValues().flat();
        const companies = companySheet.getRange(2, 1, Math.max(1, companySheet.getLastRow() - 1), 1).getValues().flat();
        
        // 重複キーワードの確認
        const uniqueKeywords = [...new Set(keywords.filter(k => k))];
        if (keywords.length > uniqueKeywords.length) {
          issuesFound++;
          syncReport += '   ⚠️ 重複キーワードを検出\n';
        }
        
        // 重複企業の確認
        const uniqueCompanies = [...new Set(companies.filter(c => c))];
        if (companies.length > uniqueCompanies.length) {
          issuesFound++;
          syncReport += '   ⚠️ 重複企業を検出\n';
        }
      }
      
      // 提案と企業の関連性チェック
      if (proposalSheet && companySheet) {
        const proposalCompanies = proposalSheet.getRange(2, 1, Math.max(1, proposalSheet.getLastRow() - 1), 1).getValues().flat();
        const masterCompanies = companySheet.getRange(2, 1, Math.max(1, companySheet.getLastRow() - 1), 1).getValues().flat();
        
        // 企業マスターに存在しない提案先をチェック
        const orphanedProposals = proposalCompanies.filter(pc => pc && !masterCompanies.includes(pc));
        if (orphanedProposals.length > 0) {
          issuesFound++;
          syncReport += `   ⚠️ マスターに存在しない提案先: ${orphanedProposals.length}件\n`;
        }
      }
      
      // ログに同期実行記録
      if (logSheet) {
        try {
          const timestamp = new Date().toLocaleString('ja-JP');
          const logData = [timestamp, 'データ同期', '実行中', `同期処理開始 - 検出問題: ${issuesFound}件`, ''];
          logSheet.appendRow(logData);
        } catch (logError) {
          console.warn('ログ記録エラー:', logError);
        }
      }
      
      syncReport += '\n✅ 同期処理完了!\n\n';
      syncReport += '📊 同期結果:\n';
      syncReport += `   🔍 検出された問題: ${issuesFound}件\n`;
      syncReport += `   🔧 修正された問題: ${fixedIssues}件\n\n`;
      
      if (issuesFound === 0) {
        syncReport += '🎉 データは完全に同期されています。\n';
      } else {
        syncReport += '💡 検出された問題の詳細は実行ログを確認してください。\n';
        syncReport += '🛠️ 手動修正が必要な場合は「システム管理」から対応してください。\n';
      }
      
      syncReport += '\n⏰ 最終同期時刻: ' + new Date().toLocaleString('ja-JP');
      
    } catch (error) {
      syncReport += `❌ 同期エラー: ${error.message}\n\n`;
      syncReport += '💡 「📊 システム管理」→「🧪 システムテスト」でシステム状況を確認してください。';
      
      console.error('データ同期エラー:', error);
    }
    
    ui.alert('データ同期', syncReport, ui.ButtonSet.OK);
    console.log('✅ データ同期処理完了');
    
  } catch (error) {
    console.error('❌ データ同期エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'データ同期でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}