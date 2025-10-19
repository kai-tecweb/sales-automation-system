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
    
    // 統合権限確認（ユーザー権限 + プラン権限）- Phase 0統合ポイント
    let isAdmin = false;
    let effectivePermissions = null;
    
    try {
      // 統合権限取得を試行
      if (typeof getEffectivePermissions === 'function') {
        effectivePermissions = getEffectivePermissions();
        isAdmin = effectivePermissions.canAccessAdminFeatures;
      } else {
        // フォールバック：既存のユーザー管理のみ
        const currentUser = getCurrentUser();
        isAdmin = currentUser.role === 'Administrator';
      }
    } catch (error) {
      console.log('権限管理未初期化 - 一般ユーザーとして処理');
    }
    
    // Phase 0: プラン統合メニュータイトル
    let menuTitle = '🚀 営業システム';
    if (effectivePermissions && effectivePermissions.planDisplayName) {
      menuTitle = `🚀 営業システム (${effectivePermissions.planDisplayName})`;
    }
    
    const mainMenu = ui.createMenu(menuTitle);
    
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
        .addItem('📊 利用統計', 'showUsageStatistics')
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
 * 🧪 システムテスト（仕様書準拠）
 */
function runSystemTest() {
  try {
    console.log('🧪 システムテスト開始（仕様書v2.0準拠）');
    
    let testResults = '🧪 システムテスト結果\n\n';
    
    // 1. シート存在確認
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ'];
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    testResults += '📋 シート確認:\n';
    requiredSheets.forEach(sheetName => {
      const exists = sheetNames.includes(sheetName);
      testResults += `${exists ? '✅' : '❌'} ${sheetName}\n`;
    });
    
    // 2. API設定確認
    testResults += '\n🔑 API設定確認:\n';
    const properties = PropertiesService.getScriptProperties().getProperties();
    testResults += `${properties.OPENAI_API_KEY ? '✅' : '❌'} OpenAI API\n`;
    testResults += `${properties.GOOGLE_SEARCH_API_KEY ? '✅' : '❌'} Google Search API\n`;
    testResults += `${properties.GOOGLE_SEARCH_ENGINE_ID ? '✅' : '❌'} Search Engine ID\n`;
    
    // 3. 仕様書準拠確認
    testResults += '\n📖 仕様書準拠確認:\n';
    if (typeof generateComplianceReport === 'function') {
      const compliance = validateMenuComplianceWithSpec();
      testResults += `${compliance.compliant ? '✅' : '❌'} メニュー構造準拠\n`;
    } else {
      testResults += '❌ 準拠チェック機能未実装\n';
    }
    
    SpreadsheetApp.getUi().alert('システムテスト', testResults, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ システムテストエラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'システムテストでエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 🔑 APIキーテスト（仕様書準拠）
 */
function testApiKeys() {
  try {
    console.log('🔑 APIキーテスト開始');
    
    // 既存のtestBasicAPI()を呼び出し（実装済み機能を活用）
    if (typeof testBasicAPI === 'function') {
      testBasicAPI();
    } else {
      SpreadsheetApp.getUi().alert(
        'APIキーテスト', 
        '❌ APIテスト機能が実装されていません。\n\n「⚙️ 設定」→「🔑 APIキー設定」で設定してください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
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

function viewActivityLog() {
  SpreadsheetApp.getUi().alert(
    '開発予定機能', 
    '📋 活動ログ\n\n仕様書v2.0で定義されていますが、まだ実装されていません。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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

/**
 * システム設定（プラン管理統合）- Phase 0
 */
function systemConfiguration() {
  try {
    console.log('🔧 システム設定表示開始');
    
    // 統合情報を取得
    let systemInfo = '🚀 営業自動化システム設定\n\n';
    
    // ライセンス情報
    try {
      if (typeof getLicenseInfo === 'function') {
        const licenseInfo = getLicenseInfo();
        systemInfo += '📋 ライセンス情報:\n';
        systemInfo += `・管理者モード: ${licenseInfo.adminMode ? '✅ 有効' : '🔴 無効'}\n`;
        systemInfo += `・システム状態: ${licenseInfo.systemLocked ? '🔒 ロック中' : '✅ 利用可能'}\n`;
        
        if (licenseInfo.remainingDays !== null) {
          systemInfo += `・残り日数: ${licenseInfo.remainingDays}営業日\n`;
        }
        systemInfo += '\n';
      }
    } catch (error) {
      systemInfo += '📋 ライセンス情報: 取得エラー\n\n';
    }
    
    // プラン情報
    try {
      if (typeof getPlanDetails === 'function') {
        const planDetails = getPlanDetails();
        systemInfo += '💰 プラン情報:\n';
        systemInfo += `・現在のプラン: ${planDetails.displayName}\n`;
        systemInfo += `・月額料金: ¥${planDetails.limits.monthlyPrice.toLocaleString()}\n`;
        systemInfo += `・キーワード生成: ${planDetails.limits.keywordGeneration ? '✅' : '❌'}\n`;
        systemInfo += `・企業検索上限: ${planDetails.limits.maxCompaniesPerDay}社/日\n`;
        systemInfo += `・AI提案生成: ${planDetails.limits.aiProposals ? '✅' : '❌'}\n`;
        
        if (planDetails.isTemporary) {
          systemInfo += `🔄 一時切り替えモード中\n`;
        }
        systemInfo += '\n';
      } else {
        systemInfo += '💰 プラン情報: プラン管理システム未初期化\n\n';
      }
    } catch (error) {
      systemInfo += '💰 プラン情報: 取得エラー\n\n';
    }
    
    // ユーザー情報
    try {
      if (typeof getCurrentUser === 'function') {
        const currentUser = getCurrentUser();
        systemInfo += '👤 ユーザー情報:\n';
        systemInfo += `・ログイン状態: ${currentUser.isLoggedIn ? '✅ ログイン中' : '🔴 未ログイン'}\n`;
        systemInfo += `・ユーザーロール: ${currentUser.role || 'Guest'}\n`;
        systemInfo += '\n';
      }
    } catch (error) {
      systemInfo += '👤 ユーザー情報: ユーザー管理システム未初期化\n\n';
    }
    
    systemInfo += '🔧 管理機能:\n';
    systemInfo += '・詳細な設定は管理者専用メニューをご利用ください\n';
    systemInfo += '・プラン変更は管理者認証後に可能です';
    
    SpreadsheetApp.getUi().alert(
      'システム設定', 
      systemInfo, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ システム設定表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'システム設定の表示でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}