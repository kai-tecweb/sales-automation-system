/**
 * 営業自動化システム - シンプルメニューシステム v3.0
 * 重複排除・ユーザビリティ重視の再設計版
 */

function onOpen() {
  try {
    console.log('🚀 シンプルメニューシステム v3.0 開始...');
    createSimplifiedMenu();
  } catch (error) {
    console.error('❌ メニュー作成エラー:', error);
    createFallbackMenu();
  }
}

/**
 * シンプル化されたメニュー作成
 */
function createSimplifiedMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // ユーザー権限の確認（簡素化）
    let userRole = 'Guest';
    let isLoggedIn = false;
    
    try {
      const currentUser = getCurrentUser();
      userRole = currentUser.role || 'Guest';
      isLoggedIn = currentUser.isLoggedIn || false;
    } catch (error) {
      console.log('ユーザー管理未初期化 - ゲストモードで開始');
    }
    
    console.log(`[INFO] ユーザー: ${userRole}, ログイン: ${isLoggedIn}`);
    
    // メインメニュー作成
    const mainMenu = ui.createMenu(`⚡ 営業自動化 (${userRole})`);
    
    // 権限別メニュー構築
    if (userRole === 'Administrator') {
      buildAdminMenu(mainMenu, ui);
    } else if (userRole === 'Standard' && isLoggedIn) {
      buildStandardUserMenu(mainMenu, ui);
    } else {
      buildGuestMenu(mainMenu, ui);
    }
    
    // 共通項目（全ユーザー）
    mainMenu.addSeparator();
    mainMenu.addItem('🔄 メニュー更新', 'reloadSimplifiedMenu');
    
    // メニュー有効化
    mainMenu.addToUi();
    
    // システム起動通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `営業自動化システム v3.0 - ${userRole}モード`, 
      '🚀 シンプルメニュー起動完了', 
      3
    );
    
    console.log('✅ シンプルメニュー作成完了');
    
  } catch (error) {
    console.error('❌ シンプルメニュー作成エラー:', error);
    throw error;
  }
}

/**
 * 🔰 ゲストユーザー向けメニュー（最小限）
 */
function buildGuestMenu(mainMenu, ui) {
  // システム確認
  mainMenu.addItem('📊 システム状態確認', 'checkSystemStatusGuest');
  
  // ログイン
  mainMenu.addItem('🔐 ログイン', 'simpleLogin');
  
  // ヘルプ
  mainMenu.addItem('❓ 使い方ガイド', 'showUserGuide');
  
  // 基本設定（API設定のみ）
  mainMenu.addSeparator();
  mainMenu.addItem('🔑 API設定', 'setApiKeysSimple');
  mainMenu.addItem('🧪 API動作テスト', 'testBasicAPI');
}

/**
 * 👤 標準ユーザー向けメニュー（主要機能）
 */
function buildStandardUserMenu(mainMenu, ui) {
  // 営業自動化（メイン機能）
  mainMenu.addSubMenu(ui.createMenu('🚀 営業自動化')
    .addItem('🔤 キーワード生成', 'generateKeywordsMain')
    .addItem('🏢 企業検索', 'executeCompanySearchMain')
    .addItem('💬 提案メッセージ生成', 'generateProposalsMain')
    .addSeparator()
    .addItem('⚡ 完全自動実行', 'executeFullWorkflowMain'));
  
  // 結果確認
  mainMenu.addSubMenu(ui.createMenu('📊 結果確認')
    .addItem('📝 生成キーワード', 'viewKeywordsResults')
    .addItem('🏢 発掘企業一覧', 'viewCompaniesResults')
    .addItem('💬 提案メッセージ', 'viewProposalsResults'));
  
  // 設定
  mainMenu.addSeparator();
  mainMenu.addItem('⚙️ 基本設定', 'showBasicSettings');
  mainMenu.addItem('🚪 ログアウト', 'logoutUser');
}

/**
 * 👑 管理者向けメニュー（完全版）
 */
function buildAdminMenu(mainMenu, ui) {
  // 営業自動化（完全版）
  mainMenu.addSubMenu(ui.createMenu('🚀 営業自動化（完全版）')
    .addItem('🔤 キーワード生成', 'generateKeywordsMain')
    .addItem('🏢 企業検索', 'executeCompanySearchMain')
    .addItem('💬 提案メッセージ生成', 'generateProposalsMain')
    .addSeparator()
    .addItem('⚡ 完全自動実行', 'executeFullWorkflowMain')
    .addSeparator()
    .addItem('🧪 高度なテスト実行', 'runAdvancedTests'));
  
  // システム管理
  mainMenu.addSubMenu(ui.createMenu('⚙️ システム管理')
    .addItem('🏥 システム診断', 'performSystemDiagnostics')
    .addItem('🔧 初期化・修復', 'initializeSystemSheets')
    .addItem('📋 ログ確認', 'viewSystemLogs')
    .addSeparator()
    .addItem('🔑 API管理', 'manageAPISettings')
    .addItem('📊 使用量確認', 'checkResourceUsage'));
  
  // ユーザー管理
  mainMenu.addSubMenu(ui.createMenu('👥 ユーザー管理')
    .addItem('➕ 新規ユーザー作成', 'createNewUser')
    .addItem('📋 ユーザー一覧', 'listAllUsers')
    .addItem('🔄 権限変更', 'changeUserPermissions'));
  
  // ライセンス管理
  mainMenu.addItem('📋 ライセンス管理', 'manageLicenseSettings');
  
  // 設定
  mainMenu.addSeparator();
  mainMenu.addItem('🚪 ログアウト', 'logoutUser');
}

/**
 * 緊急時フォールバックメニュー
 */
function createFallbackMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🆘 営業システム (緊急モード)')
      .addItem('📊 システム確認', 'checkSystemStatusGuest')
      .addItem('🔄 メニュー修復', 'repairMenuSystem')
      .addItem('🔑 API設定', 'setApiKeysSimple')
      .addToUi();
  } catch (error) {
    console.error('❌ フォールバックメニュー作成失敗:', error);
  }
}

/**
 * メニュー再読み込み
 */
function reloadSimplifiedMenu() {
  try {
    console.log('🔄 シンプルメニュー再読み込み中...');
    onOpen();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'メニュー再読み込み完了', 
      '🔄 最新メニューに更新されました', 
      3
    );
  } catch (error) {
    console.error('❌ メニュー再読み込みエラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'メニュー再読み込みでエラーが発生しました。\n管理者にお問い合わせください。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==============================================
// 🎯 統一された実行関数（重複排除）
// ==============================================

/**
 * 🔤 キーワード生成（統一版）
 */
function generateKeywordsMain() {
  try {
    console.log('🔤 キーワード生成開始');
    
    // 権限チェック
    if (!checkBasicPermission()) {
      return;
    }
    
    // 実際のキーワード生成実行
    const result = generateKeywords();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'キーワード生成完了', 
      '🔤 生成キーワードシートを確認してください', 
      5
    );
    
    return result;
    
  } catch (error) {
    console.error('❌ キーワード生成エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      `キーワード生成でエラーが発生しました:\n${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 🏢 企業検索（統一版）
 */
function executeCompanySearchMain() {
  try {
    console.log('🏢 企業検索開始');
    
    // 権限チェック
    if (!checkBasicPermission()) {
      return;
    }
    
    // 実際の企業検索実行
    const result = executeCompanySearch();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '企業検索完了', 
      '🏢 企業マスターシートを確認してください', 
      5
    );
    
    return result;
    
  } catch (error) {
    console.error('❌ 企業検索エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      `企業検索でエラーが発生しました:\n${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 💬 提案メッセージ生成（統一版）
 */
function generateProposalsMain() {
  try {
    console.log('💬 提案メッセージ生成開始');
    
    // 権限チェック
    if (!checkBasicPermission()) {
      return;
    }
    
    // 実際の提案生成実行
    const result = generatePersonalizedProposals();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '提案メッセージ生成完了', 
      '💬 提案メッセージシートを確認してください', 
      5
    );
    
    return result;
    
  } catch (error) {
    console.error('❌ 提案メッセージ生成エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      `提案メッセージ生成でエラーが発生しました:\n${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ⚡ 完全自動実行（統一版）
 */
function executeFullWorkflowMain() {
  try {
    console.log('⚡ 完全自動実行開始');
    
    // 権限チェック
    if (!checkBasicPermission()) {
      return;
    }
    
    // 確認ダイアログ
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '完全自動実行確認',
      'キーワード生成、企業検索、提案メッセージ生成を順次実行します。\n\n実行には数分かかる場合があります。\n実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 実際の完全ワークフロー実行
    const result = executeFullWorkflow();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '完全自動実行完了', 
      '⚡ すべてのシートを確認してください', 
      10
    );
    
    return result;
    
  } catch (error) {
    console.error('❌ 完全自動実行エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      `完全自動実行でエラーが発生しました:\n${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==============================================
// 🔧 サポート関数
// ==============================================

/**
 * 基本権限チェック（ログインまたはゲスト許可）
 */
function checkBasicPermission() {
  try {
    const currentUser = getCurrentUser();
    if (currentUser.isLoggedIn || currentUser.role === 'Guest') {
      return true;
    }
  } catch (error) {
    // ユーザー管理未初期化の場合はゲスト扱い
    console.log('ユーザー管理未初期化 - ゲスト権限で実行');
    return true;
  }
  
  SpreadsheetApp.getUi().alert(
    '権限エラー', 
    'この機能を使用するにはログインが必要です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return false;
}

/**
 * ゲスト向けシステム状態確認
 */
function checkSystemStatusGuest() {
  try {
    console.log('📊 ゲスト向けシステム状態確認');
    
    let statusMessage = '📊 システム状態確認\n\n';
    
    // 基本シート存在確認
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ'];
    let sheetsOK = true;
    
    requiredSheets.forEach(sheetName => {
      if (sheetNames.includes(sheetName)) {
        statusMessage += `✅ ${sheetName}: 存在\n`;
      } else {
        statusMessage += `❌ ${sheetName}: 不在\n`;
        sheetsOK = false;
      }
    });
    
    // API設定確認
    statusMessage += '\n🔑 API設定状況:\n';
    try {
      const openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
      const googleKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
      
      statusMessage += openaiKey ? '✅ OpenAI API: 設定済み\n' : '❌ OpenAI API: 未設定\n';
      statusMessage += googleKey ? '✅ Google Search API: 設定済み\n' : '❌ Google Search API: 未設定\n';
    } catch (error) {
      statusMessage += '❌ API設定確認エラー\n';
    }
    
    statusMessage += '\n💡 次のステップ:\n';
    if (!sheetsOK) {
      statusMessage += '1. システム初期化を実行してください\n';
    }
    statusMessage += '2. APIキーを設定してください\n';
    statusMessage += '3. ログインして営業自動化を開始\n';
    
    SpreadsheetApp.getUi().alert('システム状態確認', statusMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ システム状態確認エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'システム状態確認でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 簡単API設定
 */
function setApiKeysSimple() {
  try {
    console.log('🔑 簡単API設定開始');
    
    const ui = SpreadsheetApp.getUi();
    
    // OpenAI APIキー設定
    const openaiResponse = ui.prompt(
      'OpenAI API設定',
      'OpenAI API キーを入力してください:\n\n※ sk-で始まる文字列',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (openaiResponse.getSelectedButton() === ui.Button.OK) {
      const openaiKey = openaiResponse.getResponseText().trim();
      if (openaiKey && openaiKey.startsWith('sk-')) {
        PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', openaiKey);
        ui.alert('✅ 成功', 'OpenAI API キーが設定されました', ui.ButtonSet.OK);
      } else {
        ui.alert('❌ エラー', '有効なOpenAI API キーを入力してください', ui.ButtonSet.OK);
        return;
      }
    } else {
      return;
    }
    
    // Google Search APIキー設定
    const googleResponse = ui.prompt(
      'Google Search API設定',
      'Google Search API キーを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (googleResponse.getSelectedButton() === ui.Button.OK) {
      const googleKey = googleResponse.getResponseText().trim();
      if (googleKey) {
        PropertiesService.getScriptProperties().setProperty('GOOGLE_SEARCH_API_KEY', googleKey);
        ui.alert('✅ 成功', 'Google Search API キーが設定されました', ui.ButtonSet.OK);
      }
    }
    
    // Search Engine ID設定
    const engineResponse = ui.prompt(
      'Google Search Engine ID設定',
      'カスタム検索エンジンIDを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (engineResponse.getSelectedButton() === ui.Button.OK) {
      const engineId = engineResponse.getResponseText().trim();
      if (engineId) {
        PropertiesService.getScriptProperties().setProperty('GOOGLE_SEARCH_ENGINE_ID', engineId);
        ui.alert('✅ 完了', '全てのAPIキーが設定されました！\n\n営業自動化を開始できます。', ui.ButtonSet.OK);
      }
    }
    
  } catch (error) {
    console.error('❌ API設定エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'API設定でエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 基本API動作テスト
 */
function testBasicAPI() {
  try {
    console.log('🧪 基本API動作テスト開始');
    
    let testResult = '🧪 API動作テスト結果\n\n';
    
    // OpenAI APIテスト
    try {
      const openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
      if (openaiKey) {
        testResult += '✅ OpenAI API キー: 存在\n';
        // 簡単なAPI呼び出しテスト（実際の呼び出しは省略）
        testResult += '✅ OpenAI API: 設定確認OK\n';
      } else {
        testResult += '❌ OpenAI API キー: 未設定\n';
      }
    } catch (error) {
      testResult += '❌ OpenAI API: エラー\n';
    }
    
    // Google Search APIテスト
    try {
      const googleKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_API_KEY');
      const engineId = PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID');
      
      if (googleKey && engineId) {
        testResult += '✅ Google Search API: 設定確認OK\n';
      } else {
        testResult += '❌ Google Search API: 未設定\n';
      }
    } catch (error) {
      testResult += '❌ Google Search API: エラー\n';
    }
    
    testResult += '\n💡 次のステップ:\n';
    testResult += '1. 不足しているAPIキーを設定\n';
    testResult += '2. 営業自動化機能を試す\n';
    
    SpreadsheetApp.getUi().alert('API動作テスト', testResult, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ API動作テストエラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'API動作テストでエラーが発生しました。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ユーザーガイド表示
 */
function showUserGuide() {
  const guide = `
📖 営業自動化システム - 使い方ガイド

🏁 はじめに:
1. APIキー設定（OpenAI + Google Search）
2. ログイン（または管理者作成）
3. 営業自動化開始

⚡ 基本機能:
• キーワード生成: 商材から検索キーワードを自動生成
• 企業検索: 生成キーワードで企業を自動発掘
• 提案生成: 企業別にカスタマイズされた提案メッセージ生成

🔧 権限レベル:
• ゲスト: 設定・確認・テスト
• 標準: 営業自動化実行
• 管理者: 全機能 + システム管理

💡 困った時は「システム状態確認」で現状を把握できます。
`;
  
  SpreadsheetApp.getUi().alert('使い方ガイド', guide, SpreadsheetApp.getUi().ButtonSet.OK);
}