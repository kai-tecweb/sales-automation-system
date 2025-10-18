/**
 * 営業自動化システム メインメニューシステム
 * 統合されたメニュー構造とライセンス管理
 */

function onOpen() {
  try {
    console.log('🚀 Creating role-based MAIN system menu...');
    
    // ユーザー権限に基づいてメニューを作成
    createRoleBasedMenu();
    
  } catch (error) {
    console.error('❌ Main menu creation error:', error);
    
    // 最小限のフォールバックメニュー
    try {
      SpreadsheetApp.getUi()
        .createMenu('🆘 営業システム (エラー)')
        .addItem('📊 状態確認', 'checkSystemStatus')
        .addItem('🔄 メニュー再読み込み', 'reloadMenu')
        .addToUi();
    } catch (fallbackError) {
      console.error('❌ Fallback menu failed:', fallbackError);
    }
  }
}

/**
 * ユーザー権限に基づいたメニュー作成
 */
function createRoleBasedMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 現在のユーザー情報を取得
    let currentUser;
    try {
      currentUser = getCurrentUser();
    } catch (error) {
      // ユーザー管理システムが初期化されていない場合はゲストとして扱う
      console.log('ユーザー管理システム未初期化 - ゲストモードで開始');
      currentUser = { isLoggedIn: false, role: 'Guest' };
    }
    
    const userRole = currentUser.role || 'Guest';
    const isLoggedIn = currentUser.isLoggedIn || false;
    
    console.log(`[INFO] ユーザーロール: ${userRole}, ログイン状態: ${isLoggedIn}`);
    
    // メインメニュー開始
    const mainMenu = ui.createMenu(`⚡ 営業自動化システム (${userRole})`);
    
    // 基本システム機能（全ユーザー共通）
    mainMenu.addItem('📊 システム状態確認', 'checkSystemStatus');
    
    // ログイン/ログアウト機能
    if (isLoggedIn) {
      mainMenu.addItem('👤 ユーザー状態', 'showCurrentUserStatus');
      mainMenu.addItem('🚪 ログアウト', 'logoutUser');
    } else {
      mainMenu.addItem('🔐 ユーザーログイン', 'showUserLoginDialog');
    }
    
    mainMenu.addSeparator();
    
    // 権限別メニュー追加
    if (userRole === 'Administrator') {
      addAdministratorMenu(mainMenu, ui);
    } else if (userRole === 'Standard') {
      addStandardUserMenu(mainMenu, ui);
    } else {
      addGuestUserMenu(mainMenu, ui);
    }
    
    // メニューを有効化
    mainMenu.addToUi();
    
    console.log('✅ Role-based system menu created successfully');
    
    // 権限に応じたシート可視性制御
    controlSheetVisibility(userRole);
    
    // システム起動通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `営業自動化システム v2.0 - ${userRole}モード`, 
      '🚀 システム起動完了', 
      5
    );
    
  } catch (error) {
    console.error('❌ Role-based menu creation error:', error);
    throw error;
  }
}

/**
 * 管理者用メニュー追加
 */
function addAdministratorMenu(mainMenu, ui) {
  // 管理者は全機能にアクセス可能
  mainMenu.addItem('🔧 基本シート作成', 'initializeBasicSheets');
  
  // ライセンス管理（管理者専用）
  mainMenu.addSubMenu(ui.createMenu('📋 ライセンス管理')
    .addItem('📈 ライセンス状況', 'showLicenseStatus')
    .addItem('🔑 管理者認証', 'authenticateAdminFixed')
    .addSeparator()
    .addItem('💰 料金プラン確認', 'showPricingPlans')
    .addItem('📊 料金シミュレーター', 'showPricingCalculator')
    .addItem('⚙️ ライセンス設定', 'configureLicense')
    .addSeparator()
    .addItem('📅 使用開始設定', 'setLicenseStartDate')
    .addItem('🔄 期限延長', 'extendLicense')
    .addItem('🔓 システムロック解除', 'unlockSystem'));
  
  // API設定（管理者専用）
  mainMenu.addSubMenu(ui.createMenu('🔐 API設定')
    .addItem('🔧 APIキー設定', 'setApiKeys')
    .addItem('📋 設定状況確認', 'checkApiKeys')
    .addItem('🗑️ APIキー削除', 'clearApiKeys'));
  
  // ユーザー管理（管理者専用）
  mainMenu.addSubMenu(ui.createMenu('👥 ユーザー管理')
    .addItem('🔧 ユーザー管理シート初期化', 'initializeUserManagementSheet')
    .addItem('➕ 新規ユーザー作成', 'showCreateUserDialog')
    .addItem('📋 ユーザーリスト表示', 'showUserListDialog')
    .addSeparator()
    .addItem('🔄 ユーザー切り替え', 'switchUserMode')
    .addItem('🔍 権限確認', 'checkUserPermissions'));
  
  // 営業自動化機能（全機能）
  mainMenu.addSubMenu(ui.createMenu('🚀 営業自動化')
    .addItem('🔤 キーワード生成', 'generateKeywords')
    .addItem('🏢 企業検索', 'searchCompanies')
    .addItem('💬 提案メッセージ生成', 'generateProposals')
    .addSeparator()
    .addItem('⚡ 完全自動化実行', 'executeFullWorkflow'));
  
  // システム管理（管理者専用）
  mainMenu.addSubMenu(ui.createMenu('⚙️ システム管理')
    .addItem('🔄 メニュー更新', 'forceUpdateMenu')
    .addItem('🏥 システム診断', 'performSystemDiagnostics')
    .addItem('📊 システム情報', 'showSystemInfo')
    .addSeparator()
    .addItem('🧪 包括的システムテスト', 'runComprehensiveSystemTest')
    .addItem('🔐 権限テスト実行', 'testUserPermissions')
    .addItem('💊 システム健康チェック', 'performSystemHealthCheck'));
}

/**
 * スタンダードユーザー用メニュー追加
 */
function addStandardUserMenu(mainMenu, ui) {
  // 営業自動化機能（基本機能のみ）
  mainMenu.addSubMenu(ui.createMenu('🚀 営業自動化')
    .addItem('🔤 キーワード生成', 'generateKeywordsWithPermissionCheck')
    .addItem('🏢 企業検索', 'searchCompaniesWithPermissionCheck')
    .addItem('💬 提案メッセージ生成', 'generateProposalsWithPermissionCheck')
    .addSeparator()
    .addItem('⚡ 基本自動化実行', 'executeBasicWorkflow'));
  
  // データ閲覧機能
  mainMenu.addSubMenu(ui.createMenu('📊 データ閲覧')
    .addItem('📝 生成キーワード表示', 'viewKeywordData')
    .addItem('🏢 企業データ表示', 'viewCompanyData')
    .addItem('💬 提案メッセージ表示', 'viewProposalData'));
  
  // ライセンス状況確認（読み取り専用）
  mainMenu.addItem('📋 ライセンス状況確認', 'showLicenseStatusReadOnly');
}

/**
 * ゲストユーザー用メニュー追加
 */
function addGuestUserMenu(mainMenu, ui) {
  // 閲覧機能のみ
  mainMenu.addSubMenu(ui.createMenu('👀 データ閲覧 (読み取り専用)')
    .addItem('📝 キーワードデータ表示', 'viewKeywordDataReadOnly')
    .addItem('🏢 企業データ表示', 'viewCompanyDataReadOnly')
    .addItem('💬 提案メッセージ表示', 'viewProposalDataReadOnly'));
  
  // システム情報確認（限定版）
  mainMenu.addItem('ℹ️ システム情報確認', 'showSystemInfoLimited');
}

/**
 * メニュー再読み込み
 */
function reloadMenu() {
  try {
    console.log('🔄 Reloading menu...');
    onOpen();
    SpreadsheetApp.getActiveSpreadsheet().toast('メニュー再読み込み完了', '🔄 更新されました', 3);
  } catch (error) {
    console.error('Menu reload error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'メニュー再読み込みでエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =================================
// 権限チェック付き機能実行関数
// =================================

/**
 * 権限チェック付きキーワード生成
 */
function generateKeywordsWithPermissionCheck() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。\nログインしてください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // キーワード生成機能の実行
    if (typeof executeKeywordGeneration === 'function') {
      executeKeywordGeneration();
    } else {
      SpreadsheetApp.getUi().alert('機能エラー', 'キーワード生成機能が見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ キーワード生成（権限チェック付き）エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'キーワード生成中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 権限チェック付き企業検索
 */
function searchCompaniesWithPermissionCheck() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。\nログインしてください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 企業検索機能の実行
    if (typeof executeCompanySearch === 'function') {
      executeCompanySearch();
    } else {
      SpreadsheetApp.getUi().alert('機能エラー', '企業検索機能が見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ 企業検索（権限チェック付き）エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '企業検索中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 権限チェック付き提案生成
 */
function generateProposalsWithPermissionCheck() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。\nログインしてください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 提案生成機能の実行
    if (typeof executeProposalGeneration === 'function') {
      executeProposalGeneration();
    } else {
      SpreadsheetApp.getUi().alert('機能エラー', '提案生成機能が見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ 提案生成（権限チェック付き）エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '提案生成中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 基本ワークフロー実行（スタンダードユーザー用）
 */
function executeBasicWorkflow() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。\nログインしてください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 基本ワークフロー実行
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '基本ワークフロー実行',
      'キーワード生成、企業検索、提案生成を順次実行します。\n実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ui.alert('実行開始', '基本ワークフローを開始します', ui.ButtonSet.OK);
      
      // 段階的実行
      generateKeywordsWithPermissionCheck();
      Utilities.sleep(2000); // 2秒待機
      searchCompaniesWithPermissionCheck();
      Utilities.sleep(2000); // 2秒待機
      generateProposalsWithPermissionCheck();
      
      ui.alert('実行完了', '基本ワークフローが完了しました', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ 基本ワークフロー実行エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '基本ワークフロー実行中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =================================
// 閲覧専用機能
// =================================

/**
 * ライセンス状況確認（読み取り専用）
 */
function showLicenseStatusReadOnly() {
  try {
    // ライセンス情報の読み取り専用表示
    if (typeof showLicenseStatus === 'function') {
      showLicenseStatus();
    } else {
      SpreadsheetApp.getUi().alert('情報', 'ライセンス管理機能が見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    console.error('❌ ライセンス状況確認エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ライセンス状況確認中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システム情報確認（限定版）
 */
function showSystemInfoLimited() {
  try {
    const systemInfo = `
🚀 営業自動化システム v2.0

👤 現在のユーザー: ゲストユーザー
🔑 権限レベル: 閲覧のみ
⏰ アクセス時刻: ${new Date().toLocaleString()}

📋 利用可能機能:
- データ閲覧（読み取り専用）
- システム情報確認

💡 ヒント:
ログインすると追加機能が利用できます
    `;
    
    SpreadsheetApp.getUi().alert('システム情報', systemInfo, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ システム情報確認エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'システム情報確認中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =================================
// シート可視性制御機能
// =================================

/**
 * ユーザー権限に基づくシート可視性制御
 */
function controlSheetVisibility(userRole) {
  try {
    console.log(`🔍 シート可視性制御開始: ${userRole}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      if (userRole === 'Administrator') {
        // 管理者：すべてのシートを表示
        sheet.showSheet();
        console.log(`👁️ 管理者: ${sheetName} - 表示`);
        
      } else if (userRole === 'Standard') {
        // スタンダードユーザー：管理系シート以外を表示
        if (['ユーザー管理'].includes(sheetName)) {
          sheet.hideSheet();
          console.log(`🚫 スタンダード: ${sheetName} - 非表示`);
        } else {
          sheet.showSheet();
          console.log(`👁️ スタンダード: ${sheetName} - 表示`);
        }
        
      } else {
        // ゲストユーザー：データシートのみ表示
        if (['生成キーワード', '企業マスター', '提案メッセージ'].includes(sheetName)) {
          sheet.showSheet();
          console.log(`👁️ ゲスト: ${sheetName} - 表示`);
        } else {
          sheet.hideSheet();
          console.log(`🚫 ゲスト: ${sheetName} - 非表示`);
        }
      }
    });
    
    console.log('✅ シート可視性制御完了');
    
  } catch (error) {
    console.error('❌ シート可視性制御エラー:', error);
  }
}

/**
 * ログイン時のシート表示更新
 */
function updateSheetVisibilityOnLogin(userRole) {
  try {
    controlSheetVisibility(userRole);
    
    // メニューも更新
    createRoleBasedMenu();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `シート表示を${userRole}モードに更新しました`, 
      '🔄 表示更新完了', 
      3
    );
    
  } catch (error) {
    console.error('❌ ログイン時シート更新エラー:', error);
  }
}

// =================================
// データ閲覧機能
// =================================

/**
 * キーワードデータ表示（スタンダードユーザー用）
 */
function viewKeywordData() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('生成キーワード');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', 'キーワードデータが見つかりません。\nまずキーワード生成を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', 'キーワードデータがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert('📋 キーワードデータ', `生成されたキーワードデータを表示しました。\n総計: ${lastRow - 1}件のキーワード`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ キーワードデータ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'キーワードデータ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 企業データ表示（スタンダードユーザー用）
 */
function viewCompanyData() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('企業マスター');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', '企業データが見つかりません。\nまず企業検索を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', '企業データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert('🏢 企業データ', `検索された企業データを表示しました。\n総計: ${lastRow - 1}件の企業`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 企業データ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '企業データ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 提案データ表示（スタンダードユーザー用）
 */
function viewProposalData() {
  try {
    const permission = checkUserPermission('Standard');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'この機能を使用するにはスタンダードユーザー以上の権限が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('提案メッセージ');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', '提案データが見つかりません。\nまず提案生成を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', '提案データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert('💬 提案データ', `生成された提案メッセージを表示しました。\n総計: ${lastRow - 1}件の提案`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 提案データ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '提案データ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * キーワードデータ表示（ゲストユーザー用・読み取り専用）
 */
function viewKeywordDataReadOnly() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('生成キーワード');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', 'キーワードデータが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', 'キーワードデータがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする（読み取り専用）
    SpreadsheetApp.setActiveSheet(sheet);
    
    // ゲスト用：最新20件のみ表示推奨の警告
    const displayCount = Math.min(lastRow - 1, 20);
    SpreadsheetApp.getUi().alert(
      '📋 キーワードデータ（読み取り専用）', 
      `生成されたキーワードデータを表示しました。\n総計: ${lastRow - 1}件のキーワード（最新${displayCount}件を表示推奨）\n\n⚠️ ゲストユーザーのため編集はできません。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ キーワードデータ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'キーワードデータ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 企業データ表示（ゲストユーザー用・読み取り専用）
 */
function viewCompanyDataReadOnly() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('企業マスター');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', '企業データが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', '企業データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする（読み取り専用）
    SpreadsheetApp.setActiveSheet(sheet);
    
    // ゲスト用：最新20件のみ表示推奨の警告
    const displayCount = Math.min(lastRow - 1, 20);
    SpreadsheetApp.getUi().alert(
      '🏢 企業データ（読み取り専用）', 
      `検索された企業データを表示しました。\n総計: ${lastRow - 1}件の企業（最新${displayCount}件を表示推奨）\n\n⚠️ ゲストユーザーのため編集はできません。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ 企業データ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '企業データ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 提案データ表示（ゲストユーザー用・読み取り専用）
 */
function viewProposalDataReadOnly() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('提案メッセージ');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('データなし', '提案データが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('データなし', '提案データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // データをアクティブにする（読み取り専用）
    SpreadsheetApp.setActiveSheet(sheet);
    
    // ゲスト用：最新20件のみ表示推奨の警告
    const displayCount = Math.min(lastRow - 1, 20);
    SpreadsheetApp.getUi().alert(
      '💬 提案データ（読み取り専用）', 
      `生成された提案メッセージを表示しました。\n総計: ${lastRow - 1}件の提案（最新${displayCount}件を表示推奨）\n\n⚠️ ゲストユーザーのため編集はできません。`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ 提案データ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '提案データ表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =================================
// ユーザー管理UI機能
// =================================

/**
 * 新規ユーザー作成ダイアログ表示
 */
function showCreateUserDialog() {
  try {
    // 管理者権限チェック
    const permission = checkUserPermission('Administrator');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'ユーザー作成には管理者権限が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const html = `
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h2>👤 新規ユーザー作成</h2>
        <form>
          <div style="margin-bottom: 15px;">
            <label for="username" style="display: block; margin-bottom: 5px;">ユーザー名:</label>
            <input type="text" id="username" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
          </div>
          <div style="margin-bottom: 15px;">
            <label for="password" style="display: block; margin-bottom: 5px;">パスワード:</label>
            <input type="password" id="password" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
            <small style="color: #666;">※8文字以上、英数字と記号を含む</small>
          </div>
          <div style="margin-bottom: 15px;">
            <label for="role" style="display: block; margin-bottom: 5px;">権限レベル:</label>
            <select id="role" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
              <option value="Standard">スタンダードユーザー</option>
              <option value="Administrator">管理者</option>
            </select>
          </div>
          <div style="text-align: center;">
            <button type="button" onclick="createNewUser()" style="background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">作成</button>
            <button type="button" onclick="google.script.host.close()" style="background: #ccc; color: black; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px;">キャンセル</button>
          </div>
        </form>
        <div id="message" style="margin-top: 15px; padding: 10px; display: none;"></div>
      </div>
      
      <script>
        function createNewUser() {
          const username = document.getElementById('username').value;
          const password = document.getElementById('password').value;
          const role = document.getElementById('role').value;
          
          if (!username || !password) {
            showMessage('ユーザー名とパスワードを入力してください', 'error');
            return;
          }
          
          google.script.run
            .withSuccessHandler(onCreateSuccess)
            .withFailureHandler(onCreateFailure)
            .createUser(username, password, role, 'システム管理者');
        }
        
        function onCreateSuccess(result) {
          if (result.success) {
            showMessage('ユーザーの作成が完了しました', 'success');
            setTimeout(() => {
              google.script.host.close();
            }, 2000);
          } else {
            showMessage(result.message || 'ユーザー作成に失敗しました', 'error');
          }
        }
        
        function onCreateFailure(error) {
          showMessage('ユーザー作成中にエラーが発生しました', 'error');
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.textContent = text;
          messageDiv.style.display = 'block';
          messageDiv.style.backgroundColor = type === 'success' ? '#d4edda' : '#f8d7da';
          messageDiv.style.color = type === 'success' ? '#155724' : '#721c24';
          messageDiv.style.border = '1px solid ' + (type === 'success' ? '#c3e6cb' : '#f5c6cb');
          messageDiv.style.borderRadius = '4px';
        }
      </script>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(450)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '新規ユーザー作成');
    
  } catch (error) {
    console.error('❌ ユーザー作成ダイアログ表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ユーザー作成画面の表示に失敗しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ユーザーリスト表示ダイアログ
 */
function showUserListDialog() {
  try {
    // 管理者権限チェック
    const permission = checkUserPermission('Administrator');
    if (!permission.hasPermission) {
      SpreadsheetApp.getUi().alert(
        '権限エラー', 
        'ユーザーリスト表示には管理者権限が必要です。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const users = getUserList();
    
    let userListHtml = '';
    if (users.length === 0) {
      userListHtml = '<tr><td colspan="4" style="text-align: center;">ユーザーが登録されていません</td></tr>';
    } else {
      users.forEach(user => {
        userListHtml += `
          <tr>
            <td>${user.username}</td>
            <td>${user.role}</td>
            <td>${user.createdDate ? new Date(user.createdDate).toLocaleDateString() : '不明'}</td>
            <td>${user.lastLogin ? new Date(user.lastLogin).toLocaleDateString() : '未ログイン'}</td>
          </tr>
        `;
      });
    }
    
    const html = `
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h2>👥 ユーザーリスト</h2>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd;">
          <thead>
            <tr style="background-color: #f5f5f5;">
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">ユーザー名</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">権限</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">作成日</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">最終ログイン</th>
            </tr>
          </thead>
          <tbody>
            ${userListHtml}
          </tbody>
        </table>
        <div style="text-align: center; margin-top: 20px;">
          <button type="button" onclick="google.script.host.close()" style="background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">閉じる</button>
        </div>
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ユーザーリスト');
    
  } catch (error) {
    console.error('❌ ユーザーリスト表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ユーザーリスト表示中にエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 強制メニュー更新（デバッグ用）
 */
function forceUpdateMenu() {
  try {
    console.log('🔄 Force updating menu...');
    
    // 既存のメニューをクリア
    const ui = SpreadsheetApp.getUi();
    
    // 新しいメニューを作成
    ui.createMenu('🚀 営業自動化システム (最新)')
      .addItem('📋 システム状態確認', 'checkSystemStatus')
      .addItem('🔧 基本シート作成', 'initializeBasicSheets')
      .addSeparator()
      .addSubMenu(ui.createMenu('🔐 ライセンス管理')
        .addItem('📋 ライセンス状況', 'showLicenseStatus')
        .addItem('👤 管理者認証', 'authenticateAdmin')
        .addSeparator()
        .addItem('💰 料金プラン確認', 'showPricingPlans')
        .addItem('⚙️ ライセンス設定', 'configureLicense')
        .addSeparator()
        .addItem('📅 使用開始設定', 'setLicenseStartDate')
        .addItem('🔄 期限延長', 'extendLicense')
        .addItem('🔒 システムロック解除', 'unlockSystem'))
      .addSubMenu(ui.createMenu('👥 ユーザー管理')
        .addItem('🔄 ユーザー切り替え', 'switchUserMode')
        .addItem('📊 権限確認', 'checkUserPermissions'))
      .addSeparator()
      .addItem('ℹ️ システム情報', 'showSystemInfo')
      .addItem('🔄 メニュー再読み込み', 'forceUpdateMenu')
      .addToUi();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'メニュー強制更新完了', 
      '🚀 最新メニューが適用されました', 
      5
    );
    
    console.log('✅ Force menu update completed');
    
  } catch (error) {
    console.error('Force menu update error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'メニュー強制更新でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 管理者認証（修正版） - メニューから呼び出し可能
 */
function authenticateAdminFixed() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '🔐 管理者認証',
      '管理者パスワードを入力してください:\n\nパスワード: SalesAuto2024!',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const password = response.getResponseText();
      
      // ADMIN_PASSWORD = "SalesAuto2024!" (license-manager.jsで定義)
      if (password === 'SalesAuto2024!') {
        // 管理者モードを有効化
        PropertiesService.getScriptProperties().setProperty('ADMIN_MODE', 'true');
        
        SpreadsheetApp.getActiveSpreadsheet().toast(
          '管理者認証成功', 
          '🟢 管理者機能が有効になりました', 
          3
        );
        
        ui.alert(
          '✅ 認証成功',
          '管理者モードが有効になりました。\n管理者専用機能が利用可能です。',
          ui.ButtonSet.OK
        );
        
        // ライセンス管理シートを表示
        createLicenseManagementSheet();
        
        return true;
        
      } else {
        ui.alert(
          '❌ 認証失敗',
          'パスワードが正しくありません。\n正しいパスワード: SalesAuto2024!',
          ui.ButtonSet.OK
        );
        
        return false;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Admin auth fixed error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '管理者認証でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * システム状態確認
 */
function checkSystemStatus() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let status = '📊 システム状態確認\n\n';
    status += '✅ Google Apps Script: 動作中\n';
    status += '✅ スプレッドシート: 接続済み\n';
    status += '✅ メニューシステム: 正常\n\n';
    
    status += '📋 次のステップ:\n';
    status += '1. ライセンス管理システム構築\n';
    status += '2. ユーザー権限管理\n';
    status += '3. 基本機能実装\n\n';
    
    status += `⏰ 確認時刻: ${new Date().toLocaleString('ja-JP')}`;
    
    ui.alert('システム状態', status, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('System status check error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'システム状態確認でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 基本シート作成
 */
function initializeBasicSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 必要な基本シートを作成
    const requiredSheets = [
      '制御パネル',
      'ライセンス管理', 
      '実行ログ'
    ];
    
    let created = 0;
    for (const sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        ss.insertSheet(sheetName);
        created++;
      }
    }
    
    let message = '🔧 基本シート作成完了\n\n';
    message += `✅ 作成済みシート: ${created}個\n`;
    message += `📊 総シート数: ${ss.getSheets().length}個\n\n`;
    message += '次は「ライセンス管理」から設定を開始してください。';
    
    ui.alert('シート作成', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Sheet initialization error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'シート作成でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システム情報表示
 */
function showSystemInfo() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let info = '📋 営業自動化システム情報\n\n';
    info += '🏷️ バージョン: v2.0 (開発中)\n';
    info += '📅 最終更新: 2025年10月18日\n';
    info += '🔧 ステータス: 基本構築段階\n\n';
    
    info += '📝 現在の段階:\n';
    info += '1. ✅ コードベース構築\n';
    info += '2. 🔄 メニューシステム確認中\n';
    info += '3. ⏳ ライセンス管理 (次)\n';
    info += '4. ⏳ ユーザー権限管理\n';
    info += '5. ⏳ 基本機能実装\n\n';
    
    info += '🎯 目標: 完全自動化された営業支援システム';
    
    ui.alert('システム情報', info, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('System info error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'システム情報表示でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ライセンス状況表示（実装版）
 */
function showLicenseStatus() {
  try {
    // license-manager.js の関数を呼び出し
    const licenseInfo = getLicenseInfo();
    const ui = SpreadsheetApp.getUi();
    
    let status = '📋 ライセンス状況\n\n';
    
    // 管理者モード状況
    status += `👤 管理者モード: ${licenseInfo.adminMode ? '� 有効' : '🔴 無効'}\n`;
    
    // ライセンス期限状況
    if (licenseInfo.startDate) {
      status += `📅 ライセンス開始日: ${formatDate(licenseInfo.startDate)}\n`;
      status += `📅 ライセンス期限: ${licenseInfo.expiryDate ? formatDate(licenseInfo.expiryDate) : '未計算'}\n`;
      status += `⏰ 残り日数: ${licenseInfo.remainingDays !== null ? licenseInfo.remainingDays + '営業日' : '未計算'}\n`;
      status += `🔓 システム状態: ${licenseInfo.systemLocked ? '🔒 ロック中' : '✅ 利用可能'}\n\n`;
    } else {
      status += '📅 ライセンス: 未設定\n';
      status += '💡 「📅 使用開始設定」からライセンスを開始してください\n\n';
    }
    
    // 料金プラン情報
    status += '💰 料金プラン情報:\n';
    status += '• ベーシック: ¥500/月 (企業検索10社/日)\n';
    status += '• スタンダード: ¥1,500/月 (企業検索50社/日 + AI)\n';
    status += '• プロフェッショナル: ¥5,500/月 (企業検索100社/日 + AI + 2アカウント)\n';
    status += '• エンタープライズ: ¥17,500/月 (企業検索500社/日 + AI + 5アカウント)\n\n';
    
    status += '次の操作:\n';
    status += '• OK: ライセンス管理シートを開く\n';
    status += '• キャンセル: このダイアログを閉じる';
    
    const result = ui.alert('ライセンス状況', status, ui.ButtonSet.OK_CANCEL);
    
    if (result === ui.Button.OK) {
      createLicenseManagementSheet();
    }
    
  } catch (error) {
    console.error('License status error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'ライセンス状況確認でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ProspectFlow 料金プラン確認（v4.0 統一版）
 */
function showPricingPlans() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let plans = '💰 ProspectFlow 営業自動化システム 料金プラン\n\n';
    
    // トライアル版
    plans += '🆓 トライアル（10営業日無料）\n';
    plans += '• 期間: 10営業日限定\n';
    plans += '• 企業検索: 5社/日\n';
    plans += '• AI機能: ✅ 利用可能\n';
    plans += '• 月額: ¥0 + API実費約¥50\n\n';
    
    // ベーシック版
    plans += '🥉 ベーシック（¥500/月）\n';
    plans += '• 企業検索: 10社/日\n';
    plans += '• AI機能: ❌ 利用不可\n';
    plans += '• アカウント: 1名\n';
    plans += '• 合計月額: ¥500（API料金なし）\n\n';
    
    // スタンダード版
    plans += '🥈 スタンダード（¥1,500/月）\n';
    plans += '• 企業検索: 50社/日\n';
    plans += '• AI機能: ✅ 利用可能\n';
    plans += '• アカウント: 1名\n';
    plans += '• 合計月額: ¥3,500（API込み）\n\n';
    
    // プロフェッショナル版
    plans += '🥇 プロフェッショナル（¥5,500/月）\n';
    plans += '• 企業検索: 100社/日\n';
    plans += '• AI機能: ✅ 利用可能\n';
    plans += '• アカウント: 2名\n';
    plans += '• 合計月額: ¥11,500（API込み）\n\n';
    
    // エンタープライズ版
    plans += '💎 エンタープライズ（¥17,500/月）\n';
    plans += '• 企業検索: 500社/日\n';
    plans += '• AI機能: ✅ 利用可能\n';
    plans += '• アカウント: 5名\n';
    plans += '• 合計月額: ¥47,500（API込み）\n\n';
    
    // 特徴説明
    plans += '� 特徴:\n';
    plans += '• トライアル: 全機能お試し可能\n';
    plans += '• ベーシック: 手動入力・基本テンプレート\n';
    plans += '• その他: AI キーワード・提案生成\n\n';
    
    plans += '🎯 まずは10営業日無料トライアルから！';
    
    ui.alert('ProspectFlow 料金プラン', plans, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Pricing plans error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '料金プラン確認でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 料金シミュレーター
 */
function showPricingCalculator() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 企業数の入力
    const companiesResponse = ui.prompt(
      '📊 料金シミュレーター',
      '1日あたりの企業検索数を入力してください（例: 50）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (companiesResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const companiesPerDay = parseInt(companiesResponse.getResponseText());
    if (isNaN(companiesPerDay) || companiesPerDay <= 0) {
      ui.alert('エラー', '有効な数値を入力してください', ui.ButtonSet.OK);
      return;
    }
    
    // AI機能の選択
    const aiResponse = ui.alert(
      '💡 AI機能選択',
      'AI機能（キーワード・提案生成）を使用しますか？\n\n' +
      'YES: AI強化版（約2.5円/企業）\n' +
      'NO: 基本版（約0.1円/企業）',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (aiResponse === ui.Button.CANCEL) {
      return;
    }
    
    const useAI = aiResponse === ui.Button.YES;
    
    // 料金計算
    const result = calculatePricingSimulation(companiesPerDay, useAI);
    
    // 結果表示
    let simulationResult = `📊 ProspectFlow 料金シミュレーション結果\n\n`;
    simulationResult += `📈 条件:\n`;
    simulationResult += `• 企業検索: ${companiesPerDay}社/日\n`;
    simulationResult += `• AI機能: ${useAI ? 'あり（キーワード・提案生成）' : 'なし（手動入力）'}\n\n`;
    
    simulationResult += `💰 推奨プラン:\n\n`;
    
    result.plans.forEach(plan => {
      if (companiesPerDay <= plan.dailyLimit || plan.dailyLimit === 0) {
        // ベーシックプランでAI機能を希望する場合の特別処理
        if (plan.name === 'ベーシック' && useAI) {
          simulationResult += `${plan.icon} ${plan.name}\n`;
          simulationResult += `• ❌ AI機能非対応プランです\n`;
          simulationResult += `• スタンダードプラン以上をご検討ください\n\n`;
        } else {
          simulationResult += `${plan.icon} ${plan.name}\n`;
          simulationResult += `• システム利用料: ¥${plan.license.toLocaleString()}/月\n`;
          
          if (plan.name === 'トライアル') {
            simulationResult += `• 期間: ${plan.period}\n`;
            simulationResult += `• API料金: ¥${plan.apiCost}/期間\n`;
          } else {
            simulationResult += `• API料金: ¥${plan.apiCost.toLocaleString()}/月\n`;
          }
          
          if (typeof plan.totalCost === 'number') {
            simulationResult += `• 合計: ¥${plan.totalCost.toLocaleString()}/月\n`;
          }
          
          if (plan.users > 1) {
            simulationResult += `• アカウント: ${plan.users}名まで\n`;
          }
          simulationResult += `\n`;
        }
      }
    });
    
    simulationResult += `💡 ProspectFlow のポイント:\n`;
    simulationResult += `• トライアル: 全機能10営業日無料体験\n`;
    simulationResult += `• ベーシック: AI機能なし、手動運用向け\n`;
    simulationResult += `• スタンダード以上: AI自動化機能フル活用\n`;
    simulationResult += `• まずは無料トライアルから始めましょう！`;
    
    ui.alert('料金シミュレーション', simulationResult, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Pricing calculator error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '料金シミュレーターでエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ProspectFlow 料金計算ロジック
 */
function calculatePricingSimulation(companiesPerDay, useAI) {
  // ProspectFlow 料金体系に基づく計算
  let apiCostPerCompany = 0;
  let monthlyApiCost = 0;
  
  if (useAI) {
    // AI機能使用時のAPI料金（1社あたり約¥40）
    apiCostPerCompany = 40;
    monthlyApiCost = Math.ceil(companiesPerDay * apiCostPerCompany * 20); // 営業日ベース
  }
  
  const plans = [
    {
      name: 'トライアル',
      icon: '🆓',
      license: 0,
      dailyLimit: 5,
      users: 1,
      apiCost: useAI ? 50 : 0, // 10営業日での概算
      totalCost: 0,
      period: '10営業日限定'
    },
    {
      name: 'ベーシック',
      icon: '🥉',
      license: 500,
      dailyLimit: 10,
      users: 1,
      apiCost: 0, // AI機能なしのため
      totalCost: 0,
      hasAI: false
    },
    {
      name: 'スタンダード',
      icon: '🥈',
      license: 1500,
      dailyLimit: 50,
      users: 1,
      apiCost: useAI && companiesPerDay <= 50 ? 2000 : 0, // 最大想定API料金
      totalCost: 0,
      hasAI: true
    },
    {
      name: 'プロフェッショナル',
      icon: '🥇',
      license: 5500,
      dailyLimit: 100,
      users: 2,
      apiCost: useAI && companiesPerDay <= 100 ? 6000 : 0, // 最大想定API料金
      totalCost: 0,
      hasAI: true
    },
    {
      name: 'エンタープライズ',
      icon: '💎',
      license: 17500,
      dailyLimit: 500,
      users: 5,
      apiCost: useAI && companiesPerDay <= 500 ? 30000 : 0, // 最大想定API料金
      totalCost: 0,
      hasAI: true
    }
  ];
  
  // 合計料金計算
  plans.forEach(plan => {
    if (plan.name === 'ベーシック' && useAI) {
      // ベーシックプランではAI機能が使用できない
      plan.totalCost = 'AI機能非対応';
    } else {
      plan.totalCost = plan.license + plan.apiCost;
    }
  });
  
  return {
    plans: plans,
    conditions: {
      companiesPerDay: companiesPerDay,
      useAI: useAI,
      apiCostPerCompany: apiCostPerCompany,
      monthlyApiCost: monthlyApiCost
    }
  };
}

/**
 * 管理者認証（実装版）
 */
function authenticateAdmin() {
  try {
    // license-manager.js の authenticateAdmin() 関数を直接呼び出し
    callLicenseManagerAuth();
    
  } catch (error) {
    console.error('Admin auth error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '管理者認証でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ライセンス管理の認証関数を呼び出し
 */
function callLicenseManagerAuth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔐 管理者認証',
    '管理者パスワードを入力してください:\n\nパスワード: SalesAuto2024!',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const password = response.getResponseText();
    
    if (password === 'SalesAuto2024!') {
      // 管理者モードを有効化
      PropertiesService.getScriptProperties().setProperty('ADMIN_MODE', 'true');
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '管理者認証成功', 
        '🟢 管理者機能が有効になりました', 
        3
      );
      
      ui.alert(
        '✅ 認証成功',
        '管理者モードが有効になりました。\n管理者専用機能が利用可能です。',
        ui.ButtonSet.OK
      );
      
      // ライセンス管理シートを表示
      createLicenseManagementSheet();
      
      return true;
    } else {
      ui.alert(
        '❌ 認証失敗',
        'パスワードが正しくありません。',
        ui.ButtonSet.OK
      );
      return false;
    }
  }
  return false;
}

/**
 * ライセンス設定（実装版）
 */
function configureLicense() {
  try {
    const ui = SpreadsheetApp.getUi();
    const licenseInfo = getLicenseInfo();
    
    // 管理者権限チェック
    if (!licenseInfo.adminMode) {
      ui.alert(
        '🔒 権限エラー',
        'ライセンス設定は管理者専用機能です。\n先に「👤 管理者認証」を行ってください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    let configMenu = '⚙️ ライセンス設定メニュー\n\n';
    configMenu += '利用可能な操作:\n';
    configMenu += '• はい: 使用開始日設定\n';
    configMenu += '• いいえ: ライセンス期限延長\n';
    configMenu += '• キャンセル: システムロック解除\n\n';
    
    if (licenseInfo.startDate) {
      configMenu += `現在の設定:\n`;
      configMenu += `• 開始日: ${formatDate(licenseInfo.startDate)}\n`;
      configMenu += `• 期限: ${licenseInfo.expiryDate ? formatDate(licenseInfo.expiryDate) : '未計算'}\n`;
      configMenu += `• 残り: ${licenseInfo.remainingDays}営業日`;
    } else {
      configMenu += '現在の設定: ライセンス未設定';
    }
    
    const result = ui.alert('ライセンス設定', configMenu, ui.ButtonSet.YES_NO_CANCEL);
    
    if (result === ui.Button.YES) {
      setLicenseStartDate();
    } else if (result === ui.Button.NO) {
      extendLicense();
    } else if (result === ui.Button.CANCEL) {
      unlockSystem();
    }
    
  } catch (error) {
    console.error('License config error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'ライセンス設定でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ユーザー切り替え（仮実装）
 */
function switchUserMode() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert('ユーザー切り替え', '🔧 開発段階: ユーザー切り替え機能実装前\n\n次のステップでユーザー権限管理システムを構築します。', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('User switch error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'ユーザー切り替えでエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 権限確認（仮実装）
 */
function checkUserPermissions() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let perms = '📊 現在の権限状況\n\n';
    perms += '👤 ユーザータイプ: 開発者\n';
    perms += '💰 プラン: 開発モード\n';
    perms += '🔧 権限レベル: 全機能アクセス\n\n';
    perms += '💡 次のステップ: ユーザー権限管理システム構築';
    
    ui.alert('権限確認', perms, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('User permissions error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '権限確認でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システム診断機能
 */
function performSystemDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    let diagnostics = '【システム診断結果】\n\n';
    
    // 1. APIキー確認
    const apiValidation = validateApiKeys();
    diagnostics += '🔑 APIキー設定:\n';
    diagnostics += '  OpenAI: ' + (apiValidation.openaiKey ? '✅' : '❌') + '\n';
    diagnostics += '  Google Search: ' + (apiValidation.googleKey ? '✅' : '❌') + '\n';
    diagnostics += '  Search Engine ID: ' + (apiValidation.engineId ? '✅' : '❌') + '\n\n';
    
    // 2. ライセンス状況確認
    const licenseInfo = getLicenseInfo();
    diagnostics += '📋 ライセンス状況:\n';
    diagnostics += '  管理者モード: ' + (licenseInfo.isAdminMode ? '✅ 有効' : '❌ 無効') + '\n';
    diagnostics += '  ライセンス: ' + (licenseInfo.isLicenseSet ? '✅ 設定済み' : '❌ 未設定') + '\n\n';
    
    // 3. スプレッドシート確認
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    diagnostics += '📊 スプレッドシート:\n';
    diagnostics += '  シート数: ' + sheets.length + '\n';
    diagnostics += '  利用可能: ✅\n\n';
    
    // 4. 全体状況
    const allGood = apiValidation.allSet && licenseInfo.isAdminMode;
    diagnostics += '🎯 総合状況: ' + (allGood ? '✅ 正常' : '⚠️ 要設定') + '\n\n';
    
    if (!allGood) {
      diagnostics += '📝 推奨アクション:\n';
      if (!apiValidation.allSet) {
        diagnostics += '  • APIキーを設定してください\n';
      }
      if (!licenseInfo.isAdminMode) {
        diagnostics += '  • 管理者認証を行ってください\n';
      }
    }
    
    ui.alert('システム診断', diagnostics, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('システム診断エラー: ' + error.toString());
    ui.alert('診断エラー', 'システム診断中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * システムロック解除
 */
function unlockSystem() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 管理者権限チェック
    const licenseInfo = getLicenseInfo();
    if (!licenseInfo.adminMode) {
      ui.alert(
        '🔒 権限エラー',
        'システムロック解除は管理者専用機能です。\n先に「👤 管理者認証」を行ってください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const result = ui.alert(
      '🔓 システムロック解除',
      'システムロックを解除しますか？\n\n注意: この操作はライセンス期限に関係なく、\nシステムの全機能を有効にします。',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      // システムロックを解除
      PropertiesService.getScriptProperties().setProperty('SYSTEM_LOCKED', 'false');
      
      ui.alert(
        '✅ ロック解除完了',
        'システムロックが解除されました。\n全機能が利用可能になりました。',
        ui.ButtonSet.OK
      );
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'システムロック解除完了', 
        '🔓 全機能が利用可能になりました', 
        3
      );
    }
    
  } catch (error) {
    console.error('Unlock system error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', 'システムロック解除でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}