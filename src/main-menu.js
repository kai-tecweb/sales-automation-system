/**
 * 営業自動化システム メインメニューシステム
 * 統合されたメニュー構造とライセンス管理
 */

function onOpen() {
  try {
    console.log('🚀 Creating MAIN system menu...');
    
    const ui = SpreadsheetApp.getUi();
    
    // メインシステムメニュー
    ui.createMenu('🚀 営業自動化システム')
      .addItem('📋 システム状態確認', 'checkSystemStatus')
      .addItem('🔧 基本シート作成', 'initializeBasicSheets')
      .addSeparator()
      .addSubMenu(ui.createMenu('🔐 ライセンス管理')
        .addItem('📋 ライセンス状況', 'showLicenseStatus')
        .addItem('👤 管理者認証', 'authenticateAdminFixed')
        .addSeparator()
        .addItem('💰 料金プラン確認', 'showPricingPlans')
        .addItem('⚙️ ライセンス設定', 'configureLicense')
        .addSeparator()
        .addItem('📅 使用開始設定', 'setLicenseStartDate')
        .addItem('🔄 期限延長', 'extendLicense')
        .addItem('🔒 システムロック解除', 'unlockSystem'))
      .addSubMenu(ui.createMenu('🔑 API設定')
        .addItem('🔧 APIキー設定', 'setApiKeys')
        .addItem('📋 設定状況確認', 'checkApiKeys')
        .addItem('🗑️ APIキー削除', 'clearApiKeys'))
      .addSubMenu(ui.createMenu('👥 ユーザー管理')
        .addItem('🔄 ユーザー切り替え', 'switchUserMode')
        .addItem('📊 権限確認', 'checkUserPermissions'))
      .addSubMenu(ui.createMenu('⚙️ システム管理')
        .addItem('🔄 メニュー更新', 'forceUpdateMenu')
        .addItem('🏥 システム診断', 'performSystemDiagnostics')
        .addItem('� システム情報', 'showSystemInfo'))
      .addToUi();
    
    console.log('✅ MAIN system menu created successfully');
    
    // システム起動通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '営業自動化システム v2.0', 
      '🚀 メニューシステム起動完了', 
      5
    );
    
  } catch (error) {
    console.error('❌ Main menu creation error:', error);
    
    // 最小限のフォールバックメニュー
    try {
      SpreadsheetApp.getUi()
        .createMenu('🆘 営業システム (エラー)')
        .addItem('📋 状態確認', 'checkSystemStatus')
        .addItem('🔄 メニュー再読み込み', 'reloadMenu')
        .addToUi();
    } catch (fallbackError) {
      console.error('❌ Fallback menu failed:', fallbackError);
    }
  }
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
 * 料金プラン確認（仮実装）
 */
function showPricingPlans() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let plans = '💰 料金プラン一覧\n\n';
    plans += '🥉 ベーシック (¥500/月)\n';
    plans += '• 企業検索: 10社/日\n';
    plans += '• AI機能: なし\n\n';
    
    plans += '🥈 スタンダード (¥1,500/月)\n';
    plans += '• 企業検索: 50社/日\n';
    plans += '• AI機能: あり\n\n';
    
    plans += '🥇 プロフェッショナル (¥5,500/月)\n';
    plans += '• 企業検索: 100社/日\n';
    plans += '• AI機能: あり\n';
    plans += '• マルチアカウント: 2アカウント\n\n';
    
    plans += '💎 エンタープライズ (¥17,500/月)\n';
    plans += '• 企業検索: 500社/日\n';
    plans += '• AI機能: あり\n';
    plans += '• マルチアカウント: 5アカウント';
    
    ui.alert('料金プラン', plans, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Pricing plans error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '料金プラン確認でエラーが発生しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 管理者認証（実装版）
 */
function authenticateAdmin() {
  try {
    // license-manager.js の関数を呼び出し
    const result = showAdminAuthenticationDialog();
    
    if (result) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '管理者認証成功', 
        '� 管理者機能が有効になりました', 
        3
      );
      
      // ライセンス管理シートを表示
      createLicenseManagementSheet();
    }
    
  } catch (error) {
    console.error('Admin auth error:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', '管理者認証でエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
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