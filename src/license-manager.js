/**
 * ライセンス管理・管理者認証システム
 */

// 管理者パスワード（実際の運用では暗号化推奨）
const ADMIN_PASSWORD = "SalesAuto2024!";
const LICENSE_DURATION_DAYS = 10; // 営業日ベース

/**
 * ライセンス管理シートの作成
 */
function createLicenseManagementSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のライセンス管理シートを削除（再作成のため）
  const existingSheet = spreadsheet.getSheetByName('📋 ライセンス管理');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを最前面に作成
  const sheet = spreadsheet.insertSheet('📋 ライセンス管理', 0);
  
  // ヘッダー設定
  sheet.getRange(1, 1, 1, 6).setValues([['🚀 営業自動化システム - ライセンス管理', '', '', '', '', '']]);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(1, 1, 1, 6).merge();
  
  // 現在のライセンス状況
  const currentDate = new Date();
  const licenseInfo = getLicenseInfo();
  
  sheet.getRange(3, 1).setValue('📊 現在のライセンス状況');
  sheet.getRange(3, 1).setFontSize(14).setFontWeight('bold');
  
  // ライセンス情報テーブル
  const headers = ['項目', '状態', '詳細', '操作'];
  sheet.getRange(5, 1, 1, 4).setValues([headers]);
  sheet.getRange(5, 1, 1, 4).setFontWeight('bold').setBackground('#f1f3f4');
  
  const statusData = [
    ['管理者モード', licenseInfo.adminMode ? '🟢 有効' : '🔴 無効', licenseInfo.adminMode ? '管理者機能が利用可能' : '一般ユーザーモード', '👤 管理者認証'],
    ['ライセンス期限', licenseInfo.isExpired ? '🔴 期限切れ' : '🟢 有効', 
     licenseInfo.startDate ? `開始: ${formatDate(licenseInfo.startDate)}` : '未設定',
     '📅 使用開始日設定'],
    ['残り日数', licenseInfo.remainingDays !== null ? `${licenseInfo.remainingDays}営業日` : '未計算',
     licenseInfo.expiryDate ? `期限: ${formatDate(licenseInfo.expiryDate)}` : '未設定',
     '🔄 期限延長'],
    ['システム状態', licenseInfo.systemLocked ? '🔒 ロック中' : '✅ 利用可能',
     licenseInfo.systemLocked ? 'ライセンス期限により機能制限中' : '全機能が利用可能',
     '🔓 ロック解除']
  ];
  
  sheet.getRange(6, 1, statusData.length, 4).setValues(statusData);
  
  // 操作ボタン（実際にはメニューから実行）
  sheet.getRange(12, 1).setValue('🔧 管理操作');
  sheet.getRange(12, 1).setFontSize(14).setFontWeight('bold');
  
  const operations = [
    ['👤 管理者認証', '「メニュー → 📋 ライセンス管理 → 👤 管理者認証」から実行'],
    ['📅 使用開始日設定', '「メニュー → 📋 ライセンス管理 → 📅 使用開始日設定」から実行'],
    ['🔄 ライセンス更新', '「メニュー → 📋 ライセンス管理 → 🔄 ライセンス更新」から実行'],
    ['🔓 緊急解除', '「メニュー → 📋 ライセンス管理 → 🔓 緊急解除」から実行']
  ];
  
  sheet.getRange(14, 1, operations.length, 2).setValues(operations);
  sheet.getRange(14, 1, operations.length, 1).setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 150);
  
  // ライセンス履歴
  sheet.getRange(20, 1).setValue('📜 ライセンス履歴');
  sheet.getRange(20, 1).setFontSize(14).setFontWeight('bold');
  
  const historyHeaders = ['日時', 'イベント', '実行者', '詳細'];
  sheet.getRange(22, 1, 1, 4).setValues([historyHeaders]);
  sheet.getRange(22, 1, 1, 4).setFontWeight('bold').setBackground('#f1f3f4');
  
  // 最新の履歴を表示
  displayLicenseHistory(sheet);
  
  SpreadsheetApp.getUi().alert(
    'ライセンス管理シート作成完了',
    'ライセンス管理シートが最前面に作成されました。\n管理者認証を行って管理者機能を有効にしてください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * ライセンス情報の取得
 */
function getLicenseInfo() {
  const properties = PropertiesService.getScriptProperties();
  const adminMode = properties.getProperty('ADMIN_MODE') === 'true';
  const startDateStr = properties.getProperty('LICENSE_START_DATE');
  const startDate = startDateStr ? new Date(startDateStr) : null;
  
  let isExpired = false;
  let remainingDays = null;
  let expiryDate = null;
  let systemLocked = false;
  
  if (startDate) {
    expiryDate = calculateBusinessDays(startDate, LICENSE_DURATION_DAYS);
    const currentDate = new Date();
    remainingDays = calculateRemainingBusinessDays(currentDate, expiryDate);
    isExpired = remainingDays <= 0;
    systemLocked = isExpired && !adminMode;
  }
  
  // プラン管理統合ポイント - Phase 0
  let planInfo = null;
  try {
    // plan-manager.jsが利用可能な場合はプラン情報も含める
    if (typeof getUserPlan === 'function') {
      planInfo = {
        currentPlan: getUserPlan(),
        planDisplayName: getPlanDisplayName(),
        planLimits: getPlanLimits(),
        isInSwitchMode: isInSwitchMode()
      };
    }
  } catch (error) {
    console.log('プラン管理システム未初期化 - ライセンス情報のみ提供');
  }

  return {
    adminMode: adminMode,
    startDate: startDate,
    expiryDate: expiryDate,
    isExpired: isExpired,
    remainingDays: remainingDays,
    systemLocked: systemLocked,
    // Phase 0: プラン管理統合
    planInfo: planInfo
  };
}

/**
 * 営業日計算（土日祝日を除く）
 */
function calculateBusinessDays(startDate, days) {
  const result = new Date(startDate);
  let addedDays = 0;
  
  while (addedDays < days) {
    result.setDate(result.getDate() + 1);
    
    // 土日をスキップ（0=日曜, 6=土曜）
    if (result.getDay() !== 0 && result.getDay() !== 6) {
      // 日本の祝日チェック（簡易版）
      if (!isJapaneseHoliday(result)) {
        addedDays++;
      }
    }
  }
  
  return result;
}

/**
 * 残り営業日数の計算
 */
function calculateRemainingBusinessDays(currentDate, expiryDate) {
  if (!expiryDate || currentDate >= expiryDate) {
    return 0;
  }
  
  let businessDays = 0;
  const tempDate = new Date(currentDate);
  
  while (tempDate < expiryDate) {
    tempDate.setDate(tempDate.getDate() + 1);
    
    if (tempDate.getDay() !== 0 && tempDate.getDay() !== 6) {
      if (!isJapaneseHoliday(tempDate)) {
        businessDays++;
      }
    }
  }
  
  return businessDays;
}

/**
 * 日本の祝日チェック（簡易版）
 */
function isJapaneseHoliday(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  
  // 主要な固定祝日のみ（実際の運用では祝日APIの利用推奨）
  const holidays = [
    `${year}-01-01`, // 元旦
    `${year}-02-11`, // 建国記念の日
    `${year}-04-29`, // 昭和の日
    `${year}-05-03`, // 憲法記念日
    `${year}-05-04`, // みどりの日
    `${year}-05-05`, // こどもの日
    `${year}-08-11`, // 山の日
    `${year}-11-03`, // 文化の日
    `${year}-11-23`, // 勤労感謝の日
    `${year}-12-23`  // 天皇誕生日
  ];
  
  const dateStr = `${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
  return holidays.includes(dateStr);
}

/**
 * 管理者認証
 */
function authenticateAdmin() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔐 管理者認証',
    '管理者パスワードを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const password = response.getResponseText();
    
    if (password === ADMIN_PASSWORD) {
      // 管理者モードを有効化
      PropertiesService.getScriptProperties().setProperty('ADMIN_MODE', 'true');
      
      // ライセンス履歴に記録
      logLicenseEvent('管理者認証成功', Session.getActiveUser().getEmail(), '管理者モードが有効化されました');
      
      // メニューを更新
      onOpen();
      
      // ライセンス管理シートを更新
      updateLicenseManagementSheet();
      
      ui.alert(
        '✅ 認証成功',
        '管理者モードが有効になりました。\n管理者専用機能が利用可能です。',
        ui.ButtonSet.OK
      );
      
    } else {
      logLicenseEvent('管理者認証失敗', Session.getActiveUser().getEmail(), '不正なパスワードが入力されました');
      
      ui.alert(
        '❌ 認証失敗',
        'パスワードが正しくありません。',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * 使用開始日の設定
 */
function setLicenseStartDate() {
  const ui = SpreadsheetApp.getUi();
  
  // 管理者チェック
  if (!isAdminMode()) {
    ui.alert('アクセス拒否', '管理者認証が必要です。', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt(
    '📅 使用開始日設定',
    '使用開始日を入力してください (YYYY-MM-DD形式):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const dateStr = response.getResponseText();
    
    try {
      const startDate = new Date(dateStr);
      if (isNaN(startDate.getTime())) {
        throw new Error('無効な日付形式');
      }
      
      // 使用開始日を保存
      PropertiesService.getScriptProperties().setProperty('LICENSE_START_DATE', startDate.toISOString());
      
      // 期限日を計算
      const expiryDate = calculateBusinessDays(startDate, LICENSE_DURATION_DAYS);
      
      // ライセンス履歴に記録
      logLicenseEvent('使用開始日設定', Session.getActiveUser().getEmail(), 
        `開始日: ${formatDate(startDate)}, 期限: ${formatDate(expiryDate)}`);
      
      // ライセンス管理シートを更新
      updateLicenseManagementSheet();
      
      ui.alert(
        '✅ 設定完了',
        `使用開始日: ${formatDate(startDate)}\n期限日: ${formatDate(expiryDate)}\n(${LICENSE_DURATION_DAYS}営業日後)`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ui.alert(
        '❌ 設定エラー',
        `日付の形式が正しくありません。\nYYYY-MM-DD形式で入力してください。\n例: 2024-10-17`,
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * ライセンス期限延長
 */
function extendLicense() {
  const ui = SpreadsheetApp.getUi();
  
  // 管理者チェック
  if (!isAdminMode()) {
    ui.alert('アクセス拒否', '管理者認証が必要です。', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt(
    '🔄 ライセンス延長',
    '延長する営業日数を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const extensionDays = parseInt(response.getResponseText());
    
    if (isNaN(extensionDays) || extensionDays <= 0) {
      ui.alert('❌ 入力エラー', '正の整数を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 現在の開始日を取得
    const startDateStr = PropertiesService.getScriptProperties().getProperty('LICENSE_START_DATE');
    if (!startDateStr) {
      ui.alert('❌ エラー', '使用開始日が設定されていません。', ui.ButtonSet.OK);
      return;
    }
    
    const startDate = new Date(startDateStr);
    const newExpiryDate = calculateBusinessDays(startDate, LICENSE_DURATION_DAYS + extensionDays);
    
    // ライセンス履歴に記録
    logLicenseEvent('ライセンス延長', Session.getActiveUser().getEmail(), 
      `${extensionDays}営業日延長, 新期限: ${formatDate(newExpiryDate)}`);
    
    // 新しい期間を保存（簡易的に開始日を調整）
    const adjustedStartDate = calculateBusinessDays(newExpiryDate, -LICENSE_DURATION_DAYS);
    PropertiesService.getScriptProperties().setProperty('LICENSE_START_DATE', adjustedStartDate.toISOString());
    
    // ライセンス管理シートを更新
    updateLicenseManagementSheet();
    
    ui.alert(
      '✅ 延長完了',
      `ライセンスを${extensionDays}営業日延長しました。\n新しい期限: ${formatDate(newExpiryDate)}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 緊急ライセンス解除
 */
function emergencyUnlock() {
  const ui = SpreadsheetApp.getUi();
  
  const passwordResponse = ui.prompt(
    '🔓 緊急解除',
    '管理者パスワードを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (passwordResponse.getSelectedButton() === ui.Button.OK) {
    const password = passwordResponse.getResponseText();
    
    if (password === ADMIN_PASSWORD) {
      // 管理者モードを有効化
      PropertiesService.getScriptProperties().setProperty('ADMIN_MODE', 'true');
      
      // ライセンス期限をリセット（30日延長）
      const newStartDate = new Date();
      PropertiesService.getScriptProperties().setProperty('LICENSE_START_DATE', newStartDate.toISOString());
      
      // ライセンス履歴に記録
      logLicenseEvent('緊急解除', Session.getActiveUser().getEmail(), '緊急解除により30日のライセンスが付与されました');
      
      // メニューを更新
      onOpen();
      
      // ライセンス管理シートを更新
      updateLicenseManagementSheet();
      
      ui.alert(
        '✅ 緊急解除完了',
        'システムロックが解除されました。\n新しいライセンス期間: 30営業日',
        ui.ButtonSet.OK
      );
      
    } else {
      logLicenseEvent('緊急解除失敗', Session.getActiveUser().getEmail(), '不正なパスワードが入力されました');
      
      ui.alert(
        '❌ 認証失敗',
        'パスワードが正しくありません。',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * 管理者モードチェック
 */
function isAdminMode() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_MODE') === 'true';
}

/**
 * システムロックチェック
 */
function isSystemLocked() {
  const licenseInfo = getLicenseInfo();
  return licenseInfo.systemLocked;
}

/**
 * ライセンス履歴の記録
 */
function logLicenseEvent(event, user, details) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const history = JSON.parse(properties.getProperty('LICENSE_HISTORY') || '[]');
    
    history.push({
      timestamp: new Date().toISOString(),
      event: event,
      user: user,
      details: details
    });
    
    // 最新50件のみ保持
    if (history.length > 50) {
      history.splice(0, history.length - 50);
    }
    
    properties.setProperty('LICENSE_HISTORY', JSON.stringify(history));
  } catch (error) {
    Logger.log(`ライセンス履歴記録エラー: ${error.toString()}`);
  }
}

/**
 * ライセンス履歴の表示
 */
function displayLicenseHistory(sheet) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const history = JSON.parse(properties.getProperty('LICENSE_HISTORY') || '[]');
    
    if (history.length === 0) {
      sheet.getRange(23, 1).setValue('履歴なし');
      return;
    }
    
    // 最新10件を表示
    const recentHistory = history.slice(-10).reverse();
    const historyData = recentHistory.map(entry => [
      new Date(entry.timestamp).toLocaleString('ja-JP'),
      entry.event,
      entry.user,
      entry.details
    ]);
    
    sheet.getRange(23, 1, historyData.length, 4).setValues(historyData);
    
  } catch (error) {
    sheet.getRange(23, 1).setValue(`履歴表示エラー: ${error.toString()}`);
  }
}

/**
 * ライセンス管理シートの更新
 */
function updateLicenseManagementSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📋 ライセンス管理');
    if (sheet) {
      // シートを再作成（簡易的な更新）
      createLicenseManagementSheet();
    }
  } catch (error) {
    Logger.log(`ライセンス管理シート更新エラー: ${error.toString()}`);
  }
}

/**
 * 日付フォーマット
 */
function formatDate(date) {
  return date.toLocaleDateString('ja-JP', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    weekday: 'short'
  });
}
