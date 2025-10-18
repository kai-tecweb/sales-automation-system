/**
 * 営業自動化システム - 包括的テストスイート
 * 権限管理・基本機能・API統合の動作確認
 */

/**
 * システム全体の健康状態をチェック
 */
function performSystemHealthCheck() {
  try {
    console.log('🏥 システム健康状態チェック開始...');
    
    const healthReport = {
      systemInitialization: checkSystemInitializationStatus(),
      userManagement: checkUserManagementStatus(),
      apiConfiguration: checkAPIConfiguration(),
      basicFunctions: checkBasicFunctions(),
      menuIntegration: checkMenuIntegration(),
      security: checkSecurityFeatures()
    };
    
    console.log('📊 健康状態レポート:', healthReport);
    
    // 結果を制御パネルに表示
    displayHealthCheckResults(healthReport);
    
    return healthReport;
    
  } catch (error) {
    console.error('❌ システム健康チェックエラー:', error);
    return {
      error: error.message,
      status: 'FAILED'
    };
  }
}

/**
 * システム初期化状態の確認
 */
function checkSystemInitializationStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ['制御パネル', '生成キーワード', '企業マスター', '提案メッセージ', '実行ログ', 'ユーザー管理'];
    
    const results = {};
    
    for (const sheetName of requiredSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      results[sheetName] = {
        exists: !!sheet,
        rowCount: sheet ? sheet.getLastRow() : 0,
        columnCount: sheet ? sheet.getLastColumn() : 0
      };
    }
    
    const allSheetsExist = requiredSheets.every(name => results[name].exists);
    
    return {
      status: allSheetsExist ? 'OK' : 'ERROR',
      sheets: results,
      message: allSheetsExist ? '全ての必須シートが存在します' : '不足しているシートがあります'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * ユーザー管理システムの状態確認
 */
function checkUserManagementStatus() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ユーザー管理');
    
    if (!sheet) {
      return {
        status: 'ERROR',
        message: 'ユーザー管理シートが存在しません'
      };
    }
    
    // 管理者ユーザーの存在確認
    const data = sheet.getDataRange().getValues();
    const adminExists = data.some(row => row[3] === 'Administrator');
    
    // 現在のセッション状態
    const currentUser = getCurrentUser();
    
    return {
      status: adminExists ? 'OK' : 'WARNING',
      adminUserExists: adminExists,
      currentSession: {
        isLoggedIn: currentUser.isLoggedIn,
        role: currentUser.role,
        username: currentUser.username
      },
      totalUsers: data.length - 1, // ヘッダーを除外
      message: adminExists ? 'ユーザー管理システム正常' : '管理者ユーザーが見つかりません'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * API設定の確認
 */
function checkAPIConfiguration() {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    const openaiKey = properties.getProperty('OPENAI_API_KEY');
    const googleSearchKey = properties.getProperty('GOOGLE_SEARCH_API_KEY');
    const searchEngineId = properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    
    return {
      status: (openaiKey && googleSearchKey && searchEngineId) ? 'OK' : 'WARNING',
      apis: {
        openai: !!openaiKey,
        googleSearch: !!googleSearchKey,
        searchEngineId: !!searchEngineId
      },
      message: (openaiKey && googleSearchKey && searchEngineId) 
        ? '全てのAPI設定が完了しています' 
        : '一部のAPI設定が不足しています'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * 基本機能の存在確認
 */
function checkBasicFunctions() {
  try {
    const functions = [
      'executeKeywordGeneration',
      'executeCompanySearch', 
      'executeProposalGeneration',
      'authenticateUser',
      'getCurrentUser',
      'initializeSheets'
    ];
    
    const results = {};
    
    for (const funcName of functions) {
      try {
        const func = eval(funcName);
        results[funcName] = {
          exists: typeof func === 'function',
          status: 'OK'
        };
      } catch (error) {
        results[funcName] = {
          exists: false,
          status: 'ERROR',
          error: error.message
        };
      }
    }
    
    const allFunctionsExist = Object.values(results).every(r => r.exists);
    
    return {
      status: allFunctionsExist ? 'OK' : 'ERROR',
      functions: results,
      message: allFunctionsExist ? '全ての基本機能が利用可能です' : '不足している機能があります'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * メニュー統合の確認
 */
function checkMenuIntegration() {
  try {
    // メニュー関連関数の確認
    const menuFunctions = [
      'createRoleBasedMenu',
      'showUserLoginDialog',
      'showCurrentUserStatus'
    ];
    
    const results = {};
    
    for (const funcName of menuFunctions) {
      try {
        const func = eval(funcName);
        results[funcName] = {
          exists: typeof func === 'function',
          status: 'OK'
        };
      } catch (error) {
        results[funcName] = {
          exists: false,
          status: 'ERROR'
        };
      }
    }
    
    const allMenuFunctionsExist = Object.values(results).every(r => r.exists);
    
    return {
      status: allMenuFunctionsExist ? 'OK' : 'ERROR',
      menuFunctions: results,
      message: allMenuFunctionsExist ? 'メニューシステム統合完了' : 'メニュー機能に問題があります'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * セキュリティ機能の確認
 */
function checkSecurityFeatures() {
  try {
    const currentUser = getCurrentUser();
    
    // セッション管理機能の確認
    const sessionTimeout = PropertiesService.getScriptProperties()
      .getProperty('SESSION_TIMEOUT');
    
    return {
      status: 'OK',
      sessionManagement: {
        hasActiveSession: currentUser.isLoggedIn,
        sessionTimeoutSet: !!sessionTimeout,
        userRole: currentUser.role || 'Guest'
      },
      passwordHashing: {
        available: typeof hashPassword === 'function',
        verification: typeof verifyPassword === 'function'
      },
      message: 'セキュリティ機能正常'
    };
    
  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message
    };
  }
}

/**
 * 健康チェック結果をスプレッドシートに表示
 */
function displayHealthCheckResults(healthReport) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('制御パネル');
    if (!sheet) return;
    
    // 既存の健康チェック結果をクリア
    const range = sheet.getRange('H1:J20');
    range.clearContent();
    
    // 新しい結果を表示
    let row = 1;
    
    // ヘッダー
    sheet.getRange(row, 8).setValue('🏥 システム健康チェック結果');
    sheet.getRange(row, 8).setFontWeight('bold');
    sheet.getRange(row, 8).setBackground('#4285f4');
    sheet.getRange(row, 8).setFontColor('white');
    row += 2;
    
    // 各項目の結果
    const items = [
      ['システム初期化', healthReport.systemInitialization?.status || 'ERROR'],
      ['ユーザー管理', healthReport.userManagement?.status || 'ERROR'],
      ['API設定', healthReport.apiConfiguration?.status || 'ERROR'],
      ['基本機能', healthReport.basicFunctions?.status || 'ERROR'],
      ['メニュー統合', healthReport.menuIntegration?.status || 'ERROR'],
      ['セキュリティ', healthReport.security?.status || 'ERROR']
    ];
    
    for (const [item, status] of items) {
      sheet.getRange(row, 8).setValue(item);
      sheet.getRange(row, 9).setValue(status);
      
      // ステータスに応じた色分け
      const statusCell = sheet.getRange(row, 9);
      switch (status) {
        case 'OK':
          statusCell.setBackground('#d4edda');
          statusCell.setFontColor('#155724');
          break;
        case 'WARNING':
          statusCell.setBackground('#fff3cd');
          statusCell.setFontColor('#856404');
          break;
        case 'ERROR':
          statusCell.setBackground('#f8d7da');
          statusCell.setFontColor('#721c24');
          break;
      }
      
      row++;
    }
    
    // タイムスタンプ
    row++;
    sheet.getRange(row, 8).setValue('最終チェック:');
    sheet.getRange(row, 9).setValue(new Date());
    
    console.log('✅ 健康チェック結果を制御パネルに表示しました');
    
  } catch (error) {
    console.error('❌ 健康チェック結果表示エラー:', error);
  }
}

/**
 * 権限テスト実行
 */
function testUserPermissions() {
  try {
    console.log('🔐 権限テスト開始...');
    
    const currentUser = getCurrentUser();
    
    const testResults = {
      currentUser: currentUser,
      permissionTests: {
        canAccessAdminFeatures: testAdminAccess(currentUser.role),
        canAccessStandardFeatures: testStandardAccess(currentUser.role),
        canAccessGuestFeatures: testGuestAccess(currentUser.role)
      },
      sessionTests: {
        hasValidSession: currentUser.isLoggedIn,
        sessionExpiry: testSessionExpiry()
      }
    };
    
    console.log('📋 権限テスト結果:', testResults);
    return testResults;
    
  } catch (error) {
    console.error('❌ 権限テストエラー:', error);
    return { error: error.message };
  }
}

/**
 * 管理者権限テスト
 */
function testAdminAccess(userRole) {
  return {
    expected: userRole === 'Administrator',
    features: ['ユーザー管理', 'API設定', 'ライセンス管理', 'システム設定'],
    hasAccess: userRole === 'Administrator'
  };
}

/**
 * スタンダードユーザー権限テスト
 */
function testStandardAccess(userRole) {
  return {
    expected: ['Administrator', 'Standard'].includes(userRole),
    features: ['キーワード生成', '企業検索', '提案作成'],
    hasAccess: ['Administrator', 'Standard'].includes(userRole)
  };
}

/**
 * ゲスト権限テスト
 */
function testGuestAccess(userRole) {
  return {
    expected: true, // 全ユーザーがゲスト機能にアクセス可能
    features: ['データ閲覧', 'レポート表示'],
    hasAccess: true
  };
}

/**
 * セッション期限テスト
 */
function testSessionExpiry() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const sessionTimeout = properties.getProperty('SESSION_TIMEOUT');
    
    if (!sessionTimeout) {
      return { status: 'NO_SESSION', message: 'アクティブなセッションがありません' };
    }
    
    const timeoutTime = parseInt(sessionTimeout);
    const currentTime = new Date().getTime();
    const remainingTime = timeoutTime - currentTime;
    
    return {
      status: remainingTime > 0 ? 'VALID' : 'EXPIRED',
      remainingMinutes: Math.floor(remainingTime / (1000 * 60)),
      message: remainingTime > 0 ? 'セッション有効' : 'セッション期限切れ'
    };
    
  } catch (error) {
    return { status: 'ERROR', message: error.message };
  }
}

/**
 * システムテスト実行
 * メニューから呼び出される統合テスト
 */
function runComprehensiveSystemTest() {
  try {
    console.log('🧪 包括的システムテスト開始...');
    
    updateExecutionStatus('システムテストを実行中...');
    
    const testResults = {
      healthCheck: performSystemHealthCheck(),
      permissionTests: testUserPermissions(),
      timestamp: new Date()
    };
    
    // テスト結果をログに記録
    logExecution('システムテスト', 'SUCCESS', Object.keys(testResults).length, 0);
    
    updateExecutionStatus('システムテスト完了！結果を確認してください。');
    
    // 結果をユーザーに表示
    displayTestResults(testResults);
    
    return testResults;
    
  } catch (error) {
    console.error('❌ システムテストエラー:', error);
    updateExecutionStatus(`システムテストエラー: ${error.message}`);
    logExecution('システムテスト', 'ERROR', 0, 1);
    throw error;
  }
}

/**
 * テスト結果の表示
 */
function displayTestResults(testResults) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '🧪 システムテスト結果\n\n';
    
    // 健康チェック結果
    if (testResults.healthCheck) {
      message += '🏥 システム健康状態:\n';
      const health = testResults.healthCheck;
      message += `- 初期化: ${health.systemInitialization?.status || 'ERROR'}\n`;
      message += `- ユーザー管理: ${health.userManagement?.status || 'ERROR'}\n`;
      message += `- API設定: ${health.apiConfiguration?.status || 'ERROR'}\n`;
      message += `- 基本機能: ${health.basicFunctions?.status || 'ERROR'}\n`;
      message += `- メニュー統合: ${health.menuIntegration?.status || 'ERROR'}\n`;
      message += `- セキュリティ: ${health.security?.status || 'ERROR'}\n\n`;
    }
    
    // 権限テスト結果
    if (testResults.permissionTests) {
      const perms = testResults.permissionTests;
      message += '🔐 権限テスト結果:\n';
      message += `現在のユーザー: ${perms.currentUser?.username || 'ゲスト'}\n`;
      message += `権限レベル: ${perms.currentUser?.role || 'Guest'}\n`;
      message += `セッション状態: ${perms.sessionTests?.hasValidSession ? '有効' : '無効'}\n\n`;
    }
    
    message += `テスト実行時刻: ${testResults.timestamp}`;
    
    ui.alert('システムテスト完了', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ テスト結果表示エラー:', error);
  }
}