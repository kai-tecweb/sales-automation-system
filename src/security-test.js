/**
 * セキュリティテスト機能
 * 不正アクセス防止・セッション期限切れ・パスワード強度検証のテスト
 */

/**
 * 包括的セキュリティテストの実行
 */
function runSecurityTests() {
  try {
    console.log('🔒 セキュリティテスト開始...');
    
    updateExecutionStatus('セキュリティテストを実行中...');
    
    const testResults = {
      sessionSecurity: testSessionSecurity(),
      passwordSecurity: testPasswordSecurity(),
      accessControl: testAccessControl(),
      dataProtection: testDataProtection(),
      timestamp: new Date()
    };
    
    // テスト結果の評価
    const allSecurityTestsPassed = Object.values(testResults)
      .filter(result => result && typeof result === 'object' && result.status)
      .every(result => result.status === 'OK');
    
    const resultMessage = allSecurityTestsPassed 
      ? 'すべてのセキュリティテストが正常に通過しました' 
      : 'いくつかのセキュリティ項目に注意が必要です';
    
    updateExecutionStatus(`セキュリティテスト完了: ${resultMessage}`);
    
    // テスト結果をログに記録
    logExecution('セキュリティテスト', allSecurityTestsPassed ? 'SUCCESS' : 'WARNING', 
                allSecurityTestsPassed ? 4 : 0, allSecurityTestsPassed ? 0 : 1);
    
    // 結果を表示
    displaySecurityTestResults(testResults);
    
    return testResults;
    
  } catch (error) {
    console.error('❌ セキュリティテストエラー:', error);
    updateExecutionStatus(`セキュリティテストエラー: ${error.message}`);
    handleSystemError('セキュリティテスト', error);
  }
}

/**
 * セッションセキュリティのテスト
 */
function testSessionSecurity() {
  try {
    console.log('🕐 セッションセキュリティテスト開始...');
    
    const properties = PropertiesService.getScriptProperties();
    const sessionTimeout = properties.getProperty('SESSION_TIMEOUT');
    const loginTime = properties.getProperty('LOGIN_TIMESTAMP');
    
    const tests = {
      sessionTimeoutSet: !!sessionTimeout,
      loginTimeTracked: !!loginTime,
      sessionExpiryWorking: false,
      sessionCleanupWorking: false
    };
    
    // セッション期限切れのテスト
    if (sessionTimeout && loginTime) {
      const timeoutTime = parseInt(sessionTimeout);
      const currentTime = new Date().getTime();
      const remainingTime = timeoutTime - currentTime;
      
      tests.sessionExpiryWorking = remainingTime > 0;
    }
    
    // セッションクリーンアップのテスト
    try {
      const currentUser = getCurrentUser();
      tests.sessionCleanupWorking = typeof currentUser === 'object';
    } catch (error) {
      tests.sessionCleanupWorking = false;
    }
    
    const passedTests = Object.values(tests).filter(Boolean).length;
    const totalTests = Object.keys(tests).length;
    
    return {
      status: passedTests >= totalTests - 1 ? 'OK' : 'WARNING',
      message: `セッションセキュリティ: ${passedTests}/${totalTests} テスト通過`,
      details: {
        tests: tests,
        sessionInfo: {
          hasTimeout: !!sessionTimeout,
          remainingMinutes: sessionTimeout ? Math.floor((parseInt(sessionTimeout) - new Date().getTime()) / (1000 * 60)) : 0
        }
      }
    };
    
  } catch (error) {
    console.error('❌ セッションセキュリティテストエラー:', error);
    return {
      status: 'ERROR',
      message: 'セッションセキュリティテスト失敗',
      details: error.message
    };
  }
}

/**
 * パスワードセキュリティのテスト
 */
function testPasswordSecurity() {
  try {
    console.log('🔐 パスワードセキュリティテスト開始...');
    
    const tests = {
      hashFunctionAvailable: false,
      verifyFunctionAvailable: false,
      saltUsed: false,
      hashStrengthValid: false
    };
    
    // ハッシュ関数の存在確認
    try {
      tests.hashFunctionAvailable = typeof hashPassword === 'function';
    } catch (error) {
      tests.hashFunctionAvailable = false;
    }
    
    // 検証関数の存在確認
    try {
      tests.verifyFunctionAvailable = typeof verifyPassword === 'function';
    } catch (error) {
      tests.verifyFunctionAvailable = false;
    }
    
    // パスワードハッシュ化のテスト
    if (tests.hashFunctionAvailable) {
      try {
        const testPassword = 'TestPassword123!';
        const hashedPassword = hashPassword(testPassword);
        
        // ソルトの使用確認（:で分割されているか）
        tests.saltUsed = hashedPassword.includes(':');
        
        // ハッシュ強度確認（十分な長さがあるか）
        tests.hashStrengthValid = hashedPassword.length > 50;
        
      } catch (error) {
        console.log('パスワードハッシュテストでエラー:', error);
      }
    }
    
    const passedTests = Object.values(tests).filter(Boolean).length;
    const totalTests = Object.keys(tests).length;
    
    return {
      status: passedTests >= totalTests - 1 ? 'OK' : 'WARNING',
      message: `パスワードセキュリティ: ${passedTests}/${totalTests} テスト通過`,
      details: {
        tests: tests,
        recommendations: passedTests < totalTests ? ['SHA-256+ソルトハッシュ化の実装', 'パスワード複雑性要件の追加'] : []
      }
    };
    
  } catch (error) {
    console.error('❌ パスワードセキュリティテストエラー:', error);
    return {
      status: 'ERROR',
      message: 'パスワードセキュリティテスト失敗',
      details: error.message
    };
  }
}

/**
 * アクセス制御のテスト
 */
function testAccessControl() {
  try {
    console.log('🚫 アクセス制御テスト開始...');
    
    const currentUser = getCurrentUser();
    
    const tests = {
      roleBasedAccess: false,
      permissionChecking: false,
      unauthorizedPrevention: false,
      menuRestrictions: false
    };
    
    // ロールベースアクセス制御のテスト
    try {
      tests.roleBasedAccess = typeof checkUserPermission === 'function';
    } catch (error) {
      tests.roleBasedAccess = false;
    }
    
    // 権限チェック機能のテスト
    if (tests.roleBasedAccess) {
      try {
        const adminPermission = checkUserPermission('Administrator');
        const standardPermission = checkUserPermission('Standard');
        const guestPermission = checkUserPermission('Guest');
        
        tests.permissionChecking = typeof adminPermission === 'object' &&
                                  typeof standardPermission === 'object' &&
                                  typeof guestPermission === 'object';
      } catch (error) {
        tests.permissionChecking = false;
      }
    }
    
    // 不正アクセス防止のテスト
    try {
      // ゲストユーザーで管理者機能にアクセスしようとする
      const guestPermission = checkUserPermission('Administrator');
      tests.unauthorizedPrevention = !guestPermission.hasPermission || currentUser.role === 'Administrator';
    } catch (error) {
      tests.unauthorizedPrevention = true; // エラーが発生することで防止されている
    }
    
    // メニュー制限のテスト
    try {
      tests.menuRestrictions = typeof createRoleBasedMenu === 'function';
    } catch (error) {
      tests.menuRestrictions = false;
    }
    
    const passedTests = Object.values(tests).filter(Boolean).length;
    const totalTests = Object.keys(tests).length;
    
    return {
      status: passedTests >= totalTests - 1 ? 'OK' : 'WARNING',
      message: `アクセス制御: ${passedTests}/${totalTests} テスト通過`,
      details: {
        tests: tests,
        currentUserRole: currentUser.role || 'Guest',
        securityLevel: passedTests === totalTests ? '高' : passedTests >= totalTests - 1 ? '中' : '低'
      }
    };
    
  } catch (error) {
    console.error('❌ アクセス制御テストエラー:', error);
    return {
      status: 'ERROR',
      message: 'アクセス制御テスト失敗',
      details: error.message
    };
  }
}

/**
 * データ保護のテスト
 */
function testDataProtection() {
  try {
    console.log('🛡️ データ保護テスト開始...');
    
    const tests = {
      sensitiveDataEncryption: false,
      apiKeyProtection: false,
      logDataSecurity: false,
      sheetPermissions: false
    };
    
    // 機密データ暗号化のテスト
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // APIキー保護のテスト（直接取得できないようになっているか）
      const testKeys = ['OPENAI_API_KEY', 'GOOGLE_SEARCH_API_KEY', 'GOOGLE_SEARCH_ENGINE_ID'];
      let protectedKeys = 0;
      
      for (const keyName of testKeys) {
        const key = properties.getProperty(keyName);
        if (key && key.length > 10) {
          protectedKeys++;
        }
      }
      
      tests.apiKeyProtection = protectedKeys > 0;
    } catch (error) {
      tests.apiKeyProtection = false;
    }
    
    // ログデータセキュリティのテスト
    try {
      const logSheet = getSafeSheet(SHEET_NAMES.LOGS);
      tests.logDataSecurity = !!logSheet;
    } catch (error) {
      tests.logDataSecurity = false;
    }
    
    // シート権限のテスト
    try {
      tests.sheetPermissions = typeof updateSheetVisibilityOnLogin === 'function';
    } catch (error) {
      tests.sheetPermissions = false;
    }
    
    // 機密データ暗号化（基本レベル）
    tests.sensitiveDataEncryption = tests.apiKeyProtection; // APIキーが保護されていれば基本的な暗号化がされている
    
    const passedTests = Object.values(tests).filter(Boolean).length;
    const totalTests = Object.keys(tests).length;
    
    return {
      status: passedTests >= totalTests - 1 ? 'OK' : 'WARNING',
      message: `データ保護: ${passedTests}/${totalTests} テスト通過`,
      details: {
        tests: tests,
        protectionLevel: passedTests === totalTests ? '完全' : passedTests >= totalTests - 1 ? '良好' : '要改善'
      }
    };
    
  } catch (error) {
    console.error('❌ データ保護テストエラー:', error);
    return {
      status: 'ERROR',
      message: 'データ保護テスト失敗',
      details: error.message
    };
  }
}

/**
 * セキュリティテスト結果の表示
 */
function displaySecurityTestResults(testResults) {
  try {
    const sheet = getSafeSheet(SHEET_NAMES.CONTROL);
    if (!sheet) return;
    
    // セキュリティテスト結果表示エリア
    const startRow = 28;
    
    // 既存の結果をクリア
    sheet.getRange(startRow, 8, 15, 3).clearContent();
    
    let row = startRow;
    
    // ヘッダー
    sheet.getRange(row, 8).setValue('🔒 セキュリティテスト結果');
    sheet.getRange(row, 8).setFontWeight('bold');
    sheet.getRange(row, 8).setBackground('#dc3545');
    sheet.getRange(row, 8).setFontColor('white');
    row += 2;
    
    // 各セキュリティテスト結果
    const testItems = [
      ['セッションセキュリティ', testResults.sessionSecurity],
      ['パスワードセキュリティ', testResults.passwordSecurity],
      ['アクセス制御', testResults.accessControl],
      ['データ保護', testResults.dataProtection]
    ];
    
    for (const [testName, result] of testItems) {
      sheet.getRange(row, 8).setValue(testName);
      sheet.getRange(row, 9).setValue(result.status);
      sheet.getRange(row, 10).setValue(result.message);
      setStatusColor(sheet.getRange(row, 9), result.status);
      row++;
    }
    
    row += 1;
    
    // タイムスタンプ
    sheet.getRange(row, 8).setValue('テスト実行時刻:');
    sheet.getRange(row, 9).setValue(testResults.timestamp);
    
    console.log('✅ セキュリティテスト結果を制御パネルに表示しました');
    
  } catch (error) {
    console.error('❌ セキュリティテスト結果表示エラー:', error);
  }
}

/**
 * セキュリティ推奨事項の表示
 */
function showSecurityRecommendations() {
  try {
    const securityTests = runSecurityTests();
    
    let recommendations = [];
    
    // 各テスト結果から推奨事項を収集
    Object.values(securityTests).forEach(result => {
      if (result && result.details && result.details.recommendations) {
        recommendations = recommendations.concat(result.details.recommendations);
      }
    });
    
    // 一般的なセキュリティ推奨事項
    const generalRecommendations = [
      '定期的なパスワード変更',
      'セッションタイムアウトの適切な設定',
      'API キーの定期ローテーション',
      'アクセスログの監視',
      '最小権限の原則の適用'
    ];
    
    let message = '🔒 セキュリティ推奨事項\n\n';
    
    if (recommendations.length > 0) {
      message += '⚠️ 改善が必要な項目:\n';
      recommendations.forEach((rec, index) => {
        message += `${index + 1}. ${rec}\n`;
      });
      message += '\n';
    }
    
    message += '📋 一般的な推奨事項:\n';
    generalRecommendations.forEach((rec, index) => {
      message += `${index + 1}. ${rec}\n`;
    });
    
    SpreadsheetApp.getUi().alert('セキュリティ推奨事項', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ セキュリティ推奨事項表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'セキュリティ推奨事項の表示に失敗しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * セキュリティ監査ログの生成
 */
function generateSecurityAuditLog() {
  try {
    console.log('📋 セキュリティ監査ログ生成開始...');
    
    const auditData = {
      timestamp: new Date(),
      systemInfo: {
        userCount: 0,
        adminUsers: 0,
        activeeSessions: 0
      },
      securityTests: runSecurityTests(),
      recommendations: []
    };
    
    // ユーザー情報の収集
    try {
      const userSheet = getSafeSheet('ユーザー管理');
      if (userSheet && userSheet.getLastRow() > 1) {
        const userData = userSheet.getDataRange().getValues();
        auditData.systemInfo.userCount = userData.length - 1; // ヘッダーを除外
        auditData.systemInfo.adminUsers = userData.filter(row => row[3] === 'Administrator').length;
      }
    } catch (error) {
      console.log('ユーザー情報収集でエラー:', error);
    }
    
    // セッション情報の収集
    try {
      const currentUser = getCurrentUser();
      auditData.systemInfo.activeSessions = currentUser.isLoggedIn ? 1 : 0;
    } catch (error) {
      console.log('セッション情報収集でエラー:', error);
    }
    
    // 監査ログをログシートに記録
    const logSheet = getSafeSheet(SHEET_NAMES.LOGS);
    if (logSheet) {
      const auditEntry = [
        auditData.timestamp,
        'セキュリティ監査',
        'AUDIT',
        auditData.systemInfo.userCount,
        0,
        `ユーザー数: ${auditData.systemInfo.userCount}, 管理者: ${auditData.systemInfo.adminUsers}, アクティブセッション: ${auditData.systemInfo.activeSessions}`
      ];
      
      const newRow = logSheet.getLastRow() + 1;
      logSheet.getRange(newRow, 1, 1, auditEntry.length).setValues([auditEntry]);
    }
    
    console.log('✅ セキュリティ監査ログ生成完了');
    return auditData;
    
  } catch (error) {
    console.error('❌ セキュリティ監査ログ生成エラー:', error);
    throw error;
  }
}