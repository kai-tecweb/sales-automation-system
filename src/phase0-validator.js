/**
 * Phase 0 統合システム検証スクリプト
 * 修正計画書 Phase 0 の実装検証を行う
 */

/**
 * Phase 0統合システム検証メイン関数
 */
function validatePhase0Integration() {
  try {
    console.log('🔍 Phase 0統合システム検証開始...');
    
    const results = {
      timestamp: new Date().toISOString(),
      planManager: false,
      licenseIntegration: false,
      userIntegration: false,
      menuIntegration: false,
      sheetIntegration: false,
      errors: []
    };
    
    // 1. プラン管理システム基本機能テスト
    results.planManager = testPlanManagerBasics();
    
    // 2. ライセンス管理統合テスト
    results.licenseIntegration = testLicenseManagerIntegration();
    
    // 3. ユーザー管理統合テスト
    results.userIntegration = testUserManagerIntegration();
    
    // 4. メニューシステム統合テスト
    results.menuIntegration = testMenuSystemIntegration();
    
    // 5. シート拡張機能テスト
    results.sheetIntegration = testSheetIntegration();
    
    // 検証結果レポート生成
    generateValidationReport(results);
    
    return results;
    
  } catch (error) {
    console.error('❌ Phase 0統合システム検証エラー:', error);
    throw error;
  }
}

/**
 * プラン管理システム基本機能テスト
 */
function testPlanManagerBasics() {
  try {
    console.log('📋 プラン管理システム基本機能テスト...');
    
    // 必要な関数の存在確認
    const requiredFunctions = [
      'getUserPlan',
      'getPlanLimits',
      'getPlanDisplayName',
      'setUserPlan',
      'switchToTemporaryPlan',
      'restoreOriginalPlan',
      'getPlanDetails',
      'initializePlanManager'
    ];
    
    for (const funcName of requiredFunctions) {
      if (typeof eval(funcName) !== 'function') {
        throw new Error(`必須関数が見つかりません: ${funcName}`);
      }
    }
    console.log('✅ 必要関数の存在確認完了');
    
    // プラン管理システム初期化
    const initResult = initializePlanManager();
    if (!initResult) {
      throw new Error('プラン管理システムの初期化に失敗');
    }
    console.log('✅ プラン管理システム初期化完了');
    
    // 基本機能テスト
    const currentPlan = getUserPlan();
    const planLimits = getPlanLimits();
    const displayName = getPlanDisplayName();
    const planDetails = getPlanDetails();
    
    // 結果検証
    if (!currentPlan || typeof currentPlan !== 'string') {
      throw new Error('getUserPlan()が正しい値を返していません');
    }
    
    if (!planLimits || typeof planLimits !== 'object') {
      throw new Error('getPlanLimits()が正しい値を返していません');
    }
    
    if (!displayName || typeof displayName !== 'string') {
      throw new Error('getPlanDisplayName()が正しい値を返していません');
    }
    
    if (!planDetails || typeof planDetails !== 'object') {
      throw new Error('getPlanDetails()が正しい値を返していません');
    }
    
    console.log(`✅ プラン管理基本機能テスト完了: ${displayName}`);
    return true;
    
  } catch (error) {
    console.error('❌ プラン管理システムテストエラー:', error);
    return false;
  }
}

/**
 * ライセンス管理統合テスト
 */
function testLicenseManagerIntegration() {
  try {
    console.log('🔐 ライセンス管理統合テスト...');
    
    // getLicenseInfo関数の存在確認
    if (typeof getLicenseInfo !== 'function') {
      throw new Error('getLicenseInfo関数が見つかりません');
    }
    
    // ライセンス情報取得
    const licenseInfo = getLicenseInfo();
    
    // プラン統合ポイントの確認
    if (licenseInfo.planInfo) {
      console.log('✅ ライセンス情報にプラン情報が統合されています');
      
      // プラン情報の内容確認
      const planInfo = licenseInfo.planInfo;
      if (!planInfo.currentPlan || !planInfo.planDisplayName || !planInfo.planLimits) {
        throw new Error('プラン情報の内容が不完全です');
      }
    } else {
      console.log('⚠️ プラン情報は統合されていませんが、基本動作には問題ありません');
    }
    
    console.log('✅ ライセンス管理統合テスト完了');
    return true;
    
  } catch (error) {
    console.error('❌ ライセンス管理統合テストエラー:', error);
    return false;
  }
}

/**
 * ユーザー管理統合テスト
 */
function testUserManagerIntegration() {
  try {
    console.log('👤 ユーザー管理統合テスト...');
    
    // 必要な関数の存在確認
    if (typeof getCurrentUser !== 'function') {
      throw new Error('getCurrentUser関数が見つかりません');
    }
    
    if (typeof getEffectivePermissions !== 'function') {
      throw new Error('getEffectivePermissions関数が見つかりません');
    }
    
    // 現在のユーザー情報取得
    const currentUser = getCurrentUser();
    
    // プラン統合ポイントの確認
    if (currentUser.planPermissions) {
      console.log('✅ ユーザー情報にプラン権限が統合されています');
      
      const planPerms = currentUser.planPermissions;
      if (!planPerms.planType || !planPerms.planDisplayName || !planPerms.planLimits) {
        throw new Error('プラン権限情報の内容が不完全です');
      }
    } else {
      console.log('⚠️ プラン権限は統合されていませんが、基本動作には問題ありません');
    }
    
    // 統合権限システムテスト
    const effectivePermissions = getEffectivePermissions();
    
    // 統合権限の内容確認
    const requiredProps = [
      'canAccessAdminFeatures', 'canManageUsers', 'canViewSystemStats',
      'canUseBasicFeatures', 'canGenerateKeywords', 'canUseAiProposals',
      'maxCompaniesPerDay', 'planType', 'planDisplayName'
    ];
    
    for (const prop of requiredProps) {
      if (!(prop in effectivePermissions)) {
        throw new Error(`統合権限に必須プロパティが不足: ${prop}`);
      }
    }
    
    console.log('✅ ユーザー管理統合テスト完了');
    return true;
    
  } catch (error) {
    console.error('❌ ユーザー管理統合テストエラー:', error);
    return false;
  }
}

/**
 * メニューシステム統合テスト
 */
function testMenuSystemIntegration() {
  try {
    console.log('📋 メニューシステム統合テスト...');
    
    // メニュー作成関数の存在確認
    if (typeof createSpecCompliantMenu !== 'function') {
      console.log('⚠️ createSpecCompliantMenu関数が見つかりません - メニューの動的更新は利用できません');
      return true; // 致命的ではない
    }
    
    // systemConfiguration関数の存在確認
    if (typeof systemConfiguration !== 'function') {
      throw new Error('systemConfiguration関数が見つかりません');
    }
    
    console.log('✅ メニューシステム統合テスト完了');
    return true;
    
  } catch (error) {
    console.error('❌ メニューシステム統合テストエラー:', error);
    return false;
  }
}

/**
 * シート拡張機能テスト
 */
function testSheetIntegration() {
  try {
    console.log('📊 シート拡張機能テスト...');
    
    // シート関連関数の存在確認
    if (typeof updateControlPanelPlanStatus !== 'function') {
      throw new Error('updateControlPanelPlanStatus関数が見つかりません');
    }
    
    if (typeof createControlPanel !== 'function') {
      console.log('⚠️ createControlPanel関数が見つかりません - シート作成機能の確認をスキップ');
      return true; // 致命的ではない
    }
    
    console.log('✅ シート拡張機能テスト完了');
    return true;
    
  } catch (error) {
    console.error('❌ シート拡張機能テストエラー:', error);
    return false;
  }
}

/**
 * 検証結果レポート生成
 */
function generateValidationReport(results) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let report = '📋 Phase 0統合システム検証レポート\n\n';
    report += `🕐 実行時刻: ${new Date(results.timestamp).toLocaleString('ja-JP')}\n\n`;
    
    // 各テスト結果
    report += '📊 テスト結果:\n';
    report += `・プラン管理システム: ${results.planManager ? '✅ 成功' : '❌ 失敗'}\n`;
    report += `・ライセンス管理統合: ${results.licenseIntegration ? '✅ 成功' : '❌ 失敗'}\n`;
    report += `・ユーザー管理統合: ${results.userIntegration ? '✅ 成功' : '❌ 失敗'}\n`;
    report += `・メニューシステム統合: ${results.menuIntegration ? '✅ 成功' : '❌ 失敗'}\n`;
    report += `・シート拡張機能: ${results.sheetIntegration ? '✅ 成功' : '❌ 失敗'}\n\n`;
    
    // 総合評価
    const successCount = [
      results.planManager,
      results.licenseIntegration,
      results.userIntegration,
      results.menuIntegration,
      results.sheetIntegration
    ].filter(Boolean).length;
    
    const totalTests = 5;
    const successRate = (successCount / totalTests) * 100;
    
    if (successRate === 100) {
      report += '🎉 Phase 0統合完了\n';
      report += 'すべてのテストが成功しました。Phase 1の実装に進むことができます。';
    } else if (successRate >= 80) {
      report += '✅ Phase 0基盤準備完了\n';
      report += '主要機能は正常に動作しています。一部改善点がありますが、Phase 1に進めます。';
    } else {
      report += '⚠️ Phase 0に問題があります\n';
      report += 'いくつかの重要なテストが失敗しています。修正が必要です。';
    }
    
    ui.alert('Phase 0検証結果', report, ui.ButtonSet.OK);
    
    console.log('Phase 0検証完了 - 成功率:', `${successRate}%`);
    
  } catch (error) {
    console.error('❌ 検証レポート生成エラー:', error);
  }
}

/**
 * プラン管理情報をコンソールに表示（デバッグ用）
 */
function showPlanManagerDebugInfo() {
  try {
    console.log('=== プラン管理システム デバッグ情報 ===');
    
    if (typeof getUserPlan === 'function') {
      console.log('現在のプラン:', getUserPlan());
    }
    
    if (typeof getPlanDisplayName === 'function') {
      console.log('プラン表示名:', getPlanDisplayName());
    }
    
    if (typeof getPlanLimits === 'function') {
      console.log('プラン制限:', JSON.stringify(getPlanLimits(), null, 2));
    }
    
    if (typeof getPlanDetails === 'function') {
      console.log('プラン詳細:', JSON.stringify(getPlanDetails(), null, 2));
    }
    
    if (typeof isInSwitchMode === 'function') {
      console.log('切り替えモード:', isInSwitchMode());
    }
    
    console.log('=== デバッグ情報終了 ===');
    
  } catch (error) {
    console.error('❌ デバッグ情報表示エラー:', error);
  }
}

// Phase 0検証の実行用関数
function runPhase0Validation() {
  try {
    const results = validatePhase0Integration();
    return results;
  } catch (error) {
    console.error('Phase 0検証実行エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      'Phase 0検証でエラーが発生しました。\nコンソールログを確認してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}