/**
 * Phase 0・Phase 1統合機能保護システム
 * ガバナンス修正時に新しく実装した機能を保護
 */

/**
 * Phase 0・Phase 1で追加された統合機能のバックアップ
 */
function backupPhase01Integrations() {
  try {
    console.log('🛡️ Phase 0・Phase 1統合機能バックアップ開始...');
    
    const properties = PropertiesService.getScriptProperties();
    const backup = {};
    
    // 1. プラン管理システムの状態バックアップ
    backup.planManager = {
      currentPlan: properties.getProperty('USER_PLAN'),
      planHistory: properties.getProperty('PLAN_HISTORY'),
      originalPlan: properties.getProperty('ORIGINAL_PLAN'),
      switchMode: properties.getProperty('SWITCH_MODE')
    };
    
    // 2. 使用量追跡システムの状態バックアップ
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    backup.usageTracker = {
      todayUsage: properties.getProperty(`usage_${today}`),
      lastResetDate: properties.getProperty('last_usage_reset_date'),
      usageHistory: properties.getProperty('usage_history')
    };
    
    // 3. バックアップデータを保存
    const backupString = JSON.stringify(backup);
    properties.setProperty('PHASE01_BACKUP', backupString);
    properties.setProperty('BACKUP_TIMESTAMP', new Date().toISOString());
    
    console.log('✅ Phase 0・Phase 1統合機能バックアップ完了');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 0・Phase 1統合機能バックアップエラー:', error);
    return false;
  }
}

/**
 * Phase 0・Phase 1統合機能の復元
 */
function restorePhase01Integrations() {
  try {
    console.log('🔄 Phase 0・Phase 1統合機能復元開始...');
    
    const properties = PropertiesService.getScriptProperties();
    const backupString = properties.getProperty('PHASE01_BACKUP');
    
    if (!backupString) {
      console.log('⚠️ バックアップデータが見つかりません');
      return false;
    }
    
    const backup = JSON.parse(backupString);
    
    // 1. プラン管理システムの復元
    if (backup.planManager) {
      if (backup.planManager.currentPlan) {
        properties.setProperty('USER_PLAN', backup.planManager.currentPlan);
      }
      if (backup.planManager.planHistory) {
        properties.setProperty('PLAN_HISTORY', backup.planManager.planHistory);
      }
      if (backup.planManager.originalPlan) {
        properties.setProperty('ORIGINAL_PLAN', backup.planManager.originalPlan);
      }
      if (backup.planManager.switchMode) {
        properties.setProperty('SWITCH_MODE', backup.planManager.switchMode);
      }
      console.log('✅ プラン管理システム復元完了');
    }
    
    // 2. 使用量追跡システムの復元
    if (backup.usageTracker) {
      if (backup.usageTracker.todayUsage) {
        const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
        properties.setProperty(`usage_${today}`, backup.usageTracker.todayUsage);
      }
      if (backup.usageTracker.lastResetDate) {
        properties.setProperty('last_usage_reset_date', backup.usageTracker.lastResetDate);
      }
      if (backup.usageTracker.usageHistory) {
        properties.setProperty('usage_history', backup.usageTracker.usageHistory);
      }
      console.log('✅ 使用量追跡システム復元完了');
    }
    
    console.log('✅ Phase 0・Phase 1統合機能復元完了');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 0・Phase 1統合機能復元エラー:', error);
    return false;
  }
}

/**
 * Phase 0・Phase 1統合機能の健全性チェック
 */
function validatePhase01Integrations() {
  try {
    console.log('🔍 Phase 0・Phase 1統合機能健全性チェック...');
    
    const results = {
      planManager: false,
      usageTracker: false,
      limitChecker: false,
      menuIntegration: false,
      errors: []
    };
    
    // 1. プラン管理システムチェック
    try {
      if (typeof getUserPlan === 'function' && typeof getPlanLimits === 'function') {
        const plan = getUserPlan();
        const limits = getPlanLimits();
        if (plan && limits) {
          results.planManager = true;
        }
      }
    } catch (error) {
      results.errors.push('プラン管理システム: ' + error.message);
    }
    
    // 2. 使用量追跡システムチェック  
    try {
      if (typeof UsageTracker !== 'undefined' && typeof UsageTracker.getDailyUsage === 'function') {
        UsageTracker.getDailyUsage('company_search');
        results.usageTracker = true;
      }
    } catch (error) {
      results.errors.push('使用量追跡システム: ' + error.message);
    }
    
    // 3. 制限チェックシステムチェック
    try {
      if (typeof checkPlanLimit === 'function') {
        checkPlanLimit('company_search', 1);
        results.limitChecker = true;
      }
    } catch (error) {
      results.errors.push('制限チェックシステム: ' + error.message);
    }
    
    // 4. メニュー統合チェック
    try {
      if (typeof systemConfiguration === 'function') {
        results.menuIntegration = true;
      }
    } catch (error) {
      results.errors.push('メニュー統合: ' + error.message);
    }
    
    // 結果レポート
    const successCount = Object.values(results).filter(v => v === true).length;
    const totalTests = 4;
    
    console.log(`Phase 0・Phase 1統合機能健全性: ${successCount}/${totalTests} 正常`);
    
    if (results.errors.length > 0) {
      console.warn('⚠️ 検出された問題:', results.errors);
    }
    
    return results;
    
  } catch (error) {
    console.error('❌ Phase 0・Phase 1統合機能健全性チェックエラー:', error);
    return { planManager: false, usageTracker: false, limitChecker: false, menuIntegration: false, errors: [error.message] };
  }
}

/**
 * ガバナンス修正後の統合機能自動復旧
 */
function autoRecoverPhase01AfterGovernance() {
  try {
    console.log('🔧 ガバナンス修正後の自動復旧開始...');
    
    const ui = SpreadsheetApp.getUi();
    
    // 1. バックアップからの復元
    const restoreResult = restorePhase01Integrations();
    
    // 2. システム再初期化
    if (restoreResult) {
      console.log('🚀 Phase 0・Phase 1システム再初期化...');
      
      // プラン管理システム再初期化
      if (typeof initializePlanManager === 'function') {
        initializePlanManager();
      }
      
      // 使用量追跡システム再初期化
      if (typeof initializeUsageTracker === 'function') {
        initializeUsageTracker();
      }
      
      // 制限チェックシステム再初期化
      if (typeof initializeLimitCheckSystem === 'function') {
        initializeLimitCheckSystem();
      }
    }
    
    // 3. 健全性チェック
    const validation = validatePhase01Integrations();
    
    // 4. 結果レポート
    let report = '🔧 ガバナンス修正後の自動復旧完了\n\n';
    report += '📊 Phase 0・Phase 1機能状況:\n';
    report += `・プラン管理システム: ${validation.planManager ? '✅ 正常' : '❌ 問題あり'}\n`;
    report += `・使用量追跡システム: ${validation.usageTracker ? '✅ 正常' : '❌ 問題あり'}\n`;
    report += `・制限チェックシステム: ${validation.limitChecker ? '✅ 正常' : '❌ 問題あり'}\n`;
    report += `・メニュー統合: ${validation.menuIntegration ? '✅ 正常' : '❌ 問題あり'}\n\n`;
    
    if (validation.errors.length > 0) {
      report += '⚠️ 検出された問題:\n';
      validation.errors.forEach(error => {
        report += `・${error}\n`;
      });
      report += '\n手動での再設定が必要な場合があります。';
    } else {
      report += '🎉 すべてのPhase 0・Phase 1機能が正常に復旧しました！';
    }
    
    ui.alert('自動復旧完了', report, ui.ButtonSet.OK);
    
    return validation;
    
  } catch (error) {
    console.error('❌ ガバナンス修正後の自動復旧エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '自動復旧でエラーが発生しました。手動での復旧が必要です。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}

/**
 * 安全なガバナンス修正実行
 * バックアップ → ガバナンス修正 → 復旧の一貫処理
 */
function safeGovernanceRepair() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 確認ダイアログ
    const response = ui.alert(
      '安全なガバナンス修正',
      'Phase 0・Phase 1機能を保護しながらガバナンス修正を実行します。\n\n処理内容:\n1. 現在の状態をバックアップ\n2. ガバナンス違反を修正\n3. Phase 0・Phase 1機能を復旧\n\n実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return false;
    }
    
    console.log('🛡️ 安全なガバナンス修正開始...');
    
    // 1. バックアップ
    const backupResult = backupPhase01Integrations();
    if (!backupResult) {
      throw new Error('バックアップに失敗しました');
    }
    
    // 2. ガバナンス修正（プレースホルダー作成）
    if (typeof createMissingPlaceholders === 'function') {
      createMissingPlaceholders();
    } else {
      throw new Error('createMissingPlaceholders関数が見つかりません');
    }
    
    // 3. 自動復旧
    const recoveryResult = autoRecoverPhase01AfterGovernance();
    
    if (recoveryResult) {
      ui.alert(
        '安全修正完了',
        '✅ ガバナンス修正とPhase 0・Phase 1機能の復旧が完了しました！\n\n次のステップ:\n1. runPhase0Validation()実行\n2. runPhase1Validation()実行\n3. 実際の機能テスト',
        ui.ButtonSet.OK
      );
      return true;
    } else {
      throw new Error('自動復旧で問題が発生しました');
    }
    
  } catch (error) {
    console.error('❌ 安全なガバナンス修正エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '安全なガバナンス修正でエラーが発生しました。\n\n手動での修正が必要です:\n1. restorePhase01Integrations()実行\n2. ガバナンス修正の再実行',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}