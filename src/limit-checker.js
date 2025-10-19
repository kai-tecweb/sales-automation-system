/**
 * プラン制限チェック・エラーハンドリングシステム - Phase 1実装
 * システム仕様書v2.0準拠の制限チェックとユーザー通知
 */

/**
 * 機能実行前のプラン制限チェック
 * @param {string} action 実行する機能（'keyword_generation', 'company_search', 'ai_proposal'等）
 * @param {number} requestAmount 要求する使用量（デフォルト1）
 * @returns {Object} チェック結果
 */
function checkPlanLimit(action, requestAmount = 1) {
  try {
    console.log(`🔍 プラン制限チェック: ${action} (要求量: ${requestAmount})`);
    
    // プラン詳細を取得
    const planDetails = getPlanDetails();
    const limits = planDetails.limits;
    
    let result = {
      allowed: true,
      reason: 'OK',
      action: action,
      requestAmount: requestAmount,
      planType: planDetails.planType,
      planDisplayName: planDetails.displayName
    };
    
    switch (action) {
      case 'keyword_generation':
        if (!limits.keywordGeneration) {
          result.allowed = false;
          result.reason = 'FEATURE_NOT_AVAILABLE';
          result.requiredPlan = 'STANDARD';
          result.message = 'キーワード自動生成機能はスタンダードプラン以上で利用可能です';
        }
        break;
        
      case 'company_search':
        // 使用量制限チェック
        const usageCheck = UsageTracker.checkUsageLimit(USAGE_TYPES.COMPANY_SEARCH, requestAmount);
        if (!usageCheck.allowed) {
          result.allowed = false;
          result.reason = usageCheck.reason;
          result.currentUsage = usageCheck.currentUsage;
          result.limit = usageCheck.limit;
          result.remaining = usageCheck.remaining;
          
          if (usageCheck.reason === 'DAILY_LIMIT_EXCEEDED') {
            result.message = `本日の企業検索上限（${usageCheck.limit}社）に達しました`;
          }
        } else {
          result.usageInfo = usageCheck;
        }
        break;
        
      case 'ai_proposal':
        if (!limits.aiProposals) {
          result.allowed = false;
          result.reason = 'FEATURE_NOT_AVAILABLE';
          result.requiredPlan = 'STANDARD';
          result.message = 'AI提案生成機能はスタンダードプラン以上で利用可能です';
        } else {
          // API使用量チェック
          const apiUsageCheck = UsageTracker.checkUsageLimit(USAGE_TYPES.AI_PROPOSAL, requestAmount);
          if (!apiUsageCheck.allowed) {
            result.allowed = false;
            result.reason = apiUsageCheck.reason;
            result.currentUsage = apiUsageCheck.currentUsage;
            result.limit = apiUsageCheck.limit;
            result.message = 'AI提案生成の日次制限に達しました';
          } else {
            result.usageInfo = apiUsageCheck;
          }
        }
        break;
        
      case 'api_key_required':
        if (limits.requiresApiKey) {
          // APIキー設定状況を確認
          const properties = PropertiesService.getScriptProperties();
          const openaiKey = properties.getProperty('OPENAI_API_KEY');
          
          if (!openaiKey || openaiKey.trim() === '') {
            result.allowed = false;
            result.reason = 'API_KEY_REQUIRED';
            result.message = '現在のプランではChatGPT APIキーの設定が必要です';
          }
        }
        break;
        
      default:
        console.warn(`⚠️ 未知のアクション: ${action}`);
        break;
    }
    
    console.log(`${result.allowed ? '✅' : '❌'} 制限チェック結果: ${result.reason}`);
    return result;
    
  } catch (error) {
    console.error('❌ プラン制限チェックエラー:', error);
    return {
      allowed: false,
      reason: 'CHECK_ERROR',
      action: action,
      error: error.message
    };
  }
}

/**
 * 複数機能の一括制限チェック
 * @param {Array} actions 機能名の配列
 * @returns {Object} 各機能のチェック結果
 */
function checkMultiplePlanLimits(actions) {
  try {
    const results = {};
    
    actions.forEach(action => {
      results[action] = checkPlanLimit(action);
    });
    
    return results;
    
  } catch (error) {
    console.error('❌ 複数プラン制限チェックエラー:', error);
    return {};
  }
}

/**
 * 機能実行前の統合チェック（プラン制限 + ライセンス）
 * @param {string} action 実行する機能
 * @param {number} requestAmount 要求する使用量
 * @returns {Object} 統合チェック結果
 */
function performIntegratedCheck(action, requestAmount = 1) {
  try {
    console.log(`🔍 統合チェック開始: ${action}`);
    
    const result = {
      allowed: true,
      errors: [],
      warnings: [],
      action: action
    };
    
    // 1. ライセンス状態チェック
    const licenseInfo = getLicenseInfo();
    if (licenseInfo.systemLocked) {
      result.allowed = false;
      result.errors.push('システムがライセンス期限によりロックされています');
    }
    
    if (licenseInfo.isExpired) {
      result.warnings.push('ライセンスが期限切れです');
    }
    
    // 2. プラン制限チェック
    const planCheck = checkPlanLimit(action, requestAmount);
    if (!planCheck.allowed) {
      result.allowed = false;
      result.planCheckResult = planCheck;
      result.errors.push(planCheck.message || `プラン制限により${action}は利用できません`);
    }
    
    // 3. ユーザー権限チェック（統合権限システム利用）
    try {
      const permissions = getEffectivePermissions();
      
      // 基本機能アクセス権限確認
      if (!permissions.canUseBasicFeatures) {
        result.allowed = false;
        result.errors.push('基本機能の利用権限がありません');
      }
      
      // 管理者専用機能のチェック
      if (action.includes('admin') && !permissions.canAccessAdminFeatures) {
        result.allowed = false;
        result.errors.push('管理者権限が必要です');
      }
      
    } catch (error) {
      result.warnings.push('権限チェックでエラーが発生しました');
    }
    
    console.log(`${result.allowed ? '✅' : '❌'} 統合チェック完了`);
    return result;
    
  } catch (error) {
    console.error('❌ 統合チェックエラー:', error);
    return {
      allowed: false,
      errors: ['システムエラーが発生しました'],
      action: action,
      error: error.message
    };
  }
}

/**
 * プラン制限エラー表示
 * @param {Object} checkResult checkPlanLimit()の結果
 */
function showPlanLimitError(checkResult) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    let message = '🚫 プラン制限エラー\n\n';
    
    switch (checkResult.reason) {
      case 'FEATURE_NOT_AVAILABLE':
        message += `${checkResult.message}\n\n`;
        message += `現在のプラン: ${checkResult.planDisplayName}\n`;
        if (checkResult.requiredPlan) {
          const requiredPlanName = getPlanDisplayName(checkResult.requiredPlan);
          message += `必要プラン: ${requiredPlanName}以上\n\n`;
          message += 'プランをアップグレードして、より多くの機能をご利用ください。';
        }
        break;
        
      case 'DAILY_LIMIT_EXCEEDED':
        message += `${checkResult.message}\n\n`;
        message += `現在の使用量: ${checkResult.currentUsage}/${checkResult.limit}\n`;
        message += `残り使用可能数: ${checkResult.remaining}\n\n`;
        message += '明日00:00にリセットされます。\n';
        message += 'より多くの検索をご希望の場合は、プランアップグレードをご検討ください。';
        break;
        
      case 'API_KEY_REQUIRED':
        message += `${checkResult.message}\n\n`;
        message += '設定方法:\n';
        message += '1. メニューから「⚙️設定」→「🔑APIキー設定」を選択\n';
        message += '2. ChatGPT APIキーを入力してください';
        break;
        
      default:
        message += `機能「${checkResult.action}」は現在利用できません。\n\n`;
        message += `理由: ${checkResult.reason}\n`;
        if (checkResult.message) {
          message += `詳細: ${checkResult.message}`;
        }
        break;
    }
    
    ui.alert('プラン制限', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ プラン制限エラー表示失敗:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      'プラン制限の確認でエラーが発生しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * アップグレード案内ダイアログ表示
 * @param {string} feature 機能名
 * @param {string} requiredPlan 必要なプラン
 */
function showUpgradeDialog(feature, requiredPlan) {
  try {
    const ui = SpreadsheetApp.getUi();
    const currentPlan = getPlanDisplayName();
    const requiredPlanName = getPlanDisplayName(requiredPlan);
    const requiredLimits = getPlanLimits(requiredPlan);
    
    let message = `🚀 プランアップグレードのご案内\n\n`;
    message += `「${feature}」機能をご利用いただくには、${requiredPlanName}以上のプランが必要です。\n\n`;
    
    message += `📊 現在のプラン: ${currentPlan}\n`;
    message += `🎯 推奨プラン: ${requiredPlanName}\n`;
    message += `💰 料金: ¥${requiredLimits.monthlyPrice.toLocaleString()}/月\n\n`;
    
    message += `✨ ${requiredPlanName}の特徴:\n`;
    requiredLimits.features.forEach(feature => {
      message += `・${feature}\n`;
    });
    
    message += `\n管理者にお問い合わせいただくか、プラン変更をご検討ください。`;
    
    ui.alert('プランアップグレード', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ アップグレード案内表示エラー:', error);
  }
}

/**
 * 日次制限達成ダイアログ表示
 * @param {number} limit 日次制限値
 * @param {string} type 制限タイプ（オプション）
 */
function showDailyLimitDialog(limit, type = '企業検索') {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const message = `🚫 日次制限達成\n\n本日の${type}上限（${limit}社）に達しました。\n\n明日00:00にリセットされます。\n\nより多くの検索をご希望の場合は、プランアップグレードをご検討ください。`;
    
    ui.alert('使用量制限', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 日次制限ダイアログ表示エラー:', error);
  }
}

/**
 * 機能実行前の制限チェックと使用量記録を統合した関数
 * @param {string} action 機能名
 * @param {Function} executionFunction 実際の機能実行関数
 * @param {number} requestAmount 要求使用量
 * @returns {*} 機能実行結果またはnull（制限により実行されなかった場合）
 */
function executeWithLimitCheck(action, executionFunction, requestAmount = 1) {
  try {
    console.log(`🔍 制限チェック付き実行: ${action}`);
    
    // 統合チェック実行
    const checkResult = performIntegratedCheck(action, requestAmount);
    
    if (!checkResult.allowed) {
      // エラー表示
      if (checkResult.planCheckResult) {
        showPlanLimitError(checkResult.planCheckResult);
      } else {
        SpreadsheetApp.getUi().alert(
          'アクセス制限',
          checkResult.errors.join('\n'),
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return null;
    }
    
    // 警告があれば表示
    if (checkResult.warnings.length > 0) {
      console.warn('⚠️ 警告:', checkResult.warnings.join(', '));
    }
    
    // 使用量記録（実行前）
    let usageType = null;
    switch (action) {
      case 'company_search':
        usageType = USAGE_TYPES.COMPANY_SEARCH;
        break;
      case 'keyword_generation':
        usageType = USAGE_TYPES.KEYWORD_GENERATION;
        break;
      case 'ai_proposal':
        usageType = USAGE_TYPES.AI_PROPOSAL;
        break;
    }
    
    if (usageType) {
      UsageTracker.incrementUsage(usageType, requestAmount);
    }
    
    // 機能実行
    const result = executionFunction();
    
    console.log('✅ 制限チェック付き実行完了');
    return result;
    
  } catch (error) {
    console.error('❌ 制限チェック付き実行エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '機能の実行でエラーが発生しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}

/**
 * 制限チェックシステムの初期化
 */
function initializeLimitCheckSystem() {
  try {
    console.log('🚀 制限チェックシステム初期化開始...');
    
    // 依存システムの確認
    const dependencies = [
      { name: 'プラン管理', function: 'getPlanDetails' },
      { name: '使用量追跡', function: 'UsageTracker.checkUsageLimit' },
      { name: 'ライセンス管理', function: 'getLicenseInfo' }
    ];
    
    const missingDeps = [];
    dependencies.forEach(dep => {
      try {
        if (dep.function.includes('.')) {
          const [obj, method] = dep.function.split('.');
          if (typeof eval(obj)[method] !== 'function') {
            missingDeps.push(dep.name);
          }
        } else {
          if (typeof eval(dep.function) !== 'function') {
            missingDeps.push(dep.name);
          }
        }
      } catch (error) {
        missingDeps.push(dep.name);
      }
    });
    
    if (missingDeps.length > 0) {
      console.warn('⚠️ 依存システムが不完全です:', missingDeps.join(', '));
      return false;
    }
    
    console.log('✅ 制限チェックシステム初期化完了');
    return true;
    
  } catch (error) {
    console.error('❌ 制限チェックシステム初期化エラー:', error);
    return false;
  }
}