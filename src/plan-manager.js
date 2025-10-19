/**
 * 料金プラン管理システム - Phase 0 基盤実装
 * システム仕様書v2.0準拠
 */

// プラン定数（仕様書v2.0準拠）
const PLAN_TYPES = {
  TRIAL: 'TRIAL',           // 🆓 試用期間（10日間）
  BASIC: 'BASIC',           // 🥉 ベーシック（¥500/月）
  STANDARD: 'STANDARD',     // 🥈 スタンダード（¥1,500/月）
  PROFESSIONAL: 'PROFESSIONAL', // 🥇 プロフェッショナル（¥3,000/月）
  ENTERPRISE: 'ENTERPRISE'  // 💎 エンタープライズ（¥7,500/月）
};

// プラン表示名（仕様書v2.0準拠）
const PLAN_DISPLAY_NAMES = {
  TRIAL: '🆓 試用期間',
  BASIC: '🥉 ベーシック',
  STANDARD: '🥈 スタンダード',
  PROFESSIONAL: '🥇 プロフェッショナル',
  ENTERPRISE: '💎 エンタープライズ'
};

// プラン制限値定義（仕様書v2.0準拠）
const PLAN_LIMITS = {
  [PLAN_TYPES.TRIAL]: {
    keywordGeneration: true,
    maxCompaniesPerDay: 5,
    aiProposals: true,
    requiresApiKey: true,
    monthlyPrice: 0,
    features: ['全機能お試し', '10日間限定'],
    description: '無料（10日間）'
  },
  [PLAN_TYPES.BASIC]: {
    keywordGeneration: false,
    maxCompaniesPerDay: 10,
    aiProposals: false,
    requiresApiKey: false,
    monthlyPrice: 500,
    features: ['手動入力・基本テンプレートのみ', '企業検索10社/日'],
    description: '¥500/月'
  },
  [PLAN_TYPES.STANDARD]: {
    keywordGeneration: true,
    maxCompaniesPerDay: 50,
    aiProposals: true,
    requiresApiKey: true,
    monthlyPrice: 1500,
    features: ['全機能・AI活用', '企業検索50社/日'],
    description: '¥1,500/月'
  },
  [PLAN_TYPES.PROFESSIONAL]: {
    keywordGeneration: true,
    maxCompaniesPerDay: 100,
    aiProposals: true,
    requiresApiKey: true,
    monthlyPrice: 3000,
    features: ['高速処理・優先サポート', '企業検索100社/日'],
    description: '¥3,000/月'
  },
  [PLAN_TYPES.ENTERPRISE]: {
    keywordGeneration: true,
    maxCompaniesPerDay: 500,
    aiProposals: true,
    requiresApiKey: true,
    monthlyPrice: 7500,
    features: ['最大性能・専任サポート', '企業検索500社/日'],
    description: '¥7,500/月'
  }
};

/**
 * 現在のユーザープランを取得
 * 既存ライセンス管理システムとの統合
 * @returns {string} プランタイプ
 */
function getUserPlan() {
  try {
    // PropertiesServiceからプラン情報を取得
    const properties = PropertiesService.getScriptProperties();
    let userPlan = properties.getProperty('USER_PLAN');
    
    // プランが設定されていない場合は試用期間として設定
    if (!userPlan) {
      userPlan = PLAN_TYPES.TRIAL;
      properties.setProperty('USER_PLAN', userPlan);
      
      // 試用期間開始日を記録（既存ライセンス管理と連携）
      const licenseInfo = getLicenseInfo();
      if (!licenseInfo.startDate) {
        setLicenseStartDate();
      }
    }
    
    // 既存ライセンス管理との統合チェック
    const licenseInfo = getLicenseInfo();
    if (licenseInfo.isExpired && userPlan === PLAN_TYPES.TRIAL) {
      // 試用期間が終了している場合は自動的にベーシックプランに移行
      userPlan = PLAN_TYPES.BASIC;
      properties.setProperty('USER_PLAN', userPlan);
      console.log('🔄 試用期間終了により自動的にベーシックプランに移行');
    }
    
    return userPlan;
    
  } catch (error) {
    console.error('❌ プラン取得エラー:', error);
    // エラー時は安全にベーシックプランを返す
    return PLAN_TYPES.BASIC;
  }
}

/**
 * プラン制限値を取得
 * @param {string} planType プランタイプ（オプション、未指定時は現在のプランを使用）
 * @returns {Object} プラン制限値
 */
function getPlanLimits(planType = null) {
  try {
    const plan = planType || getUserPlan();
    
    if (!PLAN_LIMITS[plan]) {
      console.warn(`⚠️ 未知のプランタイプ: ${plan}, ベーシックプランの制限値を使用`);
      return PLAN_LIMITS[PLAN_TYPES.BASIC];
    }
    
    return PLAN_LIMITS[plan];
    
  } catch (error) {
    console.error('❌ プラン制限値取得エラー:', error);
    // エラー時は最も制限の厳しいベーシックプランを返す
    return PLAN_LIMITS[PLAN_TYPES.BASIC];
  }
}

/**
 * プラン表示名を取得
 * @param {string} planType プランタイプ（オプション）
 * @returns {string} プラン表示名
 */
function getPlanDisplayName(planType = null) {
  try {
    const plan = planType || getUserPlan();
    return PLAN_DISPLAY_NAMES[plan] || '🥉 ベーシック';
  } catch (error) {
    console.error('❌ プラン表示名取得エラー:', error);
    return '🥉 ベーシック';
  }
}

/**
 * ユーザープランを設定
 * @param {string} newPlan 新しいプランタイプ
 * @returns {boolean} 成功フラグ
 */
function setUserPlan(newPlan) {
  try {
    if (!PLAN_TYPES[newPlan]) {
      throw new Error(`無効なプランタイプ: ${newPlan}`);
    }
    
    const properties = PropertiesService.getScriptProperties();
    const oldPlan = properties.getProperty('USER_PLAN');
    
    // プラン変更をログに記録
    const changeLog = {
      timestamp: new Date().toISOString(),
      oldPlan: oldPlan,
      newPlan: newPlan,
      user: Session.getActiveUser().getEmail()
    };
    
    // プラン履歴に追加
    const history = JSON.parse(properties.getProperty('PLAN_HISTORY') || '[]');
    history.push(changeLog);
    
    // 履歴は最大50件まで保持
    if (history.length > 50) {
      history.splice(0, history.length - 50);
    }
    
    // プラン情報を保存
    properties.setProperties({
      'USER_PLAN': newPlan,
      'PLAN_HISTORY': JSON.stringify(history),
      'PLAN_CHANGE_DATE': new Date().toISOString()
    });
    
    console.log(`✅ プラン変更完了: ${oldPlan} → ${newPlan}`);
    return true;
    
  } catch (error) {
    console.error('❌ プラン設定エラー:', error);
    return false;
  }
}

/**
 * 管理者モード用一時的プラン切り替え
 * 元のプランを保持しつつ、テスト用プランに切り替え
 * @param {string} temporaryPlan 一時的に設定するプラン
 * @returns {boolean} 成功フラグ
 */
function switchToTemporaryPlan(temporaryPlan) {
  try {
    if (!PLAN_TYPES[temporaryPlan]) {
      throw new Error(`無効なプランタイプ: ${temporaryPlan}`);
    }
    
    const properties = PropertiesService.getScriptProperties();
    const currentPlan = getUserPlan();
    
    // 元のプランを保存
    properties.setProperty('ORIGINAL_PLAN', currentPlan);
    properties.setProperty('SWITCH_MODE', 'true');
    
    // 一時的プランに切り替え
    return setUserPlan(temporaryPlan);
    
  } catch (error) {
    console.error('❌ 一時的プラン切り替えエラー:', error);
    return false;
  }
}

/**
 * 元のプランに復元
 * @returns {boolean} 成功フラグ
 */
function restoreOriginalPlan() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const originalPlan = properties.getProperty('ORIGINAL_PLAN');
    
    if (!originalPlan) {
      console.warn('⚠️ 復元する元のプランが見つかりません');
      return false;
    }
    
    // 元のプランに復元
    const success = setUserPlan(originalPlan);
    
    if (success) {
      // 切り替えモードを解除
      properties.deleteProperty('ORIGINAL_PLAN');
      properties.deleteProperty('SWITCH_MODE');
      console.log(`✅ 元のプランに復元完了: ${originalPlan}`);
    }
    
    return success;
    
  } catch (error) {
    console.error('❌ プラン復元エラー:', error);
    return false;
  }
}

/**
 * 現在切り替えモード中かチェック
 * @returns {boolean} 切り替えモード中フラグ
 */
function isInSwitchMode() {
  try {
    const properties = PropertiesService.getScriptProperties();
    return properties.getProperty('SWITCH_MODE') === 'true';
  } catch (error) {
    console.error('❌ 切り替えモード確認エラー:', error);
    return false;
  }
}

/**
 * プランの詳細情報を取得
 * @param {string} planType プランタイプ（オプション）
 * @returns {Object} プラン詳細情報
 */
function getPlanDetails(planType = null) {
  try {
    const plan = planType || getUserPlan();
    const limits = getPlanLimits(plan);
    const displayName = getPlanDisplayName(plan);
    
    return {
      planType: plan,
      displayName: displayName,
      limits: limits,
      isTemporary: isInSwitchMode(),
      originalPlan: isInSwitchMode() ? 
        PropertiesService.getScriptProperties().getProperty('ORIGINAL_PLAN') : null
    };
    
  } catch (error) {
    console.error('❌ プラン詳細情報取得エラー:', error);
    return {
      planType: PLAN_TYPES.BASIC,
      displayName: PLAN_DISPLAY_NAMES.BASIC,
      limits: PLAN_LIMITS[PLAN_TYPES.BASIC],
      isTemporary: false,
      originalPlan: null
    };
  }
}

/**
 * プラン管理情報をシートに表示（デバッグ・管理者用）
 */
function showPlanManagementInfo() {
  try {
    const ui = SpreadsheetApp.getUi();
    const planDetails = getPlanDetails();
    const licenseInfo = getLicenseInfo();
    
    let message = `🚀 プラン管理情報\n\n`;
    message += `📊 現在のプラン: ${planDetails.displayName}\n`;
    message += `🔢 プランタイプ: ${planDetails.planType}\n`;
    message += `💰 月額料金: ¥${planDetails.limits.monthlyPrice.toLocaleString()}\n\n`;
    
    message += `🎯 制限値:\n`;
    message += `・キーワード生成: ${planDetails.limits.keywordGeneration ? '✅' : '❌'}\n`;
    message += `・企業検索上限: ${planDetails.limits.maxCompaniesPerDay}社/日\n`;
    message += `・AI提案生成: ${planDetails.limits.aiProposals ? '✅' : '❌'}\n`;
    message += `・APIキー必要: ${planDetails.limits.requiresApiKey ? '✅' : '❌'}\n\n`;
    
    if (planDetails.isTemporary) {
      message += `🔄 一時切り替えモード中\n`;
      message += `📌 元のプラン: ${getPlanDisplayName(planDetails.originalPlan)}\n\n`;
    }
    
    message += `📅 ライセンス情報:\n`;
    message += `・管理者モード: ${licenseInfo.adminMode ? '✅' : '❌'}\n`;
    message += `・ライセンス状態: ${licenseInfo.isExpired ? '🔴 期限切れ' : '🟢 有効'}\n`;
    
    if (licenseInfo.remainingDays !== null) {
      message += `・残り日数: ${licenseInfo.remainingDays}営業日\n`;
    }
    
    ui.alert('プラン管理情報', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ プラン管理情報表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'プラン情報の取得に失敗しました', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 既存ライセンス管理システムとの統合チェック
 * 既存のgetLicenseInfo()関数が存在するかチェックし、統合
 */
function checkLicenseManagerIntegration() {
  try {
    // 既存のライセンス管理関数の存在確認
    if (typeof getLicenseInfo === 'function') {
      console.log('✅ 既存ライセンス管理システムとの統合OK');
      return true;
    } else {
      console.warn('⚠️ 既存ライセンス管理システムが見つかりません');
      return false;
    }
  } catch (error) {
    console.error('❌ ライセンス管理統合チェックエラー:', error);
    return false;
  }
}

/**
 * プラン管理システム初期化
 */
function initializePlanManager() {
  try {
    console.log('🚀 プラン管理システム初期化開始...');
    
    // 既存システムとの統合確認
    const isIntegrated = checkLicenseManagerIntegration();
    
    // 初期プラン設定
    const currentPlan = getUserPlan();
    console.log(`📊 現在のプラン: ${getPlanDisplayName(currentPlan)}`);
    
    // プラン履歴の初期化
    const properties = PropertiesService.getScriptProperties();
    if (!properties.getProperty('PLAN_HISTORY')) {
      properties.setProperty('PLAN_HISTORY', JSON.stringify([]));
    }
    
    console.log('✅ プラン管理システム初期化完了');
    return true;
    
  } catch (error) {
    console.error('❌ プラン管理システム初期化エラー:', error);
    return false;
  }
}

// Phase 0で必要な基本機能をエクスポート（GAS環境では直接利用）
// export { getUserPlan, getPlanLimits, getPlanDisplayName, setUserPlan, switchToTemporaryPlan, restoreOriginalPlan, isInSwitchMode, getPlanDetails };