/**
 * 使用量追跡・制限システム - Phase 1実装
 * システム仕様書v2.0準拠の日次使用量管理
 */

// 使用量追跡定数
const USAGE_TYPES = {
  COMPANY_SEARCH: 'company_search',
  KEYWORD_GENERATION: 'keyword_generation',
  AI_PROPOSAL: 'ai_proposal',
  API_CALL: 'api_call'
};

// 使用量追跡のためのプロパティキー
const USAGE_PROPERTIES = {
  DAILY_USAGE_PREFIX: 'usage_',
  LAST_RESET_DATE: 'last_usage_reset_date',
  USAGE_HISTORY: 'usage_history'
};

/**
 * 使用量追跡クラス
 */
class UsageTracker {
  
  /**
   * 使用量をインクリメント
   * @param {string} type 使用量タイプ（USAGE_TYPES定数を使用）
   * @param {number} amount 増加量（デフォルト1）
   * @returns {number} 更新後の使用量
   */
  static incrementUsage(type, amount = 1) {
    try {
      const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      const properties = PropertiesService.getScriptProperties();
      
      const usageKey = `${USAGE_PROPERTIES.DAILY_USAGE_PREFIX}${today}`;
      const usage = JSON.parse(properties.getProperty(usageKey) || '{}');
      
      usage[type] = (usage[type] || 0) + amount;
      
      properties.setProperty(usageKey, JSON.stringify(usage));
      
      // 使用量履歴に記録
      this._recordUsageHistory(type, amount, usage[type]);
      
      console.log(`使用量更新: ${type} +${amount} (合計: ${usage[type]})`);
      return usage[type];
      
    } catch (error) {
      console.error('❌ 使用量インクリメントエラー:', error);
      return 0;
    }
  }
  
  /**
   * 日次使用量を取得
   * @param {string} type 使用量タイプ
   * @param {string} date 日付（YYYY-MM-DD、未指定時は今日）
   * @returns {number} 使用量
   */
  static getDailyUsage(type, date = null) {
    try {
      const targetDate = date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      const properties = PropertiesService.getScriptProperties();
      
      const usageKey = `${USAGE_PROPERTIES.DAILY_USAGE_PREFIX}${targetDate}`;
      const usage = JSON.parse(properties.getProperty(usageKey) || '{}');
      
      return usage[type] || 0;
      
    } catch (error) {
      console.error('❌ 日次使用量取得エラー:', error);
      return 0;
    }
  }
  
  /**
   * 全使用量タイプの日次使用量を取得
   * @param {string} date 日付（YYYY-MM-DD、未指定時は今日）
   * @returns {Object} 全使用量データ
   */
  static getAllDailyUsage(date = null) {
    try {
      const targetDate = date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      const properties = PropertiesService.getScriptProperties();
      
      const usageKey = `${USAGE_PROPERTIES.DAILY_USAGE_PREFIX}${targetDate}`;
      const usage = JSON.parse(properties.getProperty(usageKey) || '{}');
      
      // 全タイプの使用量を確実に含める
      const result = {};
      Object.values(USAGE_TYPES).forEach(type => {
        result[type] = usage[type] || 0;
      });
      
      return result;
      
    } catch (error) {
      console.error('❌ 全日次使用量取得エラー:', error);
      return {};
    }
  }
  
  /**
   * 日次使用量をリセット
   * 毎日00:00に自動実行されるべき関数
   * @param {string} date 対象日付（未指定時は今日）
   */
  static resetDailyUsage(date = null) {
    try {
      const targetDate = date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      const properties = PropertiesService.getScriptProperties();
      
      const usageKey = `${USAGE_PROPERTIES.DAILY_USAGE_PREFIX}${targetDate}`;
      
      // リセット前の使用量を履歴に保存
      const oldUsage = JSON.parse(properties.getProperty(usageKey) || '{}');
      this._archiveUsageData(targetDate, oldUsage);
      
      // 使用量をリセット
      properties.deleteProperty(usageKey);
      
      // 最後のリセット日時を記録
      properties.setProperty(USAGE_PROPERTIES.LAST_RESET_DATE, targetDate);
      
      console.log(`✅ 日次使用量リセット完了: ${targetDate}`);
      
    } catch (error) {
      console.error('❌ 日次使用量リセットエラー:', error);
    }
  }
  
  /**
   * 使用量制限チェック
   * @param {string} type 使用量タイプ
   * @param {number} requestAmount 要求使用量（デフォルト1）
   * @returns {Object} チェック結果
   */
  static checkUsageLimit(type, requestAmount = 1) {
    try {
      // 現在のプラン制限を取得
      const planLimits = getPlanLimits();
      
      // 現在の使用量を取得
      const currentUsage = this.getDailyUsage(type);
      
      // 制限値を決定
      let dailyLimit = null;
      
      switch (type) {
        case USAGE_TYPES.COMPANY_SEARCH:
          dailyLimit = planLimits.maxCompaniesPerDay;
          break;
        case USAGE_TYPES.KEYWORD_GENERATION:
          if (!planLimits.keywordGeneration) {
            return {
              allowed: false,
              reason: 'FEATURE_NOT_AVAILABLE',
              currentUsage: currentUsage,
              limit: 0,
              remaining: 0
            };
          }
          dailyLimit = 100; // 仕様書では明確な制限なし、実用的な上限
          break;
        case USAGE_TYPES.AI_PROPOSAL:
          if (!planLimits.aiProposals) {
            return {
              allowed: false,
              reason: 'FEATURE_NOT_AVAILABLE',
              currentUsage: currentUsage,
              limit: 0,
              remaining: 0
            };
          }
          dailyLimit = 200; // 仕様書では明確な制限なし、実用的な上限
          break;
        default:
          dailyLimit = 1000; // その他のタイプのデフォルト制限
          break;
      }
      
      // 制限チェック
      const remaining = dailyLimit - currentUsage;
      const wouldExceed = (currentUsage + requestAmount) > dailyLimit;
      
      return {
        allowed: !wouldExceed,
        reason: wouldExceed ? 'DAILY_LIMIT_EXCEEDED' : 'OK',
        currentUsage: currentUsage,
        limit: dailyLimit,
        remaining: Math.max(0, remaining),
        requestAmount: requestAmount
      };
      
    } catch (error) {
      console.error('❌ 使用量制限チェックエラー:', error);
      return {
        allowed: false,
        reason: 'CHECK_ERROR',
        currentUsage: 0,
        limit: 0,
        remaining: 0
      };
    }
  }
  
  /**
   * 使用量統計を取得
   * @param {number} days 過去何日分（デフォルト7日）
   * @returns {Object} 使用量統計
   */
  static getUsageStatistics(days = 7) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const stats = {
        period: days,
        dailyData: [],
        totals: {},
        averages: {}
      };
      
      // 過去指定日数分のデータを収集
      for (let i = 0; i < days; i++) {
        const date = new Date();
        date.setDate(date.getDate() - i);
        const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
        
        const dailyUsage = this.getAllDailyUsage(dateStr);
        stats.dailyData.unshift({
          date: dateStr,
          usage: dailyUsage
        });
      }
      
      // 合計値を計算
      Object.values(USAGE_TYPES).forEach(type => {
        stats.totals[type] = stats.dailyData.reduce((sum, day) => sum + (day.usage[type] || 0), 0);
        stats.averages[type] = Math.round((stats.totals[type] / days) * 100) / 100;
      });
      
      return stats;
      
    } catch (error) {
      console.error('❌ 使用量統計取得エラー:', error);
      return { period: days, dailyData: [], totals: {}, averages: {} };
    }
  }
  
  /**
   * 使用量履歴を記録（内部関数）
   * @param {string} type 使用量タイプ
   * @param {number} amount 増加量
   * @param {number} newTotal 新しい合計値
   * @private
   */
  static _recordUsageHistory(type, amount, newTotal) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const history = JSON.parse(properties.getProperty(USAGE_PROPERTIES.USAGE_HISTORY) || '[]');
      
      const record = {
        timestamp: new Date().toISOString(),
        type: type,
        amount: amount,
        total: newTotal,
        user: Session.getActiveUser().getEmail()
      };
      
      history.push(record);
      
      // 履歴は最大1000件まで保持
      if (history.length > 1000) {
        history.splice(0, history.length - 1000);
      }
      
      properties.setProperty(USAGE_PROPERTIES.USAGE_HISTORY, JSON.stringify(history));
      
    } catch (error) {
      console.error('❌ 使用量履歴記録エラー:', error);
    }
  }
  
  /**
   * 使用量データをアーカイブ（内部関数）
   * @param {string} date 日付
   * @param {Object} usageData 使用量データ
   * @private
   */
  static _archiveUsageData(date, usageData) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const archives = JSON.parse(properties.getProperty('usage_archives') || '[]');
      
      const archive = {
        date: date,
        usage: usageData,
        archivedAt: new Date().toISOString()
      };
      
      archives.push(archive);
      
      // アーカイブは最大365件（1年分）まで保持
      if (archives.length > 365) {
        archives.splice(0, archives.length - 365);
      }
      
      properties.setProperty('usage_archives', JSON.stringify(archives));
      
    } catch (error) {
      console.error('❌ 使用量アーカイブエラー:', error);
    }
  }
}

/**
 * 自動リセット用トリガーの設定
 * 毎日午前0時に日次使用量をリセットするトリガーを設定
 */
function setupDailyUsageResetTrigger() {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'executeDailyUsageReset') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新しいトリガーを作成（毎日午前0時）
    ScriptApp.newTrigger('executeDailyUsageReset')
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();
    
    console.log('✅ 日次使用量リセットトリガー設定完了');
    
  } catch (error) {
    console.error('❌ 日次使用量リセットトリガー設定エラー:', error);
  }
}

/**
 * 日次使用量リセット実行関数（トリガーから呼ばれる）
 */
function executeDailyUsageReset() {
  try {
    console.log('🔄 日次使用量自動リセット開始...');
    UsageTracker.resetDailyUsage();
    console.log('✅ 日次使用量自動リセット完了');
  } catch (error) {
    console.error('❌ 日次使用量自動リセット実行エラー:', error);
  }
}

/**
 * 使用量追跡システム初期化
 */
function initializeUsageTracker() {
  try {
    console.log('🚀 使用量追跡システム初期化開始...');
    
    // プラン管理システムの依存関係チェック
    if (typeof getPlanLimits !== 'function') {
      console.warn('⚠️ プラン管理システムが見つかりません - 基本機能のみ利用可能');
    }
    
    // トリガー設定
    setupDailyUsageResetTrigger();
    
    // 初期状態確認
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const todayUsage = UsageTracker.getAllDailyUsage();
    
    console.log('📊 今日の使用量:', JSON.stringify(todayUsage, null, 2));
    console.log('✅ 使用量追跡システム初期化完了');
    
    return true;
    
  } catch (error) {
    console.error('❌ 使用量追跡システム初期化エラー:', error);
    return false;
  }
}

/**
 * 使用量ダッシュボード表示
 */
function showUsageDashboard() {
  try {
    const ui = SpreadsheetApp.getUi();
    const stats = UsageTracker.getUsageStatistics(7);
    const todayUsage = UsageTracker.getAllDailyUsage();
    const planLimits = getPlanLimits();
    
    let dashboard = '📊 使用量ダッシュボード\n\n';
    
    // 今日の使用量
    dashboard += '📅 今日の使用量:\n';
    dashboard += `・企業検索: ${todayUsage[USAGE_TYPES.COMPANY_SEARCH]}/${planLimits.maxCompaniesPerDay}\n`;
    dashboard += `・キーワード生成: ${todayUsage[USAGE_TYPES.KEYWORD_GENERATION]}\n`;
    dashboard += `・AI提案生成: ${todayUsage[USAGE_TYPES.AI_PROPOSAL]}\n\n`;
    
    // 7日間の統計
    dashboard += '📈 過去7日間の平均:\n';
    dashboard += `・企業検索: ${stats.averages[USAGE_TYPES.COMPANY_SEARCH]}回/日\n`;
    dashboard += `・キーワード生成: ${stats.averages[USAGE_TYPES.KEYWORD_GENERATION]}回/日\n`;
    dashboard += `・AI提案生成: ${stats.averages[USAGE_TYPES.AI_PROPOSAL]}回/日\n\n`;
    
    // プラン制限情報
    dashboard += '🎯 プラン制限:\n';
    dashboard += `・企業検索上限: ${planLimits.maxCompaniesPerDay}社/日\n`;
    dashboard += `・キーワード生成: ${planLimits.keywordGeneration ? '✅ 利用可能' : '❌ 利用不可'}\n`;
    dashboard += `・AI提案生成: ${planLimits.aiProposals ? '✅ 利用可能' : '❌ 利用不可'}\n`;
    
    ui.alert('使用量ダッシュボード', dashboard, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ 使用量ダッシュボード表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '使用量ダッシュボードの表示でエラーが発生しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}