/**
 * 仕様書準拠チェックシステム
 * システム仕様書v2.0からの逸脱を防ぐためのガバナンス機能
 */

// 仕様書v2.0で定義された正式なメニュー構造
const SPEC_V2_MENU_STRUCTURE = {
  title: "🚀 営業システム",
  categories: {
    systemManagement: {
      title: "📊 システム管理",
      functions: [
        { name: "🧪 システムテスト", function: "runSystemTest" },
        { name: "🔑 APIキーテスト", function: "testApiKeys" },
        { name: "📋 基本情報", function: "showBasicInfo" },
        { name: "🔧 シート作成", function: "createBasicSheets" }
      ]
    },
    keywordManagement: {
      title: "🔍 キーワード管理",
      functions: [
        { name: "🎯 キーワード生成", function: "generateKeywords" },
        { name: "📊 使用状況分析", function: "analyzeKeywordUsage" }
      ]
    },
    companyManagement: {
      title: "🏢 企業管理", 
      functions: [
        { name: "🔍 企業検索", function: "searchCompany" },
        { name: "📊 企業分析", function: "analyzeCompany" },
        { name: "📈 マッチ度計算", function: "calculateMatching" }
      ]
    },
    proposalManagement: {
      title: "💼 提案管理",
      functions: [
        { name: "✨ 提案生成", function: "generateProposal" },
        { name: "📊 提案分析", function: "analyzeProposal" }
      ]
    },
    analyticsReports: {
      title: "📊 分析・レポート",
      functions: [
        { name: "📊 総合レポート", function: "generateComprehensiveReport" },
        { name: "📋 活動ログ", function: "viewActivityLog" }
      ]
    },
    helpDocuments: {
      title: "📚 ヘルプ・ドキュメント",
      functions: [
        { name: "🆘 基本ヘルプ", function: "showHelp" },
        { name: "📖 ユーザーガイド", function: "showUserGuide" },
        { name: "💰 料金・API設定ガイド", function: "showPricingGuide" },
        { name: "❓ よくある質問", function: "showFAQ" }
      ]
    },
    settings: {
      title: "⚙️ 設定",
      functions: [
        { name: "🔑 APIキー設定", function: "configureApiKeys" },
        { name: "📊 基本設定", function: "showBasicSettings" },
        { name: "🔧 詳細設定", function: "showAdvancedSettings" },
        { name: "🌐 システム環境", function: "showSystemEnvironment" }
      ]
    }
  },
  // 管理者専用追加項目
  adminOnlyCategories: {
    licenseManagement: {
      title: "🔐 ライセンス管理（管理者専用）",
      functions: [
        { name: "📋 ライセンス状況", function: "showLicenseStatus" },
        { name: "🔐 管理者認証", function: "authenticateAdmin" },
        { name: "🔑 APIキー管理", function: "manageApiKeys" },
        { name: "📅 使用開始日設定", function: "setLicenseStartDate" },
        { name: "🔄 期限延長", function: "extendLicense" },
        { name: "🔒 システムロック解除", function: "unlockSystem" },
        { name: "💳 課金状況管理", function: "manageBilling" }
      ]
    },
    userManagement: {
      title: "👥 ユーザー管理（管理者専用）",
      functions: [
        { name: "👤 管理者モード切替", function: "toggleAdminMode" },
        { name: "📊 利用統計", function: "showUsageStatistics" },
        { name: "🔧 システム設定", function: "systemConfiguration" }
      ]
    },
    advancedSettings: {
      title: "🔧 高度な設定（管理者専用）",
      functions: [
        { name: "📊 詳細分析設定", function: "advancedAnalytics" },
        { name: "🔄 データ同期", function: "syncAllData" },
        { name: "🛠️ システムメンテナンス", function: "systemMaintenance" }
      ]
    }
  }
};

/**
 * 仕様書準拠チェック機能（修正版）
 */
function validateMenuComplianceWithSpec() {
  try {
    console.log('🔍 仕様書v2.0準拠チェック開始...');
    
    const violations = [];
    
    // 1. 必須機能の存在チェック（修正版）
    const requiredFunctions = [
      // システム管理
      'runSystemTest', 'testApiKeys', 'showBasicInfo', 'createBasicSheets',
      // キーワード管理  
      'generateKeywords', 'analyzeKeywordUsage',
      // 企業管理
      'searchCompany', 'analyzeCompany', 'calculateMatching',
      // 提案管理
      'generateProposal', 'analyzeProposal',
      // 分析・レポート
      'generateComprehensiveReport', 'viewActivityLog',
      // ヘルプ・ドキュメント
      'showHelp', 'showUserGuide', 'showPricingGuide', 'showFAQ',
      // 設定
      'configureApiKeys', 'showBasicSettings', 'showAdvancedSettings', 'showSystemEnvironment'
    ];
    
    // 2. 関数実装チェック（スコープ問題を修正）
    requiredFunctions.forEach(funcName => {
      let isImplemented = false;
      
      try {
        // 複数の方法で関数の存在を確認
        isImplemented = (
          typeof window !== 'undefined' && typeof window[funcName] === 'function'
        ) || (
          typeof global !== 'undefined' && typeof global[funcName] === 'function'
        ) || (
          typeof this[funcName] === 'function'
        ) || (
          eval(`typeof ${funcName}`) === 'function'
        );
      } catch (error) {
        // eval失敗した場合は未実装とみなす
        isImplemented = false;
      }
      
      if (!isImplemented) {
        violations.push({
          type: 'MISSING_FUNCTION',
          function: funcName,
          category: getFunctionCategory(funcName),
          severity: 'HIGH'
        });
      }
    });
    
    // 3. 管理者機能チェック
    const adminFunctions = [
      'showLicenseStatus', 'authenticateAdmin', 'manageApiKeys', 
      'setLicenseStartDate', 'extendLicense', 'unlockSystem', 'manageBilling',
      'toggleAdminMode', 'showUsageStatistics', 'systemConfiguration',
      'advancedAnalytics', 'syncAllData', 'systemMaintenance'
    ];
    
    adminFunctions.forEach(funcName => {
      let isImplemented = false;
      
      try {
        isImplemented = (
          typeof window !== 'undefined' && typeof window[funcName] === 'function'
        ) || (
          typeof global !== 'undefined' && typeof global[funcName] === 'function'
        ) || (
          typeof this[funcName] === 'function'
        ) || (
          eval(`typeof ${funcName}`) === 'function'
        );
      } catch (error) {
        isImplemented = false;
      }
      
      if (!isImplemented) {
        violations.push({
          type: 'MISSING_ADMIN_FUNCTION',
          function: funcName,
          category: 'AdminFunctions',
          severity: 'MEDIUM'
        });
      }
    });
    
    return {
      compliant: violations.length === 0,
      violations: violations,
      checkTime: new Date().toISOString(),
      totalViolations: violations.length
    };
    
  } catch (error) {
    console.error('❌ 仕様書準拠チェックエラー:', error);
    return {
      compliant: false,
      violations: [{
        type: 'CHECK_ERROR',
        function: 'validateMenuComplianceWithSpec',
        category: 'System',
        severity: 'CRITICAL'
      }],
      checkTime: new Date().toISOString(),
      totalViolations: 1
    };
  }
}

/**
 * 関数のカテゴリを取得
 */
function getFunctionCategory(funcName) {
  const categoryMap = {
    // システム管理
    'runSystemTest': 'システム管理',
    'testApiKeys': 'システム管理', 
    'showBasicInfo': 'システム管理',
    'createBasicSheets': 'システム管理',
    // キーワード管理
    'generateKeywords': 'キーワード管理',
    'analyzeKeywordUsage': 'キーワード管理',
    // 企業管理
    'searchCompany': '企業管理',
    'analyzeCompany': '企業管理',
    'calculateMatching': '企業管理',
    // 提案管理
    'generateProposal': '提案管理',
    'analyzeProposal': '提案管理',
    // 分析・レポート
    'generateComprehensiveReport': '分析・レポート',
    'viewActivityLog': '分析・レポート',
    // ヘルプ・ドキュメント
    'showHelp': 'ヘルプ・ドキュメント',
    'showUserGuide': 'ヘルプ・ドキュメント',
    'showPricingGuide': 'ヘルプ・ドキュメント',
    'showFAQ': 'ヘルプ・ドキュメント',
    // 設定
    'configureApiKeys': '設定',
    'showBasicSettings': '設定',
    'showAdvancedSettings': '設定',
    'showSystemEnvironment': '設定'
  };
  
  return categoryMap[funcName] || '不明';
}

/**
 * 仕様書準拠レポート生成
 */
function generateComplianceReport() {
  try {
    const result = validateMenuComplianceWithSpec();
    
    let report = '📋 仕様書v2.0準拠チェックレポート\n\n';
    report += `チェック日時: ${result.checkTime}\n`;
    report += `準拠状況: ${result.compliant ? '✅ 準拠' : '❌ 違反あり'}\n\n`;
    
    if (result.violations.length > 0) {
      report += '⚠️ 検出された違反:\n\n';
      result.violations.forEach((violation, index) => {
        report += `${index + 1}. ${violation.type}\n`;
        report += `   カテゴリ: ${violation.category}\n`;
        if (violation.function) {
          report += `   機能: ${violation.function}\n`;
        }
        report += `   重要度: ${violation.severity}\n\n`;
      });
    }
    
    report += '\n💡 修正アクション:\n';
    report += '1. 仕様書v2.0に準拠したメニューの再実装\n';
    report += '2. 欠落機能の実装\n';
    report += '3. 不正な追加機能の削除\n';
    
    SpreadsheetApp.getUi().alert(
      '仕様書準拠チェック',
      report,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return result;
    
  } catch (error) {
    console.error('❌ 準拠レポート生成エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '仕様書準拠チェックでエラーが発生しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 現在のメニュー構造を取得（分析用）
 */
function getCurrentMenuStructure() {
  // 実際の実装では現在のメニュー構造を解析
  // プレースホルダー実装
  return {
    categories: {},
    violations: []
  };
}

/**
 * 指定された関数が実装されているかチェック
 */
function isFunctionImplemented(functionName) {
  try {
    // グローバルスコープで関数の存在を確認
    return typeof this[functionName] === 'function';
  } catch (error) {
    return false;
  }
}

/**
 * 仕様書駆動開発ガバナンス
 */
const SPEC_DRIVEN_GOVERNANCE = {
  
  // ルール1: メニュー変更前の必須チェック
  beforeMenuChange: function() {
    console.log('🔒 仕様書準拠チェック実行中...');
    return validateMenuComplianceWithSpec();
  },
  
  // ルール2: 新機能追加時の仕様書確認
  beforeAddingNewFunction: function(functionName, category) {
    console.log(`🔍 新機能追加チェック: ${functionName} (${category})`);
    
    // 仕様書に定義されているかチェック
    let isSpecDefined = false;
    
    Object.values(SPEC_V2_MENU_STRUCTURE.categories).forEach(cat => {
      cat.functions.forEach(func => {
        if (func.function === functionName) {
          isSpecDefined = true;
        }
      });
    });
    
    if (!isSpecDefined) {
      const message = `⚠️ 警告: 機能「${functionName}」は仕様書v2.0に定義されていません。\n\n追加する前に仕様書の更新が必要です。`;
      SpreadsheetApp.getUi().alert('仕様書違反警告', message, SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    
    return true;
  },
  
  // ルール3: 定期的な準拠性監査
  scheduleComplianceAudit: function() {
    console.log('📅 定期準拠性監査をスケジュール');
    // 実装: 定期実行トリガーの設定
  }
};

/**
 * 仕様書準拠メニュー強制実行
 */
function enforceSpecCompliantMenu() {
  try {
    console.log('🔧 仕様書準拠メニュー強制適用開始...');
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '仕様書準拠メニュー適用',
      '現在のメニューを仕様書v2.0準拠の構造に強制的に戻します。\n\n現在のカスタマイズは失われますが、標準仕様に準拠します。\n\n実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // 仕様書準拠メニューの強制実装
      createSpecCompliantMenu();
      
      ui.alert(
        '✅ 完了',
        '仕様書v2.0準拠メニューに復元されました。\n\nメニューを確認してください。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('❌ 仕様書準拠メニュー適用エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '仕様書準拠メニューの適用でエラーが発生しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 仕様書v2.0完全準拠メニューの作成
 */
function createSpecCompliantMenu() {
  try {
    console.log('🏗️ 仕様書v2.0準拠メニュー作成開始...');
    
    const ui = SpreadsheetApp.getUi();
    
    // ユーザー権限確認
    let isAdmin = false;
    try {
      const currentUser = getCurrentUser();
      isAdmin = currentUser.role === 'Administrator';
    } catch (error) {
      console.log('ユーザー管理未初期化 - 一般ユーザーとして処理');
    }
    
    // メインメニュー作成（仕様書準拠）
    const mainMenu = ui.createMenu(SPEC_V2_MENU_STRUCTURE.title);
    
    // 基本カテゴリ追加（仕様書通り）
    Object.values(SPEC_V2_MENU_STRUCTURE.categories).forEach(category => {
      const subMenu = ui.createMenu(category.title);
      
      category.functions.forEach(func => {
        subMenu.addItem(func.name, func.function);
      });
      
      mainMenu.addSubMenu(subMenu);
    });
    
    // 管理者専用機能（管理者の場合のみ）
    if (isAdmin) {
      mainMenu.addSeparator();
      
      Object.values(SPEC_V2_MENU_STRUCTURE.adminOnlyCategories).forEach(category => {
        const subMenu = ui.createMenu(category.title);
        
        category.functions.forEach(func => {
          subMenu.addItem(func.name, func.function);
        });
        
        mainMenu.addSubMenu(subMenu);
      });
    }
    
    // 仕様書準拠チェック機能を追加
    mainMenu.addSeparator();
    mainMenu.addItem('🔍 仕様書準拠チェック', 'generateComplianceReport');
    
    // メニュー有効化
    mainMenu.addToUi();
    
    // 通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '営業システム v2.0 - 仕様書準拠版', 
      '🏗️ 標準メニュー適用完了', 
      5
    );
    
    console.log('✅ 仕様書v2.0準拠メニュー作成完了');
    
  } catch (error) {
    console.error('❌ 仕様書準拠メニュー作成エラー:', error);
    throw error;
  }
}

/**
 * 開発ガバナンス: 変更前チェック
 */
function checkBeforeMenuModification() {
  console.log('🛡️ 変更前ガバナンスチェック');
  
  const compliance = validateMenuComplianceWithSpec();
  
  if (!compliance.compliant) {
    const message = `⚠️ 現在のメニューは既に仕様書から逸脱しています。\n\n違反数: ${compliance.violations.length}\n\n先に仕様書準拠メニューに戻すことを推奨します。`;
    
    SpreadsheetApp.getUi().alert('ガバナンス警告', message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  return compliance.compliant;
}