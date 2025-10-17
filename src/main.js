/**
 * 営業自動化システム - メインファイル
 * 商材起点企業発掘・提案自動生成システム
 */

/**
 * システム初期化状況の確認（スプレッドシートID指定版）
 */
function checkSystemInitializationById(spreadsheetId = null) {
  try {
    let spreadsheet;
    if (spreadsheetId) {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } else {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    if (!spreadsheet) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length > 0) {
      console.log('不足しているシート:', missingSheets);
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    console.log('✅ システム初期化済み - すべてのシートが利用可能です');
    return true;
  } catch (error) {
    console.error('❌ システム初期化チェックエラー:', error.toString());
    throw error;
  }
}

/**
 * システム初期化状況の確認
 */
function checkSystemInitialization() {
  try {
    const spreadsheet = getSafeSpreadsheet();
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length > 0) {
      console.log('不足しているシート:', missingSheets);
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    console.log('✅ システム初期化済み - すべてのシートが利用可能です');
    return true;
  } catch (error) {
    console.error('❌ システム初期化チェックエラー:', error.toString());
    throw error;
  }
}

/**
 * スプレッドシートの初期化
 */
function initializeSheets() {
  const ss = getSafeSpreadsheet();
  
  // 制御パネルシートの作成
  createControlPanel(ss);
  
  // プラン説明シートの作成
  createPlanInfoSheet(ss);
  
  // 各データシートの作成
  createKeywordsSheet(ss);
  createCompaniesSheet(ss);
  createProposalsSheet(ss);
  createLogsSheet(ss);
  
  Logger.log('システムの初期化が完了しました');
}

/**
 * 制御パネルシートの作成
 */
function createControlPanel(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONTROL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONTROL);
  }
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('営業自動化システム').setFontSize(16).setFontWeight('bold');
  
  // 入力項目設定
  const inputData = [
    ['商材名', ''],
    ['商材概要', ''],
    ['価格帯', '低価格'],
    ['対象企業規模', '中小企業'],
    ['優先地域', ''],
    ['', ''],
    ['実行ボタン', ''],
    ['', ''],
    ['', ''],
    ['検索企業数上限', 20],
    ['APIキー設定', '設定済み'],
    ['', ''],
    ['', ''],
    ['実行状況表示', '']
  ];
  
  sheet.getRange(2, 1, inputData.length, 2).setValues(inputData);
  
  // データ検証の設定
  const priceValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['低価格', '中価格', '高価格'])
    .build();
  sheet.getRange('B4').setDataValidation(priceValidation);
  
  const sizeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['個人事業主', '中小企業', '大企業', 'すべて'])
    .build();
  sheet.getRange('B5').setDataValidation(sizeValidation);
  
  // 実行ボタンの設定
  createExecutionButtons(sheet);
  
  // ダッシュボード作成
  createDashboard(sheet);
}

/**
 * 実行ボタンの作成
 */
function createExecutionButtons(sheet) {
  // キーワード生成ボタン
  sheet.getRange('A8').setValue('キーワード生成');
  sheet.getRange('B8').setValue('企業検索');
  sheet.getRange('C8').setValue('全自動実行');
  
  // ボタンの背景色設定
  sheet.getRange('A8:C8').setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
}

/**
 * ダッシュボード作成
 */
function createDashboard(sheet) {
  // ダッシュボードエリア
  sheet.getRange('E1').setValue('システムダッシュボード').setFontSize(14).setFontWeight('bold');
  
  const dashboardData = [
    ['登録済み企業数', '=COUNTA(企業マスター!A:A)-1'],
    ['生成済みキーワード数', '=COUNTA(キーワード一覧!A:A)-1'],
    ['提案メッセージ数', '=COUNTA(提案メッセージ!A:A)-1'],
    ['システム稼働日数', '=TODAY()-DATE(2024,1,1)']
  ];
  
  sheet.getRange('E2', 1, dashboardData.length, 2).setValues(dashboardData);
}

/**
 * キーワードシートの作成
 */
function createKeywordsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.KEYWORDS);
  }
  
  const headers = ['キーワード', 'カテゴリ', '優先度', '実行状況', '発見企業数', '実行日時', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e1f5fe');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 200);
}

/**
 * 企業シートの作成
 */
function createCompaniesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COMPANIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.COMPANIES);
  }
  
  const headers = [
    '企業名', 'ウェブサイト', '業界', '企業規模', '所在地', '説明',
    'マッチ度スコア', '発見キーワード', '最終更新日', '連絡状況', '備考'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e8');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
}

/**
 * 提案メッセージシートの作成
 */
function createProposalsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.PROPOSALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PROPOSALS);
  }
  
  const headers = [
    '企業名', 'メール件名', '提案メッセージ', '生成日時', '送信状況', '返信', '備考'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fff3e0');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
}

/**
 * ログシートの作成
 */
function createLogsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.LOGS);
  }
  
  const headers = ['日時', '処理', '結果', '詳細', 'エラー'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3e5f5');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 300);
}

/**
 * コントロールパネルの設定値を取得
 */
function getControlPanelSettings() {
  try {
    const sheet = getSafeSpreadsheet().getSheetByName(SHEET_NAMES.CONTROL);
    if (!sheet) {
      throw new Error('制御パネルシートが見つかりません');
    }
    
    const data = sheet.getRange('A2:B15').getValues();
    
    const settings = {
      productName: data[0][1] || '',
      productDescription: data[1][1] || '',
      priceRange: data[2][1] || '低価格',
      targetSize: data[3][1] || '中小企業',
      targetRegion: data[4][1] || '',
      searchResultsPerKeyword: data[9][1] || 20
    };
    
    console.log('制御パネル設定を取得しました:', settings);
    return settings;
  } catch (error) {
    console.error('制御パネル設定取得エラー:', error);
    throw error;
  }
}

/**
 * 結果をシートに保存
 */
function saveToSheet(sheetName, data, headers = null) {
  try {
    const sheet = getSafeSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }
    
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    if (data && data.length > 0) {
      const startRow = headers ? 2 : sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    }
    
    console.log(`${sheetName}に${data.length}件のデータを保存しました`);
  } catch (error) {
    console.error(`${sheetName}への保存エラー:`, error);
    throw error;
  }
}

/**
 * システムの健全性チェック
 */
function healthCheck() {
  const results = {
    timestamp: new Date(),
    spreadsheet: false,
    sheets: false,
    apis: false,
    errors: []
  };
  
  try {
    // スプレッドシートアクセス
    const ss = getSafeSpreadsheet();
    results.spreadsheet = true;
    
    // 必要シートの存在確認
    const requiredSheets = Object.values(SHEET_NAMES);
    const existingSheets = ss.getSheets().map(sheet => sheet.getName());
    const missingSheets = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missingSheets.length === 0) {
      results.sheets = true;
    } else {
      results.errors.push(`不足シート: ${missingSheets.join(', ')}`);
    }
    
    // API設定確認
    const properties = PropertiesService.getScriptProperties();
    const hasGoogleApi = properties.getProperty('GOOGLE_SEARCH_API_KEY') && 
                        properties.getProperty('GOOGLE_SEARCH_ENGINE_ID');
    const hasOpenAI = properties.getProperty('OPENAI_API_KEY');
    
    if (hasGoogleApi && hasOpenAI) {
      results.apis = true;
    } else {
      const missingApis = [];
      if (!hasGoogleApi) missingApis.push('Google Search API');
      if (!hasOpenAI) missingApis.push('OpenAI API');
      results.errors.push(`未設定API: ${missingApis.join(', ')}`);
    }
    
  } catch (error) {
    results.errors.push(error.toString());
  }
  
  return results;
}

/**
 * システム状況をユーザーに表示
 */
function showSystemStatus() {
  try {
    const health = healthCheck();
    
    if (health.spreadsheet && health.sheets && health.apis) {
      SpreadsheetApp.getUi().alert(
        '✅ システム正常',
        'システムは正常に動作しています。すべての機能をご利用いただけます。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      const message = `❌ システムに問題があります。\n\n解決方法:\n1. メニューバーから「拡張機能」→「Apps Script」を選択\n2. 関数一覧から「initializeSheets」を選択\n3. 実行ボタンをクリックしてシステムを初期化`;
      
      SpreadsheetApp.getUi().alert(
        'システム状況',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    console.log('システム状況:', health);
  } catch (error) {
    console.error('システム状況表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ エラー',
      `システム状況の確認中にエラーが発生しました:\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * カスタムメニューの作成
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // システム状況を確認
    const health = healthCheck();
    
    if (!health.spreadsheet || !health.sheets) {
      // システム未初期化の場合は初期化メニューのみ表示
      ui.createMenu('🔧 営業自動化システム')
        .addItem('システム初期化', 'initializeSheets')
        .addItem('システム状況確認', 'showSystemStatus')
        .addToUi();
      return;
    }
    
    // 通常メニュー
    ui.createMenu('🎯 営業自動化システム')
      .addSubMenu(ui.createMenu('🔍 企業発掘')
        .addItem('キーワード生成', 'generateKeywords')
        .addItem('企業検索', 'searchCompanies')
        .addItem('全自動実行', 'executeFullWorkflow'))
      .addSubMenu(ui.createMenu('📝 提案生成')
        .addItem('AI提案生成', 'generateProposals')
        .addItem('提案一括生成', 'batchGenerateProposals'))
      .addSubMenu(ui.createMenu('⚙️ 設定')
        .addItem('API キー設定', 'showApiKeySettings')
        .addItem('システム設定', 'showSystemSettings'))
      .addSubMenu(ui.createMenu('ℹ️ ヘルプ')
        .addItem('使用方法', 'showUserGuide')
        .addItem('システム状況', 'showSystemStatus')
        .addItem('システム初期化', 'initializeSheets'))
      .addToUi();
      
  } catch (error) {
    console.error('メニュー作成エラー:', error);
    // エラー時は基本メニューのみ表示
    SpreadsheetApp.getUi()
      .createMenu('🔧 営業自動化システム (エラー)')
      .addItem('システム初期化', 'initializeSheets')
      .addToUi();
  }
}

/**
 * API キー設定ダイアログ
 */
function showApiKeySettings() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('api-key-settings')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'API キー設定');
}

/**
 * APIキーをスクリプトプロパティに保存
 */
function saveApiKeysToProperties(googleSearchKey, searchEngineId, openaiKey) {
  const properties = PropertiesService.getScriptProperties();
  
  if (googleSearchKey) {
    properties.setProperty('GOOGLE_SEARCH_API_KEY', googleSearchKey);
  }
  if (searchEngineId) {
    properties.setProperty('GOOGLE_SEARCH_ENGINE_ID', searchEngineId);
  }
  if (openaiKey) {
    properties.setProperty('OPENAI_API_KEY', openaiKey);
  }
  
  Logger.log('API キーがプロパティに保存されました');
}

/**
 * プラン説明シートの作成
 */
function createPlanInfoSheet(ss) {
  let sheet = ss.getSheetByName('プラン説明');
  if (!sheet) {
    sheet = ss.insertSheet('プラン説明', 1); // 2番目の位置に挿入
  } else {
    sheet.clear(); // 既存の場合はクリア
  }
  
  // タイトル設定
  sheet.getRange('A1').setValue('🎯 営業自動化システム - 課金プラン説明')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white');
  sheet.getRange('A1:G1').merge();
  
  // サブタイトル
  sheet.getRange('A2').setValue('各プランの機能比較表')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.getRange('A2:G2').merge();
  
  // ヘッダー行
  const headers = ['項目', 'ベーシック\n¥500/月', 'スタンダード\n¥1,500/月', 'プロフェッショナル\n¥5,500/月', 'エンタープライズ\n¥17,500/月'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // プラン比較データ
  const planData = [
    ['💰 月額料金', '¥500', '¥1,500', '¥5,500', '¥17,500'],
    ['🔢 アカウント数', '1アカウント', '1アカウント', '2アカウント', '5アカウント'],
    ['', '', '', '', ''],
    ['🔍 基本機能', '', '', '', ''],
    ['手動企業入力', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['企業検索上限', '10社/日', '50社/日', '100社/日', '500社/日'],
    ['基本レポート', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['', '', '', '', ''],
    ['🤖 AI機能', '', '', '', ''],
    ['キーワード自動生成', '❌ 利用不可', '✅ 利用可能\n(20個/回)', '✅ 利用可能\n(40個/回)', '✅ 利用可能\n(100個/回)'],
    ['AI提案文生成', '❌ 利用不可', '✅ 利用可能\n(50件/日)', '✅ 利用可能\n(100件/日)', '✅ 利用可能\n(500件/日)'],
    ['', '', '', '', ''],
    ['📊 高度な機能', '', '', '', ''],
    ['バッチ処理', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['カスタム分析', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['データエクスポート', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['高速処理', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['マルチアカウント', '❌ 利用不可', '❌ 利用不可', '✅ 2アカウント', '✅ 5アカウント'],
    ['カスタム連携', '❌ 利用不可', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能'],
    ['', '', '', '', ''],
    ['🛠️ サポート', '', '', '', ''],
    ['メールサポート', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['優先サポート', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能', '✅ 利用可能'],
    ['電話サポート', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能', '✅ 利用可能'],
    ['専任サポート', '❌ 利用不可', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能']
  ];
  
  sheet.getRange(5, 1, planData.length, planData[0].length).setValues(planData);
  
  // セルスタイル設定
  const dataRange = sheet.getRange(5, 1, planData.length, planData[0].length);
  
  // 項目列のスタイル
  sheet.getRange(5, 1, planData.length, 1)
    .setBackground('#f8f9fa')
    .setFontWeight('bold');
  
  // プラン列のスタイル
  for (let col = 2; col <= 5; col++) {
    const columnRange = sheet.getRange(5, col, planData.length, 1);
    columnRange.setHorizontalAlignment('center');
    
    // プラン別の背景色
    switch (col) {
      case 2: // ベーシック
        columnRange.setBackground('#fff3e0');
        break;
      case 3: // スタンダード
        columnRange.setBackground('#e8f5e8');
        break;
      case 4: // プロフェッショナル
        columnRange.setBackground('#e3f2fd');
        break;
      case 5: // エンタープライズ
        columnRange.setBackground('#fce4ec');
        break;
    }
  }
  
  // ボーダー設定
  dataRange.setBorder(true, true, true, true, true, true);
  
  // 特別な行のスタイル（カテゴリ見出し行）
  const categoryRows = [7, 10, 14, 22]; // 「🔍 基本機能」「🤖 AI機能」「📊 高度な機能」「🛠️ サポート」
  categoryRows.forEach(row => {
    sheet.getRange(row + 4, 1, 1, 5) // +4はヘッダー分のオフセット
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
  });
  
  // 現在のプラン表示エリア
  sheet.getRange('A' + (5 + planData.length + 2)).setValue('📋 現在のプラン情報')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white');
  sheet.getRange('A' + (5 + planData.length + 2) + ':C' + (5 + planData.length + 2)).merge();
  
  const currentPlanRow = 5 + planData.length + 3;
  sheet.getRange('A' + currentPlanRow).setValue('現在のプラン:');
  
  // getUserPlan関数があるかチェックして安全に呼び出し
  let currentPlan = 'ベーシック';
  try {
    if (typeof getUserPlan === 'function') {
      currentPlan = getUserPlan() || 'ベーシック';
    }
  } catch (error) {
    console.log('getUserPlan関数が見つかりません。デフォルト値を使用します。');
  }
  
  sheet.getRange('B' + currentPlanRow).setValue(currentPlan);
  
  sheet.getRange('A' + (currentPlanRow + 1)).setValue('月額料金:');
  sheet.getRange('B' + (currentPlanRow + 1)).setValue('=IF(B' + currentPlanRow + '="ベーシック","¥500",IF(B' + currentPlanRow + '="スタンダード","¥1,500",IF(B' + currentPlanRow + '="プロフェッショナル","¥5,500","¥17,500")))');
  
  sheet.getRange('A' + (currentPlanRow + 2)).setValue('今日の使用量:');
  sheet.getRange('B' + (currentPlanRow + 2)).setValue('企業検索: 0回');
  
  // プランアップグレード案内
  sheet.getRange('A' + (currentPlanRow + 4)).setValue('💡 プランアップグレードについて')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#ff9800')
    .setFontColor('white');
  sheet.getRange('A' + (currentPlanRow + 4) + ':C' + (currentPlanRow + 4)).merge();
  
  const upgradeInfo = [
    ['• AI機能をご利用になりたい場合は「スタンダード」以上をお選びください'],
    ['• より多くの企業検索をお求めの場合は上位プランをご検討ください'],
    ['• 高速処理・優先サポートは「プロフェッショナル」以上で利用可能です'],
    ['• 企業向け機能・専任サポートは「エンタープライズ」プランをご利用ください']
  ];
  
  sheet.getRange(currentPlanRow + 5, 1, upgradeInfo.length, 1).setValues(upgradeInfo);
  
  // API料金情報セクション
  const apiCostStartRow = currentPlanRow + 5 + upgradeInfo.length + 2;
  
  sheet.getRange('A' + apiCostStartRow).setValue('💳 API利用料金について（別途課金）')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#d32f2f')
    .setFontColor('white');
  sheet.getRange('A' + apiCostStartRow + ':E' + apiCostStartRow).merge();
  
  const apiCostInfo = [
    ['⚠️ 重要: 以下のAPI利用料金が別途発生します', '', '', '', ''],
    ['', '', '', '', ''],
    ['API名', '料金体系', '無料枠', '超過時料金', '備考'],
    ['OpenAI ChatGPT API', 'トークン従量課金', '$5.00相当/月', '$0.002/1Kトークン', 'AI機能利用時のみ'],
    ['Google Custom Search API', 'リクエスト従量課金', '100回/日', '$5.00/1000回', '企業検索時のみ'],
    ['', '', '', '', ''],
    ['📊 プラン別API料金目安（最大想定）', '', '', '', ''],
    ['プラン', 'システム利用料', 'API料金（最大）', '合計月額', 'アカウント数'],
    ['ベーシック', '¥500', '¥0', '¥500', '1アカウント'],
    ['スタンダード', '¥1,500', '¥2,000', '¥3,500', '1アカウント'],
    ['プロフェッショナル', '¥5,500', '¥6,000', '¥11,500', '2アカウント'],
    ['エンタープライズ', '¥17,500', '¥30,000', '¥47,500', '5アカウント'],
    ['', '', '', '', ''],
    ['💡 API料金節約のコツ', '', '', '', ''],
    ['• ベーシックプランではAI機能を使わないため、OpenAI料金は¥0です', '', '', '', ''],
    ['• 企業検索は1日100回まで無料なので、Google検索料金も抑えられます', '', '', '', ''],
    ['• マルチアカウントプランでは使用量を分散して効率的に利用できます', '', '', '', ''],
    ['• API料金はGoogle・OpenAIに直接支払い（当社には支払われません）', '', '', '', ''],
    ['• 実際の料金は使用量により変動し、上記は最大想定値です', '', '', '', '']
  ];
  
  sheet.getRange(apiCostStartRow + 1, 1, apiCostInfo.length, 5).setValues(apiCostInfo);
  
  // API料金表のスタイル設定
  // ヘッダー行のスタイル
  sheet.getRange(apiCostStartRow + 3, 1, 1, 5)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(apiCostStartRow + 7, 1, 1, 5)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // API料金データ行のスタイル
  sheet.getRange(apiCostStartRow + 4, 1, 2, 5)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(apiCostStartRow + 8, 1, 4, 4)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  // 注意事項のスタイル
  sheet.getRange(apiCostStartRow + 14, 1, 1, 5)
    .setBackground('#ff9800')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // 節約コツ部分のスタイル
  sheet.getRange(apiCostStartRow + 15, 1, 4, 1)
    .setBackground('#e8f5e8')
    .setFontSize(10);
  
  // 料金に関する重要な注意喚起
  const importantNoticeRow = apiCostStartRow + apiCostInfo.length + 2;
  
  sheet.getRange('A' + importantNoticeRow).setValue('🚨 API料金に関する重要な注意事項')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#d32f2f')
    .setFontColor('white');
  sheet.getRange('A' + importantNoticeRow + ':E' + importantNoticeRow).merge();
  
  const importantNotices = [
    ['1. API料金は当社プラン料金とは別に、各APIプロバイダーから直接請求されます'],
    ['2. API料金は使用量に応じて変動するため、事前に予算をご確認ください'],
    ['3. 無料枠を超過した場合、自動的に有料課金が開始されます'],
    ['4. API料金の詳細は各プロバイダーの公式サイトでご確認ください'],
    ['   • OpenAI API: https://openai.com/pricing'],
    ['   • Google Custom Search API: https://developers.google.com/custom-search/v1/overview']
  ];
  
  sheet.getRange(importantNoticeRow + 1, 1, importantNotices.length, 1).setValues(importantNotices);
  sheet.getRange(importantNoticeRow + 1, 1, importantNotices.length, 1)
    .setBackground('#ffebee')
    .setFontSize(10)
    .setWrap(true);
  
  // 列幅調整（API料金表対応）
  sheet.setColumnWidth(1, 250); // 項目列
  sheet.setColumnWidth(2, 180); // ベーシック
  sheet.setColumnWidth(3, 120); // スタンダード
  sheet.setColumnWidth(4, 150); // プロフェッショナル
  sheet.setColumnWidth(5, 150); // エンタープライズ
  
  // 行の高さ調整
  sheet.setRowHeight(1, 40); // タイトル
  sheet.setRowHeight(4, 40); // ヘッダー
  
  Logger.log('プラン説明シートが作成されました');
}
