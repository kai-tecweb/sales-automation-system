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
    
    const requiredSheets = Obje  if (openaiKey) {
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
  const headers = ['項目', 'ベーシック\n¥500/月', 'スタンダード\n¥1,500/月', 'プロフェッショナル\n¥3,000/月', 'エンタープライズ\n¥7,500/月'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // プラン比較データ
  const planData = [
    ['💰 月額料金', '¥500', '¥1,500', '¥3,000', '¥7,500'],
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
    ['マルチユーザー', '❌ 利用不可', '❌ 利用不可', '❌ 利用不可', '✅ 利用可能'],
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
  sheet.getRange('B' + currentPlanRow).setValue('=IF(ISBLANK(B' + (currentPlanRow + 5) + '),"未設定","' + (getUserPlan() || 'ベーシック') + '")');
  
  sheet.getRange('A' + (currentPlanRow + 1)).setValue('月額料金:');
  sheet.getRange('B' + (currentPlanRow + 1)).setValue('=IF(B' + currentPlanRow + '="ベーシック","¥500",IF(B' + currentPlanRow + '="スタンダード","¥1,500",IF(B' + currentPlanRow + '="プロフェッショナル","¥3,000","¥7,500")))');
  
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
  
  // 列幅調整
  sheet.setColumnWidth(1, 200); // 項目列
  sheet.setColumnWidth(2, 150); // ベーシック
  sheet.setColumnWidth(3, 150); // スタンダード
  sheet.setColumnWidth(4, 150); // プロフェッショナル
  sheet.setColumnWidth(5, 150); // エンタープライズ
  
  // 行の高さ調整
  sheet.setRowHeight(1, 40); // タイトル
  sheet.setRowHeight(4, 40); // ヘッダー
  
  Logger.log('プラン説明シートが作成されました');
}ues(SHEET_NAMES);
    const missingSheets = [];
    
    for (const sheetName of requiredSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        missingSheets.push(sheetName);
      }
    }
    
    if (missingSheets.length > 0) {
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    return { success: true, spreadsheet: spreadsheet };
  } catch (error) {
    console.error('システム初期化確認エラー:', error.toString());
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
    const missingSheets = [];
    
    for (const sheetName of requiredSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        missingSheets.push(sheetName);
      }
    }
    
    if (missingSheets.length > 0) {
      throw new Error(`以下のシートが見つかりません: ${missingSheets.join(', ')}\n\ninitializeSheets()関数を実行してシステムを初期化してください。`);
    }
    
    return true;
  } catch (error) {
    console.error('システム初期化確認エラー:', error.toString());
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
 * ダッシュボードの作成
 */
function createDashboard(sheet) {
  sheet.getRange('D1').setValue('ダッシュボード').setFontSize(14).setFontWeight('bold');
  
  const dashboardData = [
    ['登録企業数', '=COUNTA(企業マスター!A:A)-1'],
    ['提案生成数', '=COUNTA(提案メッセージ!A:A)-1'],
    ['平均マッチ度', '=AVERAGE(企業マスター!K:K)'],
    ['最終実行日', '=MAX(実行ログ!B:B)'],
    ['今月の実行回数', ''],
    ['エラー率', '=IF(SUM(実行ログ!F:F)=0,0,SUM(実行ログ!F:F)/SUM(実行ログ!E:E))']
  ];
  
  sheet.getRange(2, 4, dashboardData.length, 2).setValues(dashboardData);
}

/**
 * 生成キーワードシートの作成
 */
function createKeywordsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.KEYWORDS);
  }
  
  const headers = [
    'キーワード', 'カテゴリ', '優先度', '戦略説明', '実行済み', 'ヒット件数', '最終実行日'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200); // キーワード
  sheet.setColumnWidth(4, 300); // 戦略説明
}

/**
 * 企業マスターシートの作成
 */
function createCompaniesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COMPANIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.COMPANIES);
  }
  
  const headers = [
    '企業ID', '会社名', '公式URL', '業界', '従業員数', '本社所在地', 
    '問合せ方法', '上場区分', '事業内容', '企業規模判定', 'マッチ度スコア', 
    '発見キーワード', '登録日時'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  
  // 列幅の調整
  sheet.setColumnWidth(2, 200); // 会社名
  sheet.setColumnWidth(3, 200); // 公式URL
  sheet.setColumnWidth(9, 300); // 事業内容
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
    '企業ID', '提案パターンA_件名', '提案パターンA_本文', '提案パターンB_件名', 
    '提案パターンB_本文', 'フォーム用メッセージ', '推奨アプローチ方法', 
    '想定課題', '提供価値', 'アプローチタイミング', '生成日時'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  
  // 列幅の調整
  sheet.setColumnWidth(3, 400); // 提案パターンA_本文
  sheet.setColumnWidth(5, 400); // 提案パターンB_本文
}

/**
 * 実行ログシートの作成
 */
function createLogsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.LOGS);
  }
  
  const headers = [
    '実行ID', '実行日時', '実行タイプ', '対象キーワード', '成功件数', 
    'エラー件数', 'エラー詳細', '処理時間', 'API使用量'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  
  // 列幅の調整
  sheet.setColumnWidth(7, 300); // エラー詳細
}

/**
 * 制御パネルから設定値を取得
 */
function getControlPanelSettings() {
  const sheet = getSheet(SHEET_NAMES.CONTROL);
  
  return {
    productName: sheet.getRange('B2').getValue(),
    productDescription: sheet.getRange('B3').getValue(),
    priceRange: sheet.getRange('B4').getValue(),
    targetSize: sheet.getRange('B5').getValue(),
    preferredRegion: sheet.getRange('B6').getValue(),
    maxCompanies: sheet.getRange('B11').getValue() || 20
  };
}

/**
 * エラーハンドリングとユーザーへの通知
 */
function handleSystemError(functionName, error) {
  const errorMessage = error.toString();
  console.error(`${functionName}エラー:`, errorMessage);
  
  // ログに記録
  try {
    logExecution(functionName, 'ERROR', 0, 1, errorMessage);
  } catch (logError) {
    console.error('ログ記録エラー:', logError.toString());
  }
  
  // UIが利用可能な場合のみアラートを表示
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      if (errorMessage.includes('getSheetByName') || errorMessage.includes('アクティブなスプレッドシート')) {
        const message = `❌ システムが初期化されていません。\n\n解決方法:\n1. メニューバーから「拡張機能」→「Apps Script」を選択\n2. 関数一覧から「initializeSheets」を選択\n3. 実行ボタンをクリックしてシステムを初期化`;
        ui.alert('初期化エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
      } else if (errorMessage.includes('以下のシートが見つかりません')) {
        ui.alert('シート不足エラー', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        ui.alert(`${functionName}エラー`, errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
      }
    }
  } catch (uiError) {
    // UIが利用できない場合（トリガーなど）はコンソールにのみ出力
    console.error('UI利用不可 - コンソールに出力:', errorMessage);
  }
  
  throw error;
}

/**
 * スプレッドシートIDを設定（ウェブアプリ用）
 */
function setSpreadsheetId(spreadsheetId) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  console.log('スプレッドシートIDが設定されました:', spreadsheetId);
}

/**
 * 設定されたスプレッドシートIDを取得
 */
function getConfiguredSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

/**
 * 安全なシート取得ヘルパー関数
 */
function getSheet(sheetName) {
  try {
    const spreadsheet = getSafeSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    return sheet;
  } catch (error) {
    console.error(`シート取得エラー (${sheetName}):`, error.toString());
    throw error;
  }
}

/**
 * 安全なスプレッドシート取得
 */
function getSafeSpreadsheet() {
  try {
    // まずアクティブスプレッドシートを試す
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSpreadsheet) {
      return activeSpreadsheet;
    }
  } catch (error) {
    console.log('アクティブスプレッドシートが利用できません');
  }
  
  // 設定されたIDから取得を試す
  const configuredId = getConfiguredSpreadsheetId();
  if (configuredId) {
    try {
      return SpreadsheetApp.openById(configuredId);
    } catch (error) {
      console.error('設定されたスプレッドシートIDでの取得に失敗:', error.toString());
    }
  }
  
  throw new Error('スプレッドシートが見つかりません。スプレッドシートにバインドするか、setSpreadsheetId()でIDを設定してください。');
}

/**
 * 実行状況の更新
 */
function updateExecutionStatus(message) {
  try {
    const sheet = getSheet(SHEET_NAMES.CONTROL);
    const timestamp = new Date().toLocaleString('ja-JP');
    sheet.getRange('B15').setValue(`${timestamp}: ${message}`);
    SpreadsheetApp.flush();
  } catch (error) {
    console.error('実行状況更新エラー:', error.toString());
  }
}

/**
 * メニューの追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('営業自動化システム')
    .addItem('システム初期化', 'initializeSheets')
    .addSeparator()
    .addItem('キーワード生成', 'generateStrategicKeywords')
    .addItem('企業検索実行', 'executeCompanySearch')
    .addSubMenu(ui.createMenu('提案メッセージ生成')
      .addItem('通常の提案生成', 'generatePersonalizedProposals')
      .addItem('詳細ログ付き提案生成', 'executeProposalGenerationEnhanced'))
    .addSeparator()
    .addItem('全自動実行', 'executeFullWorkflow')
    .addSeparator()
    .addSubMenu(ui.createMenu('診断・デバッグ')
      .addItem('詳細システム診断', 'performDetailedDiagnostics')
      .addItem('提案生成ログ確認', 'showProposalErrorLog')
      .addItem('API接続テスト', 'testAllApiConnections'))
    .addSeparator()
    .addItem('APIキー設定', 'showApiKeyDialog')
    .addToUi();
}

/**
 * APIキー設定ダイアログ
 */
function showApiKeyDialog() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px;">
      <h3>APIキー設定</h3>
      <label>Google Custom Search API Key:</label><br>
      <input type="text" id="googleSearchKey" style="width: 100%; margin-bottom: 10px;"><br>
      
      <label>Google Search Engine ID:</label><br>
      <input type="text" id="searchEngineId" style="width: 100%; margin-bottom: 10px;"><br>
      
      <label>OpenAI API Key:</label><br>
      <input type="text" id="openaiKey" style="width: 100%; margin-bottom: 20px;"><br>
      
      <button onclick="saveApiKeys()">保存</button>
    </div>
    
    <script>
      function saveApiKeys() {
        const googleSearchKey = document.getElementById('googleSearchKey').value;
        const searchEngineId = document.getElementById('searchEngineId').value;
        const openaiKey = document.getElementById('openaiKey').value;
        
        google.script.run
          .withSuccessHandler(() => {
            alert('APIキーが保存されました');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('エラー: ' + error.toString());
          })
          .saveApiKeysToProperties(googleSearchKey, searchEngineId, openaiKey);
      }
    </script>
  `).setWidth(500).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'APIキー設定');
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
  
  Logger.log('APIキーが保存されました');
}
