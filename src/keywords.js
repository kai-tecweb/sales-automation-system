/**
 * キーワード生成機能 - OpenAI API統合版
 */

/**
 * メニュー用キーワード生成関数
 */
function generateKeywords() {
  try {
    console.log('🔤 キーワード生成を開始します...');
    
    updateExecutionStatus('キーワード生成を開始します...');
    
    // 制御パネルから設定を取得
    console.log('制御パネル設定を取得中...');
    const settings = getControlPanelSettings();
    console.log('取得した設定:', settings);
    
    // 入力値チェック
    if (!settings.productName || !settings.productDescription) {
      SpreadsheetApp.getUi().alert('❌ エラー', '制御パネルで商材名と商材概要を入力してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // APIキーの確認
    console.log('APIキー確認中...');
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    console.log('APIキー取得結果:', apiKey ? '設定済み' : '未設定');
    
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('❌ エラー', 'OpenAI APIキーが設定されていません。\n\nメニュー > ⚙️ システム管理 > ⚙️ API設定管理 から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // ChatGPT APIでキーワード生成（ハイブリッド方式）
    console.log('ハイブリッド方式でキーワード生成開始');
    
    updateExecutionStatus('AI分析 + 固定パターンでキーワード生成中...');
    
    let keywords;
    try {
      // ハイブリッド方式：固定パターン + AI補強
      keywords = generateKeywordsHybrid(settings);
      
      console.log('ハイブリッド生成成功');
      console.log('keywordsの型:', typeof keywords);
      console.log('keywordsの内容:', keywords);
      console.log('keywordsは配列か:', Array.isArray(keywords));
      
      // Promiseオブジェクトかチェック（デバッグ用に詳細表示）
      if (keywords && typeof keywords.then === 'function') {
        console.error('❌ Promiseオブジェクトが返されました');
        console.error('Promiseの詳細:', keywords);
        console.error('Promiseのプロパティ:', Object.keys(keywords));
        console.error('keywords.toString():', keywords.toString());
        
        // Promiseエラーを一時的にスキップして詳細を確認
        console.log('⚠️ Promiseエラーをスキップして処理を続行します（デバッグ用）');
        
        // Promiseの場合でも処理を続行してみる
        try {
          console.log('Promise解決を試みます...');
          // 注意：これは通常避けるべきですが、デバッグ用に一時的に実行
          const resolvedValue = keywords;
          console.log('Promise内容（未解決状態）:', resolvedValue);
        } catch (promiseError) {
          console.error('Promise内容確認エラー:', promiseError);
        }
        
        // エラーメッセージを変更
        throw new Error('Promise検出: 詳細はログを確認してください');
      }
      
      // 返り値の詳細チェック
      console.log('=== 返り値詳細チェック ===');
      console.log('keywords:', keywords);
      console.log('型:', typeof keywords);
      console.log('コンストラクタ:', keywords ? keywords.constructor.name : 'null/undefined');
      if (keywords && typeof keywords === 'object') {
        console.log('オブジェクトのキー:', Object.keys(keywords));
      }
      
    } catch (apiError) {
      console.error('API呼び出しエラー:', apiError);
      throw new Error(`API呼び出しエラー: ${apiError.message}`);
    }
    

    

    
    // キーワードの検証
    if (!keywords) {
      throw new Error('キーワード生成に失敗しました - 結果がnullです');
    }
    
    if (!Array.isArray(keywords)) {
      console.error('❌ 配列ではないオブジェクト:', keywords);
      console.error('オブジェクトのプロパティ:', Object.keys(keywords || {}));
      throw new Error(`キーワード生成に失敗しました - 配列ではありません (型: ${typeof keywords})\n内容: ${JSON.stringify(keywords, null, 2).substring(0, 500)}`);
    }
    
    if (keywords.length === 0) {
      throw new Error('キーワード生成に失敗しました - 空の配列です');
    }
    
    // 結果をスプレッドシートに保存
    updateExecutionStatus('結果を保存中...');
    saveKeywordsToSheet(keywords);
    
    const successMessage = `✅ キーワード生成完了！\n\n${keywords.length}個の戦略的キーワードを生成しました。\n「生成キーワード」シートで確認できます。`;
    
    updateExecutionStatus(`キーワード生成完了: ${keywords.length}個のキーワードを生成しました`);
    logExecution('キーワード生成', 'SUCCESS', keywords.length, 0);
    
    SpreadsheetApp.getUi().alert('キーワード生成完了', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log('✅ キーワード生成処理完了');
    
  } catch (error) {
    console.error('❌ キーワード生成エラー:', error);
    console.error('エラースタック:', error.stack);
    
    updateExecutionStatus(`エラー: ${error.message}`);
    logExecution('キーワード生成', 'ERROR', 0, 1);
    SpreadsheetApp.getUi().alert('❌ エラー', `キーワード生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 制御パネル設定を取得
 */
function getControlPanelSettings() {
  const sheet = getSafeSheet(SHEET_NAMES.CONTROL);
  if (!sheet) {
    throw new Error('制御パネルシートが見つかりません。システムを初期化してください。');
  }
  
  // 制御パネルからデータを取得
  const data = sheet.getRange('A2:B6').getValues();
  
  return {
    productName: data[0][1] || '',
    productDescription: data[1][1] || '',
    priceRange: data[2][1] || '',
    targetSize: data[3][1] || '',
    preferredRegion: data[4][1] || ''
  };
}

/**
 * ChatGPT APIを使用してキーワードを生成（完全同期版）
 */
function generateKeywordsWithChatGPT(settings) {
  console.log('=== generateKeywordsWithChatGPT関数開始（完全同期版）===');
  
  try {
    // APIキーを確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!apiKey) {
      throw new Error('OpenAI APIキーが設定されていません');
    }
    
    console.log('APIキー確認完了');
  
  const prompt = createKeywordGenerationPrompt(settings);
  console.log('プロンプト作成完了');
  
  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: 'あなたは営業戦略の専門家です。商材情報から効果的な企業検索キーワードを生成してください。JSON形式で回答してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: 2000,
    temperature: 0.7
  };
  
  console.log('ペイロード準備完了');
  
  // 直接UrlFetchAppを使用して確実に同期処理
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  console.log('UrlFetchApp.fetch実行開始');
  
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  
  console.log('UrlFetchApp.fetch実行完了');
  console.log('レスポンスコード:', response.getResponseCode());
  
  if (response.getResponseCode() !== 200) {
    const errorText = response.getContentText();
    throw new Error(`OpenAI API Error (${response.getResponseCode()}): ${errorText}`);
  }
  
  const responseText = response.getContentText();
  console.log('レスポンステキスト取得完了');
  
  const data = JSON.parse(responseText);
  const messageContent = data.choices[0].message.content;
  
  console.log('メッセージ内容の型:', typeof messageContent);
  
  console.log('パース処理開始');
  const keywords = parseKeywordResponse(messageContent);
  
  console.log('パース処理完了');
  console.log('キーワードの型:', typeof keywords);
  console.log('キーワードは配列か:', Array.isArray(keywords));
  
  return keywords;
  
  } catch (error) {
    console.error('❌ generateKeywordsWithChatGPT内でエラー:', error);
    console.error('エラースタック:', error.stack);
    throw error;
  }
}

/**
 * キーワード生成用プロンプトの作成（企業直接検索特化版）
 */
function createKeywordGenerationPrompt(settings) {
  return `
商材情報：
- 商材名: ${settings.productName}
- 概要: ${settings.productDescription}
- 価格帯: ${settings.priceRange}
- 対象企業規模: ${settings.targetSize}
- 優先地域: ${settings.preferredRegion || '指定なし'}

【重要】実際に存在する企業の公式サイトや会社概要ページを直接検索するキーワードを生成してください。
記事、ニュース、解説サイトではなく、企業の公式情報を見つけることが目標です。

以下の4つのカテゴリで、実企業検索用キーワードを生成してください：

1. industryDirect: 業界名+企業規模+地域の直接検索
2. companyPages: 会社概要・企業情報ページの直接検索  
3. businessAssociations: 商工会・業界団体経由の企業検索
4. locationBased: 地域密着企業の直接検索

【必須条件】
- 各キーワードには必ず除外ワード(-wikipedia -facebook -twitter -求人 -採用 -ニュース -記事 -まとめ -解説)を含める
- 企業の公式サイトを直接見つけるためのクエリにする
- 記事やブログではなく、企業情報ページがヒットするキーワードにする

【出力例】
"製造業 中小企業 愛知県 会社概要 -wikipedia -facebook -twitter -求人 -採用 -ニュース -記事"
"IT企業 従業員数 東京 企業情報 -まとめ -解説"
"建設会社 代表取締役 大阪 site:co.jp -求人"

出力は以下のJSON形式でお願いします：
{
  "companySearchKeywords": {
    "industryDirect": {
      "purpose": "業界×規模×地域での企業直接検索",
      "keywords": ["具体的な検索クエリ1", "具体的な検索クエリ2", ...]
    },
    "companyPages": {
      "purpose": "企業情報ページの直接検索",
      "keywords": ["具体的な検索クエリ1", "具体的な検索クエリ2", ...]
    },
    "businessAssociations": {
      "purpose": "商工会・業界団体経由の企業検索",
      "keywords": ["具体的な検索クエリ1", "具体的な検索クエリ2", ...]
    },
    "locationBased": {
      "purpose": "地域密着企業の直接検索",
      "keywords": ["具体的な検索クエリ1", "具体的な検索クエリ2", ...]
    }
  }
}`;
}

/**
 * 商材説明から業界を抽出（営業適性考慮版）
 */
function extractIndustries(productDescription) {
  const industryMap = {
    '製造': ['製造業', '機械メーカー', '部品メーカー', '工場'],
    '機械': ['機械メーカー', '精密機械', '産業機械', '工作機械'],
    '自動車': ['自動車部品', '部品メーカー', '製造業'], // トヨタ自動車は除外済み
    'IT': ['IT企業', 'システム開発', 'ソフトウェア開発', 'Web制作'],
    'システム': ['システム開発', 'IT企業', 'ソフトウェア開発'],
    'ソフト': ['ソフトウェア開発', 'IT企業', 'アプリ開発'],
    '建設': ['建設業', '建築会社', '工務店', '設計事務所'],
    '建築': ['建築会社', '建設業', '設計事務所', '住宅建築'],
    '医療': ['医療機器', 'ヘルスケア', '介護サービス', 'クリニック'],
    '物流': ['運送業', '物流会社', '配送業', '倉庫業'],
    '小売': ['小売業', '卸売業', '商社', '販売店'],
    '金融': [], // 営業困難業界として除外
    '教育': ['学習塾', '専門学校', '研修会社', '教育サービス'],
    '不動産': ['不動産会社', '管理会社', '賃貸業'],
    '飲食': ['レストラン', '居酒屋', 'カフェ', '食品メーカー'],
    'サービス': ['サービス業', 'コンサルティング', 'サポートサービス'],
    '化学': ['化学工業', '製薬関連', '化学メーカー'],
    '電子': ['電子部品', '電子機器', '精密機械'],
    '繊維': ['繊維関連', '繊維メーカー', '衣料品'],
    
    // 業界拡張
    '商社': ['商社', '貿易会社', '卸売業'],
    '食品': ['食品メーカー', '食品加工', '飲料メーカー'],
    '印刷': ['印刷業', '出版業', 'デザイン会社'],
    '運輸': ['運送業', '物流会社', '交通会社'],
    '環境': ['環境関連', 'リサイクル業', '廃棄物処理'],
    '農業': ['農業関連', '農機具', '農業法人']
  };
  
  const detectedIndustries = [];
  const description = productDescription.toLowerCase();
  
  // 営業困難業界のチェック
  const difficultIndustries = ['銀行', '証券', '保険', '官公庁', '学校', '病院', '大学'];
  const hasDifficultIndustry = difficultIndustries.some(industry => 
    description.includes(industry.toLowerCase())
  );
  
  if (hasDifficultIndustry) {
    console.log('営業困難業界を検出、一般的な業界に変更');
    return ['製造業', 'IT企業', 'サービス業']; // フォールバック
  }
  
  Object.entries(industryMap).forEach(([key, industries]) => {
    if (description.includes(key.toLowerCase()) || 
        industries.some(industry => description.includes(industry.toLowerCase()))) {
      detectedIndustries.push(...industries);
    }
  });
  
  // 重複除去
  const uniqueIndustries = [...new Set(detectedIndustries)];
  
  // 業界が検出されない場合はデフォルト業界を返す
  return uniqueIndustries.length > 0 ? uniqueIndustries : ['製造業', 'IT企業', 'サービス業'];
}

/**
 * 確実性重視：固定パターンベースのキーワード生成（効率最適化版）
 */
function generateReliableKeywords(settings) {
  const targetIndustries = extractIndustries(settings.productDescription);
  const targetRegions = settings.preferredRegion ? 
    [settings.preferredRegion] : 
    ['東京', '大阪', '愛知', '神奈川', '埼玉', '千葉', '兵庫', '福岡', '静岡', '茨城', '京都', '広島']; // 地域拡張
  
  // 除外ワードの段階的緩和（質と量のバランス重視）
  const coreExclusions = '-wikipedia -求人 -採用'; // 最重要のみ（緩和版）
  const mediumExclusions = coreExclusions + ' -ニュース -記事'; // 記事系追加
  const strictExclusions = mediumExclusions + ' -商工会議所 -協会'; // 機関系追加（後処理で対応）
  
  const megaCorpExclusions = '-トヨタ自動車 -ソニー -NTT -ソフトバンク'; // 重要な大企業のみ
  const additionalMegaCorpExclusions = '-ミズノ -スタンレー電気'; // 主要大企業のみ（緩和）
  const governmentExclusions = '-go.jp -lg.jp'; // 政府サイトのみ（緩和）
  const socialMediaExclusions = '-facebook -twitter'; // 補助的
  
  // 地域別営業戦略マップ
  const regionIndustryMap = {
    '愛知': ['自動車部品', '機械加工', '製造業'],
    '大阪': ['製薬関連', '化学工業', '電子部品'],
    '東京': ['IT関連', 'システム開発', 'コンサルティング'],
    '神奈川': ['精密機械', '電子部品', '医療機器'],
    '埼玉': ['製造業', '物流', 'IT企業'],
    '千葉': ['製造業', '物流', '化学工業'],
    '兵庫': ['製造業', '機械加工', '電子部品'],
    '福岡': ['IT企業', 'サービス業', '製造業']
  };
  
  // 業界特化パターン（効率重視）
  const industryPatterns = {
    '製造業': [
      '{industry} 中小企業 {region} 工場',
      '{industry} {region} 製作所 企業情報',
      '機械加工 {region} 企業一覧',
      '部品メーカー {region} 代表取締役'
    ],
    '機械メーカー': [
      '機械メーカー 中小企業 {region}',
      '精密機械 {region} 企業情報'
    ],
    '自動車部品': [
      '自動車部品 {region} 製造',
      '部品メーカー {region} 工場'
    ],
    'IT企業': [
      '{industry} 中小企業 {region} オフィス',
      '{industry} {region} 本社 企業情報',
      'システム開発 {region} 企業一覧',
      'Web制作 {region} 会社案内'
    ],
    'システム開発': [
      'システム開発 {region} オフィス',
      'IT企業 {region} 本社'
    ],
    '建設業': [
      '{industry} 中小企業 {region} 事務所',
      '建築会社 {region} 企業情報',
      '工務店 {region} 代表取締役'
    ],
    'サービス業': [
      '{industry} 中小企業 {region} 事業所',
      'コンサルティング {region} オフィス'
    ]
  };
  
  // 高品質検索パターン（企業公式サイト直接検索）
  const premiumPatterns = [
    '{industry} 中小企業 {region} site:co.jp 会社概要',
    '{industry} {region} 企業情報 site:co.jp 代表取締役',
    '{industry} {region} site:co.jp "従業員数" -求人',
    '{industry} {region} site:co.jp "事業内容" -採用'
  ];
  
  // 効率的検索パターン
  const efficientPatterns = [
    '従業員数 10-200名 {industry} {region}',
    '中小企業 {industry} {region} site:co.jp',
    '{industry} {region} site:co.jp 会社概要',
    '{industry} 企業概要 {region}'
  ];
  
  const keywords = [];
  
  // Step 1: プレミアムパターン（企業公式サイト直接検索）- 最高品質
  premiumPatterns.forEach(pattern => {
    targetIndustries.slice(0, 3).forEach(industry => {
      targetRegions.slice(0, 4).forEach(region => {
        if (keywords.length >= 25) return;
        
        const keyword = pattern
          .replace('{industry}', industry)
          .replace('{region}', region);
        
        keywords.push({
          keyword: keyword + ' ' + coreExclusions, // 最小限の除外ワードのみ
          category: 'premiumSearch',
          priority: '最高',
          strategy: '企業公式サイト直接検索',
          executed: false,
          hitCount: 0,
          lastExecuted: ''
        });
      });
    });
  });
  
  // Step 2: 業界特化パターン - 高品質
  targetIndustries.forEach(industry => {
    const patterns = industryPatterns[industry] || industryPatterns['サービス業'];
    
    patterns.forEach(pattern => {
      targetRegions.slice(0, 4).forEach(region => {
        if (keywords.length >= 45) return;
        
        const keyword = pattern
          .replace('{industry}', industry)
          .replace('{region}', region) + ' ' + coreExclusions; // 最小限の除外ワードに変更
        
        keywords.push({
          keyword: keyword,
          category: 'industrySpecific',
          priority: '高',
          strategy: `${industry}特化検索`,
          executed: false,
          hitCount: 0,
          lastExecuted: ''
        });
      });
    });
  });
  
  // Step 3: 地域特化パターン - 地域戦略
  targetRegions.slice(0, 3).forEach(region => {
    const regionIndustries = regionIndustryMap[region] || ['製造業', 'IT企業'];
    
    regionIndustries.forEach(industry => {
      if (keywords.length >= 60) return;
      
      const keyword = `${industry} 中小企業 ${region} 企業概要 ${coreExclusions}`;
      
      keywords.push({
        keyword: keyword,
        category: 'regionSpecific',
        priority: '高',
        strategy: `${region}特化検索`,
        executed: false,
        hitCount: 0,
        lastExecuted: ''
      });
    });
  });
  
  // Step 4: 効率的検索パターン - 大企業除外強化
  efficientPatterns.forEach(pattern => {
    targetIndustries.slice(0, 2).forEach(industry => {
      targetRegions.slice(0, 3).forEach(region => {
        if (keywords.length >= 80) return;
        
        const keyword = pattern
          .replace('{industry}', industry)
          .replace('{region}', region) + ' ' + coreExclusions; // 最小限の除外ワードに変更
        
        keywords.push({
          keyword: keyword,
          category: 'efficientSearch',
          priority: '中',
          strategy: '効率的企業検索',
          executed: false,
          hitCount: 0,
          lastExecuted: ''
        });
      });
    });
  });
  
  return keywords;
}

/**
 * キーワードの営業適性スコアリング（効率最適化版）
 */
function calculateProspectScore(keyword) {
  let score = 50; // 基準点
  
  // 企業公式サイト直接検索で大幅加点（最高信頼性）
  if (keyword.includes('site:co.jp')) score += 25;
  if (keyword.includes('企業概要') || keyword.includes('会社概要')) score += 15;
  if (keyword.includes('代表取締役') || keyword.includes('事業内容')) score += 10;
  
  // 機関サイト除外で加点
  if (keyword.includes('-商工会議所') || keyword.includes('-協会')) score += 15;
  
  // 除外ワードありで加点
  if (keyword.includes('-wikipedia') || keyword.includes('-求人')) score += 10;
  
  // 企業規模指定で加点
  if (keyword.includes('中小企業') || keyword.includes('従業員数')) score += 15;
  if (keyword.includes('10-200名') || keyword.includes('1億-50億')) score += 10;
  
  // site:co.jpで加点（企業サイト限定）
  if (keyword.includes('site:co.jp')) score += 15;
  
  // 地域限定で加点
  if (keyword.includes('東京') || keyword.includes('大阪') || keyword.includes('愛知')) score += 5;
  
  // 業界特化で加点
  if (keyword.includes('製造業') || keyword.includes('IT企業')) score += 10;
  
  // 大企業除外で加点
  if (keyword.includes('-トヨタ自動車') || keyword.includes('-NTT')) score += 10;
  
  // キーワード長による効率性評価
  const keywordLength = keyword.length;
  if (keywordLength < 80) score += 5; // 短い＝効率的
  if (keywordLength > 150) score -= 10; // 長すぎる＝非効率
  
  // 営業困難パターンで減点
  if (keyword.includes('銀行') || keyword.includes('官公庁')) score -= 30;
  if (keyword.includes('大企業') || keyword.includes('上場企業')) score -= 20;
  
  return Math.min(Math.max(score, 0), 100); // 0-100の範囲に制限
}

/**
 * キーワード品質の検証
 */
function validateKeywordQuality(keywords) {
  const analysis = {
    total: keywords.length,
    withExcludeTerms: 0,
    highPriority: 0,
    industrySpecific: 0,
    averageScore: 0,
    recommendations: []
  };
  
  let totalScore = 0;
  
  keywords.forEach(kw => {
    // 除外ワード検証
    if (kw.keyword.includes('-wikipedia') || kw.keyword.includes('-求人')) {
      analysis.withExcludeTerms++;
    }
    
    // 優先度検証
    if (kw.priority === '高') {
      analysis.highPriority++;
    }
    
    // 業界特化検証
    if (kw.category === 'industrySpecific') {
      analysis.industrySpecific++;
    }
    
    // スコア計算
    const score = calculateProspectScore(kw.keyword);
    totalScore += score;
  });
  
  analysis.averageScore = Math.round(totalScore / keywords.length);
  
  // 推奨事項生成
  if (analysis.withExcludeTerms / analysis.total < 0.8) {
    analysis.recommendations.push('除外ワード付きキーワードを80%以上にすることを推奨');
  }
  
  if (analysis.averageScore < 60) {
    analysis.recommendations.push('キーワード品質が低いため、営業適性の見直しが必要');
  }
  
  if (analysis.industrySpecific / analysis.total < 0.3) {
    analysis.recommendations.push('業界特化キーワードを30%以上にすることを推奨');
  }
  
  return analysis;
}

/**
 * 検索結果の期待値設定と性能予測
 */
function predictSearchPerformance(keywords) {
  const analysis = {
    totalKeywords: keywords.length,
    expectedResults: {
      corporateSiteHitRate: 0,      // 企業公式サイトヒット率
      prospectCompanyRate: 0,       // 営業対象企業率
      articleSiteMixRate: 0,        // 記事・ブログ混入率
      averageCompaniesPerKeyword: 0 // キーワードあたり平均企業発見数
    },
    qualityDistribution: {
      premium: 0,   // 商工会・協会系
      high: 0,      // 業界・地域特化
      medium: 0,    // 効率的検索
      low: 0        // その他
    },
    recommendations: []
  };
  
  // カテゴリ別集計
  keywords.forEach(kw => {
    switch(kw.category) {
      case 'premiumSearch':
        analysis.qualityDistribution.premium++;
        break;
      case 'industrySpecific':
      case 'regionSpecific':
        analysis.qualityDistribution.high++;
        break;
      case 'efficientSearch':
        analysis.qualityDistribution.medium++;
        break;
      default:
        analysis.qualityDistribution.low++;
    }
  });
  
  // 期待値計算
  const premiumRate = analysis.qualityDistribution.premium / analysis.totalKeywords;
  const highQualityRate = analysis.qualityDistribution.high / analysis.totalKeywords;
  
  // 企業公式サイトヒット率予測（site:co.jp戦略により大幅改善）
  analysis.expectedResults.corporateSiteHitRate = Math.round(
    (premiumRate * 0.95 + highQualityRate * 0.80 + 0.70) * 100
  );
  
  // 営業対象企業率予測（機関サイト除外により改善）
  analysis.expectedResults.prospectCompanyRate = Math.round(
    (premiumRate * 0.85 + highQualityRate * 0.70 + 0.60) * 100
  );
  
  // 記事・ブログ混入率予測（強化された除外ワードにより大幅改善）
  analysis.expectedResults.articleSiteMixRate = Math.round(
    (1 - (premiumRate * 0.98 + highQualityRate * 0.90 + 0.80)) * 100
  );
  
  // キーワードあたり平均企業発見数予測
  analysis.expectedResults.averageCompaniesPerKeyword = Math.round(
    (premiumRate * 8 + highQualityRate * 5 + 3) * 10
  ) / 10;
  
  // 推奨事項（修正後戦略対応）
  if (premiumRate < 0.3) {
    analysis.recommendations.push('企業公式サイト直接検索キーワードを30%以上に増やすことを推奨');
  }
  
  if (analysis.expectedResults.corporateSiteHitRate < 80) {
    analysis.recommendations.push('企業公式サイトヒット率80%以上を目指すため、site:co.jp使用を増やすことを推奨');
  }
  
  if (analysis.expectedResults.articleSiteMixRate > 10) {
    analysis.recommendations.push('機関・記事サイト混入率を10%以下にするため、除外ワードの追加強化を推奨');
  }
  
  return analysis;
}

/**
 * ハイブリッド方式：固定パターン + AI最適化（品質検証付き）
 */
function generateKeywordsHybrid(settings) {
  console.log('ハイブリッド方式でキーワード生成開始');
  
  // Step 1: 確実な固定パターンで基本キーワード生成
  const reliableKeywords = generateReliableKeywords(settings);
  console.log(`固定パターンで${reliableKeywords.length}個のキーワード生成`);
  
  // Step 2: AIで業界特化キーワードを補強（オプション）
  let aiEnhancedKeywords = [];
  try {
    console.log('AI生成を試行中...');
    const aiResult = generateKeywordsWithChatGPT(settings);
    if (Array.isArray(aiResult)) {
      aiEnhancedKeywords = aiResult.slice(0, 20); // AI生成は最大20個まで
      console.log(`AI生成で${aiEnhancedKeywords.length}個の追加キーワード生成`);
    }
  } catch (error) {
    console.log('AI生成失敗、固定パターンのみ使用:', error.message);
  }
  
  // Step 3: 結合・重複除去
  const allKeywords = [...reliableKeywords, ...aiEnhancedKeywords];
  const uniqueKeywords = removeDuplicateKeywords(allKeywords);
  
  // Step 4: 営業適性スコアリング
  const scoredKeywords = uniqueKeywords.map(kw => ({
    ...kw,
    prospectScore: calculateProspectScore(kw.keyword)
  }));
  
  // Step 5: 品質検証
  const qualityAnalysis = validateKeywordQuality(scoredKeywords);
  console.log('キーワード品質分析:', qualityAnalysis);
  
  // Step 5.5: 性能予測
  const performancePrediction = predictSearchPerformance(scoredKeywords);
  console.log('検索性能予測:', performancePrediction);
  
  // Step 6: 高品質キーワードを優先してソート
  const finalKeywords = scoredKeywords
    .sort((a, b) => {
      // 1. 営業適性スコア優先
      if (b.prospectScore !== a.prospectScore) {
        return b.prospectScore - a.prospectScore;
      }
      // 2. 優先度で二次ソート（最高>高>中>低）
      const priorityOrder = { '最高': 4, '高': 3, '中': 2, '低': 1 };
      return (priorityOrder[b.priority] || 1) - (priorityOrder[a.priority] || 1);
    })
    .slice(0, 80); // 最大80個に制限
  
  console.log(`最終的に${finalKeywords.length}個のキーワードを生成`);
  console.log(`平均営業適性スコア: ${qualityAnalysis.averageScore}点`);
  console.log(`期待企業公式サイトヒット率: ${performancePrediction.expectedResults.corporateSiteHitRate}%`);
  console.log(`期待営業対象企業率: ${performancePrediction.expectedResults.prospectCompanyRate}%`);
  
  if (qualityAnalysis.recommendations.length > 0) {
    console.log('品質改善推奨事項:', qualityAnalysis.recommendations);
  }
  
  if (performancePrediction.recommendations.length > 0) {
    console.log('性能改善推奨事項:', performancePrediction.recommendations);
  }
  
  return finalKeywords;
}

/**
 * 結果フィルタリング強化機能：機関サイト除外＋企業サイト特定
 */
function enhancedFilterValidCompanies(searchResults) {
  if (!Array.isArray(searchResults)) {
    return [];
  }
  
  return searchResults.filter(result => {
    const url = (result.link || '').toLowerCase();
    const title = (result.title || '').toLowerCase();
    const snippet = (result.snippet || '').toLowerCase();
    
    // 除外すべきサイトパターン（機関サイト対策）
    const excludePatterns = [
      // 機関サイト
      'cci.or.jp',           // 商工会議所
      '商工会議所',
      '商工会',
      '協会',
      '組合',
      
      // リスト・検索サイト
      '会員企業',
      '企業一覧',
      'リスト',
      '名簿',
      'ポータル',
      '検索',
      'まとめ',
      '一覧',
      
      // 大企業（営業困難）
      'toyota.co.jp',
      'sony.co.jp',
      'ntt.co.jp',
      'softbank.co.jp',
      
      // 求人・採用サイト
      '求人',
      '採用',
      'recruit',
      'jobs',
      
      // ニュース・記事サイト
      'news',
      'ニュース',
      '記事',
      'article',
      'blog'
    ];
    
    const isExcluded = excludePatterns.some(pattern => 
      url.includes(pattern) || 
      title.includes(pattern) || 
      snippet.includes(pattern)
    );
    
    // 企業サイトの特徴をチェック
    const isCompanySite = (
      url.includes('.co.jp') && 
      (title.includes('株式会社') || 
       title.includes('有限会社') || 
       title.includes('合同会社') ||
       snippet.includes('代表取締役') ||
       snippet.includes('企業概要') ||
       snippet.includes('会社案内') ||
       snippet.includes('会社情報') ||
       snippet.includes('企業情報') ||
       snippet.includes('事業内容'))
    );
    
    // PDFファイルを除外
    const isPdf = url.includes('.pdf');
    
    return !isExcluded && isCompanySite && !isPdf;
  });
}

/**
 * キーワード重複除去
 */
function removeDuplicateKeywords(keywords) {
  const seen = new Set();
  return keywords.filter(kw => {
    const key = kw.keyword.toLowerCase().replace(/\s+/g, ' ').trim();
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

/**
 * OpenAI APIを呼び出し
 */
function callOpenAIAPI(payload) {
  // Google Apps Script用同期版
  for (let i = 0; i < 3; i++) {
    try {
      const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
      if (!apiKey) {
        throw new Error('OpenAI APIキーが設定されていません');
      }
      
      const options = {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload)
      };
      
      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        throw new Error(`OpenAI API Error (${response.getResponseCode()}): ${errorText}`);
      }
      
      const data = JSON.parse(response.getContentText());
      return data.choices[0].message.content;
      
    } catch (error) {
      const errorStr = error.toString();
      
      // API制限エラーの場合は指数バックオフで再試行
      if (errorStr.includes('quota') || errorStr.includes('limit') || errorStr.includes('rate_limit')) {
        const waitTime = Math.pow(2, i) * 1000; // 1秒, 2秒, 4秒
        console.log(`API制限検出。${waitTime}ms待機後に再試行します...`);
        Utilities.sleep(waitTime);
        continue;
      }
      
      // その他のエラーは即座にスロー
      throw error;
    }
  }
  
  throw new Error('API制限により処理を中止しました。時間をおいて再実行してください。');
}

/**
 * API呼び出しのリトライ機能
 */
function apiCallWithRetry(apiFunction, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return apiFunction();
    } catch (error) {
      const errorStr = error.toString();
      
      // API制限エラーの場合は指数バックオフで再試行
      if (errorStr.includes('quota') || errorStr.includes('limit') || errorStr.includes('rate_limit')) {
        const waitTime = Math.pow(2, i) * 1000; // 1秒, 2秒, 4秒
        console.log(`API制限検出。${waitTime}ms待機後に再試行します...`);
        Utilities.sleep(waitTime);
        continue;
      }
      
      // その他のエラーは即座にスロー
      throw error;
    }
  }
  
  throw new Error('API制限により処理を中止しました。時間をおいて再実行してください。');
}

/**
 * 企業検索機能（メニューから呼ばれる場合）
 */
function executeCompanySearch() {
  try {
    console.log('🏢 企業検索を開始します...');
    
    // ユーザーに明確な開始通知
    SpreadsheetApp.getUi().alert(
      '🏢 企業検索開始', 
      'executeCompanySearch関数が正常に呼び出されました！\n\n' +
      'これから実際の検索処理に移行します。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // companies.jsの実際の関数を呼び出す
    executeCompanySearchFromCompanies();
    
  } catch (error) {
    console.error('❌ 企業検索エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `企業検索でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * キーワードレスポンスのパース（企業直接検索対応版）
 */
function parseKeywordResponse(response) {
  try {
    console.log('パース開始 - レスポンス:', response);
    
    // JSON部分を抽出（markdown形式の場合に対応）
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.error('JSON形式が見つかりません。レスポンス:', response);
      throw new Error('有効なJSONレスポンスが見つかりません');
    }
    
    console.log('JSON抽出成功:', jsonMatch[0]);
    
    const data = JSON.parse(jsonMatch[0]);
    console.log('JSONパース成功:', data);
    
    // 新形式 (companySearchKeywords) と旧形式 (strategicKeywords) 両方に対応
    const keywordData = data.companySearchKeywords || data.strategicKeywords;
    console.log('keywordData:', keywordData);
    
    if (!keywordData) {
      console.error('キーワードデータが見つかりません。data:', data);
      throw new Error('レスポンスにキーワードデータが含まれていません');
    }
    
    const allKeywords = [];
    
    // 各カテゴリのキーワードを配列に変換
    Object.keys(keywordData).forEach(category => {
      const categoryData = keywordData[category];
      console.log(`カテゴリ ${category}:`, categoryData);
      
      const priority = getPriorityByCategory(category);
      
      if (!categoryData || !categoryData.keywords || !Array.isArray(categoryData.keywords)) {
        console.error(`カテゴリ ${category} のキーワードが配列ではありません:`, categoryData);
        return; // このカテゴリをスキップ
      }
      
      categoryData.keywords.forEach(keyword => {
        // 除外ワードが含まれているかチェック
        const hasExcludeTerms = keyword.includes('-wikipedia') || 
                              keyword.includes('-facebook') || 
                              keyword.includes('-求人') ||
                              keyword.includes('-採用') ||
                              keyword.includes('-ニュース') ||
                              keyword.includes('-記事');
        
        allKeywords.push({
          keyword: keyword,
          category: category,
          priority: priority,
          strategy: categoryData.purpose || categoryData.strategy || '企業直接検索',
          executed: false,
          hitCount: 0,
          lastExecuted: '',
          hasExcludeTerms: hasExcludeTerms // 除外ワード有無をフラグ化
        });
      });
    });
    
    console.log('最終的な配列:', allKeywords);
    console.log('配列の長さ:', allKeywords.length);
    
    // 除外ワードを含むキーワードの割合をチェック
    const keywordsWithExclude = allKeywords.filter(kw => kw.hasExcludeTerms);
    console.log(`除外ワード付きキーワード: ${keywordsWithExclude.length}/${allKeywords.length}`);
    
    if (allKeywords.length === 0) {
      throw new Error('キーワードを抽出できませんでした');
    }
    
    return allKeywords;
    
  } catch (error) {
    console.error('キーワードパースエラー:', error);
    console.error('レスポンス:', response);
    throw new Error(`キーワードパースエラー: ${error.toString()}`);
  }
}

/**
 * カテゴリ別優先度の決定（効率最適化対応版）
 */
function getPriorityByCategory(category) {
  const priorityMap = {
    // 新しい効率重視カテゴリ
    'premiumSearch': '最高',     // 商工会・協会系
    'industrySpecific': '高',    // 業界特化
    'regionSpecific': '高',      // 地域特化
    'efficientSearch': '中',     // 効率的検索
    
    // 既存カテゴリ
    'industryDirect': '高',
    'companyPages': '高',
    'businessAssociations': '中',
    'locationBased': '中',
    'directSearch': '中',
    
    // 旧カテゴリ（下位互換）
    'painPointHunting': '低', // 記事系なので優先度下げ
    'timingCapture': '低',    // 記事系なので優先度下げ
    'growthTargeting': '中',
    'budgetTargeting': '中'
  };
  
  return priorityMap[category] || '低';
}

/**
 * キーワードをスプレッドシートに保存
 */
function saveKeywordsToSheet(keywords) {
  const sheet = getSafeSheet(SHEET_NAMES.KEYWORDS);
  if (!sheet) {
    throw new Error('キーワードシートが見つかりません。システム初期化を実行してください。');
  }
  
  // 既存データのクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  
  // 新しいデータの挿入
  const dataToInsert = keywords.map(kw => [
    kw.keyword,
    kw.category,
    kw.priority,
    kw.strategy,
    kw.executed,
    kw.hitCount,
    kw.lastExecuted
  ]);
  
  if (dataToInsert.length > 0) {
    sheet.getRange(2, 1, dataToInsert.length, 7).setValues(dataToInsert);
  }
  
  console.log(`${keywords.length}個のキーワードを保存しました`);
}

/**
 * 提案生成（メニュー呼び出し用）
 */
function generatePersonalizedProposals() {
  try {
    console.log('💬 提案生成を開始します...');
    updateExecutionStatus('提案メッセージ生成を開始します...');
    
    // APIキーの確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      SpreadsheetApp.getUi().alert('❌ エラー', 'OpenAI APIキーが設定されていません。\n\nメニュー > ⚙️ システム管理 > ⚙️ API設定管理 から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // proposals.jsの詳細ログ付き提案生成機能を呼び出し
    executeProposalGenerationEnhanced();
    
    console.log('✅ 提案生成処理完了');
    
  } catch (error) {
    console.error('❌ 提案生成エラー:', error);
    updateExecutionStatus(`エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert('❌ エラー', `提案生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 完全自動化実行機能はworkflow.jsファイルで実装されています