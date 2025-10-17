/**
 * マッチ度スコア計算機能
 * 企業と商材の適合性を数値化
 */

/**
 * マッチ度スコアの計算
 */
function calculateMatchScore(companyData, settings) {
  let baseScore = 50;
  
  try {
    // 企業規模の適合性（±20点）
    baseScore += calculateSizeScore(companyData.companySize, settings.targetSize);
    
    // 業界適合性（±15点）
    baseScore += calculateIndustryScore(companyData.industry, settings);
    
    // 問い合わせ方法（±10点）
    baseScore += calculateContactScore(companyData.contactMethod);
    
    // 成長性・安定性（±10点）
    baseScore += calculateGrowthScore(companyData);
    
    // 情報の充実度（±5点）
    baseScore += calculateInfoScore(companyData);
    
    // スコアを0-100の範囲に正規化
    return Math.max(0, Math.min(100, Math.round(baseScore)));
    
  } catch (error) {
    Logger.log(`スコア計算エラー: ${error.toString()}`);
    return 50; // デフォルトスコア
  }
}

/**
 * 企業規模スコアの計算
 */
function calculateSizeScore(companySize, targetSize) {
  const sizeMapping = {
    '個人事業主': 1,
    'スタートアップ': 2,
    '中小企業': 3,
    '大企業': 4
  };
  
  const companySizeValue = sizeMapping[companySize] || 3;
  const targetSizeValue = sizeMapping[targetSize] || 3;
  
  if (targetSize === 'すべて') {
    return 10; // 中立的なボーナス
  }
  
  const difference = Math.abs(companySizeValue - targetSizeValue);
  
  switch (difference) {
    case 0: return 20; // 完全一致
    case 1: return 10; // 近い
    case 2: return 0;  // 普通
    case 3: return -10; // 遠い
    default: return -15;
  }
}

/**
 * 業界適合性スコアの計算
 */
function calculateIndustryScore(industry, settings) {
  // 商材に適合しやすい業界を定義
  const industryFitMap = {
    'IT': { high: ['SaaS', 'システム', 'AI', 'DX'], medium: ['コンサル', 'マーケティング'], low: [] },
    '製造業': { high: ['IoT', 'システム', '効率化'], medium: ['AI', 'DX'], low: ['マーケティング'] },
    'サービス業': { high: ['効率化', 'システム', 'マーケティング'], medium: ['AI', 'SaaS'], low: [] },
    '医療': { high: ['システム', 'AI'], medium: ['効率化'], low: ['マーケティング'] },
    '教育': { high: ['システム', 'AI', 'SaaS'], medium: ['効率化'], low: [] },
    '金融': { high: ['システム', 'AI', 'セキュリティ'], medium: ['DX'], low: [] },
    '不動産': { high: ['システム', 'マーケティング'], medium: ['AI', 'DX'], low: [] }
  };
  
  const productKeywords = extractProductKeywords(settings);
  const industryFit = industryFitMap[industry] || { high: [], medium: [], low: [] };
  
  // 商材キーワードと業界の適合性をチェック
  const hasHighFit = productKeywords.some(keyword => 
    industryFit.high.some(fit => keyword.includes(fit) || fit.includes(keyword))
  );
  
  const hasMediumFit = productKeywords.some(keyword => 
    industryFit.medium.some(fit => keyword.includes(fit) || fit.includes(keyword))
  );
  
  const hasLowFit = productKeywords.some(keyword => 
    industryFit.low.some(fit => keyword.includes(fit) || fit.includes(keyword))
  );
  
  if (hasHighFit) return 15;
  if (hasMediumFit) return 5;
  if (hasLowFit) return -5;
  
  return 0; // 中立
}

/**
 * 商材からキーワードを抽出
 */
function extractProductKeywords(settings) {
  const text = `${settings.productName} ${settings.productDescription}`.toLowerCase();
  return text.split(/[\s、。・]+/).filter(word => word.length > 1);
}

/**
 * 問い合わせ方法スコアの計算
 */
function calculateContactScore(contactMethod) {
  const contactScores = {
    'フォーム・電話': 10,
    'フォーム': 8,
    '電話': 6,
    'メール': 4,
    'その他': 2
  };
  
  return contactScores[contactMethod] || 0;
}

/**
 * 成長性スコアの計算
 */
function calculateGrowthScore(companyData) {
  let score = 0;
  
  // 事業内容から成長性キーワードを検出
  const growthKeywords = [
    'DX', 'デジタル変革', 'AI', '人工知能', '機械学習', 'IoT',
    '新規事業', 'スタートアップ', '急成長', '拡大', '展開',
    'イノベーション', '革新', '最新', '先進', '次世代'
  ];
  
  const businessText = companyData.businessDescription || '';
  const webFeatures = companyData.webFeatures || '';
  const allText = `${businessText} ${webFeatures}`.toLowerCase();
  
  const foundKeywords = growthKeywords.filter(keyword => 
    allText.includes(keyword.toLowerCase())
  );
  
  // 成長キーワードの数に応じてスコア付与
  if (foundKeywords.length >= 3) {
    score += 10; // 高成長
  } else if (foundKeywords.length >= 1) {
    score += 5; // 中成長
  }
  
  // 企業規模による調整
  if (companyData.companySize === 'スタートアップ') {
    score += 5; // スタートアップボーナス
  } else if (companyData.companySize === '大企業' && foundKeywords.length === 0) {
    score -= 3; // 大企業で成長キーワードなしの場合
  }
  
  return score;
}

/**
 * 情報充実度スコアの計算
 */
function calculateInfoScore(companyData) {
  let score = 0;
  
  // 各項目の情報充実度をチェック
  const infoChecks = [
    { field: 'employees', weight: 1 },
    { field: 'location', weight: 1 },
    { field: 'businessDescription', weight: 2 },
    { field: 'webFeatures', weight: 1 }
  ];
  
  infoChecks.forEach(check => {
    const value = companyData[check.field];
    if (value && value !== '' && value !== null) {
      if (typeof value === 'string' && value.length > 10) {
        score += check.weight; // 充実した情報
      } else if (value) {
        score += check.weight * 0.5; // 基本的な情報
      }
    }
  });
  
  return Math.min(5, score); // 最大5点
}

/**
 * スコア分析結果の取得
 */
function getScoreAnalysis(companyData, settings) {
  const scores = {
    size: calculateSizeScore(companyData.companySize, settings.targetSize),
    industry: calculateIndustryScore(companyData.industry, settings),
    contact: calculateContactScore(companyData.contactMethod),
    growth: calculateGrowthScore(companyData),
    info: calculateInfoScore(companyData)
  };
  
  const totalScore = 50 + Object.values(scores).reduce((sum, score) => sum + score, 0);
  const finalScore = Math.max(0, Math.min(100, Math.round(totalScore)));
  
  return {
    finalScore,
    breakdown: scores,
    analysis: generateScoreAnalysis(scores, companyData, settings)
  };
}

/**
 * スコア分析コメントの生成
 */
function generateScoreAnalysis(scores, companyData, settings) {
  const comments = [];
  
  // 企業規模の分析
  if (scores.size >= 15) {
    comments.push(`対象企業規模（${settings.targetSize}）と完全に一致`);
  } else if (scores.size >= 5) {
    comments.push(`対象企業規模にほぼ適合`);
  } else if (scores.size < 0) {
    comments.push(`企業規模（${companyData.companySize}）が対象と異なる`);
  }
  
  // 業界の分析
  if (scores.industry >= 10) {
    comments.push(`${companyData.industry}業界は商材と高い親和性`);
  } else if (scores.industry >= 0) {
    comments.push(`${companyData.industry}業界は商材と一定の親和性`);
  } else {
    comments.push(`${companyData.industry}業界は商材との親和性が低い可能性`);
  }
  
  // 問い合わせ方法の分析
  if (scores.contact >= 8) {
    comments.push(`問い合わせ手段（${companyData.contactMethod}）が充実`);
  } else if (scores.contact >= 4) {
    comments.push(`基本的な問い合わせ手段を確保`);
  }
  
  // 成長性の分析
  if (scores.growth >= 8) {
    comments.push('高い成長性を示すキーワードを検出');
  } else if (scores.growth >= 3) {
    comments.push('一定の成長性を示す要素を確認');
  }
  
  // 情報充実度の分析
  if (scores.info >= 4) {
    comments.push('企業情報が充実している');
  } else if (scores.info >= 2) {
    comments.push('基本的な企業情報を確認');
  } else {
    comments.push('企業情報がやや不足');
  }
  
  return comments.join('。') + '。';
}

/**
 * 高スコア企業の取得
 */
function getHighScoreCompanies(minScore = 70, maxCount = 20) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANIES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  
  return data
    .map((row, index) => ({
      rowIndex: index + 2,
      companyId: row[0],
      companyName: row[1],
      officialUrl: row[2],
      industry: row[3],
      employees: row[4],
      location: row[5],
      contactMethod: row[6],
      isPublicCompany: row[7],
      businessDescription: row[8],
      companySize: row[9],
      matchScore: row[10],
      discoveryKeyword: row[11],
      registrationDate: row[12]
    }))
    .filter(company => company.matchScore >= minScore)
    .sort((a, b) => b.matchScore - a.matchScore)
    .slice(0, maxCount);
}

/**
 * 企業マッチング統計の取得
 */
function getMatchingStatistics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANIES);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return {
      totalCompanies: 0,
      averageScore: 0,
      highScoreCount: 0,
      mediumScoreCount: 0,
      lowScoreCount: 0,
      topIndustries: []
    };
  }
  
  const scores = sheet.getRange(2, 11, lastRow - 1, 1).getValues().flat().filter(score => score);
  const industries = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat().filter(industry => industry);
  
  const highScoreCount = scores.filter(score => score >= 80).length;
  const mediumScoreCount = scores.filter(score => score >= 60 && score < 80).length;
  const lowScoreCount = scores.filter(score => score < 60).length;
  
  // 業界別集計
  const industryCount = {};
  industries.forEach(industry => {
    industryCount[industry] = (industryCount[industry] || 0) + 1;
  });
  
  const topIndustries = Object.entries(industryCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([industry, count]) => ({ industry, count }));
  
  return {
    totalCompanies: scores.length,
    averageScore: scores.length > 0 ? Math.round(scores.reduce((sum, score) => sum + score, 0) / scores.length) : 0,
    highScoreCount,
    mediumScoreCount,
    lowScoreCount,
    topIndustries
  };
}
