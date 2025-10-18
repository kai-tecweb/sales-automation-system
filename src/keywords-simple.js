/**
 * シンプルで確実なキーワード生成関数（デバッグ用）
 */
function generateKeywordsSimple() {
  try {
    console.log('=== シンプルキーワード生成開始 ===');
    
    // 制御パネルから設定を取得
    const sheet = getSafeSheet(SHEET_NAMES.CONTROL);
    const data = sheet.getDataRange().getValues();
    
    const settings = {
      productName: data[1][1] || '',
      productDescription: data[2][1] || '',
      priceRange: data[3][1] || '',
      targetSize: data[4][1] || ''
    };
    
    console.log('設定取得完了:', settings);
    
    // APIキー確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      throw new Error('OpenAI APIキーが設定されていません');
    }
    
    // シンプルなプロンプト
    const prompt = `商材「${settings.productName}」（${settings.productDescription}）に興味がありそうな企業を見つけるための検索キーワードを5個生成してください。JSON形式で以下のように返してください：
    
{
  "keywords": ["キーワード1", "キーワード2", "キーワード3", "キーワード4", "キーワード5"]
}`;

    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 500,
      temperature: 0.7
    };
    
    // 直接API呼び出し
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    console.log('API呼び出し開始...');
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`API Error: ${response.getContentText()}`);
    }
    
    const data_response = JSON.parse(response.getContentText());
    const messageContent = data_response.choices[0].message.content;
    
    console.log('APIレスポンス:', messageContent);
    
    // JSONパース
    const jsonMatch = messageContent.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('有効なJSONが見つかりません');
    }
    
    const keywordData = JSON.parse(jsonMatch[0]);
    const keywords = keywordData.keywords;
    
    if (!Array.isArray(keywords)) {
      throw new Error('キーワードが配列形式ではありません');
    }
    
    console.log('生成されたキーワード:', keywords);
    
    // シンプルな形式でシートに保存
    const keywordSheet = getSafeSheet(SHEET_NAMES.KEYWORDS);
    
    // ヘッダーがない場合は追加
    if (keywordSheet.getLastRow() === 0) {
      keywordSheet.getRange(1, 1, 1, 7).setValues([['キーワード', 'カテゴリ', '優先度', '実行状況', '発見企業数', '実行日時', '備考']]);
    }
    
    // キーワードを追加
    const newRows = keywords.map((keyword, index) => [
      keyword,
      'シンプル生成',
      'medium',
      false,
      0,
      '',
      `シンプル生成 ${new Date().toLocaleString()}`
    ]);
    
    const lastRow = keywordSheet.getLastRow();
    keywordSheet.getRange(lastRow + 1, 1, newRows.length, 7).setValues(newRows);
    
    const successMessage = `✅ シンプルキーワード生成完了！\n\n${keywords.length}個のキーワードを生成しました:\n${keywords.join('\n')}\n\n「生成キーワード」シートで確認できます。`;
    
    SpreadsheetApp.getUi().alert('キーワード生成完了', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log('✅ シンプルキーワード生成完了');
    
  } catch (error) {
    console.error('❌ シンプルキーワード生成エラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `シンプルキーワード生成でエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}