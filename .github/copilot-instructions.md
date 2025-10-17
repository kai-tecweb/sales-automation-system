# Copilot Instructions - 営業自動化システム

## 🎯 システム全体像

この **Google Apps Script (GAS) ベースの営業自動化システム** は、商材キーワードから個別最適化された提案まで、営業パイプライン全体を自動化します。すべての処理はGoogle Sheets内で外部API連携により実行されます。

### 主要コンポーネントとデータフロー
```
商材入力 → キーワード生成 (OpenAI) → 企業検索 (Google) → スコアリング → 提案生成 (OpenAI)
```

**重要なアーキテクチャファイル:**
- `src/main.js` - システム初期化とシート管理
- `src/workflow.js` - 完全自動化パイプラインの制御
- `src/globals.js` - 中央定数と安全なシート操作
- `src/menu.js` - 管理者/ユーザーモード対応の完全UIメニュー

## 🔧 必須開発パターン

### 1. シート中心アーキテクチャ
すべてのデータは事前定義されたGoogle Sheetsを通して流れます。**必ず `SHEET_NAMES` 定数を使用:**
```javascript
const SHEET_NAMES = {
  CONTROL: '制御パネル',
  KEYWORDS: '生成キーワード', 
  COMPANIES: '企業マスター',
  PROPOSALS: '提案メッセージ',
  LOGS: '実行ログ'
};
```

### 2. 指数バックオフ付きAPI呼び出しパターン
**外部API呼び出しには必ず `apiCallWithRetry()` を使用:**
```javascript
return apiCallWithRetry(() => {
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`API Error: ${response.getResponseCode()}`);
  }
  return JSON.parse(response.getContentText());
});
```

### 3. 安全なシート操作
**直接シートアクセス禁止 - 安全なゲッターを使用:**
```javascript
const sheet = getSafeSheet(SHEET_NAMES.COMPANIES);
// 禁止: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('企業マスター')
```

### 4. デュアルAPI統合パターン
システムは **OpenAI ChatGPT API**（キーワード、分析、提案）と **Google Custom Search API**（企業発掘）を統合。両方ともレート制限があり、`PropertiesService` からのAPIキーが必要。

## 🔑 必須設定

### APIキー管理
APIキーはGoogle Apps Scriptプロパティに保存:
- `OPENAI_API_KEY` - ChatGPT API呼び出し用
- `GOOGLE_SEARCH_API_KEY` - Custom Search API用
- `GOOGLE_SEARCH_ENGINE_ID` - カスタム検索エンジン識別子

### スコアリングアルゴリズム
企業は0-100点でスコア付け:
- 企業規模適合性 (±20点) - 会社規模マッチング
- 業界適合性 (±15点) - 業界関連性
- 問い合わせ方法 (±10点) - 連絡手段の有無
- 成長性 (±10点) - 成長指標
- 情報充実度 (±5点) - 情報完全性

## 🚀 開発ワークフロー

### claspローカル開発
```bash
npm install @google/clasp
clasp login
clasp push  # GASにデプロイ
```

### システム初期化
**コードデプロイ後は必ず `initializeSheets()` を実行** して必要なシート構造を作成。

### テストとデバッグ
- システム検証には `healthCheck()` を使用
- `performHealthCheck()` でAPI接続をテスト
- すべての実行は「実行ログ」シートに記録
- 管理者インターフェースからデバッグモード利用可能

### 主要実行エントリーポイント
- `executeFullWorkflow()` - 完全自動化パイプライン
- `generateStrategicKeywords()` - キーワード生成のみ
- `executeCompanySearch()` - 企業発掘のみ
- `generatePersonalizedProposals()` - 提案生成のみ

## 🛡️ エラーハンドリングとデータ品質

### 必須データ検証
処理前に必ず `validateCompanyData()` で企業データを検証。システムが期待するもの:
- `companyName` と `industry` が必須項目
- 無効な結果の自動フィルタリング（404ページ、エラーメッセージ）
- `normalizeJapaneseText()` による日本語テキスト正規化

### ライセンス管理
営業日ベースのライセンスシステム（`license-manager.js`）と管理者認証を含む。管理者機能にはパスワード認証が必要で、高度なシステム管理を可能にする。

## 🔍 プロジェクト固有の慣例

### ファイル命名と構成
- `main.js` - 中核システム機能、シート初期化
- `[機能名].js` - 機能別モジュール（keywords, companies, proposalsなど）
- `enhanced` サフィックス - 高度機能実装
- `/docs/specifications/` に包括的な日本語ドキュメント

### メニューシステムアーキテクチャ
ユーザー権限による2階層メニューシステム:
- 標準ユーザー: 基本ワークフロー機能
- 管理者モード: 高度設定、システムメンテナンス、ライセンス管理

### 日本語ファースト設計
- すべてのUIテキスト、シート名、エラーメッセージが日本語
- 全角文字と日本語テキスト正規化の適切な処理
- 日本のビジネス文脈（営業日ベースライセンス、企業規模分類）

## ⚡ 主要統合ポイント

### 外部依存関係
- Google Apps Script V8 ランタイム環境
- Google Sheets API（高度サービス有効化済み）
- OpenAI ChatGPT API（gpt-3.5-turboモデル）
- プログラマブル検索エンジン付きGoogle Custom Search API

### コンポーネント間通信
コンポーネントは直接関数呼び出しではなく、共有Google Sheetsステートを通じて通信。長時間操作中のユーザーフィードバックには必ず `updateExecutionStatus()` で進捗更新。