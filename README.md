# 営業自動化システム

商材起点企業発掘・提案自動生成システム

## 📁 プロジェクト構成

```
sales-automation-system/
├── src/                          # ソースコード
│   ├── main.js                   # メインファイル
│   ├── menu.js                   # カスタムメニュー
│   ├── keywords.js               # キーワード生成
│   ├── companies.js              # 企業管理
│   ├── proposals.js              # 提案生成
│   ├── proposals-enhanced.js     # 拡張提案機能
│   ├── help-manager.js           # ヘルプシステム
│   ├── scoring.js                # スコアリング
│   ├── utils.js                  # ユーティリティ
│   ├── workflow.js               # ワークフロー
│   ├── response-analyzer.js      # 応答解析
│   ├── license-manager.js        # ライセンス管理
│   ├── spreadsheet-config.js     # スプレッドシート設定
│   └── appsscript.json          # GAS設定
├── docs/                        # ドキュメント
│   ├── README.md                # ドキュメント構成
│   ├── specifications/          # 仕様書
│   ├── api/                     # API関連
│   └── user-guide/              # ユーザーガイド
├── archive/                     # 古いファイル
├── .clasp.json                  # Clasp設定
├── .claspignore                 # 除外ファイル設定
├── package.json                 # Node.js設定
└── README.md                    # このファイル
```

## 概要

このシステムは、Google Apps Script (GAS) を使用して開発された営業プロセス自動化ツールです。漠然とした商材情報から戦略的な企業検索キーワードを生成し、Google検索による企業情報の自動収集・分析、各企業に最適化された提案メッセージの自動生成までを一貫して行います。

## 主な機能

### 1. キーワード自動生成
- ChatGPT APIを使用した戦略的検索キーワードの生成
- 4つのカテゴリ（課題発見、成長企業、予算企業、タイミング）別の最適化

### 2. 企業検索・分析
- Google Custom Search APIによる企業検索
- ウェブサイト情報の自動抽出・分析
- マッチ度スコアリング（0-100点）

### 3. 提案メッセージ生成
- 企業ごとに個別最適化された提案文の生成
- A/Bパターンの提案（課題訴求型・成功事例型）
- 問い合わせフォーム用短縮メッセージ

### 4. 全自動実行
- キーワード生成→企業検索→提案生成の一貫実行
- 進捗状況の可視化
- 詳細な実行ログと統計情報

## ファイル構成

```
sales-automation-system/
├── appsscript.json       # GASプロジェクト設定
├── main.gs              # メイン機能・シート初期化
├── keywords.gs          # キーワード生成機能
├── companies.gs         # 企業検索・分析機能
├── scoring.gs           # マッチ度スコア計算
├── proposals.gs         # 提案メッセージ生成
├── workflow.gs          # 全自動実行制御
├── utils.gs             # ユーティリティ機能
└── README.md           # このファイル
```

## セットアップ手順

### 1. Google Apps Scriptプロジェクトの作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」を作成
3. プロジェクト名を「営業自動化システム」に変更

### 2. ファイルのアップロード

各`.gs`ファイルの内容をGoogle Apps Scriptエディタにコピー：

1. `main.gs` → `Code.gs`（デフォルトファイルを置き換え）
2. 他の`.gs`ファイルを新規作成してコピー

### 3. APIキーの取得

#### Google Custom Search API
1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクト作成
2. Custom Search API を有効化
3. APIキーを作成
4. [Programmable Search Engine](https://programmablesearchengine.google.com/) でカスタム検索エンジンを作成

#### OpenAI API
1. [OpenAI Platform](https://platform.openai.com/) でアカウント作成
2. APIキーを生成

### 4. APIキーの設定

システム初期化後、メニューから「APIキー設定」を選択して設定

### 5. システムの初期化

1. スプレッドシートを新規作成
2. 拡張機能 → Apps Script でスクリプトエディタを開く
3. 関数一覧から `initializeSheets` を実行

## 使用方法

### 基本的な流れ

1. **制御パネルで設定**
   - 商材名、商材概要を入力
   - 価格帯、対象企業規模を選択
   - 検索企業数上限を設定

2. **実行方法を選択**
   - 「全自動実行」：全工程を一括実行
   - 個別実行：キーワード生成→企業検索→提案生成を段階的に実行

3. **結果の確認**
   - 各シートで結果を確認
   - ダッシュボードで統計情報を確認

### データシート

- **制御パネル**：設定値、実行ボタン、ダッシュボード
- **生成キーワード**：戦略的検索キーワード一覧
- **企業マスター**：発見された企業情報とスコア
- **提案メッセージ**：個別最適化された提案文
- **実行ログ**：処理履歴とエラー情報

## 技術仕様

### 外部API

- **OpenAI ChatGPT API (gpt-3.5-turbo)**
  - キーワード生成
  - 企業情報分析
  - 提案メッセージ作成

- **Google Custom Search API**
  - 企業ウェブサイト検索
  - 検索結果の取得

### スコアリングアルゴリズム

マッチ度スコア（0-100点）は以下の要素で算出：

- 企業規模適合性（±20点）
- 業界適合性（±15点）
- 問い合わせ方法（±10点）
- 成長性（±10点）
- 情報充実度（±5点）

### エラーハンドリング

- API制限対応（指数バックオフ）
- データ品質チェック
- 重複データの除去
- 詳細なエラーログ

## 制限事項

### API制限
- Google Custom Search API：100回/日（無料版）
- OpenAI API：トークン数による課金
- Google Apps Script：実行時間6分制限

### データ制限
- スプレッドシート：最大1,000万セル
- 1回の実行：最大20社程度推奨

## トラブルシューティング

### よくある問題

1. **APIキーエラー**
   - 設定値を確認
   - API有効化状況を確認

2. **実行時間超過**
   - 検索企業数上限を減らす
   - 段階的実行を使用

3. **検索結果が少ない**
   - キーワード戦略を見直し
   - 価格帯・企業規模設定を調整

### デバッグ機能

- `performHealthCheck()`：システム状態確認
- `executeWorkflowStep()`：段階的実行
- ログファイルでの詳細追跡

## カスタマイズ

### スコアリング調整

`scoring.gs`の`calculateMatchScore()`関数で重み付けを調整可能

### 業界適合性調整

`scoring.gs`の`industryFitMap`で業界別適合度を調整可能

### プロンプト調整

各機能の`createPrompt()`系関数でChatGPTへの指示を調整可能

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## サポート

システムに関する質問や改善要望がある場合は、開発チームまでお問い合わせください。

---

**重要事項**
- APIキーは適切に管理し、外部に漏洩しないよう注意してください
- 検索対象企業の利用規約を遵守してください
- 生成されたメッセージは内容を確認してから使用してください
