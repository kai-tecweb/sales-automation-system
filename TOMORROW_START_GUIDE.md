# 🌅 明日の作業再開ガイド - 2025年10月18日

## ✅ 昨日（2025/10/17）の完了事項
- ✅ API仕様書 v2.0 完成
- ✅ 料金プラン体系確立
- ✅ マルチアカウント運用システム設計
- ✅ Git管理開始・コミット完了
- ✅ Google Apps Script環境構築

## 🚀 今日の作業項目（優先順）

### 1. GitHubリポジトリ作成・プッシュ 【最優先】
```bash
cd /home/iwasaki/sales-automation-system

# GitHubでリポジトリ作成後：
git remote add origin https://github.com/[username]/sales-automation-system.git
git push -u origin main
```

### 2. メイン機能実装開始
- **ファイル**: `src/main.js`
- **内容**: 
  - APIレート制限システム実装
  - OpenAI ChatGPT API接続
  - Google Custom Search API接続
  - エラーハンドリング

### 3. API設定・認証
- OpenAI API Key取得・設定
- Google Cloud Console設定
- Custom Search Engine作成
- API Keys設定（PropertiesService）

### 4. テスト実行
- 基本機能動作確認
- レート制限動作確認
- エラーハンドリング確認

## 📁 重要ファイル位置

### 主要仕様書
- `docs/specifications/API仕様書-v2.0.md` ← 昨日の主要成果
- `docs/specifications/システム仕様書-v2.0.md`
- `docs/specifications/技術仕様書-v2.0.md`

### 実装対象
- `src/main.js` ← 今日の主要実装対象
- `src/menu.js`
- `src/utils.js`

### 設定ファイル
- `.clasp.json` ← Google Apps Script設定
- `package.json` ← Node.js依存関係

## 🔑 必要なAPI設定

### OpenAI API
1. https://openai.com/ でAPI Key取得
2. PropertiesServiceに `OPENAI_API_KEY` 設定

### Google APIs
1. Google Cloud Console でプロジェクト作成
2. Custom Search API有効化
3. Search Engine ID取得
4. PropertiesServiceに設定：
   - `GOOGLE_SEARCH_API_KEY`
   - `GOOGLE_SEARCH_ENGINE_ID`

## 💡 実装時の重要ポイント

### レート制限
- OpenAI: 20秒間隔（厳密）
- Google Search: 1.5秒間隔（安全マージン）

### エラーハンドリング
- API制限エラー
- ネットワークエラー
- 認証エラー

### マルチアカウント準備
- アカウント管理システム
- 負荷分散ロジック
- 協調動作システム

## 🛠️ 開発環境確認

### Clasp (Google Apps Script CLI)
```bash
cd /home/iwasaki/sales-automation-system
npx @google/clasp push    # GASにコード同期
npx @google/clasp open    # ブラウザで開く
```

### Git管理
```bash
git status              # 変更状況確認
git add .              # 変更追加
git commit -m "message" # コミット
git push               # プッシュ
```

## 📞 困った時の対処

### 1. Claspエラー
- 認証確認: `npx @google/clasp login`
- 設定確認: `.clasp.json`の内容確認

### 2. API接続エラー
- キー設定確認
- 仕様書参照: `docs/specifications/API仕様書-v2.0.md`

### 3. Git問題
- 状況確認: `git status`
- 作業ログ確認: `WORK_LOG.md`

---

**🎯 今日のゴール**: 基本API接続とレート制限システムの動作確認まで

**📋 昨日の作業内容**: `WORK_LOG.md` 参照

**📖 詳細仕様**: `docs/specifications/` フォルダ内各種仕様書を参照
