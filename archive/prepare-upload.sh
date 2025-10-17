#!/bin/bash

# Google Apps Script アップロード準備スクリプト
# 各ファイルの内容をコピペしやすい形で出力

echo "=== Google Apps Script アップロード用ファイル出力 ==="
echo ""

# メインファイル（必須）
echo "【1. menu.gs - メインメニュー】"
echo "----------------------------------------"
cat menu.gs
echo ""
echo ""

echo "【2. appsscript.json - プロジェクト設定】"
echo "----------------------------------------"
cat appsscript.json
echo ""
echo ""

# 主要機能ファイル
echo "【3. companies.gs - 企業管理（オプション）】"
echo "----------------------------------------"
echo "# ファイルサイズ: $(wc -c < companies.gs) バイト"
echo "# 内容は companies.gs を参照してください"
echo ""

echo "【4. proposals.gs - 提案書作成（オプション）】"
echo "----------------------------------------"
echo "# ファイルサイズ: $(wc -c < proposals.gs) バイト"
echo "# 内容は proposals.gs を参照してください"
echo ""

echo "=== アップロード完了 ==="
echo "1. script.google.com を開く"
echo "2. 新しいプロジェクトを作成"
echo "3. 上記のファイル内容をコピペ"
echo "4. スプレッドシートで動作確認"
