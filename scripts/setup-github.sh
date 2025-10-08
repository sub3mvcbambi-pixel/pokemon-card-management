#!/bin/bash

# GitHub リポジトリのセットアップスクリプト

echo "🚀 ポケモンカード販売管理システム - GitHub セットアップ"
echo "=================================================="

# GitHub ユーザー名の入力
read -p "GitHub ユーザー名を入力してください: " GITHUB_USERNAME

if [ -z "$GITHUB_USERNAME" ]; then
    echo "❌ エラー: GitHub ユーザー名が入力されていません"
    exit 1
fi

# リポジトリ名
REPO_NAME="pokemon-card-management"

echo "📁 リポジトリを作成しています..."

# GitHub CLI がインストールされているかチェック
if ! command -v gh &> /dev/null; then
    echo "❌ GitHub CLI (gh) がインストールされていません"
    echo "   インストール方法: https://cli.github.com/"
    exit 1
fi

# GitHub にログインしているかチェック
if ! gh auth status &> /dev/null; then
    echo "🔐 GitHub にログインしてください"
    gh auth login
fi

# リポジトリを作成
echo "📦 リポジトリを作成中..."
gh repo create $REPO_NAME --public --description "ポケモンカード販売管理システム - Google Apps Script版" --clone=false

# リモートリポジトリを追加
echo "🔗 リモートリポジトリを追加中..."
git remote add origin https://github.com/$GITHUB_USERNAME/$REPO_NAME.git

# ブランチを main に設定
echo "🌿 ブランチを main に設定中..."
git branch -M main

# 初回プッシュ
echo "📤 初回プッシュを実行中..."
git push -u origin main

echo ""
echo "✅ GitHub リポジトリのセットアップが完了しました！"
echo ""
echo "🔧 次のステップ:"
echo "1. https://github.com/$GITHUB_USERNAME/$REPO_NAME にアクセス"
echo "2. Settings → Secrets and variables → Actions で以下を設定:"
echo "   - GAS_SCRIPT_ID: Google Apps Script のスクリプトID"
echo "   - GAS_CREDENTIALS: サービスアカウントのJSONファイルの内容"
echo "3. DEPLOYMENT.md の手順に従って Google Cloud Console を設定"
echo ""
echo "🎉 セットアップ完了！"