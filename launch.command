#!/bin/bash
cd "$(dirname "$0")"

echo "================================================"
echo "  宿泊税計算システム 起動中..."
echo "  ブラウザが開くまで少し待ってください（約5秒）"
echo "  開かない場合は手動でアクセス:"
echo "  http://localhost:8501"
echo "  終了: このウィンドウを閉じる"
echo "================================================"
echo ""

# Homebrew（Intel / Apple Silicon 両対応）やpyenvのPATHを追加
export PATH="/opt/homebrew/bin:/opt/homebrew/sbin:/usr/local/bin:/usr/local/sbin:$HOME/.pyenv/shims:$PATH"

# python3 を探す（複数の場所を順番に確認）
PYTHON3=""
for p in \
    python3 \
    /opt/homebrew/bin/python3 \
    /usr/local/bin/python3 \
    /usr/bin/python3 \
    "$HOME/.pyenv/shims/python3"
do
    if "$p" --version &>/dev/null 2>&1; then
        PYTHON3="$p"
        break
    fi
done

if [ -z "$PYTHON3" ]; then
    echo "================================================"
    echo "  [エラー] Python3 が見つかりません。"
    echo ""
    echo "  setup.command を先に実行してください。"
    echo "  それでも解決しない場合は:"
    echo "  https://www.python.org/downloads/ から"
    echo "  Python をインストールしてください。"
    echo "================================================"
    read -p "Enterキーで終了..."
    exit 1
fi

echo "Python3: $($PYTHON3 --version) ($PYTHON3)"
echo ""

# streamlit がインストール済みか確認
if ! $PYTHON3 -m streamlit --version &>/dev/null 2>&1; then
    echo "================================================"
    echo "  [エラー] streamlit が見つかりません。"
    echo "  setup.command を再実行してください。"
    echo "================================================"
    read -p "Enterキーで終了..."
    exit 1
fi

# ポート8501が使用中なら既存プロセスを終了して再起動
if lsof -ti:8501 &>/dev/null 2>&1; then
    echo "  ポート8501が使用中です。既存のプロセスを終了します..."
    lsof -ti:8501 | xargs kill -9 2>/dev/null
    sleep 2
    echo "  既存プロセスを終了しました。"
    echo ""
fi

# 5秒後にブラウザを開く（バックグラウンド）
(sleep 5 && open "http://localhost:8501") &

# Streamlit サーバー起動
$PYTHON3 -m streamlit run app.py \
    --server.headless true \
    --browser.gatherUsageStats false \
    --server.port 8501

echo ""
echo "================================================"
echo "  [エラー] サーバーが予期せず終了しました。"
echo "  上記のエラーメッセージを確認してください。"
echo "================================================"
read -p "Enterキーで閉じる..."
