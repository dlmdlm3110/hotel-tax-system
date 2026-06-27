#!/bin/bash
# 宿泊税計算システム — 初回セットアップ
APP_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$APP_DIR"

echo "================================================"
echo "  宿泊税計算システム セットアップ"
echo "================================================"
echo ""

# Python3 チェック
if ! command -v python3 &>/dev/null; then
    echo "❌ Python3 が見つかりません。"
    echo "   https://www.python.org/downloads/ からインストールしてください。"
    read -p "Enterキーで終了..."
    exit 1
fi

echo "✅ Python3: $(python3 --version)"
echo ""
echo "パッケージをインストール中..."
python3 -m pip install --upgrade pip --quiet
python3 -m pip install -r requirements.txt

# launch.command に実行権限を付与（ZIP展開後に消えることがある）
chmod +x "$APP_DIR/launch.command" 2>/dev/null
echo "✅ 起動スクリプトの権限を設定しました。"

# ── .icns アイコン作成（sips + iconutil はmacOS標準）──
ICNS_PATH="/tmp/AppIcon_hoteltax.icns"
if [ -f "$APP_DIR/宿泊税スマートロゴ.png" ]; then
    ICONSET="/tmp/AppIcon_hoteltax.iconset"
    mkdir -p "$ICONSET"
    sips -z 16   16   "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_16x16.png"      &>/dev/null
    sips -z 32   32   "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_16x16@2x.png"   &>/dev/null
    sips -z 32   32   "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_32x32.png"      &>/dev/null
    sips -z 64   64   "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_32x32@2x.png"   &>/dev/null
    sips -z 128  128  "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_128x128.png"    &>/dev/null
    sips -z 256  256  "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_128x128@2x.png" &>/dev/null
    sips -z 256  256  "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_256x256.png"    &>/dev/null
    sips -z 512  512  "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_256x256@2x.png" &>/dev/null
    sips -z 512  512  "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_512x512.png"    &>/dev/null
    sips -z 1024 1024 "$APP_DIR/宿泊税スマートロゴ.png" --out "$ICONSET/icon_512x512@2x.png" &>/dev/null
    iconutil -c icns "$ICONSET" -o "$ICNS_PATH" 2>/dev/null
    rm -rf "$ICONSET"
fi

# ── デスクトップに .app を作成 ──
echo ""
echo "デスクトップにアプリアイコンを作成中..."

APP_PATH="$HOME/Desktop/宿泊税計算システム.app"

# 既存の .app を削除して再作成
rm -rf "$APP_PATH"

cat > /tmp/_hoteltax_launcher.applescript << APPLEEOF
on run
    tell application "Terminal"
        activate
        do script "bash '${APP_DIR}/launch.command'"
    end tell
end run
APPLEEOF

osacompile -o "$APP_PATH" /tmp/_hoteltax_launcher.applescript 2>/dev/null
rm -f /tmp/_hoteltax_launcher.applescript

# カスタムアイコンを適用
if [ -d "$APP_PATH" ] && [ -f "$ICNS_PATH" ]; then
    cp "$ICNS_PATH" "$APP_PATH/Contents/Resources/AppIcon.icns"
    /usr/libexec/PlistBuddy -c "Set :CFBundleIconFile AppIcon" \
        "$APP_PATH/Contents/Info.plist" 2>/dev/null
    touch "$APP_PATH"
fi
rm -f "$ICNS_PATH"

if [ -d "$APP_PATH" ]; then
    echo "✅ デスクトップに「宿泊税計算システム」アイコンを作成しました。"
else
    echo "⚠️  アイコン作成をスキップしました（launch.command から起動してください）。"
fi

echo ""
echo "================================================"
echo "  セットアップ完了！"
echo "  デスクトップの「宿泊税計算システム」を"
echo "  ダブルクリックしてアプリを起動してください。"
echo "================================================"
read -p "Enterキーで閉じる..."
