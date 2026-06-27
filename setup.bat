@echo off
chcp 65001 > nul
cd /d "%~dp0"
set APP_DIR=%~dp0

echo ================================================
echo   宿泊税計算システム セットアップ
echo ================================================
echo.

:: ── ダウンロードブロックを解除（SmartScreen対策）──
echo セキュリティブロックを解除中...
powershell -NoProfile -Command "Get-ChildItem -Path '%APP_DIR%' | Unblock-File -ErrorAction SilentlyContinue"
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] Python が見つかりません。
    echo.
    echo セットアップ手順_Windows.txt を参照して
    echo Python をインストールしてから再実行してください。
    echo.
    pause
    exit /b 1
)

echo Python バージョン:
python --version
echo.
echo パッケージをインストール中...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt

:: ── .ico アイコン作成 ──
if exist "%APP_DIR%宿泊税スマートロゴ.png" (
    echo.
    echo アイコンファイルを作成中...
    python "%APP_DIR%_make_icon.py"
)

:: ── デスクトップにショートカットを作成 ──
echo.
echo デスクトップにアプリアイコンを作成中...
python "%APP_DIR%create_shortcut.py"
if errorlevel 1 (
    echo [警告] ショートカット作成に失敗しました。
    echo        launch.bat をダブルクリックして起動してください。
) else (
    echo デスクトップの「宿泊税計算システム」をダブルクリックして起動してください。
)

echo.
echo ================================================
echo   セットアップ完了！
echo ================================================
pause
