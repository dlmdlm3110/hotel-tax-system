@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ================================================
echo   宿泊税計算システム 起動中...
echo.
echo   起動完了後、ブラウザが自動で開きます。
echo   開かない場合は手動でアクセスしてください：
echo   http://localhost:8501
echo.
echo   終了するにはこのウィンドウを × で閉じてください。
echo ================================================
echo.

:: ── ポート8501が使用中なら既存プロセスを終了 ──
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr " LISTENING" ^| findstr ":8501 "') do (
    echo ポート8501が使用中です。既存プロセス（PID:%%a）を終了します...
    taskkill /F /PID %%a > nul 2>&1
)
timeout /t 2 /nobreak > nul

:: ── 5秒後にブラウザを開く（サーバー起動を待つ）──
start "" cmd /c "timeout /t 5 /nobreak > nul && start http://localhost:8501"

:: ── Streamlit サーバー起動 ──
python -m streamlit run app.py --server.headless true --browser.gatherUsageStats false --server.port 8501

:: エラー終了した場合のみここに来る
echo.
echo ================================================
echo   [エラー] アプリの起動に失敗しました。
echo.
echo   以下を確認してください：
echo   1. setup.bat を先に実行しましたか？
echo   2. Python が正しくインストールされていますか？
echo      （コマンドプロンプトで「python --version」を確認）
echo ================================================
pause
