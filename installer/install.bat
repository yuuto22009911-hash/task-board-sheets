@echo off
chcp 65001 > nul
setlocal

rem ============================================================
rem task-board-sheets Windows インストーラー
rem
rem 実行内容:
rem  1. ブラウザで Google スプレッドシート新規作成画面を開く
rem  2. ブラウザで Apps Script ソースコードを開く (コピー用)
rem  3. デスクトップに「タスク管理表」ショートカット作成
rem  4. 手順を表示
rem ============================================================

echo.
echo ============================================================
echo   task-board-sheets セットアップを開始します
echo ============================================================
echo.

set CODE_URL=https://raw.githubusercontent.com/yuuto22009911-hash/task-board-sheets/main/apps-script/Code.gs

echo  [1/3] 新規 Google スプレッドシートを開きます (ブラウザ起動)...
start "" "https://sheets.new"
timeout /t 3 /nobreak > nul

echo  [2/3] 別タブで Apps Script のソースコードを開きます...
start "" "%CODE_URL%"
timeout /t 2 /nobreak > nul

echo  [3/3] デスクトップに「タスク管理表」ショートカットを作成します...
powershell -ExecutionPolicy Bypass -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $sc = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\タスク管理表.url'); ^
   $sc.TargetPath = 'https://docs.google.com/spreadsheets'; ^
   $sc.Save(); ^
   Write-Host '  デスクトップに「タスク管理表.url」を作成しました'"

echo.
echo ============================================================
echo   ブラウザで開いた画面で以下の操作をしてください
echo ============================================================
echo.
echo  [A] 新規スプレッドシート (sheets.new で開いたタブ)
echo      1. 左上の「無題のスプレッドシート」を「タスク管理表」などに変更
echo      2. メニュー: 拡張機能 ^>^> Apps Script
echo      3. 出てきたエディタの中身を全部削除
echo      4. もう片方のタブ (Code.gs) からコードを全部コピー (Ctrl+A → Ctrl+C)
echo      5. Apps Script エディタに貼り付け (Ctrl+V)
echo      6. 上部のフロッピーアイコン (保存) または Ctrl+S
echo      7. 関数選択ドロップダウンから setupAll を選んで実行ボタン
echo      8. 初回のみ「権限の確認」ダイアログ → 「許可」をクリック
echo      9. 完了したらスプレッドシートに戻ると「タスク管理」メニューが出ます
echo.
echo  詳細は同梱の README.md を参照してください
echo.
pause
endlocal
