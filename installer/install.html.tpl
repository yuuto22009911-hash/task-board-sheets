<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>task-board-sheets セットアップ</title>
<style>
  :root {
    --bg: #f7f8fa;
    --card: #ffffff;
    --primary: #1a73e8;
    --primary-dark: #1558b8;
    --text: #202124;
    --muted: #5f6368;
    --border: #dadce0;
    --code-bg: #1e1e1e;
    --code-text: #d4d4d4;
    --ok: #137333;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 32px 16px 64px;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.65;
  }
  .wrap { max-width: 760px; margin: 0 auto; }
  h1 {
    font-size: 22px;
    margin: 0 0 6px;
    letter-spacing: -0.01em;
  }
  .sub {
    color: var(--muted);
    font-size: 13px;
    margin-bottom: 28px;
  }
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 22px 24px;
    margin-bottom: 16px;
  }
  .step-title {
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 600;
    font-size: 15px;
    margin: 0 0 8px;
  }
  .step-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 26px;
    height: 26px;
    background: var(--primary);
    color: #fff;
    border-radius: 50%;
    font-size: 13px;
    font-weight: 700;
    flex-shrink: 0;
  }
  .step-body {
    margin: 0 0 12px 36px;
    color: var(--text);
  }
  .step-body p { margin: 6px 0; }
  .step-body ol, .step-body ul {
    padding-left: 20px;
    margin: 6px 0;
  }
  .btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--primary);
    color: #fff;
    padding: 9px 18px;
    border: 0;
    border-radius: 8px;
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
    text-decoration: none;
    margin: 4px 6px 4px 0;
    transition: background .15s;
    font-family: inherit;
  }
  .btn:hover { background: var(--primary-dark); }
  .btn.secondary {
    background: #fff;
    color: var(--primary);
    border: 1px solid var(--primary);
  }
  .btn.secondary:hover { background: #f0f6ff; }
  .copied {
    display: inline-block;
    color: var(--ok);
    font-weight: 600;
    font-size: 12px;
    margin-left: 8px;
    opacity: 0;
    transition: opacity .2s;
  }
  .copied.show { opacity: 1; }
  details {
    margin-top: 14px;
    border-top: 1px solid var(--border);
    padding-top: 12px;
  }
  details summary {
    cursor: pointer;
    color: var(--muted);
    font-size: 12px;
    user-select: none;
  }
  details[open] summary { color: var(--text); }
  pre {
    margin: 12px 0 0;
    background: var(--code-bg);
    color: var(--code-text);
    padding: 14px;
    border-radius: 8px;
    font-family: "SF Mono", Consolas, Menlo, monospace;
    font-size: 11.5px;
    line-height: 1.5;
    max-height: 320px;
    overflow: auto;
    white-space: pre;
    -webkit-user-select: all;
    user-select: all;
  }
  code.inline {
    background: #f1f3f4;
    padding: 1px 6px;
    border-radius: 4px;
    font-family: "SF Mono", Consolas, monospace;
    font-size: 12px;
  }
  .hint {
    background: #fffbf0;
    border: 1px solid #fbe4a0;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 12.5px;
    color: #594300;
    margin: 10px 0;
  }
  .footer {
    text-align: center;
    color: var(--muted);
    font-size: 11.5px;
    margin-top: 32px;
  }
</style>
</head>
<body>
<div class="wrap">
  <h1>task-board-sheets セットアップ</h1>
  <p class="sub">Google スプレッドシート + Apps Script のタスク管理表。所要時間 約3分。</p>

  <div class="card">
    <p class="step-title"><span class="step-num">1</span>新規スプレッドシートを作成</p>
    <div class="step-body">
      <p>下のボタンから新しいスプレッドシートを開きます。タブで開かれた画面はそのままにしておいてください。</p>
      <a class="btn" href="https://sheets.new" target="_blank" rel="noopener">新規スプレッドシートを開く</a>
      <p>左上の「無題のスプレッドシート」をクリックして好きな名前(例: タスク管理表)に変更してください。</p>
    </div>
  </div>

  <div class="card">
    <p class="step-title"><span class="step-num">2</span>Apps Script エディタを開く</p>
    <div class="step-body">
      <p>スプレッドシート上部のメニューから:</p>
      <p><code class="inline">拡張機能</code> → <code class="inline">Apps Script</code></p>
      <p>新しいタブで Apps Script エディタが開きます。エディタ中央の既存コード(<code class="inline">function myFunction() {}</code> など)を <strong>すべて選択して削除</strong> (Ctrl + A → Delete) してください。</p>
    </div>
  </div>

  <div class="card">
    <p class="step-title"><span class="step-num">3</span>コードをコピーして貼り付け</p>
    <div class="step-body">
      <p>下の「コードをコピー」ボタンを押すと、Apps Script のコードが自動でクリップボードに入ります。</p>
      <button class="btn" id="copyBtn" onclick="copyCode()">コードをコピー</button>
      <span class="copied" id="copiedMsg">コピー完了</span>
      <p>その後、Apps Script エディタに戻って Ctrl + V で貼り付け、Ctrl + S で保存してください。スクリプト名を「タスク管理表」など好きな名前にして OK を押します。</p>
      <details>
        <summary>コードを直接見る (展開)</summary>
        <pre id="code">{{CODE_GS}}</pre>
      </details>
    </div>
  </div>

  <div class="card">
    <p class="step-title"><span class="step-num">4</span>初期セットアップ実行</p>
    <div class="step-body">
      <p>Apps Script エディタの上部:</p>
      <ol>
        <li>関数選択ドロップダウン (デフォルト <code class="inline">onOpen</code>) を <strong><code class="inline">setupAll</code></strong> に変更</li>
        <li>隣の <strong>実行</strong> (▶) ボタンをクリック</li>
        <li>「権限の確認」 → 自分の Google アカウントを選択</li>
        <li>「Google がこのアプリを審査していません」と出たら<br>
          → 「詳細」 → 「(プロジェクト名) (安全ではないページ) に移動」をクリック</li>
        <li>「許可」をクリック</li>
        <li>実行ログに「実行が完了しました」と出れば成功</li>
      </ol>
      <div class="hint">
        ヒント: 「安全ではない」警告は自分で書いた(自分でコピーした)コードに対する一般的な警告です。<br>
        このコードは Google スプレッドシートと Google カレンダー(読み取り)以外にはアクセスしません。
      </div>
    </div>
  </div>

  <div class="card">
    <p class="step-title"><span class="step-num">5</span>動作確認</p>
    <div class="step-body">
      <p>スプレッドシート画面に戻ります(タブを切り替え、必要ならページをリロード)。</p>
      <ul>
        <li>上部メニューに <strong>「タスク管理」</strong> が出ているはず</li>
        <li>「タスク管理」 → 「タスク追加」でタスク一覧の2行目に新規行が追加されることを確認</li>
        <li>状況プルダウンで色が変わることを確認</li>
        <li>SMTG リマインドを使う場合は「設定」シートでカレンダーIDを確認 (デフォルト <code class="inline">primary</code> は自分のメインカレンダー)</li>
      </ul>
      <a class="btn secondary" href="https://github.com/yuuto22009911-hash/task-board-sheets/blob/main/docs/SETUP_GUIDE.md" target="_blank" rel="noopener">詳しい手順 (GitHub)</a>
    </div>
  </div>

  <p class="footer">task-board-sheets / Google Cloud Console は不要 / Apps Script のみで動作</p>
</div>

<script>
function copyCode() {
  const code = document.getElementById('code').textContent;
  navigator.clipboard.writeText(code).then(() => {
    const m = document.getElementById('copiedMsg');
    m.classList.add('show');
    setTimeout(() => m.classList.remove('show'), 2000);
  }).catch((err) => {
    // フォールバック: 古いブラウザ用
    const ta = document.createElement('textarea');
    ta.value = code;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(ta);
    const m = document.getElementById('copiedMsg');
    m.classList.add('show');
    setTimeout(() => m.classList.remove('show'), 2000);
  });
}
</script>
</body>
</html>
