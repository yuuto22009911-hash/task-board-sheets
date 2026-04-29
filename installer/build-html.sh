#!/usr/bin/env bash
# install.html.tpl の {{CODE_GS}} プレースホルダを apps-script/Code.gs の内容で置換して
# install.html を生成する。Code.gs を更新したらこのスクリプトを再実行する。
set -euo pipefail
cd "$(dirname "$0")/.."

CODE_FILE="apps-script/Code.gs"
TPL="installer/install.html.tpl"
OUT="installer/install.html"

# HTML エスケープ (& < > のみ)
escape_html() {
  /usr/bin/sed -e 's/&/\&amp;/g' -e 's/</\&lt;/g' -e 's/>/\&gt;/g' "$1"
}

ESCAPED=$(escape_html "$CODE_FILE")
# 区切り文字に \x01 を使って sed で一括置換
python3 - "$TPL" "$OUT" "$CODE_FILE" <<'PY'
import sys, html, os
tpl_path, out_path, code_path = sys.argv[1], sys.argv[2], sys.argv[3]
with open(tpl_path, 'r', encoding='utf-8') as f:
    tpl = f.read()
with open(code_path, 'r', encoding='utf-8') as f:
    code = f.read()
escaped = html.escape(code, quote=False)
out = tpl.replace('{{CODE_GS}}', escaped)
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(out)
print(f'Generated {out_path} ({len(out)} bytes)')

# GitHub Pages 用にも index.html として複製
gh_pages = 'docs/index.html'
os.makedirs('docs', exist_ok=True)
with open(gh_pages, 'w', encoding='utf-8') as f:
    f.write(out)
print(f'Generated {gh_pages} ({len(out)} bytes)')
PY
