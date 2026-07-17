#!/usr/bin/env bash
# 核磁気共鳴レポートを一括ビルドする(Linux / LibreOffice 版)。
#   bash examples/nmr/build.sh
#
# 手順: (1) docx 生成 -> (2) LibreOffice で PDF 化 -> (3) 手書きスキャン PDF と結合して A4 統一。
# 出力: build/nmr_report.pdf(提出版), build/nmr.docx(編集用)。
#
# Windows + Word で組む場合は examples/planck/build.ps1 と同様に
# scripts/docx-to-pdf.ps1 を使う(体裁の再現性は Word の方が高い)。
# LibreOffice では 游明朝/游ゴシック が Noto Serif/Sans CJK JP に置換される
# (~/.config/fontconfig/fonts.conf のエイリアス設定を推奨)。
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
mkdir -p "$HERE/build"

# (1) 図と docx を生成
python3 "$HERE/make_fig1.py"
node "$HERE/make.js"

# (2) docx -> PDF
soffice --headless --convert-to pdf --outdir "$HERE/build" "$HERE/build/nmr.docx" >/dev/null

# (3) 手書きスキャン(表紙+予習+課題シート)PDF + 本文 PDF を A4 統一で結合
python3 "$REPO/scripts/merge-cover.py" \
  "$HERE/cover/cover.pdf" \
  "$HERE/build/nmr.pdf" \
  "$HERE/build/nmr_report.pdf"

echo "done: $HERE/build/nmr_report.pdf"
