# プランク定数レポートを一括ビルドする。
#   powershell -File examples/planck/build.ps1
#
# 手順: (1) docx 生成 -> (2) docx を PDF 化 -> (3) 表紙 PDF と結合して A4 統一。
# 出力: build/planck_report.pdf(提出版), build/planck.docx(編集用)。
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $here "..\..")).Path
$build = Join-Path $here "build"
New-Item -ItemType Directory -Force -Path $build | Out-Null

# (1) docx 生成
node (Join-Path $here "make.js")

# (2) docx -> PDF
powershell -File (Join-Path $repo "scripts\docx-to-pdf.ps1") `
  -In  (Join-Path $build "planck.docx") `
  -Out (Join-Path $build "planck_typed.pdf")

# (3) 表紙 PDF + 本文 PDF を A4 統一で結合
python (Join-Path $repo "scripts\merge-cover.py") `
  (Join-Path $here "cover\cover.pdf") `
  (Join-Path $build "planck_typed.pdf") `
  (Join-Path $build "planck_report.pdf")

Write-Output "done: $build\planck_report.pdf"
