# docx を PDF に変換する(Word COM 自動化)。
# Word がインストールされた Windows で実行すること。
#
#   powershell -File scripts/docx-to-pdf.ps1 -In examples/planck/build/planck.docx -Out examples/planck/build/planck_typed.pdf
#
param(
  [Parameter(Mandatory = $true)][string]$In,
  [Parameter(Mandatory = $true)][string]$Out
)

$In  = (Resolve-Path $In).Path
$Out = [System.IO.Path]::GetFullPath($Out)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
  $doc = $word.Documents.Open($In)
  # 17 = wdExportFormatPDF
  $doc.ExportAsFixedFormat($Out, 17)
  $doc.Close($false)
  Write-Output "PDF written: $Out"
}
finally {
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
