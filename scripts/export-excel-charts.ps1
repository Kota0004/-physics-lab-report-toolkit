# Excel ブックに埋め込まれたグラフを PNG として書き出す(Excel COM 自動化)。
# 図の中身をレポートにそのまま使うための手段。
# Excel がインストールされた Windows で実行すること。
#
#   powershell -File scripts/export-excel-charts.ps1 -In examples/planck/data/planck.xlsx -OutDir examples/planck/figures
#
# ブック内の各グラフを chart1.png, chart2.png, ... として OutDir に出力する。
# -Scale 2 のように倍率を指定すると、グラフとフォントを拡大して高解像度で書き出す。
param(
  [Parameter(Mandatory = $true)][string]$In,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [double]$Scale = 1.0,
  [switch]$HideTitle
)

$In = (Resolve-Path $In).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($In)
  $i = 0
  foreach ($ws in $wb.Worksheets) {
    foreach ($co in $ws.ChartObjects()) {
      $i++
      $ch = $co.Chart
      if ($HideTitle -and $ch.HasTitle) { $ch.HasTitle = $false }
      if ($Scale -ne 1.0) {
        $co.Width  = $co.Width  * $Scale
        $co.Height = $co.Height * $Scale
        try { $ch.ChartArea.Font.Size = $ch.ChartArea.Font.Size * $Scale } catch {}
        foreach ($ax in $ch.Axes()) {
          try { $ax.TickLabels.Font.Size = $ax.TickLabels.Font.Size * $Scale } catch {}
          try { if ($ax.HasTitle) { $ax.AxisTitle.Font.Size = $ax.AxisTitle.Font.Size * $Scale } } catch {}
        }
        try { if ($ch.HasLegend) { $ch.Legend.Font.Size = $ch.Legend.Font.Size * $Scale } } catch {}
      }
      $dest = Join-Path $OutDir ("chart{0}.png" -f $i)
      $ch.Export($dest, "PNG") | Out-Null
      Write-Output ("chart{0} -> {1}" -f $i, $dest)
    }
  }
  $wb.Close($false)
}
finally {
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
