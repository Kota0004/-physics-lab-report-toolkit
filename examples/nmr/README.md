# 実例:核磁気共鳴(NMR)

第6回 核磁気共鳴の実験レポートを、このツールキットで生成する一式。
課題は教科書 §4 ではなく、配布資料「核磁気共鳴『§4.課題』」の課題(1)〜(3)に準拠。
タイプ打ち部分は **課題(1)〜(3)+考察のみ**(実験結果セクション・図・表・参考文献欄なし)で、
値の導出はすべて番号付きの式で示す体裁。

## ファイル

```
make.js            レポート本文(課題(1)〜(3)・考察)の生成スクリプト
make_fig1.py       参考図(共鳴周波数 f vs 磁場 B と最小二乗直線)の生成スクリプト。
                   レポートには載せない(データ確認用)
build.sh           一括ビルド(docx生成 → LibreOffice で PDF 化 → 表紙結合)
data/nmr_measurements.csv   測定データ(f = 8.0〜11.0 MHz の 7 点の共鳴磁場)
figures/fig1_f_vs_B.png     参考図(make_fig1.py の出力・レポート非掲載)
cover/cover.pdf    手書きスキャン(表紙+予習 −1−〜−4−+課題シート −5−・磁場記入済)
build/             生成物(.gitignore 対象)
```

## ビルド

```bash
# 依存の準備(リポジトリ直下で一度だけ)
npm install
pip install -r requirements.txt
# さらに Linux では: apt-get install libreoffice-writer fonts-noto-cjk fonts-noto-cjk-extra

bash examples/nmr/build.sh
# -> examples/nmr/build/nmr_report.pdf(提出版・全9ページ)
```

Windows + Word で組む場合は `examples/planck/build.ps1` と同様に
`scripts/docx-to-pdf.ps1` を使う(体裁の再現性は Word の方が高い)。

## この実験でのポイント

- 手書き(表紙+予習 −1−〜−4−+課題シート −5−)が 6 ページなので、タイプ打ちは
  **ページ 6 から**(`buildDocument(..., { pageStart: 6 })`)、式番号も手書き原理の
  (1)〜(12) に続けて **(13) から**。
- 平均 B̄・f̄ → 偏差の積和・2乗和 → 傾き df/dB → γn → ħ → gn の順に、
  **すべて数値を代入した式(13)〜(23)で導出**(表・グラフは使わない)。
- LibreOffice の OMML 変換では `MathSum`(Σ)や `MathRoundBrackets` が
  □に化けるため、**Σ と括弧は通常文字**で数式を組んでいる(make.js 参照)。
  游明朝/游ゴシックは fontconfig で Noto Serif/Sans CJK JP に置換する。
- 主要結果: df/dB = 4.2003 × 10⁷ Hz/T(r² = 0.99998)、
  γn = 2π × df/dB = 2.6391 × 10⁸ s⁻¹・T⁻¹(理想値 2.6752 × 10⁸ の −1.35 %)、
  gn = γnħ/μN = 5.5104(理想値 5.58554 の −1.35 %)。
  ※ gn は課題(2)の 5 桁値 γn = 2.6391×10⁸ と ħ = 1.0546×10⁻³⁴ を式に代入した値
  (丸めない連鎖計算では 5.5103。差は 5 桁目のみ)。
- 使用した定数: μN = 5.0508 × 10⁻²⁷ J/T(教科書 297 ページ)、
  h = 6.626 × 10⁻³⁴ J・s(教科書巻末)→ ħ = h/2π = 1.0546 × 10⁻³⁴ J・s。
