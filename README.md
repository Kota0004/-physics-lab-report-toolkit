# 物理学実験レポート ツールキット

日本大学理工学部 物理学科の物理学実験レポートを、
**手書き表紙・予習(スキャン PDF)+ タイプ打ち本文(実験結果・課題・考察)** の
体裁で半自動生成するための仕組み一式。

タイプ打ち部分を JavaScript(`docx` パッケージ)でプログラム生成し、
Word で PDF 化したのち、手書き表紙 PDF を頭に結合して全ページを A4 に統一する。
図・表は Excel の測定ブックに作ったグラフをそのまま書き出して使う。

## できること

- 本文・見出し・段落・**番号付き数式**(分数・Σ・√)・**表**・**図**を関数で組み立て
- 図キャプションは下、表キャプションは上、という体裁ルールを自動適用
- 図・表・式の前後に空行を自動挿入、キャプションと本体がページ境界で割れない(`keepNext`)
- ページ番号・式番号を**手書きページの続き**から開始
- Excel 埋め込みグラフ → PNG 書き出し
- 手書き表紙 PDF + タイプ打ち PDF → **A4 統一・見切れなし**で結合

## 必要なもの

- Windows + **Microsoft Word / Excel**(COM 自動化で PDF 化・グラフ書き出しに使用)
- Node.js 18+
- Python 3.10+

```powershell
npm install                     # docx
pip install -r requirements.txt # pypdf, pypdfium2, pillow ほか
```

## ディレクトリ構成

```
lib/report.js              生成ライブラリ(ヘルパー + Document 組み立て)
scripts/
  export-excel-charts.ps1  Excel 埋め込みグラフを PNG 書き出し(Excel COM)
  docx-to-pdf.ps1          docx を PDF 化(Word COM)
  merge-cover.py           表紙 PDF + 本文 PDF を A4 統一で結合(pypdf)
  render-pdf.py            PDF を PNG 化(仕上がり確認用, pypdfium2)
examples/planck/           実例:プランク定数レポート一式(→ examples/planck/README.md)
```

## 使い方(新しいレポートを作る)

1. `examples/planck/` をひな型としてコピー(例: `examples/my-exp/`)。
2. `data/` に測定 Excel を置き、グラフを作る。
   `scripts/export-excel-charts.ps1` で図を PNG に書き出し、`figures/` に配置。
3. 手書き表紙・予習をスキャンして `cover/cover.pdf` に置く。
4. `make.js` を編集して本文を書く(下記 API)。手書きページ数に合わせて
   `pageStart` と式番号の開始を決める。
5. `build.ps1` を実行 → `build/xxx_report.pdf`(提出版)が出来る。

### 生成 API(`lib/report.js`)

```js
const R = require("../../lib/report");
const { T, sub, sup, m, mfrac, msum } = R;

const r = R.createReport({ assetDir: __dirname });

r.head1("実験結果");
r.body([ T("本文…"), sub("s"), T(" のように下付きも書ける。") ]);

r.table("表1 …",             // キャプション(表の上に付く)
        [2200, 2400, 2200],   // 列幅(DXA)
        rows);                // rows[0] がヘッダ行

r.figure("figures/fig1.png", 480, 320, "図1 …");  // キャプションは図の下

r.meqn([ m("h = "), mfrac([m("ea")], [m("c")]) ], 14);  // Word 数式 + 式番号(14)
r.eqn([ T("a = (1.45 ± 0.03) × 10"), sup("−6"), T(" V・m") ], 13);  // テキスト式 + 番号

// A4・フッターページ番号。手書きが 3 ページなら pageStart:4
R.pack(R.buildDocument(r.children, { pageStart: 4 }), "build/report.docx");
```

`figure` / `table` / `eqn` / `meqn` は前後に空行を自動で入れる。
主なヘルパー:

| 関数 | 用途 |
|---|---|
| `T / sub / sup` | 本文テキスト run(下付き・上付き) |
| `head1 / head2` | 大見出し・小見出し |
| `body / para` | 段落(本文は両端揃え・字下げ) |
| `table(cap, widths, rows, size?)` | 表(キャプション上) |
| `figure(file, w, h, cap, title?, type?)` | 図(キャプション下) |
| `eqn(runs, n)` / `meqn(mathParts, n)` | 番号付き式(テキスト / Word 数式) |
| `m / mfrac / msum / msub / msup / msubsup / mbr / mrad` | Word 数式パーツ |
| `buildDocument(children, {pageStart})` / `pack(doc, path)` | 文書生成・保存 |

## 体裁ルール(既定)

- 本文 Times New Roman / 游明朝 10.5pt、両端揃え・字下げ
- 見出し Arial / 游ゴシック 太字
- **図キャプションは下、表キャプションは上**
- 別行立ての式は右端に式番号。**手書きの続き番号**から
- ページ番号はフッター中央。**手書きページの続き**から(`pageStart`)
- 提出版 PDF は表紙スキャンを頭に結合し、全ページ A4(595.276×841.890 pt)に統一(見切れ禁止)

## 注意

- Word / Excel の COM 自動化を使うため **Windows + Office 必須**。
  他 OS では docx 生成(`make.js`)までは動くが、PDF 化・グラフ書き出しは動かない。
- 大学・出版社の配布資料や教科書スキャンをコミットする場合は
  **リポジトリを private** にすること。public にするならそれらを除外する。
