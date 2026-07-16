// 物理学実験レポート生成ライブラリ
// docx パッケージのラッパー。日本語レポートの体裁ルール(下記)を関数として提供する。
//
// 体裁ルール:
//   - 本文: Times New Roman / 游明朝, 10.5pt(size=21), 両端揃え・字下げ
//   - 見出し: Arial / 游ゴシック 太字
//   - 図キャプションは図の「下」、表キャプションは表の「上」
//   - 別行立ての式には式番号を右端に振る(中央タブ+右タブ)
//   - 図・表・式の前後には空行を入れる(figure/table/eqn ヘルパーが自動挿入)
//   - ページ番号はフッター中央。手書きページの続きから始める(buildDocument の pageStart)
//   - 図とキャプション/表とキャプションがページ境界で分離しないよう keepNext を付与
//
// 使い方は examples/planck/make.js を参照。

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  TabStopType, Tab, Footer, PageNumber,
  Math: OMath, MathRun, MathFraction, MathSum, MathRadical,
  MathSubScript, MathSuperScript, MathSubSuperScript, MathRoundBrackets,
} = require("docx");

// ---------- フォント ----------
const JFONT = { ascii: "Times New Roman", hAnsi: "Times New Roman", eastAsia: "游明朝" };
const GFONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "游ゴシック" };

// ---------- テキスト run ----------
function T(text, opts = {}) {
  return new TextRun({ text, font: JFONT, size: 21, ...opts });
}
function sub(text, opts = {}) { return T(text, { subScript: true, ...opts }); }
function sup(text, opts = {}) { return T(text, { superScript: true, ...opts }); }

// ---------- 段落 ----------
function para(children, opts = {}) {
  return new Paragraph({ children, spacing: { after: 80, line: 300 }, ...opts });
}
function body(children) {
  return para(children, { indent: { firstLine: 210 }, alignment: AlignmentType.BOTH });
}

// ---------- 見出し ----------
function head1(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: GFONT, size: 26, bold: true })],
    spacing: { before: 240, after: 140 },
  });
}
function head2(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: GFONT, size: 22, bold: true })],
    spacing: { before: 160, after: 100 },
  });
}

// ---------- 式(番号付き)----------
// 中央タブで数式を中央寄せ、右タブで式番号を右端に配置。
function eqn(children, num) {
  return new Paragraph({
    tabStops: [
      { type: TabStopType.CENTER, position: 4513 },
      { type: TabStopType.RIGHT, position: 9026 },
    ],
    spacing: { before: 60, after: 100 },
    children: [
      new TextRun({ children: [new Tab()], font: JFONT, size: 21 }),
      ...children,
      new TextRun({ children: [new Tab(), `(${num})`], font: JFONT, size: 21 }),
    ],
  });
}
// OMML(Word 数式)を1つ含む番号付き式
function meqn(mathChildren, num) {
  return eqn([new OMath({ children: mathChildren })], num);
}

// ---------- キャプション ----------
// above:true で表用(上・keepNext)、省略で図用(下)
function caption(text, opts = {}) {
  return para([T(text, { size: 19 })], {
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 60 },
    keepNext: !!opts.above,
  });
}

// ---------- 画像 ----------
function img(absPath, w, h, title, type = "png") {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    keepNext: true,
    children: [new ImageRun({
      type,
      data: fs.readFileSync(absPath),
      transformation: { width: w, height: h },
      altText: { title, description: title, name: title },
    })],
  });
}

// ---------- 表 ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: "444444" };
const borders = { top: border, bottom: border, left: border, right: border };

function tcellText(text, width, o = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 30, bottom: 30, left: 60, right: 60 },
    ...(o.bold ? { shading: { fill: "EFEFEF", type: ShadingType.CLEAR } } : {}),
    children: [para([T(text, { size: o.size || 20, bold: o.bold })], {
      alignment: AlignmentType.CENTER, spacing: { after: 0 },
    })],
  });
}
// rows[0] をヘッダ行(太字・網掛け)として扱う
function simpleTable(colWidths, rows, cellSize) {
  return new Table({
    alignment: AlignmentType.CENTER,
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((r, i) => new TableRow({
      children: r.map((c, j) => tcellText(c, colWidths[j], { bold: i === 0, size: cellSize })),
    })),
  });
}

// ---------- OMML(Word 数式)プリミティブ ----------
const m = (t) => new MathRun(t);
const mfrac = (num, den) => new MathFraction({ numerator: num, denominator: den });
const msum = (children, subVar = "i") => new MathSum({ children, subScript: [m(subVar)] });
const msub = (ch, s) => new MathSubScript({ children: ch, subScript: s });
const msup = (ch, s) => new MathSuperScript({ children: ch, superScript: s });
const msubsup = (ch, s, p) => new MathSubSuperScript({ children: ch, subScript: s, superScript: p });
const mbr = (children) => new MathRoundBrackets({ children });
const mrad = (children) => new MathRadical({ children });

// ---------- レポート組み立て ----------
// createReport({ assetDir }) で章立てをためていくビルダーを返す。
// figure/table/eqn/meqn は前後に空行を自動挿入する(体裁ルール)。
function createReport(opts = {}) {
  const assetDir = opts.assetDir || process.cwd();
  const children = [];

  function blankLine() {
    const last = children[children.length - 1];
    if (last && last.__isBlank) return;
    const b = new Paragraph({ children: [] });
    b.__isBlank = true;
    children.push(b);
  }

  const api = {
    children,
    push: (...els) => { els.forEach((e) => children.push(e)); return api; },
    head1: (t) => { children.push(head1(t)); return api; },
    head2: (t) => { children.push(head2(t)); return api; },
    body: (runs) => { children.push(body(runs)); return api; },
    para: (runs, o) => { children.push(para(runs, o)); return api; },
    raw: (p) => { children.push(p); return api; },
    blank: () => { blankLine(); return api; },
    // 図(前後空行・キャプション下)
    figure(file, w, h, capText, title, type = "png") {
      blankLine();
      children.push(img(path.join(assetDir, file), w, h, title || capText, type));
      children.push(caption(capText));
      blankLine();
      return api;
    },
    // 表(前後空行・キャプション上)
    table(capText, colWidths, rows, cellSize) {
      blankLine();
      children.push(caption(capText, { above: true }));
      children.push(simpleTable(colWidths, rows, cellSize));
      blankLine();
      return api;
    },
    // 番号付き式(前後空行)。omml:true なら mathChildren を OMath で包む
    eqn(childrenArr, num, ommlFlag = false) {
      blankLine();
      children.push(ommlFlag ? meqn(childrenArr, num) : eqn(childrenArr, num));
      blankLine();
      return api;
    },
    meqn(mathChildren, num) { return api.eqn(mathChildren, num, true); },
  };
  return api;
}

// A4・游明朝・フッターにページ番号。pageStart で開始ページ番号を指定(手書きの続き)。
function buildDocument(children, opts = {}) {
  const pageStart = opts.pageStart != null ? opts.pageStart : 1;
  return new Document({
    styles: { default: { document: { run: { font: JFONT, size: 21 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          pageNumbers: { start: pageStart },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: JFONT, size: 20 })],
          })],
        }),
      },
      children,
    }],
  });
}

async function pack(doc, outPath) {
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  return buf.length;
}

module.exports = {
  JFONT, GFONT,
  T, sub, sup, para, body, head1, head2, eqn, meqn, caption, img, simpleTable,
  m, mfrac, msum, msub, msup, msubsup, mbr, mrad,
  createReport, buildDocument, pack,
  AlignmentType,
};
