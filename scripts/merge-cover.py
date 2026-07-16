"""手書き表紙 PDF とタイプ打ち本文 PDF を結合し、全ページを A4 に統一する。

    python scripts/merge-cover.py cover.pdf body.pdf out.pdf

複数の入力 PDF を順に指定できる(最後の引数が出力):

    python scripts/merge-cover.py cover.pdf body1.pdf body2.pdf out.pdf

各ページはアスペクト比を保ったまま A4(595.276 x 841.890 pt)に収まるよう
拡大/縮小し、余白は中央寄せにする。これにより手書きスキャンとタイプ打ちで
ページサイズが揃い、内容が見切れることもない。
"""
import sys
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import RectangleObject

A4W, A4H = 595.276, 841.890


def main(inputs, out):
    writer = PdfWriter()
    for src in inputs:
        reader = PdfReader(src)
        for page in reader.pages:
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            s = min(A4W / w, A4H / h)
            tx = (A4W - w * s) / 2
            ty = (A4H - h * s) / 2
            page.add_transformation(Transformation().scale(s).translate(tx, ty))
            page.mediabox = RectangleObject([0, 0, A4W, A4H])
            if "/CropBox" in page:
                page.cropbox = RectangleObject([0, 0, A4W, A4H])
            writer.add_page(page)
    with open(out, "wb") as f:
        writer.write(f)
    print(f"{len(writer.pages)} pages -> {out}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("usage: python merge-cover.py <in1.pdf> [in2.pdf ...] <out.pdf>")
    main(sys.argv[1:-1], sys.argv[-1])
