"""PDF の各ページを PNG 画像に書き出す(仕上がり確認用)。

    python scripts/render-pdf.py report.pdf out_dir [--scale 1.4] [--from 5]

out_dir に page_1.png, page_2.png, ... を出力する。--from を指定すると
その番号(1 始まり)以降のページだけを書き出す。
"""
import argparse
import os
import pypdfium2 as pdfium


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out_dir")
    ap.add_argument("--scale", type=float, default=1.4)
    ap.add_argument("--from", dest="start", type=int, default=1)
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    doc = pdfium.PdfDocument(args.pdf)
    for i in range(args.start - 1, len(doc)):
        img = doc[i].render(scale=args.scale).to_pil()
        dest = os.path.join(args.out_dir, f"page_{i + 1}.png")
        img.save(dest)
        print(dest)


if __name__ == "__main__":
    main()
