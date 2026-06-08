"""Stitch rendered PNGs into contact-sheet grids for fast visual QA.
Usage: python contact_sheet.py <png_dir> <out_prefix> [cols] [rows]
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw

png_dir = Path(sys.argv[1])
out_prefix = sys.argv[2]
COLS = int(sys.argv[3]) if len(sys.argv) > 3 else 4
ROWS = int(sys.argv[4]) if len(sys.argv) > 4 else 5

pngs = sorted(png_dir.glob("*.PNG"))
cell_w, cell_h, pad, label_h = 480, 270, 8, 18
per = COLS * ROWS
sheets = 0
for start in range(0, len(pngs), per):
    chunk = pngs[start:start + per]
    W = COLS * (cell_w + pad) + pad
    H = ROWS * (cell_h + label_h + pad) + pad
    sheet = Image.new("RGB", (W, H), (24, 27, 59))
    draw = ImageDraw.Draw(sheet)
    for i, p in enumerate(chunk):
        r, c = divmod(i, COLS)
        x = pad + c * (cell_w + pad)
        y = pad + r * (cell_h + label_h + pad)
        try:
            im = Image.open(p).convert("RGB").resize((cell_w, cell_h))
            sheet.paste(im, (x, y + label_h))
        except Exception:
            pass
        draw.text((x + 2, y + 3), p.stem[:70], fill=(255, 255, 255))
    out = f"{out_prefix}_{sheets+1}.png"
    sheet.save(out)
    print("wrote", out, f"({len(chunk)} cells)")
    sheets += 1
print("total sheets", sheets)
