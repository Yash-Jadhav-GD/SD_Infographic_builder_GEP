"""Search the GEP icon index and recolor icons to brand colors.

Icons are transparent PNGs with neutral gray strokes; recolor() repaints the
opaque pixels to a target hex while preserving the alpha edges (anti-aliasing).
"""
import json, re, os
from pathlib import Path
from functools import lru_cache
from PIL import Image

HERE = Path(__file__).resolve().parent
INDEX = json.load(open(HERE / "icons_index.json", encoding="utf-8"))
RECOLOR_DIR = HERE / "icons_recolored"
RECOLOR_DIR.mkdir(exist_ok=True)

GEP = {"purple": "#48397E", "orange": "#FE860E", "navy": "#243472",
       "deepnavy": "#101B3B", "charcoal": "#111827", "steel": "#6B7280",
       "white": "#FFFFFF", "cloud": "#E5E7EB"}

STOP = set("the a an of and or to for with in on at by your our their".split())

def _tok(text):
    return [t for t in re.findall(r"[a-z][a-z0-9]+", (text or "").lower())
            if t not in STOP and len(t) > 1]

def search(query, n=8):
    """Return up to n index entries best matching the query string."""
    q = set(_tok(query))
    if not q:
        return []
    scored = []
    for e in INDEX:
        tags = set(e["tags"])
        if not tags:
            continue
        overlap = len(q & tags)
        if overlap:
            scored.append((overlap, e))
    scored.sort(key=lambda x: -x[0])
    return [e for _, e in scored[:n]]

def best(query):
    r = search(query, 1)
    return r[0] if r else None

def _hex(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def recolor(rel_file, color="navy"):
    """Recolor an icon to a GEP brand color; returns path to cached PNG."""
    hexv = GEP.get(color, color)
    rgb = _hex(hexv)
    out = RECOLOR_DIR / f"{Path(rel_file).stem}_{color.lstrip('#')}.png"
    if out.exists():
        return str(out)
    im = Image.open(HERE / rel_file).convert("RGBA")
    px = im.getdata()
    new = [(rgb[0], rgb[1], rgb[2], a) for (_, _, _, a) in px]
    im.putdata(new)
    im.save(out)
    return str(out)
