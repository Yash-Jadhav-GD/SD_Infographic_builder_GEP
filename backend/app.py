"""GEP Infographic Builder — web backend (FastAPI).

Upload/enter data -> auto-select a GEP template -> fill -> download .pptx + PNG preview.
Reuses the proven fill logic from _tools/fill_slide.py.
"""
import os, sys, json, uuid, subprocess, glob
from pathlib import Path
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import List, Optional

ROOT = Path(__file__).resolve().parent.parent          # project root
sys.path.insert(0, str(ROOT / "_tools"))
sys.path.insert(0, str(Path(__file__).resolve().parent))  # backend/ for sibling imports
import fill_slide  # noqa: E402
import icon_search  # noqa: E402

def _load_catalog():
    here = Path(__file__).parent
    curated = json.load(open(here / "catalog.json", encoding="utf-8"))["templates"]
    for t in curated:
        t.setdefault("auto", False)
    entries = list(curated)
    auto_path = here / "catalog_auto.json"
    if auto_path.exists():
        for t in json.load(open(auto_path, encoding="utf-8"))["templates"]:
            t["auto"] = True
            entries.append(t)
    return entries

CATALOG = _load_catalog()
OUT = ROOT / "backend" / "_out"
OUT.mkdir(parents=True, exist_ok=True)
EXPORT_PS1 = ROOT / "_tools" / "export_pptx.ps1"

app = FastAPI(title="GEP Infographic Builder")


class Item(BaseModel):
    label: str
    body: str = ""
    icon: Optional[str] = None  # index 'file', e.g. "icons/1_5.png"
    meta: Optional[str] = None  # extra short field, e.g. a date/year for timelines


class GenRequest(BaseModel):
    title: str
    items: List[Item]
    template_id: Optional[str] = None
    center: str = ""


def select_template(n: int, template_id: Optional[str]):
    if template_id:
        for t in CATALOG:
            if t["id"] == template_id:
                return t
    # rule-based auto-pick: only curated templates (auto entries are pick-by-name)
    pool = [t for t in CATALOG if not t.get("auto")] or CATALOG
    candidates = sorted(pool, key=lambda t: abs(t["item_count"] - n))
    return candidates[0] if candidates else None


def build_spec(tpl, req: GenRequest, out_path: str):
    name_cap = tpl.get("name_capacity", 30)
    body_cap = tpl.get("body_capacity", 200)
    text = {str(tpl["title_shape"]): req.title}
    for i, slot in enumerate(tpl["items"]):
        if i >= len(req.items):
            break
        it = req.items[i]
        label = (it.label or "").strip()
        body = (it.body or "").strip()
        if "combo_shape" in slot:
            # one text box holding label (para 0) + body (para 1)
            paras = [label] if label else []
            if body:
                paras.append(body[:body_cap])
            text[str(slot["combo_shape"])] = paras if paras else ""
            continue
        if "name_shape" in slot:
            text[str(slot["name_shape"])] = label[:name_cap]
        if "number_shape" in slot:
            text[str(slot["number_shape"])] = f"{i+1:02d}"
        if "meta_shape" in slot:
            text[str(slot["meta_shape"])] = (it.meta or "").strip()
        if "body_shape" in slot:
            # body-only slot: fold label into body
            if "name_shape" not in slot and label:
                content = f"{label} — {body}" if body else label
            else:
                content = body
            if content:
                text[str(slot["body_shape"])] = content[:body_cap]
    if "center_shape" in tpl:
        text[str(tpl["center_shape"])] = (req.center or "").strip()
    # blank out leftover decorative/placeholder shapes (garbled lorem in source deck)
    for sid in tpl.get("blank_shapes", []):
        text[str(sid)] = ""

    add_images = []
    anchors = tpl.get("icon_anchors")
    if anchors:
        color = tpl.get("icon_color", "navy")
        for i, slot_anchor in enumerate(anchors):
            if i >= len(req.items):
                break
            ic = req.items[i].icon
            if not ic:
                continue
            try:
                recolored = icon_search.recolor(ic, color)
            except Exception:
                continue
            add_images.append({"path": recolored, **slot_anchor})

    slide_spec = {"index": tpl["slide_index"], "text": text}
    if add_images:
        slide_spec["add_images"] = add_images
    if tpl.get("remove_shapes"):
        slide_spec["remove_shapes"] = tpl["remove_shapes"]
    return {
        "template": str(ROOT / tpl["template"]),
        "output": out_path,
        "slides": [slide_spec],
    }


def render_preview(pptx_path: str, out_dir: str):
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(EXPORT_PS1), "-Pptx", pptx_path, "-OutDir", out_dir],
        capture_output=True, text=True, timeout=120,
    )
    pngs = sorted(glob.glob(os.path.join(out_dir, "*.PNG")))
    return pngs[0] if pngs else None


@app.get("/api/templates")
def templates():
    return [{"id": t["id"], "name": t["name"], "item_count": t["item_count"],
             "family": t.get("family", "other"), "auto": t.get("auto", False),
             "description": t["description"]} for t in CATALOG]


@app.post("/api/generate")
def generate(req: GenRequest):
    tpl = select_template(len(req.items), req.template_id)
    if not tpl:
        return JSONResponse({"error": "No template available"}, status_code=400)
    job = uuid.uuid4().hex[:10]
    pptx_path = str(OUT / f"{job}.pptx")
    spec = build_spec(tpl, req, pptx_path)
    spec_path = OUT / f"{job}_spec.json"
    json.dump(spec, open(spec_path, "w", encoding="utf-8"))
    fill_slide.sys.argv = ["fill_slide", str(spec_path)]
    fill_slide.main()
    preview_dir = OUT / job
    preview_dir.mkdir(exist_ok=True)
    png = render_preview(pptx_path, str(preview_dir))
    return {
        "job": job,
        "template_used": tpl["name"],
        "download": f"/api/download/{job}",
        "preview": f"/api/preview/{job}" if png else None,
    }


@app.get("/api/download/{job}")
def download(job: str):
    p = OUT / f"{job}.pptx"
    return FileResponse(p, filename="GEP_Infographic.pptx") if p.exists() \
        else JSONResponse({"error": "not found"}, status_code=404)


@app.get("/api/preview/{job}")
def preview(job: str):
    pngs = sorted(glob.glob(str(OUT / job / "*.PNG")))
    return FileResponse(pngs[0]) if pngs else JSONResponse({"error": "not found"}, status_code=404)


@app.get("/api/icons")
def icons(q: str, color: str = "navy", n: int = 8):
    out = []
    for e in icon_search.search(q, n):
        rec = icon_search.recolor(e["file"], color)
        out.append({"tags": e["tags"][:6],
                    "icon": "/" + e["file"],
                    "recolored": "/icons_recolored/" + Path(rec).name})
    return out


# static mounts (must precede the catch-all "/" mount)
HERE = Path(__file__).parent
app.mount("/icons", StaticFiles(directory=str(HERE / "icons")), name="icons")
app.mount("/icons_recolored", StaticFiles(directory=str(HERE / "icons_recolored")), name="icons_rec")
# static UI (catch-all, keep last)
app.mount("/", StaticFiles(directory=str(HERE / "static"), html=True), name="static")
