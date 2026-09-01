"""PDF discovery, rendering and catalog output."""
from __future__ import annotations

import csv, hashlib, logging, os, re, time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz
import yaml
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

LOG = logging.getLogger(__name__)
HEADERS = ["年级", "学期", "PDF 文件名称", "PDF 文件所在位置", *[f"转换后的图片 {i}" for i in range(1, 6)]]

@dataclass
class Settings:
    source_root: Path; output_root: Path; grade: str; semester: str
    watermark: dict[str, Any]; render: dict[str, Any]; processing: dict[str, Any]; table: dict[str, Any]

def load_settings(path: str | Path) -> Settings:
    p = Path(path).resolve(); raw = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    def req(name):
        if not raw.get(name): raise ValueError(f"缺少配置项: {name}")
        return raw[name]
    source = Path(req("source_root")).expanduser()
    if not source.is_absolute(): source = (p.parent / source).resolve()
    out = Path(req("output_root")).expanduser()
    if not out.is_absolute(): out = (p.parent / out).resolve()
    render = {"dpi": 150, "max_pages": 5, **raw.get("render", {})}
    if not (1 <= int(render["dpi"]) <= 1200): raise ValueError("render.dpi 必须在 1~1200")
    if not (1 <= int(render["max_pages"]) <= 100): raise ValueError("render.max_pages 必须在 1~100")
    wm = {"enabled": True, "text": "幼升小", "opacity": 90, "position": "bottom_right", "font_size": 28, **raw.get("watermark", {})}
    if wm["position"] not in {"bottom_right", "bottom_left", "center"}: raise ValueError("watermark.position 不支持")
    wm["opacity"] = max(0, min(255, int(wm["opacity"])))
    table = {"xlsx": "pdf_catalog.xlsx", "csv": "pdf_catalog.csv", "include_metadata_title": False, **raw.get("table", {})}
    return Settings(source, out, str(req("grade")), str(req("semester")), wm, render, {"max_pdfs": None, **raw.get("processing", {})}, table)

def discover(root: Path) -> list[Path]:
    return sorted((p for p in root.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"), key=lambda p: str(p.relative_to(root)).lower())

def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._") or "pdf"
    return value[:100]

def _font(size: int):
    for p in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        if Path(p).exists():
            try: return ImageFont.truetype(p, size)
            except OSError: pass
    return ImageFont.load_default()

def _fingerprint(pdf: Path, settings: Settings) -> str:
    st = pdf.stat(); payload = f"{st.st_size}:{st.st_mtime_ns}:{settings.render['dpi']}:{settings.render['max_pages']}:{settings.watermark}"; return hashlib.sha1(payload.encode()).hexdigest()[:12]

def process_pdf(pdf: Path, settings: Settings, watermark=True) -> tuple[str, list[str], str | None]:
    rel = pdf.relative_to(settings.source_root); category = rel.parts[0] if len(rel.parts) > 1 else "uncategorized"
    token = hashlib.sha1(str(rel).encode("utf-8")).hexdigest()[:8]
    folder = settings.output_root / "images" / safe_name(category) / (safe_name(pdf.stem) + "-" + token); folder.mkdir(parents=True, exist_ok=True)
    try:
        doc = fitz.open(pdf); title = (doc.metadata.get("title") or "").strip() if settings.table.get("include_metadata_title") else ""
        title = title or pdf.stem; count = min(len(doc), int(settings.render["max_pages"])); paths=[]; fp=_fingerprint(pdf, settings)
        marker = folder / ".fingerprint"
        for i in range(count):
            out = folder / f"page-{i+1:02d}.png"
            if out.exists() and marker.exists() and marker.read_text() == fp: paths.append(str(out.relative_to(settings.output_root)).replace(os.sep, "/")); continue
            pix = doc.load_page(i).get_pixmap(matrix=fitz.Matrix(float(settings.render["dpi"])/72, float(settings.render["dpi"])/72), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            if watermark and settings.watermark.get("enabled") and settings.watermark.get("text"):
                layer=Image.new("RGBA", img.size, (0,0,0,0)); d=ImageDraw.Draw(layer); f=_font(int(settings.watermark["font_size"])); box=d.textbbox((0,0), settings.watermark["text"], font=f); w,h=box[2]-box[0],box[3]-box[1]; margin=max(1,int(min(img.size)*.03)); pos={"bottom_right":(img.width-w-margin,img.height-h-margin),"bottom_left":(margin,img.height-h-margin),"center":((img.width-w)//2,(img.height-h)//2)}[settings.watermark["position"]]; d.text(pos, settings.watermark["text"], font=f, fill=(255,0,0,int(settings.watermark["opacity"]))) ; img=Image.alpha_composite(img.convert("RGBA"),layer).convert("RGB")
            tmp=out.with_suffix(".tmp.png"); img.save(tmp, "PNG"); os.replace(tmp,out); paths.append(str(out.relative_to(settings.output_root)).replace(os.sep,"/"))
        marker.write_text(fp, encoding="ascii"); doc.close(); return title, paths, None
    except Exception as exc:
        return pdf.stem, [], f"{type(exc).__name__}: {exc}"

def write_tables(rows: list[dict[str, Any]], settings: Settings, errors: list[dict[str,str]], elapsed: float) -> None:
    settings.output_root.mkdir(parents=True, exist_ok=True); xlsx=settings.output_root/settings.table["xlsx"]; csvp=settings.output_root/settings.table["csv"]
    wb=Workbook(); ws=wb.active; ws.title="PDF目录"; ws.append(HEADERS); ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:{get_column_letter(len(HEADERS))}{len(rows)+1}"
    for c in ws[1]: c.font=Font(bold=True); c.alignment=Alignment(horizontal="center")
    for row in rows: ws.append([row.get(h,"") for h in HEADERS])
    for col in range(1,len(HEADERS)+1): ws.column_dimensions[get_column_letter(col)].width = 22 if col<4 else 48
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.alignment=Alignment(vertical="top", wrap_text=True)
    wb.save(xlsx)
    with csvp.open("w", newline="", encoding="utf-8-sig") as f: w=csv.DictWriter(f, fieldnames=HEADERS); w.writeheader(); w.writerows(rows)
    with (settings.output_root/"errors.csv").open("w", newline="", encoding="utf-8-sig") as f: w=csv.DictWriter(f, fieldnames=["path","stage","error"]); w.writeheader(); w.writerows(errors)
    (settings.output_root/"run.log").write_text(f"发现数: {len(rows)}\n成功数: {len(rows)-len(errors)}\n失败数: {len(errors)}\n耗时秒: {elapsed:.2f}\n", encoding="utf-8")

def run(settings: Settings, mode="run", limit=None, no_watermark=False, max_pages=None, dry_run=False) -> dict[str,int]:
    if not settings.source_root.exists(): raise FileNotFoundError(f"源目录不存在: {settings.source_root}")
    files=discover(settings.source_root); lim = limit if limit is not None else settings.processing.get("max_pdfs"); files=files[:lim] if lim is not None else files
    rows=[]; errors=[]; start=time.time()
    for pdf in files:
        if dry_run: continue
        if mode == "scan": title, paths, err = pdf.stem, [], None
        else: title, paths, err = process_pdf(pdf, settings, not no_watermark)
        if err: errors.append({"path":str(pdf),"stage":"process","error":err})
        row={"年级":settings.grade,"学期":settings.semester,"PDF 文件名称":title,"PDF 文件所在位置":str(pdf)}
        for i,h in enumerate(HEADERS[4:]): row[h]=paths[i] if i<len(paths) else ""
        rows.append(row)
    if not dry_run: write_tables(rows, settings, errors, time.time()-start)
    return {"发现数":len(files),"成功数":len(rows)-len(errors),"失败数":len(errors),"生成图片数":sum(sum(bool(r[h]) for h in HEADERS[4:]) for r in rows)}
