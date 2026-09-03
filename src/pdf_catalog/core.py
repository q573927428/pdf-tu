"""PDF discovery, rendering and catalog output."""
from __future__ import annotations

import csv, hashlib, logging, os, re, time, shutil, json, urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pymupdf
import yaml
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

LOG = logging.getLogger(__name__)
HEADERS = ["年级", "学期", "科目", "分类", "PDF 文件名称", "PDF 文件所在位置", "生成文案", "封面图链接", *[f"转换后的图片 {i}" for i in range(1, 6)]]

@dataclass
class Settings:
    source_root: Path; output_root: Path; grade: str; semester: str
    watermark: dict[str, Any]; render: dict[str, Any]; processing: dict[str, Any]; table: dict[str, Any]; ai: dict[str, Any] = field(default_factory=dict)

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
    ai = {"enabled": False, "endpoint": "https://ark.cn-beijing.volces.com/api/v3", "api_key": "", "model": "", "generate_copy": False, "generate_cover": False, "timeout": 60, **raw.get("ai", {})}
    return Settings(source, out, str(req("grade")), str(req("semester")), wm, render, {"max_pdfs": None, **raw.get("processing", {})}, table, ai)

def discover(root: Path) -> list[Path]:
    return sorted((p for p in root.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"), key=lambda p: str(p.relative_to(root)).lower())

def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._") or "pdf"
    return value[:100]

def safe_dir_name(value: str) -> str:
    """保留中文目录名，仅替换 Windows 不允许的字符。"""
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    return value[:100] or "未分类"

def directory_fields(pdf: Path, settings: Settings) -> tuple[str, str, str]:
    """从 PDF 相对路径提取学期、科目和分类目录。"""
    parts = list(pdf.relative_to(settings.source_root).parts[:-1])
    semester = settings.semester
    if parts and parts[0] in {"上册", "下册"}:
        semester, parts = parts[0], parts[1:]
    return semester, (parts[0] if parts else ""), (parts[1] if len(parts) > 1 else "")

def _font(size: int):
    for p in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        if Path(p).exists():
            try: return ImageFont.truetype(p, size)
            except OSError: pass
    return ImageFont.load_default()

def _fingerprint(pdf: Path, settings: Settings) -> str:
    st = pdf.stat(); payload = f"{st.st_size}:{st.st_mtime_ns}:{settings.render['dpi']}:{settings.render['max_pages']}:{settings.watermark}"; return hashlib.sha1(payload.encode()).hexdigest()[:12]

def process_pdf(pdf: Path, settings: Settings, sequence: int, watermark=True) -> tuple[str, list[str], str | None]:
    rel = pdf.relative_to(settings.source_root); semester, subject, category = directory_fields(pdf, settings)
    token = hashlib.sha1(str(rel).encode("utf-8")).hexdigest()[:8]
    folder = settings.output_root / "images" / safe_dir_name(settings.grade) / safe_dir_name(semester) / safe_dir_name(subject or "未分类") / safe_dir_name(category or "未分类") / f"{sequence:06d}"; folder.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(pdf, folder / pdf.name)
        doc = pymupdf.open(pdf); title = (doc.metadata.get("title") or "").strip() if settings.table.get("include_metadata_title") else ""
        title = title or pdf.stem; count = min(len(doc), int(settings.render["max_pages"])); paths=[]; fp=_fingerprint(pdf, settings)
        marker = folder / ".fingerprint"
        for i in range(count):
            out = folder / f"page-{i+1:02d}.png"
            if out.exists() and marker.exists() and marker.read_text() == fp: paths.append(str(out.relative_to(settings.output_root)).replace(os.sep, "/")); continue
            pix = doc.load_page(i).get_pixmap(matrix=pymupdf.Matrix(float(settings.render["dpi"])/72, float(settings.render["dpi"])/72), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            if watermark and settings.watermark.get("enabled") and settings.watermark.get("text"):
                layer=Image.new("RGBA", img.size, (0,0,0,0)); d=ImageDraw.Draw(layer); f=_font(int(settings.watermark["font_size"])); box=d.textbbox((0,0), settings.watermark["text"], font=f); w,h=box[2]-box[0],box[3]-box[1]; margin=max(1,int(min(img.size)*.03)); pos={"bottom_right":(img.width-w-margin,img.height-h-margin),"bottom_left":(margin,img.height-h-margin),"center":((img.width-w)//2,(img.height-h)//2)}[settings.watermark["position"]]; d.text(pos, settings.watermark["text"], font=f, fill=(255,0,0,int(settings.watermark["opacity"]))) ; img=Image.alpha_composite(img.convert("RGBA"),layer).convert("RGB")
            tmp=out.with_suffix(".tmp.png"); img.save(tmp, "PNG"); os.replace(tmp,out); paths.append(str(out.relative_to(settings.output_root)).replace(os.sep,"/"))
        marker.write_text(fp, encoding="ascii"); doc.close(); return title, paths, None
    except Exception as exc:
        return pdf.stem, [], f"{type(exc).__name__}: {exc}"

def _doubao(settings: Settings, messages: list[dict[str, str]], *, image=False) -> Any:
    """调用豆包 Ark 的 OpenAI 兼容接口；未配置时明确报错。"""
    if not settings.ai.get("enabled") or not settings.ai.get("api_key") or not settings.ai.get("model"):
        raise RuntimeError("未配置 ai.enabled、ai.api_key 或 ai.model")
    if image:
        endpoint = str(settings.ai.get("image_endpoint") or settings.ai.get("endpoint", "")).rstrip("/") + "/images/generations"
        payload = {"model": settings.ai["model"], "prompt": messages[-1].get("content", ""), "size": settings.ai.get("image_size", "1024x1536"), "n": 1}
    else:
        endpoint = str(settings.ai.get("endpoint", "")).rstrip("/") + "/chat/completions"
        payload = {"model": settings.ai["model"], "messages": messages}
    req = urllib.request.Request(endpoint, data=json.dumps(payload, ensure_ascii=False).encode(), headers={"Authorization": "Bearer " + settings.ai["api_key"], "Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=float(settings.ai.get("timeout", 60))) as res:
        return json.loads(res.read().decode("utf-8"))

def generate_copy(settings: Settings, filename: str, category: str, grade_subject: str) -> str:
    prompt = f"根据文件名“{filename}”、资料类型“{category or '学习资料'}”、年级科目“{grade_subject}”，按儿童教辅营销要求写一条简体中文文案：100字以内，口语化有画面感，至少两种急迫感，紧扣文件名具体内容，结尾催促保存/打印。只输出文案。"
    data = _doubao(settings, [{"role": "user", "content": prompt}])
    content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
    if isinstance(content, list): content = "".join(x.get("text", "") for x in content if isinstance(x, dict))
    return str(content).strip()[:100]

def generate_cover(settings: Settings, copy: str, name: str, category: str, grade_subject: str) -> str:
    prompt = f"生成竖版儿童学习资料封面图。严格使用奶油白浅燕麦纸纹、雾蓝暖粉鹅黄低饱和手账风，无水印无logo。顶部居中标题取文案最抓眼一句不超过10字并加鹅黄波浪线；中部卡片写资料名称和亮点；底部10%写打印版·趁早存好。文案：{copy}；资料名称：{name}；亮点：{category or '查漏补缺·每日一练'}；年级科目：{grade_subject}。文字清晰。"
    data = _doubao(settings, [{"role": "user", "content": prompt}], image=True)
    def find(v):
        if isinstance(v, str) and (v.startswith("http://") or v.startswith("https://")): return v
        if isinstance(v, dict):
            for x in v.values():
                got = find(x)
                if got: return got
        if isinstance(v, list):
            for x in v:
                got = find(x)
                if got: return got
        return ""
    return find(data)

def write_tables(rows: list[dict[str, Any]], settings: Settings, errors: list[dict[str,str]], elapsed: float, details: list[str] | None = None) -> None:
    settings.output_root.mkdir(parents=True, exist_ok=True); xlsx=settings.output_root/settings.table["xlsx"]; csvp=settings.output_root/settings.table["csv"]
    wb=Workbook(); ws=wb.active; ws.title="PDF目录"; ws.append(HEADERS); ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:{get_column_letter(len(HEADERS))}{len(rows)+1}"
    for c in ws[1]: c.font=Font(bold=True); c.alignment=Alignment(horizontal="center")
    for row in rows: ws.append([row.get(h,"") for h in HEADERS])
    for col in range(1,len(HEADERS)+1): ws.column_dimensions[get_column_letter(col)].width = 22 if col<4 else 48
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.alignment=Alignment(vertical="top", wrap_text=True)
    try:
        wb.save(xlsx)
        with csvp.open("w", newline="", encoding="utf-8-sig") as f: w=csv.DictWriter(f, fieldnames=HEADERS); w.writeheader(); w.writerows(rows)
        with (settings.output_root/"errors.csv").open("w", newline="", encoding="utf-8-sig") as f: w=csv.DictWriter(f, fieldnames=["path","stage","error"]); w.writeheader(); w.writerows(errors)
        summary = f"发现数: {len(rows)}\n成功数: {len(rows)-len(errors)}\n失败数: {len(errors)}\n耗时秒: {elapsed:.2f}\n"
        detail_text = "\n".join(details or [])
        (settings.output_root/"run.log").write_text((detail_text + "\n\n" if detail_text else "") + summary, encoding="utf-8")
    except PermissionError as exc:
        raise PermissionError(f"输出文件被占用，请关闭 Excel/编辑器后重试: {exc.filename}") from exc

def run(settings: Settings, mode="run", limit=None, no_watermark=False, max_pages=None, dry_run=False, start=1, end=None, ai_start=None, ai_end=None, ai_limit=None, generate_copy_flag=False, generate_cover_flag=False) -> dict[str,int]:
    if not settings.source_root.exists(): raise FileNotFoundError(f"源目录不存在: {settings.source_root}")
    LOG.info("开始扫描 PDF 文件: %s", settings.source_root)
    discovered = discover(settings.source_root)
    LOG.info("文件扫描完成: 共发现 %d 个 PDF", len(discovered))
    lim = limit if limit is not None else settings.processing.get("max_pdfs")
    range_start = max(1, int(start or 1)); end = int(end) if end is not None else None
    files = discovered[range_start - 1:end]
    if lim is not None: files = files[:int(lim)]
    if lim is not None:
        LOG.info("文件筛选完成: 限制数量=%d，实际处理 %d 个 PDF", lim, len(files))
    else:
        LOG.info("文件筛选完成: 未设置数量限制，实际处理 %d 个 PDF", len(files))
    rows=[]; errors=[]; details=[]; run_start=time.time()
    start_detail = f"开始处理: 源目录={settings.source_root}，待处理 PDF {len(files)} 个，模式={mode}"
    LOG.info(start_detail)
    details.append(start_detail)
    for idx, pdf in enumerate(files):
        if dry_run:
            detail = f"[{idx + 1}/{len(files)}] 扫描完成: {pdf}"
            LOG.info(detail)
            details.append(detail)
            continue
        if mode in {"scan", "ai"}: title, paths, err = pdf.stem, [], None
        else: title, paths, err = process_pdf(pdf, settings, idx + 1, not no_watermark)
        if err:
            errors.append({"path":str(pdf),"stage":"process","error":err})
            detail = f"[{idx + 1}/{len(files)}] 失败: {pdf} ({err})"
            LOG.error(detail)
        else:
            detail = f"[{idx + 1}/{len(files)}] 完成: {pdf}，生成图片 {len(paths)} 张"
            LOG.info(detail)
        details.append(detail)
        semester, subject, category = directory_fields(pdf, settings)
        row={"年级":settings.grade,"学期":semester,"科目":subject,"分类":category,"PDF 文件名称":title,"PDF 文件所在位置":str(pdf),"生成文案":"","封面图链接":""}
        for i,h in enumerate(HEADERS[8:]): row[h]=paths[i] if i<len(paths) else ""
        ai_index = range_start + idx
        in_ai_range = (ai_start is None or ai_index >= int(ai_start)) and (ai_end is None or ai_index <= int(ai_end))
        if ai_limit is not None and idx >= int(ai_limit): in_ai_range = False
        if in_ai_range and (generate_copy_flag or settings.ai.get("generate_copy")):
            try: row["生成文案"] = generate_copy(settings, pdf.name, category, f"{settings.grade}{subject}")
            except Exception as exc: errors.append({"path":str(pdf),"stage":"copy","error":f"{type(exc).__name__}: {exc}"})
        if in_ai_range and (generate_cover_flag or settings.ai.get("generate_cover")):
            try:
                cover_copy = row["生成文案"] or title
                row["封面图链接"] = generate_cover(settings, cover_copy, title, category, f"{settings.grade}{subject}")
            except Exception as exc: errors.append({"path":str(pdf),"stage":"cover","error":f"{type(exc).__name__}: {exc}"})
        rows.append(row)
    if not dry_run:
        LOG.info("开始写入目录文件: %s", settings.output_root)
        write_tables(rows, settings, errors, time.time()-run_start, details)
        LOG.info("目录文件写入完成: %s", settings.output_root)
    return {"发现数":len(files),"成功数":len(rows)-len(errors),"失败数":len(errors),"生成图片数":sum(sum(bool(r[h]) for h in HEADERS[8:]) for r in rows)}
