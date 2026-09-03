"""Command line interface for PDF catalog builder."""
from __future__ import annotations
import argparse, logging
import pymupdf
from .core import load_settings, run, serve

def main() -> None:
    parser = argparse.ArgumentParser(prog="pdf-catalog", description="扫描 PDF 并生成目录")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("scan", "convert", "run", "ai", "serve"):
        p = sub.add_parser(name)
        p.add_argument("--config", required=True)
        if name == "serve":
            p.add_argument("--host", default="127.0.0.1")
            p.add_argument("--port", type=int, default=8765)
            continue
        p.add_argument("--limit", default=None, help="数量或 none")
        p.add_argument("--start", "--start-index", dest="start", type=int, default=1, help="扫描起始序号（从1开始，包含）")
        p.add_argument("--end", "--end-index", dest="end", type=int, help="扫描结束序号（包含）")
        p.add_argument("--ai-start", "--generate-start", dest="ai_start", type=int, help="AI生成起始序号（从1开始）")
        p.add_argument("--ai-end", "--generate-end", dest="ai_end", type=int, help="AI生成结束序号（包含）")
        p.add_argument("--ai-limit", "--generate-count", dest="ai_limit", type=int, help="AI生成最多数量")
        p.add_argument("--generate-copy", action="store_true", help="调用豆包生成营销文案")
        p.add_argument("--generate-cover", action="store_true", help="调用豆包生成封面图并写入链接")
        p.add_argument("--no-watermark", action="store_true")
        p.add_argument("--max-pages", type=int)
        p.add_argument("--dry-run", action="store_true")
        p.add_argument("--log-level", default="INFO")
    args = parser.parse_args()
    log_level = getattr(args, "log_level", "INFO")
    logging.basicConfig(level=getattr(logging, log_level.upper(), logging.INFO), format="%(asctime)s %(levelname)s %(message)s")
    # PDF 解析仍会自动修复部分结构问题，但不把 MuPDF 的底层诊断刷到终端。
    pymupdf.TOOLS.mupdf_display_errors(False)
    try:
        settings = load_settings(args.config)
        if args.command == "serve":
            serve(settings, args.host, args.port)
            return
        if args.max_pages is not None: settings.render["max_pages"] = args.max_pages
        limit = None if args.limit in (None, "none", "None") else int(args.limit)
        generate_copy = args.generate_copy or args.command == "ai"
        stats = run(settings, args.command, limit, args.no_watermark, args.max_pages, args.dry_run, args.start, args.end, args.ai_start, args.ai_end, args.ai_limit, generate_copy, args.generate_cover)
        print("; ".join(f"{k}: {v}" for k, v in stats.items()))
    except Exception as exc:
        parser.error(str(exc))

if __name__ == "__main__":
    main()
