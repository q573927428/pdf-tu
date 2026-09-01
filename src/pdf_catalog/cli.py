"""Command line interface for PDF catalog builder."""
from __future__ import annotations
import argparse, logging
from .core import load_settings, run

def main() -> None:
    parser = argparse.ArgumentParser(prog="pdf-catalog", description="扫描 PDF 并生成目录")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("scan", "convert", "run"):
        p = sub.add_parser(name)
        p.add_argument("--config", required=True)
        p.add_argument("--limit", default=None, help="数量或 none")
        p.add_argument("--no-watermark", action="store_true")
        p.add_argument("--max-pages", type=int)
        p.add_argument("--dry-run", action="store_true")
        p.add_argument("--log-level", default="INFO")
    args = parser.parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level.upper(), logging.INFO), format="%(asctime)s %(levelname)s %(message)s")
    try:
        settings = load_settings(args.config)
        if args.max_pages is not None: settings.render["max_pages"] = args.max_pages
        limit = None if args.limit in (None, "none", "None") else int(args.limit)
        stats = run(settings, args.command, limit, args.no_watermark, args.max_pages, args.dry_run)
        print("; ".join(f"{k}: {v}" for k, v in stats.items()))
    except Exception as exc:
        parser.error(str(exc))

if __name__ == "__main__":
    main()
