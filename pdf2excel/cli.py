"""Command-line interface: pdf2excel input.pdf [-o output.xlsx]."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .converter import convert


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf2excel",
        description="Extract tables from PDF files into Excel (.xlsx) workbooks.",
    )
    p.add_argument("inputs", nargs="+", type=Path, help="PDF file(s) or folder(s) of PDFs")
    p.add_argument("-o", "--output", type=Path,
                   help="Output .xlsx file (one input) or folder (several inputs)")
    p.add_argument("-p", "--pages", help="Pages to convert, e.g. '1,3-5' (default: all)")
    p.add_argument("-s", "--single-sheet", action="store_true",
                   help="Put all tables on one sheet instead of one sheet per table")
    p.add_argument("--keep-text", action="store_true",
                   help="Keep numbers as text instead of converting them to numeric cells")
    p.add_argument("--no-text-fallback", action="store_true",
                   help="Skip pages with no tables instead of writing their plain text")
    p.add_argument("--password", help="Password for encrypted PDFs")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def _collect(inputs: list[Path]) -> list[Path]:
    files: list[Path] = []
    for path in inputs:
        if path.is_dir():
            files.extend(sorted(f for f in path.iterdir() if f.suffix.lower() == ".pdf"))
        else:
            files.append(path)
    return files


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    pdfs = _collect(args.inputs)
    if not pdfs:
        print("No PDF files found.", file=sys.stderr)
        return 1

    failures = 0
    for pdf in pdfs:
        if args.output is None:
            out = None
        elif len(pdfs) == 1 and args.output.suffix.lower() == ".xlsx":
            out = args.output
        else:
            out = args.output / pdf.with_suffix(".xlsx").name
        try:
            path, count = convert(
                pdf, out,
                pages=args.pages,
                single_sheet=args.single_sheet,
                convert_numbers=not args.keep_text,
                text_fallback=not args.no_text_fallback,
                password=args.password,
            )
            print(f"{pdf} -> {path} ({count} table{'s' if count != 1 else ''})")
        except Exception as exc:  # report and continue with the next file
            failures += 1
            print(f"Error converting {pdf}: {exc}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
