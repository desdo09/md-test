"""
main.py
-------
CLI entry point for the document → Markdown + PDF pipeline.

Usage
-----
  python main.py                          # convert all files in ./input/
  python main.py path/to/file.docx        # convert a single file
  python main.py path/to/dir/ -o exports  # custom output folder
  python main.py path/to/file.pdf --no-pdf  # Markdown only, skip PDF

The two worker modules are:
  md_converter.py  – document  → Markdown  (via MarkItDown)
  pdf_converter.py – Markdown  → PDF       (via fpdf2 + python-bidi)
"""

import argparse
import io
import sys
from pathlib import Path

# Make stdout UTF-8 safe on Windows (Hebrew filenames, etc.)
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from md_converter import convert_to_markdown, extract_markdown
from pdf_converter import build_pdf


# ---------------------------------------------------------------------------
# Single-file conversion
# ---------------------------------------------------------------------------

def convert_file(
    src: Path,
    output_dir: Path,
    generate_pdf: bool = True,
) -> dict:
    """Convert *src* to Markdown (and optionally PDF) inside *output_dir*.

    Returns a dict with keys:
        ``md_path``  – Path to the generated .md file, or None on error.
        ``pdf_path`` – Path to the generated .pdf file, or None.
        ``error``    – Error message string, or None on success.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Step 1: extract Markdown ─────────────────────────────────────────────
    try:
        md_text = extract_markdown(src)
    except Exception as exc:
        return {"md_path": None, "pdf_path": None, "error": str(exc)}

    md_path = output_dir / f"{src.stem}.md"
    md_path.write_text(md_text, encoding="utf-8")

    # ── Step 2: build PDF (optional) ─────────────────────────────────────────
    pdf_path = None
    if generate_pdf:
        try:
            pdf_path = output_dir / f"{src.stem}.pdf"
            build_pdf(md_text, pdf_path)
        except Exception as exc:
            print(f"  [WARNING] PDF generation failed for {src.name}: {exc}")
            pdf_path = None

    return {"md_path": md_path, "pdf_path": pdf_path, "error": None}


# ---------------------------------------------------------------------------
# Directory conversion
# ---------------------------------------------------------------------------

def convert_directory(
    input_dir: Path,
    output_dir: Path,
    generate_pdf: bool = True,
) -> None:
    """Convert every file in *input_dir* to Markdown (and optionally PDF)."""
    files = [f for f in sorted(input_dir.iterdir()) if f.is_file()]

    if not files:
        print(f"No files found in {input_dir}. Place documents there and re-run.")
        return

    for src in files:
        print(f"\nConverting: {src.name}")
        info = convert_file(src, output_dir, generate_pdf=generate_pdf)

        if info["error"]:
            print(f"  [ERROR] {info['error']}")
        else:
            print(f"  Markdown -> {info['md_path']}")
            if info["pdf_path"]:
                print(f"  PDF      -> {info['pdf_path']}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert documents to Markdown and PDF using MarkItDown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "source",
        nargs="?",
        help="File or directory to convert. Defaults to ./input/",
    )
    parser.add_argument(
        "--output", "-o",
        default="output",
        metavar="DIR",
        help="Output directory (default: ./output/)",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Generate Markdown only; skip PDF.",
    )
    args = parser.parse_args()

    src_path = Path(args.source) if args.source else Path("input")
    out_dir  = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    if src_path.is_dir():
        convert_directory(src_path, out_dir, generate_pdf=not args.no_pdf)

    elif src_path.is_file():
        print(f"\nConverting: {src_path.name}")
        info = convert_file(src_path, out_dir, generate_pdf=not args.no_pdf)
        if info["error"]:
            print(f"  [ERROR] {info['error']}")
            sys.exit(1)
        print(f"  Markdown -> {info['md_path']}")
        if info["pdf_path"]:
            print(f"  PDF      -> {info['pdf_path']}")

    else:
        print(f"[ERROR] Source not found: {src_path}")
        sys.exit(1)


if __name__ == "__main__":
    main()
