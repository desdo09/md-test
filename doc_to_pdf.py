"""
doc_to_pdf.py
-------------
Converts an original document file **directly** to PDF — no Markdown
intermediate file is saved.

Conversion strategy per format
--------------------------------
| Format          | Engine                                      |
|-----------------|---------------------------------------------|
| .pdf            | Already a PDF — copied to the output folder |
| .docx           | docx2pdf  (drives MS Word via COM/LibreOffice) |
| .html / .htm    | WeasyPrint (CSS-based HTML → PDF renderer)  |
| Everything else | MarkItDown → pdf_converter fallback         |

Public API
----------
  convert_to_pdf(src, output_dir) -> Path
      Convert *src* to a PDF file in *output_dir* and return the path.
"""

from __future__ import annotations

import shutil
from pathlib import Path


# ---------------------------------------------------------------------------
# Per-format converters
# ---------------------------------------------------------------------------

def _pdf_copy(src: Path, pdf_path: Path) -> None:
    """Source is already a PDF — just copy it."""
    shutil.copy2(src, pdf_path)


def _docx_to_pdf(src: Path, pdf_path: Path) -> None:
    """Convert a Word document to PDF using docx2pdf.

    On Windows this drives Microsoft Word via COM automation, producing a
    pixel-perfect PDF that matches what Word would print.
    On macOS/Linux it uses LibreOffice if available.
    """
    from docx2pdf import convert
    convert(str(src), str(pdf_path))


def _html_to_pdf(src: Path, pdf_path: Path) -> None:
    """Convert an HTML file to PDF using WeasyPrint.

    WeasyPrint produces high-fidelity CSS-rendered PDFs but requires the
    GTK3 runtime on Windows (https://doc.courtbouillon.org/weasyprint).
    If the native libraries are not found, the function falls back to the
    MarkItDown → pdf_converter pipeline automatically.
    """
    try:
        from weasyprint import HTML
        HTML(filename=str(src)).write_pdf(str(pdf_path))
    except Exception as exc:
        if "cannot load library" in str(exc) or "OSError" in type(exc).__name__:
            # GTK3 runtime not installed — fall back to Markdown pipeline
            print(
                "  [INFO] WeasyPrint requires GTK3 (not found). "
                "Falling back to Markdown pipeline for HTML."
            )
            _markdown_pipeline_fallback(src, pdf_path)
        else:
            raise


def _markdown_pipeline_fallback(src: Path, pdf_path: Path) -> None:
    """Fallback: MarkItDown → Markdown text → pdf_converter → PDF.

    Used for formats that have no dedicated direct converter
    (.xlsx, .csv, .json, .pptx, …).
    """
    from md_converter import extract_markdown
    from pdf_converter import build_pdf

    md_text = extract_markdown(src)
    build_pdf(md_text, pdf_path)


# ---------------------------------------------------------------------------
# Dispatch table
# ---------------------------------------------------------------------------

# Maps lowercase file suffix → converter function
_CONVERTERS = {
    ".pdf":  _pdf_copy,
    ".docx": _docx_to_pdf,
    ".doc":  _docx_to_pdf,
    ".html": _html_to_pdf,
    ".htm":  _html_to_pdf,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def convert_to_pdf(src: Path, output_dir: Path) -> Path:
    """Convert the document at *src* directly to a PDF file in *output_dir*.

    The output file is named ``<src.stem>.pdf``.

    Args:
        src:        Path to the source document.
        output_dir: Directory where the PDF will be saved.

    Returns:
        Path to the generated PDF file.

    Raises:
        FileNotFoundError: if *src* does not exist.
        Exception:         propagated from the underlying converter on failure.
    """
    if not src.exists():
        raise FileNotFoundError(f"Source file not found: {src}")

    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = output_dir / f"{src.stem}.pdf"

    converter = _CONVERTERS.get(src.suffix.lower(), _markdown_pipeline_fallback)
    converter(src, pdf_path)

    return pdf_path


# ---------------------------------------------------------------------------
# CLI (python doc_to_pdf.py <file_or_dir> [-o output_dir])
# ---------------------------------------------------------------------------

def _cli() -> None:
    import argparse
    import io
    import sys

    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )

    parser = argparse.ArgumentParser(
        description="Convert documents directly to PDF (no Markdown saved)."
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
    args = parser.parse_args()

    src_path = Path(args.source) if args.source else Path("input")
    out_dir  = Path(args.output)

    if src_path.is_file():
        files = [src_path]
    elif src_path.is_dir():
        files = sorted(f for f in src_path.iterdir() if f.is_file())
        if not files:
            print(f"No files found in {src_path}.")
            sys.exit(0)
    else:
        print(f"[ERROR] Source not found: {src_path}")
        sys.exit(1)

    for src in files:
        print(f"\nConverting: {src.name}")
        try:
            pdf_path = convert_to_pdf(src, out_dir)
            print(f"  PDF -> {pdf_path}")
        except Exception as exc:
            print(f"  [ERROR] {exc}")


if __name__ == "__main__":
    _cli()
