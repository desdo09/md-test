"""
convert.py
----------
Converts documents in the `input/` folder to Markdown (and optionally PDF)
using Microsoft's MarkItDown library.

Supported input formats (handled automatically by MarkItDown):
  PDF, Word (.docx), Excel (.xlsx/.xls), PowerPoint (.pptx),
  HTML, CSV, JSON, XML, Images, Audio, EPubs, ZIP, and more.

Usage:
  python convert.py                        # converts all files in ./input/
  python convert.py path/to/file.docx      # converts a single file
  python convert.py path/to/file.pdf --no-pdf   # skip PDF output
"""

import argparse
import re
import sys
from pathlib import Path

from markitdown import MarkItDown


# ---------------------------------------------------------------------------
# PDF generation (optional)
# ---------------------------------------------------------------------------

def _build_pdf(md_text: str, pdf_path: Path) -> None:
    """Render a Markdown string to a PDF file using fpdf2.

    Supports Unicode (Hebrew, Arabic, CJK, …) as long as the system font
    covers the characters.  RTL scripts (Hebrew / Arabic) are handled with
    python-bidi so text is displayed in the correct visual order.
    """
    from fpdf import FPDF

    # ── RTL helpers ──────────────────────────────────────────────────────────
    _RTL_RANGES = (
        (0x0590, 0x05FF),   # Hebrew
        (0x0600, 0x06FF),   # Arabic
        (0x0750, 0x077F),   # Arabic Supplement
        (0xFB00, 0xFDFF),   # Alphabetic Presentation / Arabic Presentation-A
        (0xFE70, 0xFEFF),   # Arabic Presentation-B
    )

    def _is_rtl(text: str) -> bool:
        """Return True if the text contains any RTL characters."""
        for ch in text:
            cp = ord(ch)
            if any(lo <= cp <= hi for lo, hi in _RTL_RANGES):
                return True
        return False

    def _bidi(text: str) -> str:
        """Apply the Unicode BiDi algorithm for correct visual ordering."""
        try:
            from bidi.algorithm import get_display
            return get_display(text)
        except ImportError:
            return text  # graceful degradation if python-bidi is missing

    # Locate Unicode-capable TTF fonts on this machine (Windows / macOS / Linux).
    def _find_unicode_fonts() -> dict:
        """Return dict with keys regular/bold/italic/mono pointing to TTF paths."""
        import os

        def first(paths):
            return next((p for p in paths if os.path.exists(p)), "")

        # ---- Windows ----
        win_mono = first([
            r"C:\Windows\Fonts\consola.ttf",   # Consolas
            r"C:\Windows\Fonts\lucon.ttf",     # Lucida Console
        ])
        win = {
            "regular": first([r"C:\Windows\Fonts\arial.ttf",
                               r"C:\Windows\Fonts\calibri.ttf"]),
            "bold":    first([r"C:\Windows\Fonts\arialbd.ttf",
                               r"C:\Windows\Fonts\calibrib.ttf"]),
            "italic":  first([r"C:\Windows\Fonts\ariali.ttf",
                               r"C:\Windows\Fonts\calibrii.ttf"]),
            "mono":    win_mono,
        }
        if win["regular"]:
            return win

        # ---- macOS ----
        mac = {
            "regular": first(["/Library/Fonts/Arial.ttf",
                               "/System/Library/Fonts/Helvetica.ttc"]),
            "bold":    first(["/Library/Fonts/Arial Bold.ttf"]),
            "italic":  first(["/Library/Fonts/Arial Italic.ttf"]),
            "mono":    first(["/Library/Fonts/Courier New.ttf"]),
        }
        if mac["regular"]:
            return mac

        # ---- Linux (DejaVu) ----
        dv_base = "/usr/share/fonts/truetype/dejavu"
        linux = {
            "regular": f"{dv_base}/DejaVuSans.ttf",
            "bold":    f"{dv_base}/DejaVuSans-Bold.ttf",
            "italic":  f"{dv_base}/DejaVuSans-Oblique.ttf",
            "mono":    f"{dv_base}/DejaVuSansMono.ttf",
        }
        if os.path.exists(linux["regular"]):
            return linux

        return {"regular": "", "bold": "", "italic": "", "mono": ""}

    _FONTS = _find_unicode_fonts()

    class MarkdownPDF(FPDF):
        """Minimal Markdown → PDF renderer (headings, lists, bold, italic, code)."""

        CODE_BG = (240, 240, 240)
        HEADING_SIZES = {1: 20, 2: 16, 3: 13, 4: 11}

        # Font family names (set once in __init__)
        _BODY_FAMILY = "Body"
        _MONO_FAMILY = "Mono"

        def __init__(self):
            super().__init__()
            self.set_auto_page_break(auto=True, margin=15)

            # Register Unicode TTF fonts when available
            if _FONTS["regular"]:
                reg  = _FONTS["regular"]
                bold = _FONTS["bold"]  or reg
                ital = _FONTS["italic"] or reg
                mono = _FONTS["mono"]  or reg
                self.add_font(self._BODY_FAMILY, style="",  fname=reg)
                self.add_font(self._BODY_FAMILY, style="B", fname=bold)
                self.add_font(self._BODY_FAMILY, style="I", fname=ital)
                self.add_font(self._MONO_FAMILY, style="",  fname=mono)
            else:
                # Fall back to built-in (Latin-only) fonts
                self._BODY_FAMILY = "Helvetica"
                self._MONO_FAMILY = "Courier"

            self.add_page()
            self.set_margins(20, 20, 20)
            self._set_body()

        # ---- font helpers ----

        def _set_body(self):
            self.set_font(self._BODY_FAMILY, size=10)
            self.set_text_color(30, 30, 30)

        def _set_heading(self, level: int):
            size = self.HEADING_SIZES.get(level, 10)
            self.set_font(self._BODY_FAMILY, style="B", size=size)
            self.set_text_color(0, 0, 0)

        def _set_code(self):
            self.set_font(self._MONO_FAMILY, size=9)
            self.set_text_color(50, 50, 50)

        # ---- inline markup ----

        def _rtl_cell(self, text: str, line_height: int = 5) -> None:
            """Render a single RTL line right-aligned on the full text width."""
            self.set_x(self.l_margin)
            self.multi_cell(self.w - self.l_margin - self.r_margin,
                            line_height, _bidi(text), align="R")

        def _write_inline(self, text: str):
            """Write text with basic bold/italic/code inline markup.

            For RTL scripts (Hebrew, Arabic) the bidi algorithm is applied and
            the text is placed in a right-aligned cell instead of using write()
            so that the visual direction is correct.
            """
            if _is_rtl(text):
                # RTL path: emit the whole line as a right-aligned cell
                self._rtl_cell(text)
                return

            # LTR path: inline markup with write()
            pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*|`.*?`)")
            parts = pattern.split(text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    self.set_font(self._BODY_FAMILY, style="B", size=self.font_size)
                    self.write(5, part[2:-2])
                    self.set_font(self._BODY_FAMILY, size=self.font_size)
                elif part.startswith("*") and part.endswith("*"):
                    self.set_font(self._BODY_FAMILY, style="I", size=self.font_size)
                    self.write(5, part[1:-1])
                    self.set_font(self._BODY_FAMILY, size=self.font_size)
                elif part.startswith("`") and part.endswith("`"):
                    self.set_font(self._MONO_FAMILY, size=self.font_size)
                    self.write(5, part[1:-1])
                    self.set_font(self._BODY_FAMILY, size=self.font_size)
                else:
                    self.write(5, part)

        # ---- block rendering ----

        def render_markdown(self, md: str):
            lines = md.splitlines()
            in_code_block = False
            code_lines: list[str] = []

            for raw_line in lines:
                # ── fenced code block ──────────────────────────────────────
                if raw_line.startswith("```"):
                    if not in_code_block:
                        in_code_block = True
                        code_lines = []
                    else:
                        in_code_block = False
                        self._render_code_block(code_lines)
                    continue

                if in_code_block:
                    code_lines.append(raw_line)
                    continue

                line = raw_line.rstrip()

                # ── horizontal rule ────────────────────────────────────────
                if re.match(r"^(-{3,}|_{3,}|\*{3,})$", line):
                    self.ln(3)
                    self.set_draw_color(180, 180, 180)
                    self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
                    self.ln(3)
                    self._set_body()
                    continue

                # ── headings ───────────────────────────────────────────────
                heading_match = re.match(r"^(#{1,4})\s+(.*)", line)
                if heading_match:
                    level = len(heading_match.group(1))
                    text = heading_match.group(2)
                    self.ln(4)
                    self._set_heading(level)
                    if _is_rtl(text):
                        self._rtl_cell(text, line_height=8)
                    else:
                        self.multi_cell(0, 8, text)
                    self.ln(1)
                    self._set_body()
                    continue

                # ── unordered list item ────────────────────────────────────
                list_match = re.match(r"^(\s*)[-*+]\s+(.*)", line)
                if list_match:
                    indent = len(list_match.group(1))
                    text = list_match.group(2)
                    self._set_body()
                    if _is_rtl(text):
                        # RTL: bullet on right, text right-aligned
                        self._rtl_cell(_bidi(text) + " \u2022")
                    else:
                        self.set_x(20 + indent * 3)
                        self.write(5, "\u2022 ")
                        self._write_inline(text)
                        self.ln()
                    continue

                # ── ordered list item ──────────────────────────────────────
                ol_match = re.match(r"^(\s*)\d+\.\s+(.*)", line)
                if ol_match:
                    indent = len(ol_match.group(1))
                    num_text = line.lstrip()
                    self._set_body()
                    if _is_rtl(num_text):
                        self._rtl_cell(num_text)
                    else:
                        self.set_x(20 + indent * 3)
                        self._write_inline(num_text)
                        self.ln()
                    continue

                # ── blockquote ─────────────────────────────────────────────
                if line.startswith("> "):
                    content = line[2:]
                    self.set_font(self._BODY_FAMILY, style="I", size=9)
                    self.set_text_color(100, 100, 100)
                    if _is_rtl(content):
                        self._rtl_cell(content)
                    else:
                        self.set_x(25)
                        self.multi_cell(0, 5, content)
                    self._set_body()
                    continue

                # ── blank line → paragraph spacing ─────────────────────────
                if not line.strip():
                    self.ln(4)
                    continue

                # ── normal paragraph text ──────────────────────────────────
                self._set_body()
                self.set_x(20)
                self._write_inline(line)
                if not _is_rtl(line):   # _write_inline already adds ln for RTL
                    self.ln()

        def _render_code_block(self, lines: list[str]):
            """Render a fenced code block with a shaded background."""
            text = "\n".join(lines)
            self.ln(2)
            self._set_code()
            self.set_fill_color(*self.CODE_BG)
            self.multi_cell(0, 5, text, fill=True)
            self.ln(2)
            self._set_body()

    pdf = MarkdownPDF()
    pdf.render_markdown(md_text)
    pdf.output(str(pdf_path))


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------

def convert_file(src: Path, output_dir: Path, generate_pdf: bool = True) -> dict:
    """Convert a single file to Markdown (and optionally PDF).

    Returns a dict with keys: 'md_path', 'pdf_path' (or None), 'error' (or None).
    """
    md_converter = MarkItDown()

    try:
        result = md_converter.convert(str(src))
    except Exception as exc:
        return {"md_path": None, "pdf_path": None, "error": str(exc)}

    md_text = result.text_content

    # Save Markdown
    md_path = output_dir / f"{src.stem}.md"
    md_path.write_text(md_text, encoding="utf-8")

    pdf_path = None
    if generate_pdf:
        try:
            pdf_path = output_dir / f"{src.stem}.pdf"
            _build_pdf(md_text, pdf_path)
        except Exception as exc:
            print(f"  [WARNING] PDF generation failed for {src.name}: {exc}")
            pdf_path = None

    return {"md_path": md_path, "pdf_path": pdf_path, "error": None}


def convert_directory(input_dir: Path, output_dir: Path, generate_pdf: bool = True):
    """Convert all supported files found in input_dir."""
    # MarkItDown will reject files it cannot handle, so we try everything
    # that is not already a plain .md or .txt (those it can handle too, but
    # there's little point converting them here).
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
# CLI entry-point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert documents to Markdown (and PDF) using MarkItDown."
    )
    parser.add_argument(
        "source",
        nargs="?",
        help="Path to a single file OR a directory. Defaults to ./input/",
    )
    parser.add_argument(
        "--output", "-o",
        default="output",
        help="Output directory (default: ./output/)",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Skip PDF generation.",
    )
    args = parser.parse_args()

    src_path = Path(args.source) if args.source else Path("input")
    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    if src_path.is_dir():
        convert_directory(src_path, out_dir, generate_pdf=not args.no_pdf)
    elif src_path.is_file():
        print(f"\nConverting: {src_path.name}")
        info = convert_file(src_path, out_dir, generate_pdf=not args.no_pdf)
        if info["error"]:
            print(f"  [ERROR] {info['error']}")
            sys.exit(1)
        else:
            print(f"  Markdown -> {info['md_path']}")
            if info["pdf_path"]:
                print(f"  PDF      -> {info['pdf_path']}")
    else:
        print(f"[ERROR] Source not found: {src_path}")
        sys.exit(1)


if __name__ == "__main__":
    main()
