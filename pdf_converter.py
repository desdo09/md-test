"""
pdf_converter.py
----------------
Renders a Markdown string to a PDF file using fpdf2.

Features
--------
- Unicode fonts (auto-detected on Windows, macOS, Linux).
- Full RTL support for Hebrew and Arabic via python-bidi.
- Renders headings, paragraphs, lists, blockquotes, inline code, fenced code
  blocks, horizontal rules, and bold/italic markup.

Public API
----------
  build_pdf(md_text, pdf_path)
      Render *md_text* (a Markdown string) and write the PDF to *pdf_path*.
"""

from __future__ import annotations

import os
import re
from pathlib import Path

# ---------------------------------------------------------------------------
# RTL helpers (module-level so they are computed once)
# ---------------------------------------------------------------------------

_RTL_RANGES = (
    (0x0590, 0x05FF),   # Hebrew
    (0x0600, 0x06FF),   # Arabic
    (0x0750, 0x077F),   # Arabic Supplement
    (0xFB00, 0xFDFF),   # Alphabetic Presentation / Arabic Presentation-A
    (0xFE70, 0xFEFF),   # Arabic Presentation-B
)


def _is_rtl(text: str) -> bool:
    """Return True if *text* contains any RTL (Hebrew / Arabic) characters."""
    for ch in text:
        cp = ord(ch)
        if any(lo <= cp <= hi for lo, hi in _RTL_RANGES):
            return True
    return False


def _bidi(text: str) -> str:
    """Apply the Unicode Bidirectional Algorithm for correct visual ordering.

    Requires the ``python-bidi`` package.  Falls back gracefully if missing.
    """
    try:
        from bidi.algorithm import get_display
        return get_display(text)
    except ImportError:
        return text


# ---------------------------------------------------------------------------
# Font discovery (module-level, computed once)
# ---------------------------------------------------------------------------

def _find_unicode_fonts() -> dict[str, str]:
    """Return a dict with keys ``regular``, ``bold``, ``italic``, ``mono``
    pointing to TTF file paths on the current OS.  Empty strings mean the
    key was not found; the caller should fall back to a built-in font."""

    def first(paths: list[str]) -> str:
        return next((p for p in paths if os.path.exists(p)), "")

    # ---- Windows --------------------------------------------------------
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

    # ---- macOS ----------------------------------------------------------
    mac = {
        "regular": first(["/Library/Fonts/Arial.ttf",
                           "/System/Library/Fonts/Helvetica.ttc"]),
        "bold":    first(["/Library/Fonts/Arial Bold.ttf"]),
        "italic":  first(["/Library/Fonts/Arial Italic.ttf"]),
        "mono":    first(["/Library/Fonts/Courier New.ttf"]),
    }
    if mac["regular"]:
        return mac

    # ---- Linux (DejaVu) -------------------------------------------------
    dv = "/usr/share/fonts/truetype/dejavu"
    linux = {
        "regular": f"{dv}/DejaVuSans.ttf",
        "bold":    f"{dv}/DejaVuSans-Bold.ttf",
        "italic":  f"{dv}/DejaVuSans-Oblique.ttf",
        "mono":    f"{dv}/DejaVuSansMono.ttf",
    }
    if os.path.exists(linux["regular"]):
        return linux

    return {"regular": "", "bold": "", "italic": "", "mono": ""}


_FONTS = _find_unicode_fonts()


# ---------------------------------------------------------------------------
# MarkdownPDF renderer
# ---------------------------------------------------------------------------

class _MarkdownPDF:
    """Internal PDF renderer.  Instantiate via ``build_pdf()``."""

    CODE_BG = (240, 240, 240)
    HEADING_SIZES = {1: 20, 2: 16, 3: 13, 4: 11}

    def __init__(self):
        from fpdf import FPDF

        self._pdf = FPDF()
        pdf = self._pdf

        pdf.set_auto_page_break(auto=True, margin=15)

        # Font family names used throughout rendering
        self._body = "Body"
        self._mono = "Mono"

        if _FONTS["regular"]:
            reg  = _FONTS["regular"]
            bold = _FONTS["bold"]  or reg
            ital = _FONTS["italic"] or reg
            mono = _FONTS["mono"]  or reg
            pdf.add_font(self._body, style="",  fname=reg)
            pdf.add_font(self._body, style="B", fname=bold)
            pdf.add_font(self._body, style="I", fname=ital)
            pdf.add_font(self._mono, style="",  fname=mono)
        else:
            # Fall back to built-in Latin-only fonts
            self._body = "Helvetica"
            self._mono = "Courier"

        pdf.add_page()
        pdf.set_margins(20, 20, 20)
        self._set_body()

    # ------------------------------------------------------------------
    # Font helpers
    # ------------------------------------------------------------------

    def _set_body(self):
        self._pdf.set_font(self._body, size=10)
        self._pdf.set_text_color(30, 30, 30)

    def _set_heading(self, level: int):
        size = self.HEADING_SIZES.get(level, 10)
        self._pdf.set_font(self._body, style="B", size=size)
        self._pdf.set_text_color(0, 0, 0)

    def _set_code(self):
        self._pdf.set_font(self._mono, size=9)
        self._pdf.set_text_color(50, 50, 50)

    # ------------------------------------------------------------------
    # RTL helper
    # ------------------------------------------------------------------

    def _rtl_cell(self, text: str, line_height: int = 5) -> None:
        """Render *text* right-aligned across the full printable width."""
        pdf = self._pdf
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(
            pdf.w - pdf.l_margin - pdf.r_margin,
            line_height,
            _bidi(text),
            align="R",
        )

    # ------------------------------------------------------------------
    # Inline markup (LTR only; RTL goes to _rtl_cell)
    # ------------------------------------------------------------------

    def _write_inline(self, text: str) -> None:
        """Write *text* with bold/italic/code markup, handling RTL direction."""
        if _is_rtl(text):
            self._rtl_cell(text)
            return

        pdf = self._pdf
        pattern = re.compile(r"(\*\*.*?\*\*|\*.*?\*|`.*?`)")
        for part in pattern.split(text):
            if part.startswith("**") and part.endswith("**"):
                pdf.set_font(self._body, style="B", size=pdf.font_size)
                pdf.write(5, part[2:-2])
                pdf.set_font(self._body, size=pdf.font_size)
            elif part.startswith("*") and part.endswith("*"):
                pdf.set_font(self._body, style="I", size=pdf.font_size)
                pdf.write(5, part[1:-1])
                pdf.set_font(self._body, size=pdf.font_size)
            elif part.startswith("`") and part.endswith("`"):
                pdf.set_font(self._mono, size=pdf.font_size)
                pdf.write(5, part[1:-1])
                pdf.set_font(self._body, size=pdf.font_size)
            else:
                pdf.write(5, part)

    # ------------------------------------------------------------------
    # Block rendering
    # ------------------------------------------------------------------

    def render(self, md: str) -> None:
        """Render the full Markdown string into the PDF."""
        pdf = self._pdf
        lines = md.splitlines()
        in_code_block = False
        code_lines: list[str] = []

        for raw_line in lines:
            # ── fenced code block ─────────────────────────────────────
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

            # ── horizontal rule ───────────────────────────────────────
            if re.match(r"^(-{3,}|_{3,}|\*{3,})$", line):
                pdf.ln(3)
                pdf.set_draw_color(180, 180, 180)
                pdf.line(pdf.get_x(), pdf.get_y(),
                         pdf.get_x() + 170, pdf.get_y())
                pdf.ln(3)
                self._set_body()
                continue

            # ── headings ──────────────────────────────────────────────
            m = re.match(r"^(#{1,4})\s+(.*)", line)
            if m:
                level, text = len(m.group(1)), m.group(2)
                pdf.ln(4)
                self._set_heading(level)
                if _is_rtl(text):
                    self._rtl_cell(text, line_height=8)
                else:
                    pdf.multi_cell(0, 8, text)
                pdf.ln(1)
                self._set_body()
                continue

            # ── unordered list item ───────────────────────────────────
            m = re.match(r"^(\s*)[-*+]\s+(.*)", line)
            if m:
                indent, text = len(m.group(1)), m.group(2)
                self._set_body()
                if _is_rtl(text):
                    self._rtl_cell(_bidi(text) + " \u2022")
                else:
                    pdf.set_x(20 + indent * 3)
                    pdf.write(5, "\u2022 ")
                    self._write_inline(text)
                    pdf.ln()
                continue

            # ── ordered list item ─────────────────────────────────────
            m = re.match(r"^(\s*)\d+\.\s+(.*)", line)
            if m:
                indent   = len(m.group(1))
                num_text = line.lstrip()
                self._set_body()
                if _is_rtl(num_text):
                    self._rtl_cell(num_text)
                else:
                    pdf.set_x(20 + indent * 3)
                    self._write_inline(num_text)
                    pdf.ln()
                continue

            # ── blockquote ────────────────────────────────────────────
            if line.startswith("> "):
                content = line[2:]
                pdf.set_font(self._body, style="I", size=9)
                pdf.set_text_color(100, 100, 100)
                if _is_rtl(content):
                    self._rtl_cell(content)
                else:
                    pdf.set_x(25)
                    pdf.multi_cell(0, 5, content)
                self._set_body()
                continue

            # ── blank line → paragraph spacing ───────────────────────
            if not line.strip():
                pdf.ln(4)
                continue

            # ── normal paragraph text ─────────────────────────────────
            self._set_body()
            pdf.set_x(20)
            self._write_inline(line)
            if not _is_rtl(line):   # _write_inline already adds newline for RTL
                pdf.ln()

    def _render_code_block(self, lines: list[str]) -> None:
        """Render a fenced code block with a shaded background."""
        pdf = self._pdf
        text = "\n".join(lines)
        pdf.ln(2)
        self._set_code()
        pdf.set_fill_color(*self.CODE_BG)
        pdf.multi_cell(0, 5, text, fill=True)
        pdf.ln(2)
        self._set_body()

    def save(self, path: Path) -> None:
        """Write the PDF to *path*."""
        self._pdf.output(str(path))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_pdf(md_text: str, pdf_path: Path) -> None:
    """Render *md_text* (Markdown string) and save the PDF to *pdf_path*.

    Supports Unicode content including Hebrew and Arabic (RTL).

    Args:
        md_text:  Markdown-formatted text to render.
        pdf_path: Destination path for the generated PDF file.

    Raises:
        Exception: propagated from fpdf2 if rendering fails.
    """
    renderer = _MarkdownPDF()
    renderer.render(md_text)
    renderer.save(pdf_path)
