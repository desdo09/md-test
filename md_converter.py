"""
md_converter.py
---------------
Converts a document file to Markdown using Microsoft's MarkItDown library.

Supported input formats (handled automatically):
  PDF, Word (.docx), Excel (.xlsx/.xls), PowerPoint (.pptx),
  HTML, CSV, JSON, XML, Images, Audio, EPubs, ZIP, and more.

Public API
----------
  convert_to_markdown(src, output_dir) -> Path
      Converts src to a .md file in output_dir and returns the path.

  extract_markdown(src) -> str
      Returns the raw Markdown text without writing any file.
"""

from pathlib import Path

from markitdown import MarkItDown

# Single shared MarkItDown instance (stateless, safe to reuse).
_md = MarkItDown()


def extract_markdown(src: Path) -> str:
    """Extract Markdown text from *src* and return it as a string.

    Raises:
        Exception: propagated from MarkItDown if the file cannot be converted.
    """
    result = _md.convert(str(src))
    return result.text_content


def convert_to_markdown(src: Path, output_dir: Path) -> Path:
    """Convert *src* to a Markdown file inside *output_dir*.

    The output file is named ``<src.stem>.md``.

    Returns:
        Path to the generated ``.md`` file.

    Raises:
        Exception: propagated from MarkItDown if the file cannot be converted.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    md_text = extract_markdown(src)
    md_path = output_dir / f"{src.stem}.md"
    md_path.write_text(md_text, encoding="utf-8")
    return md_path
