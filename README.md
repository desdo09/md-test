# md-test — MarkItDown Conversion Test

Convert documents (PDF, Word, Excel, PowerPoint, HTML, CSV, JSON, …) to
**Markdown** and optionally to **PDF** using Microsoft's
[MarkItDown](https://github.com/microsoft/markitdown) library.

Supports Unicode content including **Hebrew, Arabic**, and other RTL languages.

## Quick Start

```powershell
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate sample input files (optional)
python create_sample_docs.py

# 3. Place your own documents in the input/ folder (or use the samples)

# ── Full pipeline: document → Markdown + PDF ──────────────────────────────
python main.py                        # convert all files in ./input/
python main.py input\my_document.docx # convert a single file
python main.py -o exports             # custom output folder
python main.py --no-pdf               # Markdown only, skip PDF

# ── Direct to PDF: document → PDF only (no .md saved) ────────────────────
python doc_to_pdf.py                        # convert all files in ./input/
python doc_to_pdf.py input\my_document.docx # convert a single file
python doc_to_pdf.py -o exports             # custom output folder
```

## Project structure

```
md-test/
├── main.py               ← full pipeline: document → Markdown + PDF
├── doc_to_pdf.py         ← direct conversion: document → PDF only
├── md_converter.py       ← document → Markdown  (MarkItDown)
├── pdf_converter.py      ← Markdown  → PDF       (fpdf2 + python-bidi)
├── create_sample_docs.py ← generates sample input files for testing
├── test_hebrew.py        ← end-to-end Hebrew / RTL test
├── requirements.txt
├── input/                ← put your source documents here
└── output/               ← generated files land here
```

## Module overview

| File | Responsibility |
|------|---------------|
| `main.py` | CLI — document → `.md` **and** `.pdf` |
| `doc_to_pdf.py` | CLI — document → `.pdf` only (direct, no `.md` saved) |
| `md_converter.py` | `extract_markdown(src)` / `convert_to_markdown(src, out_dir)` |
| `pdf_converter.py` | `build_pdf(md_text, pdf_path)` — Unicode + RTL rendering |

## Supported input formats

| Format | Notes |
|--------|-------|
| PDF | via `pdfminer-six` |
| Word (`.docx`) | via `mammoth` |
| Excel (`.xlsx` / `.xls`) | via `openpyxl` / `xlrd` |
| PowerPoint (`.pptx`) | via `python-pptx` |
| HTML | built-in |
| CSV / JSON / XML | built-in |
| Images | EXIF metadata + OCR |
| Audio | speech transcription |
| EPub | built-in |
| ZIP | iterates over contents |

## Output

For each file `input/foo.docx` the pipeline creates:
- `output/foo.md` — extracted Markdown text
- `output/foo.pdf` — PDF rendered from the Markdown (omit with `--no-pdf`)

## Hebrew and RTL languages

Hebrew, Arabic, and other RTL languages are fully supported:

- **Markdown extraction** — MarkItDown preserves all Unicode characters, so
  Hebrew text in Word, HTML, PDF, etc. is extracted correctly into the `.md` file.
- **PDF rendering** — `python-bidi` applies the Unicode Bidirectional Algorithm
  so Hebrew and Arabic lines are rendered right-to-left. Headings, lists,
  blockquotes, and paragraphs all respect text direction automatically.
- **Font** — Arial (Windows) / DejaVu (Linux) is loaded as a Unicode TTF,
  covering the full Hebrew and Arabic Unicode blocks.

Mixed LTR/RTL documents work correctly: each line is rendered in its own
natural direction.
