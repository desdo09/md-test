# md-test — MarkItDown Conversion Test

Convert documents (PDF, Word, Excel, PowerPoint, HTML, CSV, JSON, …) to
**Markdown** and optionally to **PDF** using Microsoft's
[MarkItDown](https://github.com/microsoft/markitdown) library.

## Quick Start

```powershell
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate sample input files (optional)
python create_sample_docs.py

# 3. Place your own documents in the input/ folder (or use the samples)

# 4. Convert everything in input/ → output/
python convert.py

# 5. Convert a single file
python convert.py input\my_document.docx

# 6. Convert without generating a PDF
python convert.py --no-pdf
```

## Folder layout

```
md-test/
├── input/           ← put your source documents here
├── output/          ← generated .md and .pdf files land here
├── convert.py       ← main conversion script
├── create_sample_docs.py  ← generates test input files
└── requirements.txt
```

## Supported input formats

| Format | MarkItDown flag |
|--------|----------------|
| PDF | `[pdf]` |
| Word (.docx) | `[docx]` |
| Excel (.xlsx / .xls) | `[xlsx]` / `[xls]` |
| PowerPoint (.pptx) | `[pptx]` |
| HTML | built-in |
| CSV / JSON / XML | built-in |
| Images (EXIF + OCR) | `[all]` |
| Audio (transcription) | `[audio-transcription]` |
| EPub | built-in |

## Output

For each file `input/foo.docx` the script creates:
- `output/foo.md` — Markdown text
- `output/foo.pdf` — PDF rendered from the Markdown (pass `--no-pdf` to skip)

## Hebrew / RTL languages

Yes — Hebrew, Arabic, and other RTL languages are fully supported:

- **Markdown extraction**: MarkItDown preserves all Unicode characters, so Hebrew text in Word, HTML, PDF, etc. comes out correctly in the `.md` file.
- **PDF rendering**: The `python-bidi` library applies the Unicode Bidirectional Algorithm so Hebrew lines are rendered right-to-left. The Windows Arial font (used automatically) includes the full Hebrew Unicode block.

Mixed Hebrew + English documents work correctly: LTR lines render left-to-right and RTL lines render right-to-left.
