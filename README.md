# Google Slides Dump Tool

A Windows toolkit that captures a presentation you own or are allowed to copy
and exports it to PDF, DOCX, or a searchable PDF with automatic bookmarks.
The searchable-PDF workflow is fully local and uses Tesseract OCR.

## Intended use and legal notice

Use this project only for lawful personal archiving or for content you own or
have explicit permission to reproduce. You are responsible for complying with
copyright, licences, applicable law, and the source platform's terms. The
project does not bypass access controls and is not affiliated with Google LLC.
"Google Slides" is a trademark of Google LLC. This notice is not legal advice.

## Requirements

- Windows and PowerShell
- Python 3.9 or newer
- The Python packages in `requirements.txt`
- Tesseract for searchable PDF exports

Install the Python packages:

```powershell
python -m pip install -r requirements.txt
```

Install Tesseract on Windows:

```powershell
winget install UB-Mannheim.TesseractOCR
```

For French slides, place `fra.traineddata` in
`C:\Program Files\Tesseract-OCR\tessdata\`. The script automatically uses
`fra+eng` when both languages are installed, then falls back to `fra` or `eng`.
Set `OCR_LANG` to override this choice.

If PowerShell blocks the launcher, allow scripts only for the current window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Quick start

Run:

```powershell
.\screenshot-recorder.ps1
```

The menu can:

1. Capture a presentation now.
2. Reuse `captures/Page_N.*` images that already exist.
3. Process an existing PDF without taking new screenshots.

For new captures, open the presentation in Present mode first. The recorder
waits five seconds, takes the requested number of screenshots, and advances
with the right-arrow key.

Available exports are PDF, DOCX, both, or searchable PDF with an automatic
table of contents. The generated document opens when processing finishes.

## Robust local title detection

The searchable-PDF workflow performs full-page Tesseract OCR for the hidden
search layer and separate targeted title scans. It analyses all pages before
building the bookmarks so repeated logos, headers, and footers can be detected
across the complete document.

Title priority is:

1. Manual correction from `titles.json`.
2. Dominant centered title on a sparse cover, divider, or section slide.
3. Upper-left title on a normal content slide.
4. Centered or full-width upper header when no upper-left title is reliable.
5. First sentence of the first readable paragraph.
6. `Slide N` when no meaningful text is available.

The detector supports wrapped titles, centered multiline titles, cover slides,
two-column layouts, top-center headers, side panels, sparse slides, pages
without titles, PDF-native text, and OCR confidence differences. Visible title
decoration is retained, so `[AGENDA]` remains exactly `[AGENDA]`.

Use `--debug-titles` to print ranked candidates, strategies, confidence values,
and repeated-template penalties.

## Existing PDFs

The source PDF is never overwritten. Pages and page dimensions are preserved.
For image-based pages, the embedded full-page image is extracted at its original
pixel dimensions. A vector page without such an image is rendered only for
local title analysis.

If a PDF already has a usable text layer, that layer is reused and full-page
re-OCR is skipped. Tesseract still performs a small targeted scan of the title
areas; this can recover a heading that the older full-page OCR missed without
rebuilding the PDF text layer. Missing text layers receive a hidden OCR layer.

Direct examples:

```powershell
python img-2-searchable-pdf.py --captures-dir captures
python img-2-searchable-pdf.py --input-pdf slides.pdf --output slides-searchable.pdf
python img-2-searchable-pdf.py --input-pdf slides.pdf --debug-titles --output checked.pdf
```

## Smart dark mode

Searchable PDF exports offer an optional smart dark mode. It is disabled by
default. A page is converted only when sampling confirms all of the following:

- the background is predominantly light and neutral;
- enough dark neutral ink is present;
- no complex photographic region dominates the slide.

Neutral tones are inverted, turning white backgrounds black and black text
white, while saturated colors are kept. Existing dark slides, photographs, and
colored designs are skipped. Capture images and embedded PDF images keep their
original pixel dimensions, and PDF page dimensions never change. A vector-only
page has no original pixel resolution; if dark mode applies to it, the page is
rasterized at `--render-dpi` (200 by default) while retaining its PDF dimensions.

```powershell
python img-2-searchable-pdf.py --captures-dir captures --dark-mode
python img-2-searchable-pdf.py --input-pdf slides.pdf --dark-mode --output slides-dark.pdf
```

## Manual bookmark corrections

Create an optional UTF-8 `titles.json` next to the scripts. Page numbers start
at 1:

```json
{
  "2": "Quarterly overview",
  "12": "Appendix"
}
```

Manual corrections have the highest priority. Invalid JSON, a non-string value,
or an empty title stops generation before an existing output can be replaced.
Out-of-range pages produce a warning and are ignored.

```powershell
python img-2-searchable-pdf.py --titles-file titles.json
```

## Reliability guarantees

- Processing is local and requires no network connection or account.
- Output is written to a temporary PDF and atomically replaces the destination
  only after every page, bookmark, and metadata step succeeds.
- An earlier output remains untouched after a failure.
- The input PDF cannot be selected as its own output.
- Full-page OCR, targeted title OCR, the hidden search layer, page count,
  bookmarks, and dark-mode decisions are covered by deterministic tests.

## Command-line options

```text
--captures-dir PATH   Reuse Page_N images from a chosen folder
--input-pdf FILE      Process an existing PDF
--output FILE         Choose the destination PDF
--titles-file FILE    Load manual bookmark titles
--dark-mode           Enable selective black-on-white conversion
--render-dpi DPI      Vector-PDF analysis/render DPI, from 96 to 600
--debug-titles        Print detailed local title ranking
```

## Files

| File | Purpose |
| --- | --- |
| `screenshot-recorder.ps1` | Captures slides, selects a source, and drives exports. |
| `img-2-pdf.py` | Compiles captured images into a PDF. |
| `img-2-docx.py` | Compiles captured images into a DOCX. |
| `img-2-searchable-pdf.py` | Builds the searchable PDF, bookmarks, and optional dark render. |
| `slide_title_detection.py` | Pure local layout and title-ranking engine. |
| `tests/` | Deterministic title, PDF pipeline, dark-mode, and failure tests. |

## License

Released under the [MIT License](./LICENSE).
