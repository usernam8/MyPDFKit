# Toast Docs vs. an online converter (measured 2026-11-05)

Same public sample files from this repo, same PC and internet connection. Toast Docs times are from `tests/bench.py`
(best of 3). iLovePDF times run from the moment the file was handed to the site until its download page started
loading; downloading the result comes on top of that.

| Job | Toast Docs 1.9.3 | Toast Docs 1.9.2 | iLovePDF (online) |
|---|---:|---:|---:|
| PDF to Word, 4-page report with tables | **2.2 s** | 16.2 s | 4.4 s |
| Word file to PDF | **0.23 s** | 0.23 s | 6.9 s |
| Combine the same 10 PDFs | **0.09 s** | 0.09 s | 6.9 s |
| PDF to Word, 159-page book | **10.4 s** | 15.2 s | not tested (the book isn't public) |

Adobe Acrobat isn't installed on this PC, so it couldn't be timed. Microsoft Word's own PDF import hangs on a
hidden prompt when it's driven automatically, so it's left out (`python tests/bench.py --word` tries it).

## What changed in 1.9.3 to get there

PDF to Word gives exactly the same Word files as before, byte for byte (checked on 30 files), just faster:

- python-docx found each table cell by rebuilding the whole table grid every time: quadratic work, 20 of 28 seconds
  on the report. Toast Docs now finds the cell from its own row.
- PyMuPDF's rectangle class parsed its arguments the slow, general way 2.2 million times for the book.
  The common forms now take a direct path.
