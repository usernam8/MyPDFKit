# Toast Docs

A free PDF app for Windows: open, edit, sign, fill in, print and convert PDFs, Word, Excel, PowerPoint and pictures, with an AI helper that runs on your own computer.

Download: [Microsoft Store](https://apps.microsoft.com/detail/9MWH8JPBPRLL) or the [releases page](https://github.com/usernam8/MyPDFKit-releases/releases/latest).

## License

Copyright (C) 2026 usernam8.

Toast Docs is free software under the GNU Affero General Public License version 3 (see [LICENSE](LICENSE)).
You can use, study and change it. If you share a changed version, you have to share its full source code under the same license.

The name "Toast Docs" and its icon are not covered by the license. Please use a different name and icon for your own version.

## Built with

| Part | What it does | License |
|---|---|---|
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | Reads, draws and edits PDFs | AGPL-3.0 |
| [pdf.js](https://github.com/mozilla/pdf.js) | Shows pages in the window | Apache-2.0 |
| [Electron](https://www.electronjs.org/) | The app window | MIT |
| [pypdfium2](https://github.com/pypdfium2-team/pypdfium2) / PDFium | Printing and page pictures | Apache-2.0 / BSD-3-Clause |
| [llama.cpp](https://github.com/ggml-org/llama.cpp) | Runs the built-in AI | MIT |
| [Qwen3.5-2B](https://huggingface.co/Qwen) | The built-in AI model | Apache-2.0 |
| [pyHanko](https://github.com/MatthiasValvekens/pyHanko) | Digital signatures | MIT |
| [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) | PDF to Word | MIT |
| [python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx), [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl), [xlrd](https://github.com/python-excel/xlrd) | Office files | MIT / BSD |
| [Pillow](https://github.com/python-pillow/Pillow), [pillow-heif](https://github.com/bigcat88/pillow_heif) (libheif, libde265, x265) | Pictures, iPhone photos | MIT-CMU / BSD-3-Clause (LGPL-3.0, LGPL-3.0, GPL-2.0) |
| [ezdxf](https://github.com/mozman/ezdxf) | CAD drawings | MIT |
| [Flask](https://github.com/pallets/flask) | The engine's local connection | BSD-3-Clause |

Each part keeps its own license; the full texts come with the installed app.
