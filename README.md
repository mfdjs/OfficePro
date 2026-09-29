# OfficePro

**An all-in-one, cross-platform office suite built with Electron** — word processing, spreadsheets, presentations, PDF tools, document scanning and a lightweight database, all in a single desktop app.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Electron](https://img.shields.io/badge/Electron-22-47848F.svg)](https://www.electronjs.org/)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D4.svg)](#download)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

English | [简体中文](README.zh-CN.md)

> **This project is looking for contributors.** OfficePro is a working, usable app, but it is a solo-developer project with a large single-file UI and limited tests. If you are good at Electron, front-end architecture, file-format parsing, or testing, your help would be very welcome. See [Help wanted](#-help-wanted) and open a [discussion](https://github.com) / pull request.

---

## Features

- **Word** — rich-text editing, headings/styles, tables, images, tracked changes, comments, footnotes/endnotes, citations & table of contents, mail merge, and local smart-writing assistants.
- **Spreadsheet** — formulas and a function library, charts, data validation, conditional formatting, sort/filter, freeze panes, grouping, named ranges, subtotals, and sheet protection.
- **Presentation** — slides, shapes, text boxes, WordArt, animations and transitions, themes & variants, masters, presenter view, and a full slideshow player.
- **PDF Tools** — viewing, merge/split, compress, encrypt/decrypt, page extraction, redaction, conversion to/from other formats, and OCR.
- **Scan** — camera/import scanning, batch processing, OCR, filters, rotate/flip, brightness/contrast/sharpen, and old-photo restoration.
- **Database** — lightweight tables and queries.
- **International UI** — switch the interface between 10 languages at runtime.

## Supported file formats

OfficePro tries to open virtually any document you throw at it, including:

- **Documents:** doc, docx, docm, dot, dotx, odt, rtf, wps, pages, txt, md
- **Spreadsheets:** xls, xlsx, xlsm, xlsb, csv, tsv, ods, numbers, dif, prn
- **Presentations:** ppt, pptx, pps, ppsx, pot, potx, odp, key
- **PDF & e-books:** pdf, xps, oxps, epub, mobi, azw, azw3, fb2
- **Images:** jpg, png, bmp, gif, webp, tiff, svg, ico, avif, heic
- **Text / source code:** json, xml, yaml, html, css, js, ts, py, java, c, cpp, cs, go, rs, php, sh, sql, and many more
- **Archives & media:** zip, tar, gz, 7z, rar, eml, msg, fonts, mp3, wav, mp4, mkv, …

Encrypted or corrupted files are detected and reported with a clear message instead of showing garbled output.

## Languages

The interface supports: Chinese (default), English, Japanese, Korean, Spanish, French, German, Russian, Portuguese, and Arabic (with full right-to-left layout). Translations live in [`locales/`](locales) and are applied through [`i18n.js`](i18n.js).

## Getting started (development)

Requirements: **Node.js 16+** and **npm**.

```bash
# 1. Clone
git clone https://github.com/<your-username>/OfficePro.git
cd OfficePro

# 2. Install dependencies
npm install

# 3. Run in development
npm start
```

To enable the DevTools protocol on a fixed port:

```bash
npx electron . --no-sandbox --remote-debugging-port=9222
```

### OCR language data

The Chinese-simplified OCR trained data is not committed (around 19 MB). Tesseract.js can download it automatically on first use; to place it manually:

```bash
# create the folder and download chi_sim
mkdir -p tessdata
curl -L -o tessdata/chi_sim.traineddata.gz \
  https://tessdata.projectnaptha.com/4.0.0/chi_sim.traineddata.gz
```

## Build

```bash
npm run build
```

This produces an NSIS installer, a portable `.exe`, and an unpacked directory under `dist/`.

## Tech stack

- [Electron](https://www.electronjs.org/) — desktop shell
- Vanilla HTML/CSS/JavaScript — no front-end framework
- [mammoth](https://github.com/mwilliamson/mammoth.js) / [word-extractor](https://github.com/ArtifexSoftware/word-extractor) — Word parsing
- [SheetJS (xlsx)](https://sheetjs.com/) — spreadsheet parsing
- [JSZip](https://stuk.github.io/jszip/) — Office Open XML / zip handling
- [pdf-parse](https://gitlab.com/autokent/pdf-parse) — PDF text
- [tesseract.js](https://github.com/naptha/tesseract.js) — OCR
- [docx](https://github.com/dolanmiu/docx) — document generation

## Project structure

```
OfficePro/
├── index.html           # Main UI (all modules)
├── main.js              # Electron main process
├── preload.js           # Secure bridge
├── file-parser.js       # Multi-format file parsing
├── officepro-enhance.js # Layout / shortcuts / OCR hooks
├── officepro-real.js    # Concrete feature implementations
├── i18n.js              # Internationalization runtime
├── locales/             # UI translations (9 languages)
├── tessdata/            # OCR data (not committed)
└── package.json
```

## Roadmap

- [ ] Split the large `index.html` into maintainable modules/components
- [ ] Add a proper unit-test suite and CI
- [ ] Improve Office Open XML parsing/rendering fidelity
- [ ] macOS and Linux builds
- [ ] Large-file performance and streaming
- [ ] Accessibility (a11y) pass
- [ ] Proofread and extend community translations

## Help wanted

Specifically, the project would benefit from help with:

1. **Architecture** — refactoring the single-file UI into components and separating concerns.
2. **Tests & CI** — introducing unit tests and automated build/release pipelines.
3. **File-format fidelity** — handling more edge cases in Word/Excel/PPT/PDF.
4. **Cross-platform** — verifying and packaging for macOS/Linux.
5. **Translations** — reviewing and improving the bundled language packs.
6. **Code modernization** — TypeScript, ESM, and a component-based view layer.

Please read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a pull request.

## Contributing

Contributions, bug reports, and feature ideas are welcome. Fork the repo, create a branch, and open a PR. See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## License

Released under the [MIT License](LICENSE).
