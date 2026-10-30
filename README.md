# Just a File Converter

A private, browser-local document toolkit inspired by the practical workflow of PDF utility sites, with its own visual identity.

## Features

- Code to PDF and editable Word (.docx): C, C++, Java, Python, JavaScript, TypeScript, C#, Verilog, Arduino, HTML/CSS, SQL and other source/text formats.
- Upload up to 10 source files or paste code. Reorder files, preview text, add a title and line numbers, choose A4/Letter, portrait/landscape, font size and tab width. Each source file starts on a new page.
- Images to Word: JPG, PNG, WebP, GIF and BMP in a `.docx` with one image per page, optional filenames, page size and orientation. No OCR. Animated files use one still frame. Images are normalized locally, limited to 40 megapixels each and resized to at most 3,000 pixels on their longest side.
- PPTX to Word/PDF: handout-style exports of slide text and embedded JPG/PNG images in presentation order. Each slide starts a new page; long slides may span pages. Original slide layouts, backgrounds, charts, SmartArt, notes, animations and media are not reproduced. Tables become plain text; unsupported images are reported. Up to 100 slides, 200 image references, 400,000 text characters and 200 MB unpacked. Older `.ppt` and encrypted presentations are not supported.
- Word to PDF: DOCX text, headings, tables, supported images; legacy Word 97–2003 DOC text extraction.
- PDF to Word: editable selectable text, or full-page images in a DOCX.
- Merge, split, reorder, delete and rotate PDF pages.
- Image-based compression with original-file fallback if the output grows.
- JPG/PNG to PDF; PDF pages to JPG (ZIP for multiple pages).
- Watermarks, page numbers and text extraction.
- Search, category filters, favorite tools, session file manager and local activity history.
- Real sample files in every workflow.

No documents are uploaded. Libraries, fonts, PDF worker resources and samples are served from the same site. Session files are held in memory and cleared on refresh. Local storage contains only favorite tool IDs and recent tool names/times/file counts.

Created by **Preetam RC** and **Prajwal**.

## Limits

Code conversion formats source text only; it does not execute code or capture program output. Files must be UTF-8, or UTF-16 with a byte-order mark. Code tools accept up to 1 MB/file, 10 files, 5,000 source lines, 400,000 characters and 8,000 wrapped lines per task. Tabs expand to configurable spaces; long lines wrap without removing text. PDF uses the bundled DejaVu Sans Mono font and reports unsupported characters rather than silently dropping them. Word preserves Unicode text; display depends on the reader's available fonts. The font license is in `dist/fonts/LICENSE-DejaVu.txt`.

Conversions are approximate, not Microsoft Word or Acrobat rendering. DOCX complex layouts, custom fonts, headers and footers may differ; legacy DOC is text only. PDF-to-Word text mode does not reconstruct columns, tables or images. Scanned text cannot be edited: OCR is not included. Page-image Word mode preserves visual content without editable text. Compression flattens pages to JPEG, removing searchable text, forms and links. Password-protected documents are rejected. The UI explains these limits before each operation.

50 MB/file; 150 MB/task; 30 files/task. 200-page organizer/Word conversion, 100-page raster conversion, 300-page individual split export.

## Development

Requires Node.js 22+.

```sh
npm ci
npm test
npm run build
```

The deployment serves `dist/`. `build.mjs` uses esbuild to create lazy-loaded browser bundles and copies bundled PDF.js resources and Inter fonts. Source files are in `src/`; styles and the HTML entry are in `dist/`. Sample documents can be recreated with `node scripts/make-samples.mjs` (the static sample JPEG is retained).

The tracked `.openai/hosting.json` retains the existing Sites project identity. Source synchronization and deployments use the Sites workflow.

For Netlify, deploy branch `main`, leave the base directory blank, use `npm run build`, publish `dist`, and set `NODE_VERSION=22`. The source stylesheet, HTML, sample files and DejaVu font remain in `dist`; the build recreates dependency bundles and Inter fonts.

## Verification

Node tests cover source encodings, tab stops, code wrapping, PDF pagination and font coverage, editable Word code, range validation, file limits, merging, extraction, organization/rotation, watermarks, numbering, image-to-PDF, reading order, and ANSI/Unicode legacy DOC text with field codes. Browser QA exercises conversions, page previews, controls and result screens with real sample documents.
