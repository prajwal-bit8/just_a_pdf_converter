# Just a File Converter

A private, browser-local document toolkit inspired by the practical workflow of PDF utility sites, with its own visual identity.

## Features

- Word to PDF: DOCX text, headings, tables, supported images; legacy Word 97–2003 DOC text extraction.
- PDF to Word: editable selectable text, or full-page images in a DOCX.
- Merge, split, reorder, delete and rotate PDF pages.
- Image-based compression with original-file fallback if the output grows.
- JPG/PNG to PDF; PDF pages to JPG (ZIP for multiple pages).
- Watermarks, page numbers and text extraction.
- Search, category filters, favorite tools, session file manager and local activity history.
- Real sample files in every workflow.

No documents are uploaded. Libraries, fonts, PDF worker resources and samples are served from the same site. Session files are held in memory and cleared on refresh. Local storage contains only favorite tool IDs and recent tool names/times/file counts.

## Limits

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

The tracked `.openai/hosting.json` retains the site's private hosting identity. Source synchronization and deployments use the Sites workflow.

## Verification

Node tests cover range validation, file limits, merging, extraction, organization/rotation, watermarks, numbering, image-to-PDF, reading order, and ANSI/Unicode legacy DOC text with field codes. Browser QA exercises conversions, page previews, controls and result screens with real sample documents.
