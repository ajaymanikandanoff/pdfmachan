# PDF Machan / Image Machan

A full-stack file toolbox: merge, split, compress, edit, watermark, rotate and
password-protect PDFs, plus Word / Excel / PowerPoint / image conversions.

The frontend is the original single-page UI. Every file operation is now handled
by the Node.js backend, so the tools work consistently and the previously broken
features (PDF → Word, PDF → Excel, PPT → PDF, PDF → PPT, JPG → PNG, PNG → JPG)
actually convert the files instead of renaming them.

## Run it

```bash
npm install
npm start
```

Open http://localhost:3000.

For development with automatic restarts:

```bash
npm run dev
```

## Test everything

```bash
npm test
```

The test suite generates its own PDF / DOCX / XLSX / PPTX / JPG / PNG fixtures and
calls all 19 endpoints end to end, validating every output file.

You can also create fixtures for manual browser testing with:

```bash
node tests/make-fixtures.js
```

## How it works

```
index.html            UI (unchanged tools, now calls the backend)
server.js             Express server, uploads, static files, error handling
lib/processing.js     All 19 processing handlers
lib/office.js         Optional Word / Excel / PowerPoint native conversion bridge
tests/api.test.js     End-to-end API tests
```

Uploads stay in memory (`multer.memoryStorage()`), are processed, streamed back,
and then discarded. The optional native Office bridge writes short-lived files to
the OS temp folder so Word / Excel / PowerPoint can open them, and always deletes
that folder before returning. Nothing is retained between requests.

| Tool | Engine |
| --- | --- |
| Merge / Split / Rotate / Protect / Page numbers / Editor / Watermark | `@cantoo/pdf-lib` |
| Compress PDF | `@cantoo/pdf-lib` image-stream optimisation + `pdfjs-dist` fallback |
| PDF → images | `pdfjs-dist` + `@napi-rs/canvas` |
| Word → PDF | Microsoft Word / LibreOffice, falling back to `mammoth` + `pdfkit` |
| PDF → Word | `pdfjs-dist` + `docx`, or Microsoft Word when `OFFICE_PDF2WORD=1` |
| Excel → PDF | Microsoft Excel / LibreOffice, falling back to `xlsx` + `pdfkit` |
| PDF → Excel | `pdfjs-dist` + `xlsx` |
| PPT → PDF | Microsoft PowerPoint / LibreOffice, falling back to `jszip` + `pdfkit` |
| PDF → PPT | `pdfjs-dist` + `pptxgenjs` |
| Images → PDF | `@cantoo/pdf-lib` + `sharp` |
| JPG ↔ PNG, image compression | `sharp` |

### Microsoft Office / LibreOffice conversions

The pure-JavaScript converters only approximate complex Office layouts. When the
machine has Microsoft Office (Windows) or LibreOffice installed, PDF Machan hands
`.docx/.doc/.xlsx/.xls/.pptx/.ppt` files to the real application and exports the
PDF with its own layout engine — the same approach used by Stirling-PDF and
Gotenberg. PDF → Word also uses Microsoft Word's PDF reflow when available.

Detection is automatic and best-effort; if the native application is missing or
fails, the JavaScript converter is used without interrupting the request. Set:

- `OFFICE_ENGINE=auto|msoffice|libreoffice|none` to force a preference
- `MSOFFICE_DIR` to point at a non-standard Microsoft Office folder
- `SOFFICE_PATH` to point at a non-standard LibreOffice executable
- `OFFICE_PDF2WORD=1` to opt in to Microsoft Word's PDF reflow for PDF → Word
- `DEBUG_OFFICE=1` to log native-conversion fallbacks

## API

`POST /api/tools/:tool` — `multipart/form-data`

- `files` — one or more uploaded files
- `meta` — JSON string with tool options (page ranges, passwords, watermark
  layers, editor changes, quality, …)

The response is the processed file. The filename is returned in the
`X-Output-Filename` header and `Content-Disposition`. Compression responses also
include `X-Original-Size`, `X-Output-Size` and `X-Compression-Strategy`.

Common `meta` options:

- `merge`: `order` — array of `{ fileIndex, pageIndex }` (both 0-based) describing
  the exact page order and which pages to keep
- `split`: `mode` (`select`, `each`, `range`), `pages` (1-based page numbers) or
  `start` / `end` for a range
- `pagenum`: `position`, `format`, `startNumber`, `fontSize`, `colorHex`, `bold`
- `compress`: `quality` from `0.2` (smallest) to `0.95` (highest quality)

`GET /api/health` lists the available tools and upload limits.

### Limits

- 100 MB per file
- 20 files per request
- 10 minute processing timeout

## Deploy

Set `PORT` if your host requires it; the server binds to `0.0.0.0`. Any Node host
that supports native modules (`sharp`, `@napi-rs/canvas`) works, including
Render, Railway, Fly.io and a plain VPS.

## Notes

- The PDF text editor uses `pdf-lib`'s standard fonts, so characters outside
  WinAnsi (for example Tamil) are replaced with `?` when writing edited text back
  into a PDF. Latin text is unaffected.
- Compressing a PDF first re-encodes and downscales its embedded images, so text
  and vector art stay sharp. A raster fallback is used only when it produces a
  clearly smaller result or when the images cannot be rewritten. The server
  always returns the smallest safe candidate, never a file larger than the input.
- Scanned/image-only PDFs have no extractable text, so PDF → Word / Excel will
  note the empty page instead of inventing text.
