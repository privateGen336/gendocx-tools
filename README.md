# GenDocX Tools

[![build](https://github.com/privateGen336/gendocx-tools/actions/workflows/build.yml/badge.svg)](https://github.com/privateGen336/gendocx-tools/actions/workflows/build.yml)

**Six document tools that run entirely in your browser. Nothing is uploaded, nothing is stored, no account required.**

Convert images to Word, turn Markdown into a styled `.docx`, merge and sign PDFs, stamp watermarks on DOCX/PDF, compare two versions of a contract, and edit a `.docx` — all with the network panel open and silent.

[中文说明](./README.zh-CN.md) · [Hosted version](https://gendocx.com/tools)

![Six document tools illustrated inside a browser window](./docs/images/browser-document-tools.jpg)

---

## Hosted versions

If you would rather not install anything, the same tools run on the GenDocX site. Same code, same "files stay local" behaviour for the standalone tools.

| Tool | Source | Hosted |
|---|---|---|
| Images → Word | `src/components/ImageToWord.tsx` | [Image to Word converter](https://gendocx.com/tools/images-to-docx) |
| Markdown → Word | `src/components/MarkdownToWord.tsx` | [Markdown to Word](https://gendocx.com/tools/markdown-to-docx) |
| PDF toolbox | `src/components/PdfToolbox.tsx` | [Merge, split and sign PDFs](https://gendocx.com/tools/pdf) |
| Watermark | `src/components/WatermarkTool.tsx` | [Add a watermark to PDF or Word](https://gendocx.com/tools/watermark) |
| Compare | `src/components/WordCompare.tsx` | [Compare two Word documents](https://gendocx.com/tools/compare) |
| Editor | `src/components/WordEditor.tsx` | [Edit a Word document online](https://gendocx.com/tools/editor) |

## Why this exists

Most "free online converter" sites upload your file to a server you cannot see. For a signed contract, an ID scan or an internal memo, that is a real risk — and you have no way to check what happens after the upload completes.

These tools take the opposite approach: decode, transform and repack in the tab.

- Images are decoded with `createImageBitmap` and packed into OOXML by the `docx` library.
- PDFs are parsed by `pdf.js` in a Web Worker and edited with `pdf-lib`.
- DOCX comparison runs on `docxodus` (a .NET engine compiled to WebAssembly) in a worker.
- The only network requests are the static files that load this page.

The code is MIT-licensed. Read it, audit it, fork it, or ship it inside your own product.

## What's inside

### Images → Word

- JPG, PNG, WebP, GIF, AVIF, BMP — up to 50 images at once
- Drag to reorder, rotate, delete; page breaks preview live as you change settings
- Paper size (A4 / Letter), orientation (auto or locked), images per page (1 / 2 / 4)
- Margins (none / 1 cm / 2 cm / split top-bottom), quality (fit or original pixels), JPEG or PNG encoding
- Output is a real `.docx` — Word, Pages, WPS, Google Docs and LibreOffice all open it

### PDF toolbox

- **Merge** several files in the order you choose
- **Split** at a visual cut line, or every N pages
- **Extract** and **delete** individual pages from a page wall
- **Rotate** and **reorder** by dragging
- **Flatten** a document when the source allows it
- **Sign** — draw a signature and place it on any page
- Undo / redo, thumbnail preview before committing, ZIP download for multi-file results

### PDF & DOCX watermark

- Works on both `.pdf` and `.docx`
- Text or image watermark, solid or outline, above or below the content
- Nine anchor positions, adjustable size and opacity
- Live variables — `{date}`, `{time}`, `{datetime}`, `{year}`, `{filename}`, `{pages}`
- Scans an existing file and can strip watermarks you added yourself

Removing a watermark only makes sense on documents you own. Don't use it to pass off someone else's work.

### Compare Word documents

- Drop an original and a modified `.docx`
- Every revision is listed with its type — insert, delete, format, move
- Jump between changes in the document preview
- Export a redline `.docx` with tracked changes intact

### Markdown to Word

- Paste Markdown or open a `.md` file, pick one of four layouts, download a `.docx`
- Headings, tables, code blocks, lists, links and images land as native Word elements, not a screenshot
- LaTeX `$...$` and `$$...$$` become editable Word equations
- Standard, Compact, Report and Official presets — the last follows GB/T 9704 margins, 仿宋 body at 16pt and a two-character first-line indent
- The preview renders the generated `.docx` back into the page, so it is the file you download

### Word editor

- Open a `.docx` and edit it in place: text, tables, lists, images
- Print or export the result
- Runs on the `@docx-editor.dev` engine (Apache-2.0), loaded lazily

## Run it locally

```bash
npm install
npm run dev        # http://localhost:5173
```

```bash
npm run build      # static output in dist/
npm run preview    # serve the production build
```

Node 20.19+ (or 22.12+).

## Deploy it anywhere

`npm run build` emits a plain static site — no Node runtime, no server functions. `base` is `./`, so it also works from a sub-path: GitHub Pages, Cloudflare Pages, Netlify, Vercel (static), S3, nginx, Caddy.

## How the no-upload claim holds

There is no upload endpoint to disable — the code never calls one:

- `src/lib/generate-docx.ts` takes `File[]` and returns a `Blob`
- `src/lib/pdf-ops.ts` takes `Uint8Array` and returns `Uint8Array`
- `src/lib/watermark/*` rewrites the file bytes in memory
- `src/lib/docx-compare.worker.ts` runs the comparison engine off the main thread

The only external URL in the UI is the link back to the hosted product.

## Relationship to GenDocX

These six tools are the browser-side core of [GenDocX](https://gendocx.com) — a document workspace where you edit a Word file and talk to an AI agent in the same window. The agent can read your uploaded files, produce a `.docx`, and hand it back for review, without a separate converter step.

The difference is scope, not trust:

- **This repository** is the standalone toolbox. One file in, one file out, no account, nothing saved.
- **[gendocx.com](https://gendocx.com)** adds the workspace around them — sessions, file history, an AI agent, and templates. If you want to see what that looks like, [start with the tools page](https://gendocx.com/tools) or read [the docs](https://gendocx.com/docs).

If you only need to convert a batch of scans once, clone this repo and never think about it again. If you need to keep working on the document afterwards, that is what the hosted product is for.

## Tech stack

React 19 · TypeScript · Vite · `docx` · `pdf-lib` · `pdfjs-dist` · `docxodus` · `@docx-editor.dev` · `fflate` · `lucide-react`

## FAQ

**Do my files get uploaded?**
No. There is no upload code path. Files are read locally and written back with `URL.createObjectURL`.

**Can I use it offline?**
Yes. Build once and serve `dist/` from localhost or any static host; after the page loads, nothing else is fetched.

**Does it work with scanned PDFs?**
For merging, splitting, rotating, deleting, reordering and watermarking — yes. There is no OCR, so scanned text is not searchable.

**Can it open password-protected PDFs?**
No. Encrypted files are rejected rather than partially processed.

**Is the output a real .docx?**
Yes — OOXML generated by the `docx` library, not HTML renamed to `.docx`.

**Why is the bundle large?**
`docxodus` and the editor engine are WebAssembly and only load when you open Compare or Editor. The other three tools load a much smaller bundle.

**Can I use this commercially?**
Yes, MIT. Attribution is appreciated but not required.

## License

MIT © GenDocX
