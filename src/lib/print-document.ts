"use client";

/**
 * Print / "Save as PDF" for the standalone editor, entirely in the browser.
 *
 * The editor only exposes `save(): ArrayBuffer` (DOCX). To get a faithful print
 * we reuse the same DOCX→HTML + PaginationEngine path the compare page already
 * ships (both run the local Docxodus WASM, no upload), then hand the paginated
 * iframe to the browser's print dialog. Choosing "Save as PDF" there produces a
 * vector, selectable-text PDF — not a screenshot.
 */

// Off-screen but still laid out: `display:none` / `visibility:hidden` iframes
// print blank in some browsers. Do not add opacity either.
const FRAME_STYLE = "position:fixed;top:0;left:-10000px;width:1200px;height:1600px;border:0";

interface PageSize { widthPt: number; heightPt: number }

function printCss(size: PageSize): string {
  const width = size.widthPt > 0 ? size.widthPt : 612;
  const height = size.heightPt > 0 ? size.heightPt : 792;
  return `
@page { size: ${width}pt ${height}pt; margin: 0; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
@media print {
  html, body {
    background: #fff !important; margin: 0 !important; padding: 0 !important;
    overflow: visible !important;
  }
  .compare-pages { width: auto !important; min-width: 0 !important; margin: 0 !important; }
  .compare-paper-box {
    box-shadow: none !important; margin: 0 !important;
    break-after: page; page-break-after: always;
  }
  .compare-paper-box:last-child { break-after: auto; page-break-after: auto; }
  [data-compare-focus] { outline: none !important; background-color: transparent !important; }
}
`;
}

/** Resolve after the print dialog has been handed the paginated document. */
export async function printDocx(bytes: ArrayBuffer, title: string): Promise<void> {
  const [{ checkedHtml, loadCompareRuntime }, { paginatePreview }] = await Promise.all([
    import("./docx-compare-runtime"),
    import("./paginate-preview"),
  ]);
  const exports = await loadCompareRuntime();
  // Same conversion settings as the compare preview, which are already proven
  // against real documents. `additionalCss` is injected later into the iframe
  // head so the print rules win over the preview stylesheet.
  const html = checkedHtml(exports.DocumentConverter.ConvertDocxToHtmlComplete(
    new Uint8Array(bytes), title || "Document", "docx-", true, "", -1, "comment-",
    0, 1, "page-", false, 0, "annot-", false, false, true, true, true, false, null, false,
  ));

  const frame = document.createElement("iframe");
  // allow-same-origin is required to reach contentDocument for pagination;
  // allow-modals is required for window.print() inside a sandboxed frame.
  // allow-scripts is deliberately absent — the HTML comes from a DOCX file.
  frame.setAttribute("sandbox", "allow-same-origin allow-modals");
  frame.setAttribute("aria-hidden", "true");
  frame.title = title || "print";
  frame.style.cssText = FRAME_STYLE;
  frame.srcdoc = html;

  return new Promise<void>((resolve, reject) => {
    let settled = false;
    const cleanup = () => frame.remove();
    const fail = (error: unknown) => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };

    frame.addEventListener("error", () => fail(new Error("Print frame failed to load")), { once: true });
    frame.addEventListener("load", () => {
      void (async () => {
        const doc = frame.contentDocument;
        const win = frame.contentWindow;
        if (!doc || !win) throw new Error("Print frame is not accessible");
        const { pages, pageSize } = await paginatePreview(doc);
        if (!pages.isConnected) throw new Error("Pagination produced no pages");
        const style = doc.createElement("style");
        style.textContent = printCss(pageSize);
        doc.head.appendChild(style);
        await doc.fonts.ready;
        await Promise.all(Array.from(doc.images).filter(image => !image.complete).map(image =>
          new Promise<void>(done => {
            image.addEventListener("load", () => done(), { once: true });
            image.addEventListener("error", () => done(), { once: true });
          })));
        // Give the injected print rules one frame before the browser snapshots.
        await new Promise<void>(done => requestAnimationFrame(() => done()));
        win.addEventListener("afterprint", cleanup, { once: true });
        win.focus();
        win.print();
        // Some browsers never fire afterprint; don't leave the frame behind.
        window.setTimeout(cleanup, 10_000);
        if (!settled) { settled = true; resolve(); }
      })().catch(fail);
    }, { once: true });

    document.body.appendChild(frame);
  });
}
