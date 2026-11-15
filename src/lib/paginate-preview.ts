// Keep the document's own section sizes and margins. The viewport only zooms
// these fixed-size pages, so resizing cannot reflow the text into a new width.
/** Page size of the first section, in points, for the print dialog's `@page size`.
 *  Compare only uses `pages`/`width`; the editor's print path uses this too. */
export interface PreviewPageSize { widthPt: number; heightPt: number }

export interface PaginateOptions {
  /**
   * Let content that is wider than the sheet paint outside it.
   *
   * The engine writes `contain: layout paint` on every page box, and paint
   * containment clips descendants to the box no matter what `overflow` says.
   * A table whose fixed column widths run past the text column therefore loses
   * its right-most columns even after the container is widened. Compare turns
   * this on because reading the whole table matters more than a tidy sheet;
   * print leaves it off, where the extra width would land on the next page.
   */
  releaseOverflowClip?: boolean
}

export async function paginatePreview(doc: Document, options: PaginateOptions = {}): Promise<{
  pages: HTMLElement; width: number; pageSize: PreviewPageSize;
}> {
  const { PaginationEngine } = await import("docxodus");
  await doc.fonts.ready;
  const staging = document.createElement("div");
  for (const child of Array.from(doc.body.childNodes)) staging.appendChild(document.importNode(child, true));
  const pages = document.createElement("div");
  pages.className = "compare-pages";
  const style = doc.createElement("style");
  style.textContent = `
    html{background:#e8edf3}
    body{margin:0!important;padding:12px!important;box-sizing:border-box;background:transparent!important;overflow:auto}
    .compare-pages{width:max-content;min-width:0;margin:0 auto}
    .compare-paper-box{background:white;box-shadow:0 2px 12px #18284018;margin:0 auto 18px}
    /* Word lets a table's fixed column widths run past the text column. The
       pagination engine clips every page box and content area to the paper
       width, which silently cut the right-most columns off. Let that overflow
       paint outside the sheet instead; the iframe's own overflow:auto then
       exposes it with a horizontal scrollbar, matching the chat preview.
       overflow alone is not enough — the engine also sets paint containment on
       the box, which clips regardless. See the release step below. */
    .compare-paper-box,.compare-paper-content{overflow:visible!important}
    [data-compare-focus]{outline:3px solid #5471e7;outline-offset:4px;background-color:#edf1ff!important}
  `;
  doc.head.appendChild(style);
  doc.body.replaceChildren(staging, pages);
  // Create staging in the host realm so the engine's native element checks
  // also work for tables inside the same-origin iframe.
  const result = new PaginationEngine(staging, pages, {
    scale: 1,
    cssPrefix: "compare-paper-",
    showPageNumbers: false,
    fragmentParagraphs: true,
  }).paginate();
  if (!result.pages.length) {
    // A valid empty Word document still has one sheet.
    const section = staging.querySelector<HTMLElement>("[data-section-index]");
    const width = Number(section?.dataset.pageWidth) || 612;
    const height = Number(section?.dataset.pageHeight) || 792;
    const blank = document.createElement("div");
    blank.className = "compare-paper-box";
    blank.style.width = `${width}pt`;
    blank.style.height = `${height}pt`;
    pages.appendChild(blank);
    staging.remove();
    return { pages, width: width * 96 / 72, pageSize: { widthPt: width, heightPt: height } };
  }
  staging.remove();
  // An over-wide table paints outside the sheet (see the injected style), but
  // the absolutely positioned content area does not contribute to the scroll
  // width — the iframe would still clip it at its own viewport edge. Widen the
  // pages container to the table's right edge so body{overflow:auto} can
  // actually scroll to the last column.
  const pagesRect = pages.getBoundingClientRect();
  // The sheets are centred with margin:auto, so the container has to grow
  // symmetrically around its centre — measuring from the left edge would shift
  // the centre and leave the table's right edge outside the new width anyway.
  const pagesCenter = pagesRect.left + pagesRect.width / 2;
  let halfWidth = pagesRect.width / 2;
  for (const table of pages.querySelectorAll("table")) {
    halfWidth = Math.max(halfWidth, table.getBoundingClientRect().right - pagesCenter);
  }
  let fitWidth = 0;
  if (halfWidth * 2 > pagesRect.width + 1) {
    fitWidth = Math.ceil(halfWidth * 2 + 24);
    pages.style.minWidth = `${fitWidth}px`;
  }
  // Fit the zoom to the sheet *plus* any table that runs past its right edge,
  // so the default view shows every column instead of starting scrolled. The
  // min-width above remains as the fallback when the user zooms back in.
  const boxes = Array.from(pages.querySelectorAll<HTMLElement>(".compare-paper-box"));
  const contentLeft = boxes.length
    ? Math.min(...boxes.map(box => box.getBoundingClientRect().left))
    : pagesRect.left;
  const contentRight = Math.max(
    ...boxes.map(box => box.getBoundingClientRect().right),
    ...Array.from(pages.querySelectorAll("table")).map(table => table.getBoundingClientRect().right),
  );
  const width = Math.max(contentRight - contentLeft, fitWidth, 1);
  // Drop the paint clip on the sheets that actually overflow, and only those.
  // Every page box is clipped by `contain: layout paint`, which no amount of
  // `overflow: visible` can undo; releasing it document-wide would unclip
  // headers and footers that are supposed to be trimmed at the page edge.
  //
  // Two signals, because neither is complete on its own: `scrollWidth` covers
  // anything that reaches past the sheet (text frames, floating images) whereas
  // the table rects cover the Word case this file was written for no matter how
  // the browser computes overflow on a non-scrolling box.
  if (options.releaseOverflowClip) {
    for (const box of boxes) {
      const boxRect = box.getBoundingClientRect();
      const spills =
        box.scrollWidth > box.clientWidth + 1 ||
        Array.from(box.querySelectorAll("table")).some(table => {
          const rect = table.getBoundingClientRect();
          return rect.right > boxRect.right + 1 || rect.left < boxRect.left - 1;
        });
      if (spills) box.style.contain = "layout";
    }
  }
  const first = result.pages[0];
  return {
    pages,
    width,
    // A document can declare several sections with different sizes; the print
    // dialog only accepts one @page size, so follow the first section.
    pageSize: { widthPt: first.dimensions.pageWidth, heightPt: first.dimensions.pageHeight },
  };
}
