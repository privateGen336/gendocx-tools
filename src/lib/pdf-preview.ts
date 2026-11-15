import type { PDFDocumentProxy } from "pdfjs-dist";
// Vite emits the worker as a hashed asset next to the bundle. The Next.js build
// served the same file from `public/`; `?url` is the portable equivalent.
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

type PDFPage = Awaited<ReturnType<PDFDocumentProxy["getPage"]>>;

export const WORKER_SRC = workerUrl;

let workerConfigured = false;

/**
 * pdf.js is imported lazily: it costs ~1 MB and the merge mode never needs it.
 * The worker stays out of the initial bundle too.
 */
export async function openPdf(bytes: Uint8Array): Promise<PDFDocumentProxy> {
  const pdfjs = await import("pdfjs-dist");
  if (!workerConfigured) {
    pdfjs.GlobalWorkerOptions.workerSrc = WORKER_SRC;
    workerConfigured = true;
  }
  // pdf.js transfers (detaches) the underlying ArrayBuffer to its worker.
  // Hand it a copy or the caller's bytes become unusable for pdf-lib later.
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return pdfjs.getDocument({ data: copy }).promise;
}

export interface RenderedSize {
  width: number;
  height: number;
}

export interface RenderTaskHandle {
  cancel: () => void;
}

/**
 * pdf.js applies the page's own /Rotate by default. Passing `rotation` to
 * getViewport replaces (not adds to) it, so the delta must be accumulated on
 * top of page.rotate — the same rule the exporter uses with pdf-lib.
 */
export function pageRotation(page: PDFPage, delta = 0): number {
  return (((page.rotate + delta) % 360) + 360) % 360;
}

/** Viewport including a pending user rotation, at any scale. */
export function pageViewport(page: PDFPage, scale: number, delta = 0) {
  return page.getViewport({ scale, rotation: pageRotation(page, delta) });
}

export async function renderPage(
  doc: PDFDocumentProxy,
  /** 1-based physical page number. */
  pageNumber: number,
  canvas: HTMLCanvasElement,
  scale: number,
  /** Lets the caller cancel an in-flight render before reusing the canvas. */
  onTask?: (task: RenderTaskHandle) => void,
  /** Extra clockwise degrees shown in the preview (rotate mode). */
  rotationDelta = 0,
): Promise<RenderedSize> {
  const page = await doc.getPage(pageNumber);
  const viewport = pageViewport(page, scale, rotationDelta);
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  canvas.width = Math.floor(viewport.width * dpr);
  canvas.height = Math.floor(viewport.height * dpr);
  // 只钉宽度，高度交给 CSS 的 height:auto。设了内联高度后，窄屏上
  // `max-width` 把 canvas 压窄时高度不会跟着缩，图会被拉长变形。
  canvas.style.width = `${viewport.width}px`;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("canvas context unavailable");
  const task = page.render({
    canvasContext: context,
    viewport,
    transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0],
  });
  onTask?.(task);
  await task.promise;
  return { width: viewport.width, height: viewport.height };
}
