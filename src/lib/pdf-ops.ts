import { PDFDocument, degrees } from "pdf-lib";
import { computePdfSignatureRect } from "./signature-rect";

// Re-export keeps the public surface (and existing tests) unchanged while the
// editor can import the pure geometry module directly.
export { computeSignatureRect } from "./signature-rect";
export type { SignatureRect } from "./signature-rect";

export type ActionId =
  | "merge"
  | "split"
  | "extract"
  | "rotate"
  | "delete"
  | "reorder"
  | "flatten"
  | "sign";

export interface PdfSource {
  name: string;
  bytes: Uint8Array;
}

export interface PdfOutput {
  name: string;
  bytes: Uint8Array;
  pages: number;
}

export interface PdfResult {
  files: PdfOutput[];
  /** Non-fatal notice, e.g. form fields dropped by copyPages. */
  warning?: string;
  /** copyPages rebuilds the document: bookmarks, metadata and annotations are lost. */
  rebuilt?: boolean;
}

export type PdfErrorCode =
  | "encrypted"
  | "invalid"
  | "range"
  | "empty"
  | "noForm"
  | "tooLarge"
  | "tooManyFiles"
  | "tooManyPages";

export class PdfError extends Error {
  constructor(public code: PdfErrorCode, message?: string) {
    super(message ?? code);
  }
}

/** Guard rails for browser memory: every PDF is loaded whole into an ArrayBuffer. */
export const MAX_FILE_BYTES = 80 * 1024 * 1024;
export const MAX_TOTAL_BYTES = 120 * 1024 * 1024;
export const MAX_PAGES = 400;
export const MAX_FILES = 20;

const baseName = (name: string) => name.replace(/\.pdf$/i, "");
const normalizeAngle = (angle: number) => (((angle % 360) + 360) % 360);

async function load(bytes: Uint8Array): Promise<PDFDocument> {
  try {
    // `ignoreEncryption: true` would read a partially unusable document;
    // surface the problem instead of silently producing a broken file.
    return await PDFDocument.load(bytes);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/encrypt|password|decrypt/i.test(message)) throw new PdfError("encrypted", message);
    throw new PdfError("invalid", message);
  }
}

/**
 * Advanced "by page number" input only. The primary UI lets people click pages
 * directly, so this is the fallback path rather than the default.
 */
export function parseRanges(input: string, count: number): number[] {
  const text = input.trim();
  if (!text || text === "all") return Array.from({ length: count }, (_, i) => i);
  const out = new Set<number>();
  for (const part of text.split(/[,，\s]+/).filter(Boolean)) {
    const match = /^(\d+)\s*(?:[-–]\s*(\d+))?$/.exec(part);
    if (!match) throw new PdfError("range", part);
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : start;
    if (start < 1 || end > count || start > end) throw new PdfError("range", part);
    for (let page = start; page <= end; page += 1) out.add(page - 1);
  }
  if (!out.size) throw new PdfError("range", text);
  return [...out].sort((a, b) => a - b);
}

function checkSize(sources: PdfSource[]): void {
  let total = 0;
  for (const source of sources) {
    if (source.bytes.byteLength > MAX_FILE_BYTES) throw new PdfError("tooLarge", source.name);
    total += source.bytes.byteLength;
  }
  if (total > MAX_TOTAL_BYTES) throw new PdfError("tooLarge", "total");
}

function checkPages(count: number): void {
  if (count > MAX_PAGES) throw new PdfError("tooManyPages");
}

/**
 * Upload-time guard. `runAction` only needs the active file for single-file
 * modes, so callers pass one source there — this check belongs at the point
 * where all files are actually accepted, otherwise the limit would silently
 * block unrelated operations later (e.g. extracting one page from file #1
 * when 21 files are open).
 */
export function checkUpload(sources: PdfSource[]): void {
  if (sources.length > MAX_FILES) throw new PdfError("tooManyFiles");
  checkSize(sources);
}

async function save(doc: PDFDocument, name: string): Promise<PdfOutput> {
  const bytes = await doc.save({ useObjectStreams: true });
  return { name, bytes, pages: doc.getPageCount() };
}

/** Detect AcroForm fields that copyPages will silently drop. */
function hasFormFields(doc: PDFDocument): boolean {
  try {
    return doc.getForm().getFields().length > 0;
  } catch {
    return false;
  }
}

async function clonePages(source: PDFDocument, indices: number[]): Promise<PDFDocument> {
  const out = await PDFDocument.create();
  const copied = await out.copyPages(source, indices);
  copied.forEach((page) => out.addPage(page));
  return out;
}

export interface SignaturePlacement {
  /** Source page index the signature belongs to. */
  pageIndex: number;
  /** Left edge of the signature box, as a fraction of page width. */
  u: number;
  /** Top edge of the signature box, as a fraction of page height. */
  v: number;
  /** Signature width as a fraction of page width. */
  widthRatio: number;
  png: Uint8Array;
}

export interface ComposeOptions {
  /** Source page indices in the exact order the user arranged on the page wall. */
  order: number[];
  /** Extra clockwise rotation keyed by source page index. */
  rotations?: Record<number, number>;
  sign?: SignaturePlacement;
  name?: string;
}

interface ComposeOutput extends PdfOutput {
  warning?: string;
  rebuilt?: boolean;
}

/**
 * The single primitive behind extract / delete / rotate / reorder / sign.
 * Rebuilding from `order` means what the user saw on the page wall is exactly
 * what comes out — no page-number bookkeeping in the UI.
 */
export async function compose(source: PdfSource, options: ComposeOptions): Promise<ComposeOutput> {
  checkSize([source]);
  checkPages(options.order.length);
  if (!options.order.length) throw new PdfError("empty");
  const doc = await load(source.bytes);
  const hadForm = hasFormFields(doc);
  const out = await clonePages(doc, options.order);

  for (let position = 0; position < options.order.length; position += 1) {
    const delta = options.rotations?.[options.order[position]];
    if (!delta) continue;
    const page = out.getPage(position);
    // copyPages preserves /Rotate, so accumulate rather than overwrite.
    page.setRotation(degrees(normalizeAngle(page.getRotation().angle + delta)));
  }

  if (options.sign) {
    const position = options.order.indexOf(options.sign.pageIndex);
    if (position >= 0) {
      const page = out.getPage(position);
      const media = page.getMediaBox();
      const crop = page.getCropBox();
      const x = Math.max(media.x, crop.x);
      const y = Math.max(media.y, crop.y);
      const width = Math.min(media.x + media.width, crop.x + crop.width) - x;
      const height = Math.min(media.y + media.height, crop.y + crop.height) - y;
      const view = width > 0 && height > 0 ? { x, y, width, height } : media;
      const rotation = page.getRotation();
      const png = await out.embedPng(options.sign.png);
      page.drawImage(png, { ...computePdfSignatureRect(view, rotation.angle, options.sign, png.width / png.height), rotate: rotation });
    }
  }

  const name = options.name ?? `${baseName(source.name)}-edited.pdf`;
  return {
    ...(await save(out, name)),
    warning: hadForm ? "form" : undefined,
    rebuilt: true,
  };
}

/**
 * Split at visual boundaries. `cuts` holds 1-based positions in `order`:
 * a cut at n means "end the current part after the nth page on the wall".
 */
export async function splitAtBoundaries(
  source: PdfSource,
  order: number[],
  cuts: number[],
): Promise<ComposeOutput[]> {
  const boundaries = [...new Set(cuts.map((n) => Math.round(n)))]
    .filter((n) => n > 0 && n < order.length)
    .sort((a, b) => a - b);
  if (!boundaries.length) throw new PdfError("empty");
  checkSize([source]);
  checkPages(order.length);
  const parts: number[][] = [];
  let start = 0;
  for (const cut of boundaries) {
    parts.push(order.slice(start, cut));
    start = cut;
  }
  parts.push(order.slice(start));
  // Parse the source exactly once. Calling compose() per part re-ran
  // PDFDocument.load for every output file — splitting 400 pages into 400
  // parts meant 400 full parses and 400 copies of the source in memory.
  const doc = await load(source.bytes);
  const hadForm = hasFormFields(doc);
  const files: ComposeOutput[] = [];
  for (let index = 0; index < parts.length; index += 1) {
    const out = await clonePages(doc, parts[index]);
    files.push({
      ...(await save(out, `${baseName(source.name)}-${index + 1}.pdf`)),
      warning: hadForm ? "form" : undefined,
      rebuilt: true,
    });
  }
  return files;
}

export async function merge(sources: PdfSource[]): Promise<PdfResult> {
  if (sources.length < 2) throw new PdfError("empty", "merge");
  checkSize(sources);
  const out = await PDFDocument.create();
  let pages = 0;
  let formWarning = false;
  for (const source of sources) {
    const doc = await load(source.bytes);
    checkPages(pages + doc.getPageCount());
    pages += doc.getPageCount();
    if (hasFormFields(doc)) formWarning = true;
    const copied = await out.copyPages(doc, doc.getPageIndices());
    copied.forEach((page) => out.addPage(page));
  }
  return {
    files: [await save(out, "merged.pdf")],
    warning: formWarning ? "form" : undefined,
    rebuilt: true,
  };
}

export async function flatten(source: PdfSource): Promise<PdfResult> {
  checkSize([source]);
  const doc = await load(source.bytes);
  if (!hasFormFields(doc)) throw new PdfError("noForm");
  // Flatten must run on the source document. copyPages never copies form
  // fields, so flattening a composed document would be a no-op.
  doc.getForm().flatten();
  return { files: [await save(doc, `${baseName(source.name)}-flattened.pdf`)] };
}

/** Read page count and form presence once, for the page wall. */
export async function inspect(source: PdfSource): Promise<{ pages: number; hasForm: boolean }> {
  checkSize([source]);
  const doc = await load(source.bytes);
  checkPages(doc.getPageCount());
  return { pages: doc.getPageCount(), hasForm: hasFormFields(doc) };
}

export interface RunParams {
  order: number[];
  selected: number[];
  cuts: number[];
  rotation: number;
  sign?: SignaturePlacement;
}

export async function runAction(
  action: ActionId,
  sources: PdfSource[],
  params: RunParams,
): Promise<PdfResult> {
  if (!sources.length) throw new PdfError("empty");
  if (sources.length > MAX_FILES) throw new PdfError("tooManyFiles");
  const first = sources[0];
  const { order, selected, cuts, rotation, sign } = params;
  const chosen = new Set(selected);
  const fromCompose = (file: ComposeOutput): PdfResult => ({
    files: [file],
    warning: file.warning,
    rebuilt: file.rebuilt,
  });

  switch (action) {
    case "merge":
      return merge(sources);
    case "flatten":
      return flatten(first);
    case "split": {
      const files = await splitAtBoundaries(first, order, cuts);
      return {
        files,
        warning: files.some((file) => file.warning) ? "form" : undefined,
        rebuilt: files.some((file) => file.rebuilt),
      };
    }
    case "extract": {
      const kept = order.filter((index) => chosen.has(index));
      if (!kept.length) throw new PdfError("empty");
      return fromCompose(await compose(first, { order: kept, name: `${baseName(first.name)}-extract.pdf` }));
    }
    case "delete": {
      const kept = order.filter((index) => !chosen.has(index));
      if (!kept.length) throw new PdfError("range", "all");
      return fromCompose(await compose(first, { order: kept, name: `${baseName(first.name)}-trimmed.pdf` }));
    }
    case "rotate": {
      const rotations: Record<number, number> = {};
      for (const index of selected) rotations[index] = rotation;
      return fromCompose(await compose(first, { order, rotations, name: `${baseName(first.name)}-rotated.pdf` }));
    }
    case "reorder": {
      return fromCompose(await compose(first, { order, name: `${baseName(first.name)}-reordered.pdf` }));
    }
    case "sign": {
      if (!sign) throw new PdfError("empty", "sign");
      return fromCompose(await compose(first, { order, sign, name: `${baseName(first.name)}-signed.pdf` }));
    }
  }
}
