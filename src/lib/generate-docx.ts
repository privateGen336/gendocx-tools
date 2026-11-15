import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  PageBreak,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  VerticalAlign,
  WidthType,
} from "docx";

export type PaperSize = "fit" | "a4" | "letter";
export type Orientation = "auto" | "portrait" | "landscape";
/** What the generator actually draws with — "auto" is resolved by the caller. */
export type ResolvedOrientation = Exclude<Orientation, "auto">;
export type PerPage = 1 | 2 | 4;
export type MarginMm = 0 | 10 | 20;
export type ImageEncoding = "auto" | "jpg" | "png";

/** Distinguishes an undecodable file from a generic generation failure. */
export class ImageDecodeError extends Error {
  constructor(public fileName: string) {
    super(`Cannot decode image: ${fileName}`);
    this.name = "ImageDecodeError";
  }
}

/**
 * "Fit to image" derives the page from the first image. A 90°/270° turn swaps
 * the axes; scaling is proportional so a 4:3 photo never becomes square, and
 * each edge stays inside what Word accepts as a page (5–50 cm).
 */
export function computeFitSize(width: number, height: number, rotation: number): { width: number; height: number } {
  if (!width || !height) return { width: 210, height: 297 };
  const swapped = rotation % 180 !== 0;
  let w = (swapped ? height : width) / PX_PER_MM;
  let h = (swapped ? width : height) / PX_PER_MM;
  const shrink = Math.min(1, 500 / Math.max(w, h));
  w *= shrink;
  h *= shrink;
  const shortSide = Math.min(w, h);
  if (shortSide > 0 && shortSide < 50) {
    const grow = Math.min(50 / shortSide, 500 / Math.max(w, h));
    w *= grow;
    h *= grow;
  }
  const round = (value: number) => Math.round(value * 10) / 10;
  return { width: round(w), height: round(h) };
}

export type SourceImage = {
  file: File;
  /** Clockwise degrees: 0, 90, 180 or 270. */
  rotation: number;
};

export type GenerateOptions = {
  paper: PaperSize;
  /** Used when paper is "fit": the first image's size in millimetres. */
  fitSize?: { width: number; height: number };
  orientation: ResolvedOrientation;
  perPage: PerPage;
  /** Left and right margin, in millimetres. */
  marginXMm: MarginMm;
  /** Top and bottom margin, in millimetres. */
  marginYMm: MarginMm;
  /** "standard" rasterises at screen density; "original" up to 350 dpi. */
  quality: "standard" | "original";
  /** auto = PNG for images with transparency, JPEG otherwise. */
  encoding: ImageEncoding;
  /** Allow small images to be enlarged to fill their slot (matches the preview). */
  upscale: boolean;
  onProgress?: (done: number, total: number) => void;
};

export const PAPER_MM: Record<Exclude<PaperSize, "fit">, { width: number; height: number }> = {
  a4: { width: 210, height: 297 },
  letter: { width: 215.9, height: 279.4 },
};

export const PX_PER_MM = 96 / 25.4; // CSS pixels at 96 dpi
const TWIP_PER_MM = 56.692913;
const SLOT_GAP_PX = 10;
const JPEG_QUALITY = 0.9;
/** "original" stops here: beyond ~350 dpi a Word page gains nothing visible. */
const OUTPUT_DPI = 350;
/** Browsers reject canvases beyond roughly 16k px per side; stay well below. */
const MAX_CANVAS_EDGE = 8192;

const mmToTwip = (mm: number) => Math.round(mm * TWIP_PER_MM);
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = {
  top: noBorder,
  bottom: noBorder,
  left: noBorder,
  right: noBorder,
  insideHorizontal: noBorder,
  insideVertical: noBorder,
};

/** Cheap alpha probe: downscale to 64px and look for any non-opaque pixel. */
function hasTransparency(bitmap: ImageBitmap): boolean {
  const probe = document.createElement("canvas");
  probe.width = Math.min(64, bitmap.width);
  probe.height = Math.min(64, bitmap.height);
  const context = probe.getContext("2d", { willReadFrequently: true });
  if (!context || !probe.width || !probe.height) return false;
  context.drawImage(bitmap, 0, 0, probe.width, probe.height);
  const data = context.getImageData(0, 0, probe.width, probe.height).data;
  for (let i = 3; i < data.length; i += 4) {
    if (data[i] < 250) return true;
  }
  return false;
}

async function renderImage(
  image: SourceImage,
  slotWidth: number,
  slotHeight: number,
  quality: "standard" | "original",
  encoding: ImageEncoding,
  upscale: boolean,
): Promise<{ data: Uint8Array; width: number; height: number; type: "jpg" | "png" }> {
  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(image.file);
  } catch {
    throw new ImageDecodeError(image.file.name);
  }
  const rotation = ((image.rotation % 360) + 360) % 360;
  const swapped = rotation === 90 || rotation === 270;
  const sourceWidth = swapped ? bitmap.height : bitmap.width;
  const sourceHeight = swapped ? bitmap.width : bitmap.height;
  const usePng = encoding === "png" || (encoding === "auto" && hasTransparency(bitmap));

  // Layout size on the page: fit inside the slot. The preview always scales
  // up to fill its cell, so "fit page" must be able to do the same.
  const displayScale = Math.min(slotWidth / sourceWidth, slotHeight / sourceHeight, upscale ? Infinity : 1);
  const displayWidth = Math.max(1, Math.round(sourceWidth * displayScale));
  const displayHeight = Math.max(1, Math.round(sourceHeight * displayScale));

  // "standard" rasterises at 96 dpi (one pixel per layout pixel). "original"
  // caps at 350 dpi of the printed size — keeping 4000 px for a 90 mm cell
  // quadruples the file without a visible gain; the canvas guard is the only
  // other limit.
  const targetWidthPx = (displayWidth / 96) * OUTPUT_DPI;
  const pixelScale = quality === "original"
    ? Math.min(1, targetWidthPx / sourceWidth, MAX_CANVAS_EDGE / Math.max(sourceWidth, sourceHeight))
    : displayScale;
  const drawWidth = Math.max(1, Math.round(sourceWidth * pixelScale));
  const drawHeight = Math.max(1, Math.round(sourceHeight * pixelScale));

  const canvas = document.createElement("canvas");
  canvas.width = drawWidth;
  canvas.height = drawHeight;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas is unavailable");

  // JPEG has no alpha channel; a white matte keeps transparent PNGs from
  // turning black. PNG output keeps its alpha intact and is lossless for
  // screenshots and line art.
  if (!usePng) {
    context.fillStyle = "#FFFFFF";
    context.fillRect(0, 0, drawWidth, drawHeight);
  }
  context.translate(drawWidth / 2, drawHeight / 2);
  context.rotate((rotation * Math.PI) / 180);
  const bitmapWidth = bitmap.width * pixelScale;
  const bitmapHeight = bitmap.height * pixelScale;
  context.drawImage(bitmap, -bitmapWidth / 2, -bitmapHeight / 2, bitmapWidth, bitmapHeight);
  bitmap.close?.();

  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob(
      (value) => (value ? resolve(value) : reject(new Error("Image encoding failed"))),
      usePng ? "image/png" : "image/jpeg",
      usePng ? undefined : JPEG_QUALITY,
    );
  });

  return {
    data: new Uint8Array(await blob.arrayBuffer()),
    width: displayWidth,
    height: displayHeight,
    type: usePng ? "png" : "jpg",
  };
}

function imageParagraph(
  image: { data: Uint8Array; width: number; height: number; type: "jpg" | "png" },
): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [
      new ImageRun({
        type: image.type,
        data: image.data,
        transformation: { width: image.width, height: image.height },
      }),
    ],
  });
}

export async function generateImageDocx(
  images: SourceImage[],
  options: GenerateOptions,
): Promise<{ blob: Blob; pages: number }> {
  const { paper, fitSize, orientation, perPage, marginXMm, marginYMm, quality, encoding, upscale, onProgress } = options;
  const paperSize = paper === "fit" ? (fitSize ?? PAPER_MM.a4) : PAPER_MM[paper];
  // A fitted page already carries the image's own orientation; rotating it by
  // `orientation` again would stand a landscape photo on its side.
  const pageWidthMm = paper === "fit" ? paperSize.width : orientation === "landscape" ? paperSize.height : paperSize.width;
  const pageHeightMm = paper === "fit" ? paperSize.height : orientation === "landscape" ? paperSize.width : paperSize.height;

  const contentWidthPx = (pageWidthMm - marginXMm * 2) * PX_PER_MM;
  const contentHeightPx = (pageHeightMm - marginYMm * 2) * PX_PER_MM;
  const columns = perPage === 4 ? 2 : 1;
  const rows = perPage === 4 ? 2 : perPage;
  const slotWidth = (contentWidthPx - SLOT_GAP_PX * (columns - 1)) / columns;
  const slotHeight = (contentHeightPx - SLOT_GAP_PX * (rows - 1)) / rows;
  const pages = Math.ceil(images.length / perPage);

  const children: Array<Paragraph | Table> = [];
  let done = 0;

  for (let start = 0; start < images.length; start += perPage) {
    if (start > 0) children.push(new Paragraph({ children: [new PageBreak()] }));

    const batch = images.slice(start, start + perPage);
    const rendered: Array<{ data: Uint8Array; width: number; height: number; type: "jpg" | "png" }> = [];
    for (const image of batch) {
      rendered.push(await renderImage(image, slotWidth, slotHeight, quality, encoding, upscale));
      done += 1;
      onProgress?.(done, images.length);
    }

    if (perPage === 4) {
      const cells = rendered.map((image) => new TableCell({
        width: { size: 50, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [imageParagraph(image)],
      }));
      while (cells.length < 4) {
        cells.push(new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({ children: [] })],
        }));
      }
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorders,
        rows: [
          new TableRow({ children: cells.slice(0, 2) }),
          new TableRow({ children: cells.slice(2, 4) }),
        ],
      }));
    } else {
      // No paragraph spacing: slotHeight already reserves SLOT_GAP_PX between
      // rows, and an extra gap pushed two-per-page layouts onto a third page
      // in Word even though the preview showed two.
      rendered.forEach((image) => {
        children.push(imageParagraph(image));
      });
    }
  }

  const documentFile = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: mmToTwip(pageWidthMm), height: mmToTwip(pageHeightMm) },
          margin: {
            top: mmToTwip(marginYMm),
            bottom: mmToTwip(marginYMm),
            left: mmToTwip(marginXMm),
            right: mmToTwip(marginXMm),
          },
        },
      },
      children,
    }],
  });

  return { blob: await Packer.toBlob(documentFile), pages };
}

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
