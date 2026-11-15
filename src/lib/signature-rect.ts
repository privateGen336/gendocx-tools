/**
 * Pure geometry shared by the signature editor (preview) and the exporter
 * (PDF drawing). Keeping it free of pdf-lib lets the editor import it without
 * pulling the whole PDF writer into its chunk.
 */
export interface SignatureRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Map the preview's displayed page back into the PDF's default user space. */
export function computePdfSignatureRect(
  page: SignatureRect,
  rotation: number,
  placement: Parameters<typeof computeSignatureRect>[2],
  pngRatio: number,
): SignatureRect {
  const angle = ((rotation % 360) + 360) % 360;
  const { x, y, width, height } = page;
  const sideways = angle === 90 || angle === 270;
  const rect = computeSignatureRect(sideways ? height : width, sideways ? width : height, placement, pngRatio);
  switch (angle) {
    case 90: return { ...rect, x: x + width - rect.y, y: y + rect.x };
    case 180: return { ...rect, x: x + width - rect.x, y: y + height - rect.y };
    case 270: return { ...rect, x: x + rect.y, y: y + height - rect.x };
    default: return { ...rect, x: x + rect.x, y: y + rect.y };
  }
}

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));

/**
 * Convert a preview-space box (u,v measured from the top-left, in page
 * fractions) into PDF coordinates, whose origin is bottom-left.
 */
export function computeSignatureRect(
  pageWidth: number,
  pageHeight: number,
  placement: { u: number; v: number; widthRatio: number },
  pngRatio: number,
): SignatureRect {
  const ratio = clamp(placement.widthRatio, 0.05, 0.9);
  let height = (pageWidth * ratio) / pngRatio;
  if (height > pageHeight * 0.5) height = pageHeight * 0.5;
  const width = height * pngRatio;
  const x = clamp(placement.u * pageWidth, 0, Math.max(0, pageWidth - width));
  const y = clamp(pageHeight - placement.v * pageHeight - height, 0, Math.max(0, pageHeight - height));
  return { x, y, width, height };
}
