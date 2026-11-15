import { zipSync } from "fflate";
import type { PdfOutput } from "./pdf-ops";

function trigger(bytes: Uint8Array, name: string, mime: string): void {
  // Copy into a fresh buffer so the Blob never aliases a WASM-owned view.
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  const blob = new Blob([copy], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  // Give the browser a tick to start the download before revoking.
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * A single output downloads directly. Split produces several files, which
 * must be zipped: browsers block a burst of programmatic downloads.
 */
export function downloadOutputs(files: PdfOutput[], zipName: string): void {
  if (files.length === 1) {
    trigger(files[0].bytes, files[0].name, "application/pdf");
    return;
  }
  const entries: Record<string, Uint8Array> = {};
  for (const file of files) entries[file.name] = file.bytes;
  trigger(zipSync(entries, { level: 0 }), zipName, "application/zip");
}

export async function toSources(files: File[]): Promise<{ name: string; bytes: Uint8Array }[]> {
  return Promise.all(
    files.map(async (file) => ({
      name: file.name,
      bytes: new Uint8Array(await file.arrayBuffer()),
    })),
  );
}
