import { compareDocx, loadCompareRuntime, type DocxComparison } from "./docx-compare-runtime";

let lastResult: { original: Uint8Array; modified: Uint8Array; result: DocxComparison; details: boolean } | undefined;

function sameBytes(a: Uint8Array, b: Uint8Array) {
  return a.length === b.length && a.every((byte, index) => byte === b[index]);
}

self.addEventListener("message", async (event: MessageEvent<{
  id: number;
  original?: ArrayBuffer;
  modified?: ArrayBuffer;
  details?: boolean;
}>) => {
  const { id, original, modified, details = false } = event.data;
  try {
    await loadCompareRuntime();
    if (!original || !modified) return;

    const before = new Uint8Array(original);
    const after = new Uint8Array(modified);
    if (!lastResult || (details && !lastResult.details) || !sameBytes(lastResult.original, before) || !sameBytes(lastResult.modified, after)) {
      const result = await compareDocx(before, after, details);
      lastResult = { original: before, modified: after, result, details };
    }
    self.postMessage({ id, ...lastResult.result });
  } catch (error) {
    self.postMessage({ id, error: error instanceof Error ? error.message : String(error) });
  }
});
