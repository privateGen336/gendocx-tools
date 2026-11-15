// Narrow adapter for Docxodus 9.9.0's existing WASM exports.
// Derived from Docxodus's MIT-licensed dist/index.js. See docxodus-LICENSE.txt.
import type { Revision } from "docxodus";

interface CompareExports {
  DocxDiffBridge: {
    Compare(before: Uint8Array, after: Uint8Array, settings: string): Uint8Array;
  };
  DocumentConverter: {
    ConvertDocxToHtmlComplete(
      bytes: Uint8Array, pageTitle: string, cssPrefix: string, fabricateClasses: boolean,
      additionalCss: string, commentRenderMode: number, commentCssClassPrefix: string,
      paginationMode: number, paginationScale: number, paginationCssClassPrefix: string,
      renderAnnotations: boolean, annotationLabelMode: number, annotationCssClassPrefix: string,
      renderFootnotesAndEndnotes: boolean, renderHeadersAndFooters: boolean,
      renderTrackedChanges: boolean, showDeletedContent: boolean, renderMoveOperations: boolean,
      renderUnsupportedContentPlaceholders: boolean, documentLanguage: string | null,
      stampAnchors: boolean,
    ): string;
  };
  DocumentComparer: {
    GetRevisionsJson(compared: Uint8Array): string;
    CompareDocumentsToHtmlWithOptions(
      before: Uint8Array, after: Uint8Array, author: string, tracked: boolean, engine: number,
    ): string;
  };
}
let runtime: Promise<CompareExports> | undefined;

export function loadCompareRuntime(moduleUrl?: string): Promise<CompareExports> {
  if (runtime) return runtime;
  const path = moduleUrl ?? new URL("/wasm/_framework/dotnet.js", self.location.origin).href;
  runtime = (async () => {
    const { dotnet } = await import(/* webpackIgnore: true */ /* @vite-ignore */ path);
    const { getAssemblyExports, getConfig } = await dotnet.withDiagnosticTracing(false).create();
    const exports = await getAssemblyExports(getConfig().mainAssemblyName);
    return exports.DocxodusWasm as CompareExports;
  })().catch(error => {
    runtime = undefined;
    throw error;
  });
  return runtime;
}

export function checkedHtml(html: string): string {
  // The WASM bridge serializes conversion failures as JSON strings.
  if (html.trimStart().startsWith("{")) {
    let result;
    try { result = JSON.parse(html); } catch { return html; }
    if (result && typeof result === "object" && "Error" in result) {
      throw new Error(String(result.Error));
    }
  }
  return html;
}

export interface DocxComparison {
  html: string;
  revisions?: Revision[];
  redline?: Uint8Array;
}

export async function compareDocx(before: Uint8Array, after: Uint8Array, includeDetails = false): Promise<DocxComparison> {
  const exports = await loadCompareRuntime();
  const redline = exports.DocxDiffBridge.Compare(before, after, JSON.stringify({
    authorForRevisions: "Compare",
    formatComparison: 1, // DocxDiffFormatComparison.Full in 9.9.0
    trackBlockFormatChanges: true,
    preserveInputRevisions: true,
  }));
  if (!redline.length) throw new Error("DocxDiff comparison failed - empty result");
  let html: string;
  try {
    html = checkedHtml(exports.DocumentConverter.ConvertDocxToHtmlComplete(
      redline, "Document", "docx-", true, "", -1, "comment-", 0, 1, "page-",
      false, 0, "annot-", false, false, true, true, true, false, null, false,
    ));
  } catch (error) {
    if (!(error instanceof Error) || !error.message.includes("InvalidOperation_EmptyStack")) throw error;
    html = checkedHtml(exports.DocumentComparer.CompareDocumentsToHtmlWithOptions(
      before, after, "Compare", true, 0, // WmlComparer compatibility path
    ));
  }
  if (!includeDetails) return { html };
  const data = JSON.parse(checkedHtml(exports.DocumentComparer.GetRevisionsJson(redline)));
  // This 9.9.0 bridge returns PascalCase records. Normalize once at the boundary.
  const revisions: Revision[] = (data.Revisions ?? data.revisions ?? []).map((r: Record<string, unknown>) => {
    const format = (r.FormatChange ?? r.formatChange) as Record<string, unknown> | undefined;
    return {
      revisionType: r.RevisionType ?? r.revisionType,
      text: r.Text ?? r.text ?? "",
      author: r.Author ?? r.author ?? "",
      date: r.Date ?? r.date ?? "",
      moveGroupId: r.MoveGroupId ?? r.moveGroupId,
      isMoveSource: r.IsMoveSource ?? r.isMoveSource,
      formatChange: format ? {
        oldProperties: format.OldProperties ?? format.oldProperties,
        newProperties: format.NewProperties ?? format.newProperties,
        changedPropertyNames: format.ChangedPropertyNames ?? format.changedPropertyNames,
        scope: format.Scope ?? format.scope,
      } : undefined,
    } as Revision;
  });
  return { html, revisions, redline };
}

export async function compareDocxHtml(before: Uint8Array, after: Uint8Array): Promise<string> {
  return (await compareDocx(before, after)).html;
}
