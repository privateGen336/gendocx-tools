"use client";

/**
 * 首屏页面缩略图。
 *
 * 加水印时需要看到「这页长什么样」：只画一行斜字，用户既认不出是哪个文件，
 * 也判断不了水印压上去会不会盖住标题。所以这里取第一页：
 *
 * - PDF 用 pdf.js 真渲染（同一份 worker 和渲染函数，PDF 工具箱已经在用）。
 * - DOCX 不做整版排版（Docxodus + 分页引擎要拉几 MB WASM，只为看一眼首屏不值），
 *   改成从 document.xml 里按段落取真实文字，铺在按 sectPr 页面尺寸算出的纸上。
 *   字号按纸宽换算，看起来就是一页密的、或一页稀的真实文档，而不是假文。
 */

import { strFromU8, unzipSync } from "fflate";
import { openPdf, renderPage } from "../pdf-preview";

export type PagePreview =
  | { kind: "pdf"; image: string; ratio: number }
  | { kind: "docx"; paragraphs: string[]; ratio: number };

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** 缩略图只用第一页，渲染宽度够 2 倍屏就停，再大只是白烧内存。 */
const RENDER_WIDTH = 720;
const MAX_PARAGRAPHS = 18;
const MAX_CHARS = 900;
/** A4 竖向 595×842pt，取不到 sectPr 时的兜底。 */
const A4_RATIO = 595 / 842;

export async function renderPagePreview(bytes: Uint8Array, format: "pdf" | "docx"): Promise<PagePreview> {
  if (format === "docx") return docxPreview(bytes);

  const doc = await openPdf(bytes);
  try {
    const first = await doc.getPage(1);
    const base = first.getViewport({ scale: 1 });
    const scale = Math.min(2.2, RENDER_WIDTH / Math.max(1, base.width));
    const canvas = document.createElement("canvas");
    await renderPage(doc, 1, canvas, scale);
    return { kind: "pdf", image: canvas.toDataURL("image/png"), ratio: base.width / base.height };
  } finally {
    void doc.destroy();
  }
}

function docxPreview(bytes: Uint8Array): PagePreview {
  const zip = unzipSync(bytes, { filter: (file) => file.name === "word/document.xml" });
  const xml = zip["word/document.xml"];
  if (!xml) return { kind: "docx", paragraphs: [], ratio: A4_RATIO };

  const doc = new DOMParser().parseFromString(strFromU8(xml), "application/xml");
  const paragraphs: string[] = [];
  let chars = 0;

  for (const paragraph of Array.from(doc.getElementsByTagNameNS(WORD_NS, "p"))) {
    // 同一个段落里可能被拆成多个 run，拼起来才是这一段的文字。
    const text = Array.from(paragraph.getElementsByTagNameNS(WORD_NS, "t"))
      .map((node) => node.textContent ?? "")
      .join("")
      .replace(/\s+/g, " ")
      .trim();
    if (!text) continue;

    paragraphs.push(text);
    chars += text.length;
    if (paragraphs.length >= MAX_PARAGRAPHS || chars >= MAX_CHARS) break;
  }

  return { kind: "docx", paragraphs, ratio: pageRatio(doc) };
}

/** sectPr 里的 pgSz 以缇（1/20 磅）为单位，取不到就用 A4。 */
function pageRatio(doc: Document): number {
  const size = doc.getElementsByTagNameNS(WORD_NS, "pgSz")[0];
  if (!size) return A4_RATIO;
  const width = Number(size.getAttributeNS(WORD_NS, "w") ?? size.getAttribute("w:w") ?? 0);
  const height = Number(size.getAttributeNS(WORD_NS, "h") ?? size.getAttribute("w:h") ?? 0);
  if (!width || !height) return A4_RATIO;
  const ratio = width / height;
  return ratio > 0.3 && ratio < 3 ? ratio : A4_RATIO;
}
