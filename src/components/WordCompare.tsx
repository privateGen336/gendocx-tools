import { useEffect, useMemo, useRef, useState } from "react";
import { useLang, useT } from "../i18n";
import { ArrowDownToLine, ArrowLeftRight, ArrowRight, Check, FolderOpen, LoaderCircle, LockKeyhole, Minus, PencilLine, Plus, RotateCcw, Search, Upload, X } from "lucide-react";
import { storeEditorDocument } from "../lib/local-document-transfer";
import { useHoverTooltip } from "./use-hover-tooltip";
import type { Revision } from "docxodus";
import type { DocxComparison } from "../lib/docx-compare-runtime";
import hero from "../styles/tool-hero.module.css";
import styles from "../styles/word-compare.module.css";
import { paginatePreview } from "../lib/paginate-preview";

const kinds = ["all", "Inserted", "Deleted", "FormatChanged", "Moved"] as const;
const properties = ["fontSize", "bold", "italic", "fontFamily", "color", "underline", "alignment", "spacing", "shell", "grid"];
type Translate = (key: string, values?: Record<string, string | number>) => string;

/** Bring the finished comparison into view: the result section is the last block on the page. */
function scrollToPageBottom() {
  const doc = document.documentElement;
  const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
  let cancelled = false;
  const cancel = () => { cancelled = true; };
  const cleanup = () => {
    window.removeEventListener("wheel", cancel);
    window.removeEventListener("touchstart", cancel);
  };
  // Hand control back the moment the user scrolls, so the corrections never fight them.
  window.addEventListener("wheel", cancel, { passive: true, once: true });
  window.addEventListener("touchstart", cancel, { passive: true, once: true });

  // The global stylesheet sets `scroll-behavior: smooth`, so a plain
  // `scrollTo` animates and reads `scrollHeight` at call time. Corrections have
  // to be instant or they stop at the pre-render page height.
  const snap = () => {
    const previous = doc.style.scrollBehavior;
    doc.style.scrollBehavior = "auto";
    doc.scrollTop = doc.scrollHeight;
    doc.style.scrollBehavior = previous;
  };

  if (reduced) snap();
  else window.scrollTo({ top: doc.scrollHeight, behavior: "smooth" });

  // The result pane keeps growing after its first paint (iframe, fonts, table
  // layout), so the first scroll lands short. Re-align until the page settles.
  const deadline = Date.now() + 1500;
  const topUp = () => {
    if (cancelled) { cleanup(); return; }
    if (doc.scrollHeight - window.scrollY - window.innerHeight > 4) snap();
    if (Date.now() < deadline) window.setTimeout(topUp, 150);
    else cleanup();
  };
  window.setTimeout(topUp, 600);
}

function readable(value: string | undefined, property: string, t: Translate) {
  if (value === undefined || value === "") return t("default");
  if (property === "fontSize" && /^\d+(\.\d+)?$/.test(value)) return `${Number(value) / 2} pt`;
  if (value === "true") return t("yes");
  if (value === "false") return t("no");
  if (value.trimStart().startsWith("<")) {
    if (property === "bold" || property === "italic") return /(?:w:)?val=["'](?:0|false|off)["']/.test(value) ? t("no") : t("yes");
    const attribute = value.match(/(?:w:)?val=["']([^"']+)["']/)?.[1];
    if (attribute) return attribute;
    return t("adjusted");
  }
  return value;
}

function changeLabel(revision: Revision, t: Translate) {
  const scope = revision.formatChange?.scope;
  if (revision.revisionType === "FormatChanged" && scope && scope !== "run") {
    return t(`scopes.${scope}`);
  }
  return kinds.includes(revision.revisionType as typeof kinds[number]) ? t(`kinds.${revision.revisionType}`) : revision.revisionType;
}

function FileSlot({ file, title, onFile, disabled }: {
  file: File | null; title: string; onFile: (file: File) => void; disabled: boolean;
}) {
  const t = useT("wordCompare");
  const input = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  return (
    <div className={`${styles.fileSlot} ${dragging ? styles.dragging : ""}`}
      onDragOver={event => { event.preventDefault(); if (!disabled) setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={event => {
        event.preventDefault(); setDragging(false);
        if (!disabled && event.dataTransfer.files[0]) onFile(event.dataTransfer.files[0]);
      }}>
      <span className={styles.slotLabel}>{title}</span>
      <button className={styles.fileButton} disabled={disabled} onClick={() => input.current?.click()}>
        <span className={styles.fileIcon}>{file ? <i className="ri-file-word-line" aria-hidden="true" /> : <Upload size={22} />}</span>
        <span className={styles.fileName}><strong>{file?.name ?? t("drop")}</strong>
          <small>{file ? t("replace", { size: (file.size / 1024).toFixed(1) }) : t("limit")}</small>
        </span>
        {file ? <Check size={18} className={styles.fileCheck} /> : <FolderOpen size={18} />}
      </button>
      <input ref={input} type="file" accept=".docx" aria-label={t("select", { title })} hidden
        onChange={event => { const selected = event.target.files?.[0]; if (selected) onFile(selected); event.target.value = ""; }} />
    </div>
  );
}

export default function WordCompare() {
  const t = useT("wordCompare");
  const locale = useLang().lang;
  const [before, setBefore] = useState<File | null>(null);
  const [after, setAfter] = useState<File | null>(null);
  const [result, setResult] = useState<DocxComparison | null>(null);
  const [busy, setBusy] = useState(false);
  const [demoBusy, setDemoBusy] = useState(false);
  const [error, setError] = useState("");
  const [filter, setFilter] = useState("all");
  const [query, setQuery] = useState("");
  const [active, setActive] = useState<number | null>(null);
  const [elapsed, setElapsed] = useState(0);
  const [progress, setProgress] = useState(0);
  const tip = useHoverTooltip();
  const tShared = useT("shared");
  const editDocumentLabel = `${t("editDocument")} (${tShared("nav.newTab")})`;
  const [zoom, setZoom] = useState(100);
  const [openingEditor, setOpeningEditor] = useState(false);
  const worker = useRef<Worker | null>(null);
  const requestId = useRef(0);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const frame = useRef<HTMLIFrameElement>(null);
  const sampleCache = useRef<{ original: File; modified: File; result: DocxComparison } | null>(null);
  const previewPages = useRef<HTMLElement | null>(null);
  const paperWidth = useRef(816);
  const manualZoom = useRef(false);

  // The comparison runs inside a worker with no progress channel, so the bar
  // follows elapsed time: a quick read phase, then a slowing approach to 90%.
  // It never claims 100% before the result actually lands.
  useEffect(() => {
    if (!busy) return;
    const start = performance.now();
    const tick = () => {
      const ms = performance.now() - start;
      setElapsed(ms);
      setProgress(ms < 1000 ? (ms / 1000) * 25 : 25 + 65 * (1 - Math.exp(-(ms - 1000) / 9000)));
    };
    const frame = window.requestAnimationFrame(tick);
    const timer = window.setInterval(tick, 120);
    return () => {
      window.cancelAnimationFrame(frame);
      window.clearInterval(timer);
    };
  }, [busy]);

  useEffect(() => {
    if (previewPages.current) previewPages.current.style.zoom = String(zoom / 100);
  }, [zoom]);

  useEffect(() => {
    const iframe = frame.current;
    if (!iframe) return;
    const observer = new ResizeObserver(() => {
      if (manualZoom.current || !previewPages.current) return;
      const available = iframe.contentDocument?.documentElement?.clientWidth || iframe.clientWidth;
      const scale = Math.max(25, Math.min(100, Math.floor((available - 24) / paperWidth.current * 100)));
      previewPages.current.style.zoom = String(scale / 100);
      setZoom(scale);
    });
    observer.observe(iframe);
    return () => observer.disconnect();
  }, [result]);

  function changeZoom(value: number) {
    manualZoom.current = true;
    setZoom(value);
  }

  async function preparePreview(iframe: HTMLIFrameElement) {
    const doc = iframe.contentDocument;
    if (!doc) return;
    if (doc.body) doc.body.style.visibility = "hidden";
    previewPages.current = null;
    try {
      const { pages, width } = await paginatePreview(doc, { releaseOverflowClip: true });
      if (frame.current?.contentDocument !== doc) return;
      previewPages.current = pages;
      paperWidth.current = width;
      manualZoom.current = false;
      const available = doc.documentElement?.clientWidth || iframe.clientWidth;
      const scale = Math.max(25, Math.min(100, Math.floor((available - 24) / width * 100)));
      pages.style.zoom = String(scale / 100);
      setZoom(scale);
      if (doc.body) doc.body.style.visibility = "visible";
    } catch {
      if (frame.current?.contentDocument === doc) setError(t("errors.compare"));
    }
  }
  useEffect(() => () => {
    requestId.current++;
    worker.current?.terminate();
    if (timer.current) clearTimeout(timer.current);
  }, []);

  function getWorker() {
    if (!worker.current) {
      worker.current = new Worker(new URL("../lib/docx-compare.worker.ts", import.meta.url), { type: "module" });
      worker.current.postMessage({ id: 0 });
    }
    return worker.current;
  }

  function selectFile(file: File, side: "before" | "after") {
    if (!file.name.toLowerCase().endsWith(".docx")) { setError(t("errors.type")); return; }
    if (file.size > 20 * 1024 * 1024) { setError(t("errors.size")); return; }
    requestId.current++;
    if (side === "before") setBefore(file); else setAfter(file);
    setResult(null); setError(""); setActive(null);
    try { getWorker(); } catch { /* A comparison will surface worker startup errors. */ }
  }

  async function compare(original = before, modified = after) {
    if (!original || !modified) return;
    const id = ++requestId.current;
    setBusy(true); setProgress(0); setError(""); setActive(null); setFilter("all"); setQuery("");
    const start = performance.now();
    try {
      const cached = sampleCache.current;
      if (cached && cached.original === original && cached.modified === modified) {
        setResult(cached.result);
        setElapsed(performance.now() - start);
        setTimeout(scrollToPageBottom, 50);
        return;
      }
      const engine = getWorker();
      const comparison = new Promise<DocxComparison>((resolve, reject) => {
        engine.onmessage = ({ data }: MessageEvent<DocxComparison & { id: number; error?: string }>) => {
          if (data.id !== id) return;
          if (data.error) reject(new Error(data.error)); else resolve(data);
        };
        engine.onerror = () => {
          engine.terminate(); worker.current = null;
          reject(new Error(t("errors.engine")));
        };
        timer.current = setTimeout(() => {
          engine.terminate(); worker.current = null;
          reject(new Error(t("errors.timeout")));
        }, 90_000);
      });
      // Attach the rejection handler before reading files so a timeout cannot
      // create an unhandled rejection while file IO is in flight.
      const documents = Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
      const [data] = await Promise.all([comparison, documents.then(([a, b]) => {
        engine.postMessage({ id, original: a, modified: b, details: true }, [a, b]);
      })]);
      if (requestId.current !== id) return;
      setResult(data); setElapsed(performance.now() - start);
      setTimeout(scrollToPageBottom, 50);
    } catch (failure) {
      if (requestId.current === id) setError(failure instanceof Error ? failure.message : t("errors.compare"));
    } finally {
      if (timer.current) clearTimeout(timer.current);
      if (requestId.current === id) setBusy(false);
    }
  }

  async function demo() {
    const id = ++requestId.current;
    const start = performance.now();
    setDemoBusy(true); setError("");
    try {
      const load = async (name: string) => {
        const response = await fetch(`/demo/compare/${locale}/${name}`);
        if (!response.ok) throw new Error(t("errors.demo"));
        return response;
      };
      // The fixed samples are generated using the production comparison adapter.
      // Viewing them does not need to download or initialize the WASM runtime.
      const [original, modified, redline, preview] = await Promise.all([
        load("before.docx").then(response => response.arrayBuffer()),
        load("after.docx").then(response => response.arrayBuffer()),
        load("redline.docx").then(response => response.arrayBuffer()),
        load("result.json").then(response => response.json()) as Promise<DocxComparison>,
      ]);
      if (requestId.current !== id) return;
      const originalFile = new File([original], t("demoBefore"));
      const modifiedFile = new File([modified], t("demoAfter"));
      const result = { html: preview.html, revisions: preview.revisions, redline: new Uint8Array(redline) };
      sampleCache.current = { original: originalFile, modified: modifiedFile, result };
      setBefore(originalFile);
      setAfter(modifiedFile);
      setResult(result);
      setFilter("all"); setQuery(""); setActive(null);
      setElapsed(performance.now() - start);
      setTimeout(scrollToPageBottom, 50);
    } catch (failure) {
      if (requestId.current === id) setError(failure instanceof Error ? failure.message : t("errors.demo"));
    } finally {
      if (requestId.current === id) setDemoBusy(false);
    }
  }

  const revisions = useMemo(() => (result?.revisions ?? []).filter(revision =>
    !(revision.revisionType === "Moved" && revision.isMoveSource === true && revision.moveGroupId !== undefined)
  ), [result]);
  const visible = revisions.map((revision, index) => ({ revision, index })).filter(({ revision }) =>
    (filter === "all" || revision.revisionType === filter) &&
    `${revision.text} ${changeLabel(revision, t)} ${revision.formatChange?.changedPropertyNames?.join(" ") ?? ""}`.toLowerCase().includes(query.toLowerCase())
  );

  function locate(revision: Revision, index: number) {
    setActive(index);
    const doc = frame.current?.contentDocument;
    if (!doc) return;
    doc.querySelectorAll("[data-compare-focus]").forEach(node => node.removeAttribute("data-compare-focus"));
    const text = revision.text.replace(/\s/g, "");
    const candidates = Array.from(doc.querySelectorAll("ins, del, .rev-format-change, p, td, tr, table, section"));
    const target = candidates.find(node => text && node.textContent?.replace(/\s/g, "").includes(text)) ??
      (revision.revisionType === "FormatChanged" ? doc.querySelector(".rev-format-change") : null);
    if (target) {
      target.setAttribute("data-compare-focus", "true");
      target.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }

  function download() {
    if (!result?.redline) return;
    const url = URL.createObjectURL(new Blob([new Uint8Array(result.redline).buffer], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }));
    const anchor = document.createElement("a"); anchor.href = url; anchor.download = t("exportName"); anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function openEditor() {
    if (!after || openingEditor) return;
    // Open synchronously so browsers retain the user gesture across IndexedDB IO.
    const tab = window.open("about:blank", "_blank");
    if (!tab) { setError(t("errors.editor")); return; }
    tab.opener = null;
    setOpeningEditor(true);
    try {
      const id = await storeEditorDocument(after);
      tab.location.replace(`/${locale}/tools/editor?document=${encodeURIComponent(id)}`);
    } catch {
      tab.close();
      setError(t("errors.editor"));
    } finally { setOpeningEditor(false); }
  }

  return (
    <div className={styles.page}>
      <main>
        <section className={hero.layout}>
          <div className={hero.copy}>
            <h1>{t("title")}</h1>
            <p>{t("intro")}</p>
            <button className={hero.demo} onClick={demo} disabled={busy || demoBusy}>{demoBusy ? <LoaderCircle size={16} className={styles.spin} /> : <i className="ri-file-word-line" aria-hidden="true" />}{t("demo")}<ArrowRight size={16} /></button>
          </div>
          <div className={hero.panel}>
            <div className={hero.panelHeading}><h2>{t("choose")}</h2><span>.docx</span></div>
            <FileSlot title={t("original")} file={before} disabled={busy || demoBusy} onFile={file => selectFile(file, "before")} />
            <div className={styles.swapRow}><span /><button aria-label={t("swap")} disabled={!before || !after || busy || demoBusy} onClick={() => { setBefore(after); setAfter(before); setResult(null); }}><ArrowLeftRight size={15} /></button><span /></div>
            <FileSlot title={t("modified")} file={after} disabled={busy || demoBusy} onFile={file => selectFile(file, "after")} />
            {busy ? (
              <div className={styles.progress} role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(progress)} aria-label={t("working")}>
                <div className={styles.progressMeta}>
                  <span className={styles.progressLabel}>{t("working")}</span>
                  <span className={styles.progressValue}>{(elapsed / 1000).toFixed(1)}s · {Math.round(progress)}%</span>
                </div>
                <div className={styles.progressTrack}>
                  <div className={styles.progressFill} style={{ width: `${progress}%` }} />
                </div>
              </div>
            ) : (
              <button className={styles.primary} disabled={!before || !after || demoBusy} onClick={() => compare()}>
                {demoBusy ? <><LoaderCircle size={18} className={styles.spin} />{t("working")}</> : <>{t("start")}<ArrowRight size={18} /></>}
              </button>
            )}
            <p className={hero.privacy}><LockKeyhole size={14} />{t("privacy")}</p>
          </div>
        </section>
        {error && <div role="alert" className={styles.error}><span>{error}</span><button onClick={() => setError("")} aria-label={t("closeError")}><X size={18} /></button></div>}
        {result ? <section className={styles.results}>
          <div className={styles.resultHeading}><h2><i className="ri-check-double-line" aria-hidden="true" />{t("review")}</h2>
            <div className={styles.resultActions}><button data-tip={t("retry")} {...tip.triggerProps} aria-label={t("retry")} onClick={() => compare()} disabled={busy}><RotateCcw size={15} /><span className={styles.actionLabel}>{t("retry")}</span></button><button data-tip={t("download")} {...tip.triggerProps} aria-label={t("download")} onClick={download} disabled={!result.redline}><ArrowDownToLine size={15} /><span className={styles.actionLabel}>{t("download")}</span></button></div>
          </div>
          <div className={styles.reviewGrid}>
            <div className={styles.previewPane}><div className={styles.previewBar}>
              <div className={styles.previewIdentity}>
                <span className={styles.previewTitle}><i className="ri-file-word-line" aria-hidden="true" /></span>
                <span className={styles.comparisonNames}>
                  <span title={before?.name}>{before?.name}</span>
                  <ArrowRight size={13} aria-hidden="true" />
                  <span title={after?.name}>{after?.name}</span>
                </span>
              </div>
              <div className={styles.previewTools}>
                <div className={styles.previewLegend}><span className={styles.insertKey}>{t("kinds.Inserted")}</span><span className={styles.deleteKey}>{t("kinds.Deleted")}</span><span className={styles.formatKey}>{t("kinds.FormatChanged")}</span></div>
                <div className={styles.zoomControls}>
                  <button data-tip={t("zoomOut")} {...tip.triggerProps} aria-label={t("zoomOut")} disabled={zoom <= 25} onClick={() => changeZoom(Math.max(25, zoom - 10))}><Minus size={14} /></button>
                  <button data-tip={t("resetZoom")} {...tip.triggerProps} aria-label={t("resetZoom")} className={styles.zoomValue} onClick={() => changeZoom(100)}>{zoom}%</button>
                  <button data-tip={t("zoomIn")} {...tip.triggerProps} aria-label={t("zoomIn")} disabled={zoom >= 200} onClick={() => changeZoom(Math.min(200, zoom + 10))}><Plus size={14} /></button>
                </div>
                <span className={styles.toolSeparator} aria-hidden="true" />
                <button className={styles.editButton} data-tip={editDocumentLabel} {...tip.triggerProps} aria-label={editDocumentLabel} disabled={!after || openingEditor} onClick={openEditor}>
                  {openingEditor ? <LoaderCircle size={14} className={styles.spin} /> : <PencilLine size={14} />}
                  <span>{t("editDocument")}</span>
                </button>
              </div>
            </div>
              <iframe ref={frame} title={t("preview")} sandbox="allow-same-origin" srcDoc={result.html}
                onLoad={event => { void preparePreview(event.currentTarget); }} />
            </div>
            <aside className={styles.changesPane}>
              <div className={styles.changesHeading}><h3>{t("changes")}</h3><span>{t("count", { count: revisions.length })}</span></div>
              <div className={styles.stats}>{kinds.slice(1).map(kind => <button key={kind} className={styles[kind]} data-active={filter === kind || undefined} aria-pressed={filter === kind} onClick={() => setFilter(filter === kind ? "all" : kind)}><strong>{revisions.filter(r => r.revisionType === kind).length}</strong>{t(`kinds.${kind}`)}</button>)}</div>
              <div className={styles.changeControls}>
                <label className="input input-sm w-full">
                  <Search size={15} className="opacity-50" />
                  <input type="search" aria-label={t("search")} placeholder={t("search")} value={query} onChange={event => setQuery(event.target.value)} />
                </label>
              </div>
              <div className={styles.changesList}>
                {visible.map(({ revision, index }) => <button key={index} className={`${styles.change} ${active === index ? styles.activeChange : ""}`} aria-current={active === index ? "true" : undefined} onClick={() => locate(revision, index)}>
                  <div className={styles.changeTop}><span className={styles[revision.revisionType] ?? ""}>{changeLabel(revision, t)}</span><span>#{String(index + 1).padStart(2, "0")}</span></div>
                  <p className={revision.revisionType === "Deleted" ? styles.deletedText : ""}>{revision.text.trim() || (revision.revisionType === "FormatChanged" ? t("structural") : t("nonText"))}</p>
                  {revision.formatChange?.changedPropertyNames?.map(property => <div key={property} className={styles.property}><span>{properties.includes(property === "font" ? "fontFamily" : property) ? t(`properties.${property === "font" ? "fontFamily" : property}`) : property}</span><div><del>{readable(revision.formatChange?.oldProperties?.[property], property, t)}</del><ArrowRight size={12} /><strong>{readable(revision.formatChange?.newProperties?.[property], property, t)}</strong></div></div>)}
                </button>)}
                {!visible.length && <div className={styles.empty}><Check size={24} /><strong>{revisions.length ? t("noMatch") : t("noChanges")}</strong><p>{revisions.length ? t("trySearch") : t("tryFiles")}</p></div>}
              </div>
            </aside>
          </div>
        </section> : <>
          <section className={styles.explainer}>
            <h2>{t("explainTitle")}</h2>
            <div><article><span className={styles.exampleInsert}>{t("exampleInsert")}</span><h3>{t("textTitle")}</h3><p>{t("textBody")}</p></article>
            <article><span className={styles.exampleFormat}>{t("exampleFormat")}</span><h3>{t("formatTitle")}</h3><p>{t("formatBody")}</p></article>
            <article><span className={styles.exampleReview}><Check size={14} />{t("exampleReview")}</span><h3>{t("exportTitle")}</h3><p>{t("exportBody")}</p></article></div>
          </section>
          <div className="mx-auto w-full max-w-[1180px] px-8"></div>
        </>}
      </main>
      {tip.tooltip}
    </div>
  );
}
