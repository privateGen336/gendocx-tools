import { Fragment, useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useHoverTooltip } from "./use-hover-tooltip";
import { useT } from "../i18n";
import {
  ArrowDown,
  ArrowDownToLine,
  ArrowRight,
  ArrowUp,
  Check,
  ChevronLeft,
  ChevronRight,
  Copy,
  FileSearch,
  FileText,
  ListOrdered,
  LoaderCircle,
  LockKeyhole,
  PenLine,
  Redo2,
  RotateCcw,
  Scissors,
  ShieldCheck,
  Trash2,
  Undo2,
  X,
  Zap,
  type LucideIcon,
} from "lucide-react";
import type { PDFDocumentProxy } from "pdfjs-dist";
import { downloadOutputs } from "../lib/download";
import { openPdf } from "../lib/pdf-preview";
import type { ActionId, PdfResult, SignaturePlacement } from "../lib/pdf-ops";
import PageThumb from "./PageThumb";
import SignatureEditor from "./SignatureEditor";
import hero from "../styles/tool-hero.module.css";
import styles from "../styles/pdf-tools.module.css";

const MODES: { id: ActionId; Icon: LucideIcon }[] = [
  { id: "merge", Icon: Copy },
  { id: "split", Icon: Scissors },
  { id: "extract", Icon: FileSearch },
  { id: "delete", Icon: Trash2 },
  { id: "rotate", Icon: RotateCcw },
  { id: "reorder", Icon: ListOrdered },
  { id: "flatten", Icon: LockKeyhole },
  { id: "sign", Icon: PenLine },
];

const SELECT_MODES = new Set<ActionId>(["extract", "delete", "rotate"]);
const ERROR_CODES = [
  "encrypted", "invalid", "range", "empty", "noForm", "tooLarge", "tooManyFiles", "tooManyPages",
];
/** Modes that rebuild the PDF via copyPages, so document-level data is lost. */
const REBUILT_MODES = new Set<ActionId>(["merge", "split", "extract", "delete", "rotate", "reorder", "sign"]);

interface Entry {
  id: string;
  name: string;
  bytes: Uint8Array;
  doc: PDFDocumentProxy;
  pages: number;
  hasForm: boolean;
}

/** Everything the page editor can undo: page order, split points, selection, signature. */
type Snapshot = {
  activeId: string;
  entryIds: string[];
  orders: Record<string, number[]>;
  selected: number[];
  cuts: number[];
  signature: SignaturePlacement | null;
};

type SignaturePreview = {
  pageIndex: number;
  u: number;
  v: number;
  widthRatio: number;
  /** Actual PNG aspect ratio, read from its IHDR header. */
  ratio: number;
  url: string;
};

/** PNG IHDR stores width/height as big-endian uint32 at offsets 16 / 20. */
function pngAspect(bytes: Uint8Array): number {
  if (bytes.byteLength < 24) return 280 / 96;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  return height > 0 ? width / height : 280 / 96;
}

export default function PdfToolbox() {
  const tip = useHoverTooltip("top");
  const t = useT("pdfTools");

  const [entries, setEntries] = useState<Entry[]>([]);
  const [activeId, setActiveId] = useState("");
  const [orders, setOrders] = useState<Record<string, number[]>>({});
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [cuts, setCuts] = useState<Set<number>>(new Set());
  const [mode, setMode] = useState<ActionId | null>(null);
  const [rotation, setRotation] = useState(90);
  const [signature, setSignature] = useState<SignaturePlacement | null>(null);
  const [signingPage, setSigningPage] = useState<number | null>(null);
  const [advanced, setAdvanced] = useState(false);
  const [advancedRanges, setAdvancedRanges] = useState("");
  const [draggingPosition, setDraggingPosition] = useState<number | null>(null);
  const [dropHint, setDropHintState] = useState<{ position: number; mode: "swap" | "before" | "after" } | null>(null);
  const [draggingFile, setDraggingFile] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [warnings, setWarnings] = useState<string[]>([]);
  const [result, setResult] = useState<PdfResult | null>(null);
  const [history, setHistory] = useState<{ past: Snapshot[]; future: Snapshot[] }>({ past: [], future: [] });

  const fileInput = useRef<HTMLInputElement>(null);
  const nextId = useRef(0);
  const resultRef = useRef<HTMLDivElement>(null);
  const wallRef = useRef<HTMLDivElement>(null);
  const workspaceRef = useRef<HTMLElement>(null);
  const hadFiles = useRef(false);
  const hintRef = useRef<{ position: number; mode: "swap" | "before" | "after" } | null>(null);

  // dragover fires continuously; only touch state when the target actually changes.
  const setDropHint = useCallback((next: { position: number; mode: "swap" | "before" | "after" } | null) => {
    const prev = hintRef.current;
    if (prev === next) return;
    if (prev && next && prev.position === next.position && prev.mode === next.mode) return;
    hintRef.current = next;
    setDropHintState(next);
  }, []);

  // Ctrl/Cmd+Z / Shift+Cmd+Z / Ctrl+Y, skipped while typing in the range input.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey)) return;
      const target = event.target as HTMLElement | null;
      if (target && /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName)) return;
      const key = event.key.toLowerCase();
      if (key === "z") {
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
      } else if (key === "y") {
        event.preventDefault();
        redo();
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  // The signature is baked into the PDF only on export, so the page card gets a
  // live overlay: object URL from the PNG bytes, position straight from the
  // placement fractions the editor already uses.
  const [signaturePreview, setSignaturePreview] = useState<SignaturePreview | null>(null);
  useEffect(() => {
    // Object URLs must be created and revoked in an effect, not during render
    // (StrictMode discards renders and would leak). The derived state is set
    // here because this effect owns that external resource's lifetime.
    const url = signature
      ? URL.createObjectURL(new Blob([new Uint8Array(signature.png)], { type: "image/png" }))
      : null;
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setSignaturePreview(url && signature ? {
      pageIndex: signature.pageIndex,
      u: signature.u,
      v: signature.v,
      widthRatio: signature.widthRatio,
      ratio: pngAspect(signature.png),
      url,
    } : null);
    return () => { if (url) URL.revokeObjectURL(url); };
  }, [signature]);

  // The hero now stays above the editor, so moving between the two needs a
  // stable offset: 80px = 64px fixed header + 16px breathing room, the same
  // rule the images-to-docx tool uses for its "start editing" jump.
  const scrollToEditor = useCallback(() => {
    const node = workspaceRef.current;
    if (!node) return;
    const top = node.getBoundingClientRect().top + window.scrollY - 80;
    window.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
  }, []);

  // First batch of files: drop the view to the editor. Without this the user
  // stays on the hero, sees no editing controls and thinks nothing happened.
  useEffect(() => {
    if (entries.length > 0 && !hadFiles.current) {
      hadFiles.current = true;
      scrollToEditor();
    } else if (entries.length === 0) {
      hadFiles.current = false;
      // Removing the last file drops the editor, so the page is shorter than the
      // current offset — go back to the landing block instead of being stranded
      // in the middle of the SEO copy.
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }, [entries.length, scrollToEditor]);

  // Same rule as the images-to-docx tool: follow the result only when it would
  // otherwise land off-screen, and centre it — no jump when it is already there.
  useEffect(() => {
    if (!result) return;
    const banner = resultRef.current;
    if (!banner) return;
    const rect = banner.getBoundingClientRect();
    if (rect.top >= 0 && rect.bottom <= window.innerHeight) return;
    banner.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [result]);

  const active = entries.find((entry) => entry.id === activeId) ?? entries[0];
  const order = active ? (orders[active.id] ?? Array.from({ length: active.pages }, (_, i) => i)) : [];
  const cutList = useMemo(() => [...cuts].sort((a, b) => a - b), [cuts]);

  function resetOutcome() {
    setError("");
    setWarnings([]);
    setResult(null);
  }

  function capture(): Snapshot {
    return {
      activeId,
      entryIds: entries.map((entry) => entry.id),
      orders: Object.fromEntries(Object.entries(orders).map(([key, value]) => [key, [...value]])),
      selected: [...selected],
      cuts: [...cuts],
      signature,
    };
  }

  function pushHistory() {
    const snapshot = capture();
    setHistory((current) => ({ past: [...current.past, snapshot].slice(-50), future: [] }));
  }

  function restore(snapshot: Snapshot) {
    setActiveId(snapshot.activeId);
    setOrders(snapshot.orders);
    setSelected(new Set(snapshot.selected));
    setCuts(new Set(snapshot.cuts));
    setSignature(snapshot.signature);
    // File order is only visual and safe to reapply; skip it when files changed.
    const known = new Set(entries.map((entry) => entry.id));
    if (snapshot.entryIds.length === entries.length && snapshot.entryIds.every((id) => known.has(id))) {
      const byId = new Map(entries.map((entry) => [entry.id, entry]));
      setEntries(snapshot.entryIds.map((id) => byId.get(id)!));
    }
    resetOutcome();
  }

  function undo() {
    const previous = history.past[history.past.length - 1];
    if (!previous) return;
    const current = capture();
    setHistory({ past: history.past.slice(0, -1), future: [current, ...history.future].slice(0, 50) });
    restore(previous);
  }

  function redo() {
    const next = history.future[0];
    if (!next) return;
    const current = capture();
    setHistory({ past: [...history.past, current].slice(-50), future: history.future.slice(1) });
    restore(next);
  }

  async function addFiles(list: FileList | File[]) {
    const pdfs = Array.from(list).filter(
      (file) => /\.pdf$/i.test(file.name) || file.type === "application/pdf",
    );
    if (!pdfs.length) {
      setError(t("errors.invalid"));
      return;
    }
    setBusy(true);
    setError("");
    try {
      const [{ inspect, checkUpload }, loaded] = await Promise.all([
        import("../lib/pdf-ops"),
        Promise.all(pdfs.map(async (file) => ({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) }))),
      ]);
      // Enforce the limits at upload time (against the whole open set), not at
      // export time — otherwise a single-file action fails because unrelated
      // files happen to be open.
      checkUpload([
        ...entries.map(({ name, bytes }) => ({ name, bytes })),
        ...loaded,
      ]);
      const added: Entry[] = [];
      const newOrders: Record<string, number[]> = {};
      for (const source of loaded) {
        const info = await inspect(source);
        const doc = await openPdf(source.bytes);
        const id = `pdf-${nextId.current++}`;
        added.push({ id, ...source, doc, ...info });
        newOrders[id] = Array.from({ length: info.pages }, (_, i) => i);
      }
      setEntries((current) => [...current, ...added]);
      setOrders((current) => ({ ...current, ...newOrders }));
      setHistory({ past: [], future: [] });
      if (!activeId) setActiveId(added[0].id);
      resetOutcome();
    } catch (cause) {
      const code = (cause as { code?: string })?.code ?? "";
      setError(ERROR_CODES.includes(code) ? t(`errors.${code}`) : t("errors.unknown"));
    } finally {
      setBusy(false);
    }
  }

  function removeEntry(id: string) {
    const entry = entries.find((item) => item.id === id);
    if (entry) void entry.doc.destroy();
    const next = entries.filter((item) => item.id !== id);
    setEntries(next);
    setOrders((current) => {
      const remaining = { ...current };
      delete remaining[id];
      return remaining;
    });
    setHistory({ past: [], future: [] });
    if (activeId === id) {
      setActiveId(next[0]?.id ?? "");
      setSelected(new Set());
      setCuts(new Set());
      setSignature(null);
    }
    resetOutcome();
  }

  function moveFile(index: number, delta: number) {
    const target = index + delta;
    if (target < 0 || target >= entries.length) return;
    pushHistory();
    const next = [...entries];
    [next[index], next[target]] = [next[target], next[index]];
    setEntries(next);
    resetOutcome();
  }

  function switchMode(next: ActionId) {
    setMode(next);
    setSelected(new Set());
    setCuts(new Set());
    setDraggingPosition(null);
    resetOutcome();
  }

  function switchFile(id: string) {
    setActiveId(id);
    setSelected(new Set());
    setCuts(new Set());
    setSignature(null);
    resetOutcome();
  }

  function toggleSelected(pageIndex: number) {
    pushHistory();
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(pageIndex)) next.delete(pageIndex);
      else next.add(pageIndex);
      return next;
    });
    resetOutcome();
  }

  function toggleCut(position: number) {
    pushHistory();
    setCuts((current) => {
      const next = new Set(current);
      if (next.has(position)) next.delete(position);
      else next.add(position);
      return next;
    });
    resetOutcome();
  }

  function applyAdvanced() {
    if (!active) return;
    void (async () => {
      const { parseRanges } = await import("../lib/pdf-ops");
      // Only record a snapshot when the input actually applies: a rejected
      // range used to push a no-op undo step and clear the redo stack.
      let parsed: number[];
      try {
        parsed = parseRanges(advancedRanges, order.length);
      } catch {
        setError(t("errors.range"));
        return;
      }
      pushHistory();
      setSelected(new Set(parsed));
      setError("");
    })();
  }

  function movePage(from: number, to: number) {
    if (from === to || !active) return;
    pushHistory();
    const next = [...order];
    const [moved] = next.splice(from, 1);
    next.splice(to, 0, moved);
    setOrders((current) => ({ ...current, [active.id]: next }));
    resetOutcome();
  }

  function swapPages(a: number, b: number) {
    if (a === b || !active) return;
    pushHistory();
    const next = [...order];
    [next[a], next[b]] = [next[b], next[a]];
    setOrders((current) => ({ ...current, [active.id]: next }));
    resetOutcome();
  }

  /** `gap` is a slot in the pre-removal order (0..length). */
  function insertPage(from: number, gap: number) {
    if (gap === from || gap === from + 1) return;
    movePage(from, gap > from ? gap - 1 : gap);
  }

  /** Middle of a card = swap, its edges or the gutter = insert. */
  const resolveDrop = (clientX: number, clientY: number) => {
    const wall = wallRef.current;
    if (!wall) return null;
    const cells = Array.from(wall.querySelectorAll<HTMLElement>("[data-wall-position]"));
    let best = -1;
    let bestRect: DOMRect | null = null;
    let bestDistance = Infinity;
    for (const cell of cells) {
      const rect = cell.getBoundingClientRect();
      const position = Number(cell.dataset.wallPosition);
      const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0;
      const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0;
      const distance = Math.hypot(dx, dy);
      if (distance < bestDistance) {
        bestDistance = distance;
        best = position;
        bestRect = rect;
      }
    }
    if (!bestRect || best < 0) return null;
    const inside = bestDistance === 0;
    const edge = Math.min(bestRect.width * 0.2, 26);
    const x = clientX - bestRect.left;
    if (inside) {
      if (x < edge) return { position: best, mode: "before" as const };
      if (x > bestRect.width - edge) return { position: best, mode: "after" as const };
      return { position: best, mode: "swap" as const };
    }
    return {
      position: best,
      mode: clientX < bestRect.left + bestRect.width / 2 ? ("before" as const) : ("after" as const),
    };
  };

  const canRun = (() => {
    if (!active || !mode || busy) return false;
    switch (mode) {
      case "merge": return entries.length >= 2;
      case "split": return cutList.length > 0;
      case "extract":
      case "delete":
      case "rotate": return selected.size > 0 && (mode !== "delete" || selected.size < order.length);
      case "reorder": return true;
      case "flatten": return active.hasForm;
      case "sign": return signature !== null;
    }
  })();

  const summary = (() => {
    if (!active || !mode) return t("summary.chooseMode");
    switch (mode) {
      case "merge": return t("summary.merge", { files: entries.length });
      case "split": return cutList.length ? t("summary.split", { parts: cutList.length + 1 }) : t("summary.splitEmpty");
      case "extract":
      case "delete":
      case "rotate": return selected.size ? t("summary.selected", { pages: selected.size }) : t("summary.selectPages");
      case "reorder": return t("summary.reorder");
      case "flatten": return t("summary.flatten");
      case "sign": return signature ? t("summary.sign", { page: order.indexOf(signature.pageIndex) + 1 }) : t("summary.signEmpty");
    }
  })();

  async function run() {
    if (!canRun || !active || !mode) return;
    setBusy(true);
    resetOutcome();
    try {
      const ops = await import("../lib/pdf-ops");
      // Single-file modes operate on the active document only. Passing every
      // open file made runAction's MAX_FILES guard reject unrelated actions
      // (e.g. extracting one page while 21 files are open).
      const sources = (mode === "merge" ? entries : [active]).map(({ name, bytes }) => ({ name, bytes }));
      const output = await ops.runAction(mode, sources, {
        order,
        selected: [...selected],
        cuts: cutList,
        rotation,
        sign: signature ?? undefined,
      });
      setResult(output);
      const notes: string[] = [];
      if (output.warning === "form") notes.push(t("warning.form"));
      if (output.rebuilt ?? REBUILT_MODES.has(mode)) notes.push(t("warning.rebuilt"));
      setWarnings(notes);
    } catch (cause) {
      const code = (cause as { code?: string })?.code ?? "";
      setError(ERROR_CODES.includes(code) ? t(`errors.${code}`) : t("errors.unknown"));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className={styles.page}>
      {tip.tooltip}
      <input
        hidden
        ref={fileInput}
        type="file"
        accept=".pdf,application/pdf"
        multiple
        aria-label={t("browse")}
        onChange={(event) => {
          if (event.target.files?.length) void addFiles(event.target.files);
          event.target.value = "";
        }}
      />

      <main className={entries.length > 0 ? styles.editorMain : undefined}>
        {/* Kept mounted after a file loads: this is the page the user starts
            from, and the docx-stitch tool behaves the same way — scrolling up
            always returns to it instead of losing the landing block. */}
        <section className={hero.layout}>
          <div className={hero.copy}>
            <h1>{t("title")}</h1>
            <p>{t("intro")}</p>
          </div>
          <div
            className={`${hero.panel} ${draggingFile ? styles.dragging : ""}`}
            onDragOver={(event) => { event.preventDefault(); setDraggingFile(true); }}
            onDragLeave={() => setDraggingFile(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDraggingFile(false);
              if (event.dataTransfer.files.length) void addFiles(event.dataTransfer.files);
            }}
          >
            <div className={hero.panelHeading}>
              <h2>{t("choose")}</h2>
              <span>{entries.length > 0 ? t("heroFiles", { files: entries.length }) : ".pdf"}</span>
            </div>
            <button type="button" className={styles.dropzone} onClick={() => fileInput.current?.click()} disabled={busy}>
              <span className={styles.fileIcon}>
                {busy ? <LoaderCircle size={24} className={styles.spin} /> : <i className="ri-file-pdf-2-line" aria-hidden="true" />}
              </span>
              <span className={styles.dropText}>
                <strong>{t("dropTitle")}</strong>
                <small>{t("dropHint")}</small>
              </span>
            </button>
            {entries.length > 0 && (
              <button type="button" className={styles.primary} onClick={scrollToEditor}>
                {t("startEditing")}<ArrowDown size={17} aria-hidden="true" />
              </button>
            )}
            <p className={hero.privacy}><LockKeyhole size={14} aria-hidden="true" />{t("privacy")}</p>
          </div>
        </section>

        {entries.length > 0 && (
          <section className={styles.workspaceIntro} ref={workspaceRef}>
            <div className={styles.workspaceTop}>
            <div className={styles.fileBar}>
              {entries.map((entry, index) => (
                <div key={entry.id} className={`${styles.fileChip} ${entry.id === active?.id ? styles.fileChipActive : ""}`}>
                  <button type="button" className={styles.fileChipName} onClick={() => switchFile(entry.id)} data-tip={entry.name} {...tip.triggerProps}>
                    <i className="ri-file-pdf-2-line" aria-hidden="true" /><span>{entry.name}</span><small>{t("pagesOnly", { pages: entry.pages })}</small>
                  </button>
                  {mode === "merge" && (
                    <span className={styles.chipActions}>
                      <button type="button" aria-label={t("moveUp")} data-tip={t("moveUp")} {...tip.triggerProps} disabled={index === 0} onClick={() => moveFile(index, -1)}><ArrowUp size={13} /></button>
                      <button type="button" aria-label={t("moveDown")} data-tip={t("moveDown")} {...tip.triggerProps} disabled={index === entries.length - 1} onClick={() => moveFile(index, 1)}><ArrowDown size={13} /></button>
                    </span>
                  )}
                  <button type="button" className={styles.chipRemove} aria-label={t("removeFile")} data-tip={t("removeFile")} {...tip.triggerProps} onClick={() => removeEntry(entry.id)}><X size={13} /></button>
                </div>
              ))}
              <button type="button" className={styles.fileChipAdd} onClick={() => fileInput.current?.click()}>+ {t("addMore")}</button>
            </div>

              <div className={styles.historyGroup}>
                <button type="button" className={styles.historyButton} onClick={undo} disabled={history.past.length === 0} aria-label={t("undo")} data-tip={t("undo")} {...tip.triggerProps}>
                  <Undo2 size={16} aria-hidden="true" />
                </button>
                <button type="button" className={styles.historyButton} onClick={redo} disabled={history.future.length === 0} aria-label={t("redo")} data-tip={t("redo")} {...tip.triggerProps}>
                  <Redo2 size={16} aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Mode and the action it produces share one row: pick on the left,
                run on the right. The button used to sit above the modes it
                depends on, which read backwards. */}
            <div className={styles.actionRow}>
              <div className={styles.modeBar}>
                {MODES.map(({ id, Icon }) => {
                  const unavailable = id === "flatten" && !!active && !active.hasForm;
                  return (
                    <button
                      key={id}
                      type="button"
                      className={`${styles.modeButton} ${mode === id ? styles.modeButtonActive : ""}`}
                      aria-disabled={unavailable || undefined}
                      data-tip={unavailable ? t("flattenUnavailable") : undefined} {...tip.triggerProps}
                      onClick={() => { if (!unavailable) switchMode(id); }}
                    >
                      <Icon size={14} aria-hidden="true" />{t(`modes.${id}`)}
                    </button>
                  );
                })}
              </div>
              <div className={styles.actionGroup}>
                {mode && <span className={styles.actionSummaryText}>{summary}</span>}
                <button type="button" className={`${styles.primary} ${styles.runButton}`} disabled={!canRun} onClick={() => void run()}>
                  {busy ? <><LoaderCircle size={17} className={styles.spin} />{t("working")}</> : <>{mode ? t(`run.${mode}`) : t("run.default")}<ArrowRight size={17} /></>}
                </button>
              </div>
            </div>

            <div className={styles.modeHints}>
              <span>{mode ? t(`modeHint.${mode}`) : t("summary.chooseMode")}</span>
                {mode === "rotate" && (
                  <span className={styles.segmented}>
                    {[90, 180, 270].map((value) => (
                      <button key={value} type="button" className={rotation === value ? styles.segmentActive : ""} onClick={() => { setRotation(value); resetOutcome(); }}>{value}°</button>
                    ))}
                  </span>
                )}
                {mode === "split" && (
                  <button type="button" className={styles.linkButton} onClick={() => { pushHistory(); setCuts(new Set(order.slice(0, -1).map((_, i) => i + 1))); resetOutcome(); }}>{t("splitEveryPage")}</button>
                )}
                {mode && SELECT_MODES.has(mode) && (
                  <button type="button" className={styles.linkButton} onClick={() => setAdvanced((value) => !value)}>{t("advanced")}</button>
                )}
              <small className={styles.hintsPrivacy}><LockKeyhole size={12} aria-hidden="true" />{t("privacy")}</small>
            </div>

            {advanced && mode && SELECT_MODES.has(mode) && (
              <div className={styles.advancedRow}>
                <input value={advancedRanges} onChange={(event) => setAdvancedRanges(event.target.value)} placeholder="1,3-5" aria-label={t("advancedLabel")} />
                <button type="button" className={styles.ghostButton} onClick={applyAdvanced}>{t("advancedApply")}</button>
                <span className={styles.hint}>{t("advancedHint")}</span>
              </div>
            )}

            <div
              ref={wallRef}
              className={styles.wall}
              role="list"
              aria-label={t("pageWall")}
              onDragEnd={() => { setDraggingPosition(null); setDropHint(null); }}
              onDragOver={(event) => {
                if (mode !== "reorder") return;
                event.preventDefault();
                event.dataTransfer.dropEffect = "move";
                setDropHint(resolveDrop(event.clientX, event.clientY));
              }}
              onDragLeave={(event) => {
                const next = event.relatedTarget as Node | null;
                if (!next || !event.currentTarget.contains(next)) setDropHint(null);
              }}
              onDrop={(event) => {
                event.preventDefault();
                const hint = hintRef.current;
                const from = draggingPosition;
                setDraggingPosition(null);
                setDropHint(null);
                if (mode !== "reorder" || from === null || !hint) return;
                if (hint.mode === "swap") swapPages(from, hint.position);
                else insertPage(from, hint.mode === "before" ? hint.position : hint.position + 1);
              }}
            >
              {active && order.map((pageIndex, position) => (
                <Fragment key={pageIndex}>
                  {mode === "split" && position > 0 && (
                    <button
                      type="button"
                      className={`${styles.splitGap} ${cuts.has(position) ? styles.splitGapActive : ""}`}
                      aria-label={t("splitPoint")} data-tip={t("splitPoint")} {...tip.triggerProps}
                      aria-pressed={cuts.has(position)}
                      onClick={() => toggleCut(position)}
                    >
                      <span /><i>{cuts.has(position) ? <Check size={12} /> : "+"}</i>
                    </button>
                  )}
                  <div
                    role="listitem"
                    data-wall-position={position}
                    className={[
                      styles.wallItem,
                      dropHint?.position === position && dropHint.mode === "swap" ? styles.wallSwap : "",
                      dropHint?.position === position && dropHint.mode === "before" ? styles.wallInsertBefore : "",
                      dropHint?.position === position && dropHint.mode === "after" ? styles.wallInsertAfter : "",
                    ].filter(Boolean).join(" ")}
                  >
                    <PageThumb
                      doc={active.doc}
                      pageIndex={pageIndex}
                      position={position + 1}
                      selected={selected.has(pageIndex)}
                      signed={signature?.pageIndex === pageIndex}
                      signaturePreview={signaturePreview && signaturePreview.pageIndex === pageIndex ? signaturePreview : null}
                      draggable={mode === "reorder"}
                      onActivate={() => {
                        if (!mode) return;
                        if (SELECT_MODES.has(mode)) toggleSelected(pageIndex);
                        else if (mode === "sign") setSigningPage(pageIndex);
                      }}
                      onDragStart={(event) => {
                        setDraggingPosition(position);
                        event.dataTransfer.effectAllowed = "move";
                        // Firefox does not start a drag unless dataTransfer has data.
                        event.dataTransfer.setData("text/plain", String(position));
                      }}
                      action={mode === "delete" ? "delete" : mode === "rotate" ? "rotate" : null}
                      actionLabel={mode === "delete" ? t("deletePage", { page: position + 1 }) : t("rotatePage", { page: position + 1 })}
                      rotationDelta={mode === "rotate" && selected.has(pageIndex) ? rotation : 0}
                      pageLabel={t("pageLabel", { page: position + 1 })}
                      onAction={() => toggleSelected(pageIndex)}
                    />
                    {mode === "reorder" && (
                      <div className={styles.reorderControls}>
                        <button type="button" disabled={position === 0} aria-label={t("movePageLeft")} data-tip={t("movePageLeft")} {...tip.triggerProps} onClick={() => movePage(position, position - 1)}><ChevronLeft size={14} /></button>
                        <button type="button" disabled={position === order.length - 1} aria-label={t("movePageRight")} data-tip={t("movePageRight")} {...tip.triggerProps} onClick={() => movePage(position, position + 1)}><ChevronRight size={14} /></button>
                      </div>
                    )}
                  </div>
                </Fragment>
              ))}
            </div>
          </section>
        )}

        {error && (
          <div role="alert" className={styles.error}>
            <span>{error}</span>
            <button type="button" aria-label={t("dismiss")} data-tip={t("dismiss")} {...tip.triggerProps} onClick={() => setError("")}><X size={18} /></button>
          </div>
        )}
        {!error && warnings.map((note) => <div key={note} className={styles.notice}>{note}</div>)}
        {result && !error && (
          <div className={styles.result} role="status" ref={resultRef}>
            <Check size={20} aria-hidden="true" />
            <div className={styles.resultText}>
              <strong>{t("resultTitle")}</strong>
              <small>{t("resultMeta", { files: result.files.length, pages: result.files.reduce((sum, file) => sum + file.pages, 0) })}</small>
            </div>
            <button type="button" onClick={() => downloadOutputs(result.files, `gendocx-${mode ?? "pdf"}.zip`)}>
              <ArrowDownToLine size={16} />{t(result.files.length === 1 ? "download" : "downloadZip")}
            </button>
          </div>
        )}

        {/* Marketing only belongs on the landing state. Once a file is loaded
            the page is the editor, and these three cards would crowd the result. */}
        {entries.length === 0 && (
          <section className={styles.features}>
            {[ShieldCheck, Zap, FileText].map((Icon, index) => (
              <article key={index}>
                <Icon size={22} aria-hidden="true" />
                <h2>{t(`feature${index + 1}`)}</h2>
                <p>{t(`body${index + 1}`)}</p>
              </article>
            ))}
          </section>
        )}

        {/* Kept mounted in the editor state: the page's FAQPage JSON-LD is
            server-rendered, so the visible FAQ must exist in the DOM too. */}
      </main>

      {signingPage !== null && active && (
        <SignatureEditor
          key={signingPage}
          doc={active.doc}
          pageIndex={signingPage}
          initial={signature?.pageIndex === signingPage ? signature : null}
          onCancel={() => setSigningPage(null)}
          onConfirm={(placement) => {
            pushHistory();
            setSignature(placement);
            setSigningPage(null);
            resetOutcome();
          }}
          labels={{
            title: t("signEditor.title"),
            draw: t("signEditor.draw"),
            hint: t("signEditor.hint"),
            clear: t("signEditor.clear"),
            undo: t("signEditor.undo"),
            size: t("signEditor.size"),
            preview: t("signEditor.preview", { page: order.indexOf(signingPage) + 1 }),
            cancel: t("signEditor.cancel"),
            confirm: t("signEditor.confirm"),
            empty: t("signEditor.empty"),
            disclaimer: t("signEditor.disclaimer"),
          }}
        />
      )}
    </div>
  );
}
