import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { useT } from "../i18n";
import { useHoverTooltip } from "./use-hover-tooltip";
import {
  ArrowDownToLine,
  ArrowRight,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  FileText,
  Scan,
  ImagePlus,
  LockKeyhole,
  RotateCw,
  ShieldCheck,
  Trash2,
  X,
} from "lucide-react";
import hero from "../styles/tool-hero.module.css";
import styles from "../styles/image-to-word.module.css";
import {
  PAPER_MM,
  PX_PER_MM,
  ImageDecodeError,
  computeFitSize,
  formatBytes,
  generateImageDocx,
  type ImageEncoding,
  type MarginMm,
  type Orientation,
  type PaperSize,
  type PerPage,
  type ResolvedOrientation,
} from "../lib/generate-docx";

const MAX_IMAGES = 50;
/** The app header is fixed and 64px tall; leave a little air under it. */
const HEADER_OFFSET = 80;
const ACCEPTED_TYPES = new Set([
  "image/jpeg", "image/png", "image/webp", "image/gif", "image/avif", "image/bmp", "image/x-ms-bmp",
]);

type Item = {
  id: string;
  file: File;
  url: string;
  rotation: number;
  /** Intrinsic pixels, filled in right after the object URL loads. */
  width: number;
  height: number;
};
type Result = { blob: Blob; pages: number; name: string };
/** `swap` = dropped onto another thumbnail; `before`/`after` = dropped in the gutter. */
type DropHint = { index: number; mode: "swap" | "before" | "after" };

/**
 * A 90°/270° turn swaps the image's bounding box, so the "contain" fit computed
 * for the unrotated image overflows the frame and gets clipped. Returns the
 * extra scale needed to keep the turned image inside `boxWidth × boxHeight`
 * (only the box's ratio matters, the unit is arbitrary).
 */
function rotationScale(
  sourceWidth: number,
  sourceHeight: number,
  rotation: number,
  boxWidth: number,
  boxHeight: number,
): number {
  if (!sourceWidth || !sourceHeight || rotation % 180 === 0) return 1;
  const base = Math.min(boxWidth / sourceWidth, boxHeight / sourceHeight);
  const turned = Math.min(boxWidth / sourceHeight, boxHeight / sourceWidth);
  return base > 0 && turned > 0 ? turned / base : 1;
}

function Choice<T extends string | number>({
  value,
  options,
  onChange,
  disabled,
}: {
  value: T;
  options: Array<{ value: T; label: string }>;
  onChange: (value: T) => void;
  disabled?: boolean;
}) {
  return (
    <div className={styles.segmented}>
      {options.map((option) => (
        <button
          key={option.value}
          type="button"
          disabled={disabled}
          data-active={option.value === value || undefined}
          onClick={() => onChange(option.value)}
        >
          {option.label}
        </button>
      ))}
    </div>
  );
}

export default function ImageToWord() {
  const tip = useHoverTooltip("top");
  const t = useT("imageToWord");
  const [images, setImages] = useState<Item[]>([]);
  const imagesRef = useRef<Item[]>([]);
  // Keep the ref in sync after commit, not during render: the paste handler and
  // the unmount cleanup both need the latest list without re-subscribing.
  useEffect(() => { imagesRef.current = images; }, [images]);

  const [paper, setPaper] = useState<PaperSize>("fit");
  const [orientation, setOrientation] = useState<Orientation>("auto");
  const [perPage, setPerPage] = useState<PerPage>(1);
  const [marginX, setMarginX] = useState<MarginMm>(20);
  const [marginY, setMarginY] = useState<MarginMm>(20);
  // Asymmetric margins are rare: one row by default, a second one on demand.
  const [splitMargins, setSplitMargins] = useState(false);
  const [quality, setQuality] = useState<"standard" | "original">("original");
  const [encoding, setEncoding] = useState<ImageEncoding>("auto");
  // Fill the slot by default: the preview has always scaled images up, so
  // generating at original size used to disagree with what the user saw.
  const [upscale, setUpscale] = useState<"fit" | "original">("fit");

  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [error, setError] = useState("");
  const [result, setResult] = useState<Result | null>(null);
  const [dragIndex, setDragIndex] = useState<number | null>(null);
  const [dropHint, setDropHintState] = useState<DropHint | null>(null);
  const [dropActive, setDropActive] = useState(false);
  const [previewPage, setPreviewPage] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const gridCardRef = useRef<HTMLDivElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const workspaceRef = useRef<HTMLElement>(null);

  // dragover fires continuously; only touch state when the resolved target
  // actually changes, otherwise every mousemove re-renders the whole grid.
  const hintRef = useRef<DropHint | null>(null);
  const setDropHint = useCallback((next: DropHint | null) => {
    const prev = hintRef.current;
    if (prev === next) return;
    if (prev && next && prev.index === next.index && prev.mode === next.mode) return;
    hintRef.current = next;
    setDropHintState(next);
  }, []);

  // "Fit to image" takes the first image's own size; rotating 90°/270° swaps
  // the axes. Scaled proportionally into what Word accepts as a page (5–50 cm)
  // — clamping each edge on its own would turn a 4:3 photo into a square.
  const fitSize = useMemo(() => {
    const first = images[0];
    return computeFitSize(first?.width ?? 0, first?.height ?? 0, first?.rotation ?? 0);
  }, [images]);

  const pageCount = Math.max(1, Math.ceil(images.length / perPage));
  // Clamp while rendering instead of in an effect: removing an image or raising
  // "images per page" can shrink the page count below the page being previewed.
  const currentPage = Math.min(previewPage, pageCount - 1);
  const pageImages = images.slice(currentPage * perPage, currentPage * perPage + perPage);
  const paperSize = paper === "fit" ? fitSize : PAPER_MM[paper];
  // Judge the source aspect, not the clamped page: a 4000×3000 photo must stay
  // portrait even when its page is capped at 500 mm.
  const firstSource = images[0];
  const sourceAspect = firstSource?.width && firstSource?.height
    ? (firstSource.rotation % 180 !== 0
      ? firstSource.height / firstSource.width
      : firstSource.width / firstSource.height)
    : fitSize.width / fitSize.height;
  const autoLandscape = sourceAspect > 1.02;
  // "Fit" always follows the image; on a fixed paper "auto" does the same and
  // an explicit portrait/landscape wins.
  const landscape = paper === "fit"
    ? autoLandscape
    : orientation === "auto"
      ? autoLandscape
      : orientation === "landscape";
  const resolvedOrientation: ResolvedOrientation = landscape ? "landscape" : "portrait";
  // A fitted page already carries the image's own orientation; swapping it by
  // `landscape` would stand a 16:9 photo on its side.
  const paperWidth = paper === "fit"
    ? fitSize.width
    : landscape ? paperSize.height : paperSize.width;
  const paperHeight = paper === "fit"
    ? fitSize.height
    : landscape ? paperSize.width : paperSize.height;
  // Keep the paper's aspect ratio but let it shrink so tall images (long
  // screenshots) cannot push the card off screen.
  const previewHeight = Math.min(200, (landscape ? 172 : 148) * (paperHeight / paperWidth));
  const previewWidth = (previewHeight * paperWidth) / paperHeight;
  const previewPxPerMm = previewWidth / paperWidth;
  // Ratio of one preview cell, used to keep turned images inside their slot.
  const cellWidthMm = (paperWidth - marginX * 2) / (perPage === 4 ? 2 : 1);
  const cellHeightMm = (paperHeight - marginY * 2) / (perPage === 4 ? 2 : perPage);
  const percent = progress.total ? Math.round((progress.done / progress.total) * 100) : 0;

  // Rough docx size: document scans land around 0.10–0.12 bytes per output
  // pixel at q0.9 (photos run higher, hence the “about”). Shows nothing until
  // the intrinsic sizes are known, so it can never lie wildly.
  const estimatedBytes = useMemo(() => {
    const contentWidthPx = Math.max(1, (paperWidth - marginX * 2) * PX_PER_MM);
    const contentHeightPx = Math.max(1, (paperHeight - marginY * 2) * PX_PER_MM);
    const columns = perPage === 4 ? 2 : 1;
    const rows = perPage === 4 ? 2 : perPage;
    const slotWidth = Math.max(1, (contentWidthPx - 10 * (columns - 1)) / columns);
    const slotHeight = Math.max(1, (contentHeightPx - 10 * (rows - 1)) / rows);
    let bytes = 0;
    for (const image of images) {
      if (!image.width || !image.height) continue;
      const swapped = image.rotation % 180 !== 0;
      const sourceWidth = swapped ? image.height : image.width;
      const sourceHeight = swapped ? image.width : image.height;
      const displayScale = Math.min(slotWidth / sourceWidth, slotHeight / sourceHeight, upscale === "fit" ? Infinity : 1);
      const displayWidth = sourceWidth * displayScale;
      const displayHeight = sourceHeight * displayScale;
      const outWidth = quality === "original"
        ? Math.min(sourceWidth, (displayWidth / 96) * 350)
        : displayWidth;
      const outHeight = quality === "original"
        ? Math.min(sourceHeight, (displayHeight / 96) * 350)
        : displayHeight;
      // Calibrated against real exports: JPEG q0.9 lands near 0.145 B/px,
      // PNG is several times larger. Still an estimate, hence "about".
      bytes += outWidth * outHeight * (encoding === "png" ? 0.62 : 0.145);
    }
    return bytes;
  }, [images, paperWidth, paperHeight, marginX, marginY, perPage, quality, upscale, encoding]);

  const downloadUrl = useMemo(
    () => (result ? URL.createObjectURL(result.blob) : null),
    [result],
  );
  useEffect(() => () => {
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  }, [downloadUrl]);

  // Generating happens inside the preview card, so it is usually already on
  // screen; only move the page when the download would land below the fold.
  useEffect(() => {
    if (!result) return;
    const card = previewRef.current;
    if (!card) return;
    const rect = card.getBoundingClientRect();
    if (rect.top >= 0 && rect.bottom <= window.innerHeight) return;
    card.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [result]);

  useEffect(() => () => {
    imagesRef.current.forEach((image) => URL.revokeObjectURL(image.url));
  }, []);

  const addFiles = useCallback((incoming: File[]) => {
    if (busy || incoming.length === 0) return;
    const accepted: Item[] = [];
    let rejected = false;
    let overflow = false;
    for (const file of incoming) {
      if (!ACCEPTED_TYPES.has(file.type)) {
        rejected = true;
        continue;
      }
      if (imagesRef.current.length + accepted.length >= MAX_IMAGES) {
        overflow = true;
        break;
      }
      accepted.push({
        id: crypto.randomUUID(),
        file,
        url: URL.createObjectURL(file),
        rotation: 0,
        width: 0,
        height: 0,
      });
    }
    if (accepted.length > 0) {
      setImages((prev) => [...prev, ...accepted]);
      setResult(null);
      // Read intrinsic size once: "fit to image" and the size estimate need it,
      // but the grid must render immediately, so this never blocks the upload.
      for (const item of accepted) {
        const probe = new Image();
        probe.onload = () => {
          const { naturalWidth: width, naturalHeight: height } = probe;
          if (!width || !height) return;
          setImages((prev) => prev.map((entry) => (
            entry.id === item.id ? { ...entry, width, height } : entry
          )));
        };
        probe.onerror = () => setError(t("errorDecode"));
        probe.src = item.url;
      }
    }
    if (overflow) setError(t("errorTooMany"));
    else if (rejected) setError(t("errorType"));
  }, [busy, t]);

  useEffect(() => {
    const onPaste = (event: ClipboardEvent) => {
      const files = Array.from(event.clipboardData?.files ?? []);
      if (files.length === 0) return;
      event.preventDefault();
      addFiles(files);
    };
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, [addFiles]);

  const removeAt = (index: number) => {
    if (busy) return;
    // Revoke outside the state updater: React may call an updater twice in
    // StrictMode, and side effects must not live in render-phase callbacks.
    const removed = imagesRef.current[index];
    setImages((prev) => prev.filter((_, itemIndex) => itemIndex !== index));
    if (removed) URL.revokeObjectURL(removed.url);
    setResult(null);
  };

  const move = (from: number, to: number) => {
    if (busy || to < 0 || to >= imagesRef.current.length || from === to) return;
    setImages((prev) => {
      const next = [...prev];
      const [item] = next.splice(from, 1);
      next.splice(to, 0, item);
      return next;
    });
    setResult(null);
  };

  /** Drop a thumbnail into a gap, where `gap` is an index in the pre-removal list. */
  const insertAt = (from: number, gap: number) => {
    const count = imagesRef.current.length;
    if (busy || gap < 0 || gap > count) return;
    if (gap === from || gap === from + 1) return;
    move(from, gap > from ? gap - 1 : gap);
  };

  const swap = (a: number, b: number) => {
    if (busy || a === b) return;
    setImages((prev) => {
      if (a >= prev.length || b >= prev.length) return prev;
      const next = [...prev];
      [next[a], next[b]] = [next[b], next[a]];
      return next;
    });
    setResult(null);
  };

  /** Which thumbnail is under the pointer, and does that mean swap or insert? */
  const resolveDrop = (clientX: number, clientY: number): DropHint | null => {
    const grid = gridRef.current;
    if (!grid) return null;
    const cells = Array.from(grid.querySelectorAll<HTMLElement>("[data-thumb-index]"));
    let bestIndex = -1;
    let bestRect: DOMRect | null = null;
    let bestDistance = Infinity;
    for (const cell of cells) {
      const rect = cell.getBoundingClientRect();
      const index = Number(cell.dataset.thumbIndex);
      const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0;
      const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0;
      const distance = Math.hypot(dx, dy);
      if (distance < bestDistance) {
        bestDistance = distance;
        bestIndex = index;
        bestRect = rect;
      }
    }
    if (!bestRect || bestIndex < 0) return null;

    const inside = bestDistance === 0;
    const edge = Math.min(bestRect.width * 0.2, 26);
    const x = clientX - bestRect.left;
    if (inside) {
      // Narrow strips at the sides insert; the middle of a thumbnail swaps.
      if (x < edge) return { index: bestIndex, mode: "before" };
      if (x > bestRect.width - edge) return { index: bestIndex, mode: "after" };
      return { index: bestIndex, mode: "swap" };
    }
    // Pointer in the gutter between two thumbnails: insert on the nearer side.
    return {
      index: bestIndex,
      mode: clientX < bestRect.left + bestRect.width / 2 ? "before" : "after",
    };
  };

  const rotate = (index: number) => {
    if (busy) return;
    setImages((prev) => prev.map((item, itemIndex) =>
      itemIndex === index ? { ...item, rotation: (item.rotation + 90) % 360 } : item,
    ));
    setResult(null);
  };

  const clearAll = () => {
    if (busy) return;
    imagesRef.current.forEach((image) => URL.revokeObjectURL(image.url));
    setImages([]);
    setResult(null);
    setError("");
    setPreviewPage(0);
  };

  const stepPage = (delta: number) => {
    setPreviewPage(Math.min(pageCount - 1, Math.max(0, currentPage + delta)));
  };

  /** Any parameter change invalidates the generated file, so the CTA returns. */
  function setting<T>(setter: (value: T) => void) {
    return (value: T) => {
      setter(value);
      setResult(null);
    };
  }

  /** Scroll from the upload panel down to the workspace, clearing the fixed header. */
  const startEditing = () => {
    // On narrow screens `.side` has order:-1, so the workspace section starts
    // with the preview card — scroll to the arrange grid instead, which is what
    // the button label promises.
    const target = gridCardRef.current ?? workspaceRef.current;
    if (!target) return;
    const top = window.scrollY + target.getBoundingClientRect().top - HEADER_OFFSET;
    window.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
  };

  async function generate() {
    if (busy || images.length === 0) return;
    setBusy(true);
    setError("");
    setResult(null);
    setProgress({ done: 0, total: images.length });
    try {
      const generated = await generateImageDocx(
        images.map((image) => ({ file: image.file, rotation: image.rotation })),
        {
          paper,
          fitSize,
          orientation: resolvedOrientation,
          perPage,
          marginXMm: marginX,
          marginYMm: marginY,
          quality,
          encoding,
          upscale: upscale === "fit",
          onProgress: (done, total) => setProgress({ done, total }),
        },
      );
      const date = new Date().toISOString().slice(0, 10);
      setResult({ blob: generated.blob, pages: generated.pages, name: `images-${date}.docx` });
    } catch (cause) {
      console.error("Image DOCX generation failed:", cause);
      setError(cause instanceof ImageDecodeError ? t("errorDecode") : t("errorGenerate"));
    } finally {
      setBusy(false);
    }
  }

  const onPreviewKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === "ArrowLeft") {
      event.preventDefault();
      stepPage(-1);
    } else if (event.key === "ArrowRight") {
      event.preventDefault();
      stepPage(1);
    }
  };

  const activeHint = (index: number) => (dropHint?.index === index ? dropHint.mode : null);

  return (
    <div className={styles.page}>
      {tip.tooltip}
      <main>
        <section className={hero.layout}>
          <div className={hero.copy}>
            <h1>{t("title")}</h1>
            <p>{t("intro")}</p>
          </div>

          <div className={hero.panel}>
            <div className={hero.panelHeading}>
              <h2>{t("choose")}</h2>
              <span>{images.length} / {MAX_IMAGES}</span>
            </div>
            <button
              type="button"
              className={`${styles.dropZone} ${dropActive ? styles.dropZoneActive : ""}`}
              disabled={busy}
              onClick={() => inputRef.current?.click()}
              onDragOver={(event) => { event.preventDefault(); setDropActive(true); }}
              onDragLeave={() => setDropActive(false)}
              onDrop={(event) => {
                event.preventDefault();
                setDropActive(false);
                addFiles(Array.from(event.dataTransfer.files));
              }}
            >
              <ImagePlus size={26} aria-hidden="true" />
              <strong>{t("drop")}</strong>
              <small>{t("limit")}</small>
            </button>
            <input
              ref={inputRef}
              type="file"
              accept="image/jpeg,image/png,image/webp,image/gif,image/avif,image/bmp,.avif,.bmp"
              multiple
              hidden
              onChange={(event) => {
                addFiles(Array.from(event.target.files ?? []));
                event.target.value = "";
              }}
            />

            {images.length > 0 && (
              <div className={styles.summary}>
                <span>{t("count", { count: images.length })}</span>
                <span>{t("pagesOnly", { pages: pageCount })}</span>
              </div>
            )}

            {/* First-touch CTA. It does not build the file — it drops the view
                to the editing workspace, where arranging and generating live. */}
            <button
              type="button"
              className={styles.primary}
              disabled={images.length === 0 || busy}
              onClick={startEditing}
            >
              {t("startEditing")}
              <ArrowRight size={18} aria-hidden="true" />
            </button>
            <p className={hero.privacy}><LockKeyhole size={14} aria-hidden="true" />{t("privacy")}</p>
          </div>
        </section>

        {error && (
          <div role="alert" className={styles.error}>
            <span>{error}</span>
            <button onClick={() => setError("")} aria-label={t("closeError")} data-tip={t("closeError")} {...tip.triggerProps}><X size={18} /></button>
          </div>
        )}

        {images.length > 0 && (
          <section className={styles.workspace} ref={workspaceRef}>
            <div className={styles.gridCard} ref={gridCardRef}>
              <div className={styles.gridHead}>
                <h2>{t("arrange")}</h2>
                <div>
                  <span className={styles.countText}>{t("count", { count: images.length })}</span>
                  <button className={styles.ghostButton} onClick={() => inputRef.current?.click()} disabled={busy}>
                    <ImagePlus size={13} aria-hidden="true" />{t("addMore")}
                  </button>
                  <button className={styles.ghostButton} onClick={clearAll} disabled={busy}>
                    <Trash2 size={13} aria-hidden="true" />{t("clear")}
                  </button>
                </div>
              </div>
              <div
                ref={gridRef}
                className={styles.grid}
                onDragOver={(event) => {
                  if (dragIndex === null) return;
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
                  const from = dragIndex;
                  setDragIndex(null);
                  setDropHint(null);
                  if (from === null || !hint) return;
                  if (hint.mode === "swap") swap(from, hint.index);
                  else insertAt(from, hint.mode === "before" ? hint.index : hint.index + 1);
                }}
              >
                {images.map((image, index) => {
                  const hint = activeHint(index);
                  return (
                    <div
                      key={image.id}
                      data-thumb-index={index}
                      className={[
                        styles.thumb,
                        dragIndex === index ? styles.thumbDragging : "",
                        hint === "swap" ? styles.thumbSwap : "",
                        hint === "before" ? styles.thumbInsertBefore : "",
                        hint === "after" ? styles.thumbInsertAfter : "",
                      ].filter(Boolean).join(" ")}
                      onClick={(event) => {
                        // Buttons inside the card keep their own action.
                        if ((event.target as HTMLElement).closest("button")) return;
                        setPreviewPage(Math.floor(index / perPage));
                        const card = previewRef.current;
                        if (!card) return;
                        const rect = card.getBoundingClientRect();
                        if (rect.top < 0 || rect.bottom > window.innerHeight) {
                          card.scrollIntoView({ behavior: "smooth", block: "center" });
                        }
                      }}
                      draggable={!busy}
                      onDragStart={(event) => {
                        setDragIndex(index);
                        event.dataTransfer.effectAllowed = "move";
                      }}
                      onDragEnd={() => { setDragIndex(null); setDropHint(null); }}
                    >
                      <img
                        className={styles.thumbMedia}
                        src={image.url}
                        alt=""
                        style={{
                          transform: `rotate(${image.rotation}deg) scale(${rotationScale(image.width, image.height, image.rotation, 210, 297)})`,
                        }}
                      />
                      <span className={styles.thumbIndex}>{index + 1}</span>
                      <button
                        className={styles.thumbRemove}
                        onClick={() => removeAt(index)}
                        aria-label={t("remove", { index: index + 1 })} data-tip={t("remove", { index: index + 1 })} {...tip.triggerProps}
                        disabled={busy}
                      >
                        <X size={12} />
                      </button>
                      <div className={styles.thumbBar}>
                        <button onClick={() => move(index, index - 1)} disabled={busy || index === 0} aria-label={t("moveUp")} data-tip={t("moveUp")} {...tip.triggerProps}>
                          <ChevronUp size={15} />
                        </button>
                        <button onClick={() => move(index, index + 1)} disabled={busy || index === images.length - 1} aria-label={t("moveDown")} data-tip={t("moveDown")} {...tip.triggerProps}>
                          <ChevronDown size={15} />
                        </button>
                        <button onClick={() => rotate(index)} disabled={busy} aria-label={t("rotate")} data-tip={t("rotate")} {...tip.triggerProps}>
                          <RotateCw size={15} />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
              <div className={styles.gridFooter}>
                <p className={styles.gridHint}>{t("reorderHint")}</p>
              </div>
            </div>

            <aside className={styles.side}>
              <div className={styles.previewCard} ref={previewRef}>
                <div className={styles.previewHead}>
                  <h2>{t("previewTitle")}</h2>
                  <span className={styles.previewCounter}>{t("pageCount", { current: currentPage + 1, total: pageCount })}</span>
                </div>
                <div
                  className={styles.previewNav}
                  role="group"
                  aria-label={t("previewTitle")}
                  tabIndex={0}
                  onKeyDown={onPreviewKeyDown}
                >
                  <button
                    type="button"
                    className={styles.previewArrow}
                    onClick={() => stepPage(-1)}
                    disabled={currentPage === 0}
                    aria-label={t("prevPage")} data-tip={t("prevPage")} {...tip.triggerProps}
                  >
                    <ChevronLeft size={16} />
                  </button>
                  <div className={styles.previewFrame}>
                    <div
                      className={styles.paper}
                      style={{
                        width: previewWidth,
                        aspectRatio: `${paperWidth} / ${paperHeight}`,
                        padding: `${marginY * previewPxPerMm}px ${marginX * previewPxPerMm}px`,
                      }}
                    >
                      <div
                        className={styles.paperGrid}
                        data-per-page={perPage}
                        style={{ "--paper-rows": perPage === 4 ? 2 : perPage } as CSSProperties}
                      >
                        {pageImages.map((image) => (
                          <div className={styles.paperCell} key={image.id}>
                            <img
                              src={image.url}
                              alt=""
                              style={{
                                transform: `rotate(${image.rotation}deg) scale(${rotationScale(image.width, image.height, image.rotation, cellWidthMm, cellHeightMm)})`,
                              }}
                            />
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                  <button
                    type="button"
                    className={styles.previewArrow}
                    onClick={() => stepPage(1)}
                    disabled={currentPage >= pageCount - 1}
                    aria-label={t("nextPage")} data-tip={t("nextPage")} {...tip.triggerProps}
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>

                {busy ? (
                  <div
                    className={styles.progress}
                    role="progressbar"
                    aria-valuemin={0}
                    aria-valuemax={100}
                    aria-valuenow={percent}
                  >
                    <div className={styles.progressMeta}>
                      <span className={styles.progressLabel}>{t("generating", { done: progress.done, total: progress.total })}</span>
                      <span className={styles.progressValue}>{percent}%</span>
                    </div>
                    <div className={styles.progressTrack}>
                      <div className={styles.progressFill} style={{ width: `${percent}%` }} />
                    </div>
                  </div>
                ) : result && downloadUrl ? (
                  <div className={styles.downloadBlock}>
                    {/* One line: name, then pages · size. The name ellipsises
                        first so the numbers never wrap. */}
                    <p className={styles.resultLine}>
                      <span className={styles.resultName}>{result.name}</span>
                      <span className={styles.resultMeta}>{t("resultMeta", { pages: result.pages, size: formatBytes(result.blob.size) })}</span>
                    </p>
                    <a className={styles.downloadButton} href={downloadUrl} download={result.name}>
                      <ArrowDownToLine size={16} aria-hidden="true" />{t("download")}
                    </a>
                  </div>
                ) : (
                  <>
                    <div className={styles.estimate}>
                      <span>{t("estimateLabel")}</span>
                      <strong>
                        {estimatedBytes > 0
                          ? t("estimateValue", { pages: pageCount, size: formatBytes(estimatedBytes) })
                          : t("pagesOnly", { pages: pageCount })}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className={styles.primary}
                      onClick={() => void generate()}
                      disabled={images.length === 0}
                    >
                      {t("generateWithPages", { pages: pageCount })}
                      <ArrowRight size={18} aria-hidden="true" />
                    </button>
                  </>
                )}
              </div>

              <div className={styles.settingsCard}>
                <h2>{t("settings")}</h2>
                <div className={styles.settingRow}>
                  <span>{t("paper")}</span>
                  <Choice
                    value={paper}
                    onChange={setting(setPaper)}
                    disabled={busy}
                    options={[
                      { value: "fit", label: t("paperFit") },
                      { value: "a4", label: "A4" },
                      { value: "letter", label: "Letter" },
                    ]}
                  />
                </div>
                <div className={styles.settingRow} title={paper === "fit" ? t("orientationLocked") : undefined}>
                  <span>{t("orientation")}</span>
                  <Choice
                    value={paper === "fit" ? "auto" : orientation}
                    onChange={setting(setOrientation)}
                    // A fitted page already derives its orientation from the image.
                    disabled={busy || paper === "fit"}
                    options={[
                      { value: "auto", label: t("orientationAuto") },
                      { value: "portrait", label: t("portrait") },
                      { value: "landscape", label: t("landscape") },
                    ]}
                  />
                </div>
                <div className={styles.settingRow}>
                  <span>{t("perPage")}</span>
                  <Choice
                    value={perPage}
                    onChange={setting(setPerPage)}
                    disabled={busy}
                    options={[1, 2, 4].map((value) => ({ value: value as PerPage, label: t("perPageCount", { count: value }) }))}
                  />
                </div>
                <div className={styles.settingRow}>
                  <span>{splitMargins ? t("marginX") : t("margin")}</span>
                  <Choice
                    value={marginX}
                    onChange={(value) => {
                      setMarginX(value);
                      // Until the user splits them, both axes move together.
                      if (!splitMargins) setMarginY(value);
                      setResult(null);
                    }}
                    disabled={busy}
                    options={[
                      { value: 0 as MarginMm, label: t("marginNone") },
                      { value: 10 as MarginMm, label: t("marginValue", { value: 1 }) },
                      { value: 20 as MarginMm, label: t("marginValue", { value: 2 }) },
                    ]}
                  />
                </div>
                {splitMargins && (
                  <div className={styles.settingRow}>
                    <span>{t("marginY")}</span>
                    <Choice
                      value={marginY}
                      onChange={setting(setMarginY)}
                      disabled={busy}
                      options={[
                        { value: 0 as MarginMm, label: t("marginNone") },
                        { value: 10 as MarginMm, label: t("marginValue", { value: 1 }) },
                        { value: 20 as MarginMm, label: t("marginValue", { value: 2 }) },
                      ]}
                    />
                  </div>
                )}
                <button
                  type="button"
                  className={styles.splitMargins}
                  aria-pressed={splitMargins}
                  onClick={() => setSplitMargins((value) => !value)}
                  disabled={busy}
                >
                  {splitMargins ? t("marginTogether") : t("marginSplit")}
                </button>
                <div className={styles.settingRow}>
                  <span>{t("quality")}</span>
                  <Choice
                    value={quality}
                    onChange={setting(setQuality)}
                    disabled={busy}
                    options={[
                      { value: "standard", label: t("qualityStandard") },
                      { value: "original", label: t("qualityOriginal") },
                    ]}
                  />
                </div>
                <div className={styles.settingRow}>
                  <span>{t("upscale")}</span>
                  <Choice
                    value={upscale}
                    onChange={setting(setUpscale)}
                    disabled={busy}
                    options={[
                      { value: "fit", label: t("upscaleFit") },
                      { value: "original", label: t("upscaleOriginal") },
                    ]}
                  />
                </div>
                <div className={styles.settingRow}>
                  <span>{t("encoding")}</span>
                  <Choice
                    value={encoding}
                    onChange={setting(setEncoding)}
                    disabled={busy}
                    options={[
                      { value: "auto", label: t("encodingAuto") },
                      { value: "jpg", label: t("encodingJpg") },
                      { value: "png", label: t("encodingPng") },
                    ]}
                  />
                </div>
              </div>
            </aside>
          </section>
        )}

        {/* Highlights are marketing and disappear once a result is ready. */}
        {!busy && !result && (
          <section className={styles.highlights}>
            <h2>{t("featureTitle")}</h2>
            <div>
              {[ShieldCheck, Scan, FileText].map((Icon, index) => (
                <article key={index}>
                  <span className={styles.featureIcon}><Icon size={22} aria-hidden="true" /></span>
                  <h3>{t(`feature${index + 1}Title`)}</h3>
                  <p>{t(`feature${index + 1}Body`)}</p>
                </article>
              ))}
            </div>
          </section>
        )}
        {/* The FAQ must stay in the DOM: its FAQPage JSON-LD is server-rendered,
            and structured data has to describe content that is actually visible. */}
      </main>
    </div>
  );
}
