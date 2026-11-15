import { useEffect, useRef, useState } from "react";
import { Check, LoaderCircle, RotateCw, X } from "lucide-react";
import type { PDFDocumentProxy } from "pdfjs-dist";
import { pageViewport, renderPage, type RenderTaskHandle } from "../lib/pdf-preview";
import { computeSignatureRect } from "../lib/signature-rect";
import { useHoverTooltip } from "./use-hover-tooltip";
import styles from "../styles/pdf-tools.module.css";

/** 必须与 pdf-tools.module.css 里 .thumbCanvas 的宽度一致（桌面 172px）。 */
const THUMB_WIDTH = 172;

interface Props {
  doc: PDFDocumentProxy;
  /** 0-based index in the source PDF. */
  pageIndex: number;
  /** Position on the page wall (1-based). */
  position: number;
  selected: boolean;
  signed: boolean;
  draggable: boolean;
  /** Per-page shortcut shown only in delete / rotate mode. */
  action: "delete" | "rotate" | null;
  actionLabel: string;
  /** Pending clockwise rotation, rendered into the thumbnail itself. */
  rotationDelta: number;
  /** Readable label for screen readers: "Page 3", not just "3". */
  pageLabel: string;
  onAction?: () => void;
  /** Live overlay so a placed signature shows up without re-rendering the PDF. */
  signaturePreview: { u: number; v: number; widthRatio: number; ratio: number; url: string } | null;
  onActivate: () => void;
  onDragStart?: (event: React.DragEvent) => void;
  onDragOver?: (event: React.DragEvent) => void;
  onDrop?: (event: React.DragEvent) => void;
}

export default function PageThumb({
  doc,
  pageIndex,
  position,
  selected,
  signed,
  draggable,
  action,
  actionLabel,
  rotationDelta,
  pageLabel,
  onAction,
  signaturePreview,
  onActivate,
  onDragStart,
  onDragOver,
  onDrop,
}: Props) {
  const tip = useHoverTooltip("top");
  const host = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const [status, setStatus] = useState<"idle" | "loading" | "ready" | "failed">("idle");
  const [pageSize, setPageSize] = useState({ width: 595, height: 842 });

  useEffect(() => {
    const element = host.current;
    if (!element) return;
    setStatus("idle");
    let cancelled = false;
    let task: RenderTaskHandle | null = null;
    const observer = new IntersectionObserver(
      (entries) => {
        if (!entries[0]?.isIntersecting) return;
        observer.disconnect();
        void (async () => {
          if (cancelled) return;
          setStatus("loading");
          try {
            const page = await doc.getPage(pageIndex + 1);
            const full = page.getViewport({ scale: 1 });
            // Signature coordinates stay in the unrotated media box, matching
            // the exporter's drawImage space.
            setPageSize({ width: full.width, height: full.height });
            // Fit the *rotated* viewport into the fixed thumbnail width. Using
            // the unrotated width for 90°/270° would clip the turned page.
            const rotated = pageViewport(page, 1, rotationDelta);
            const scale = THUMB_WIDTH / rotated.width;
            if (canvas.current && !cancelled) {
              await renderPage(
                doc, pageIndex + 1, canvas.current, scale,
                (started) => { task = started; }, rotationDelta,
              );
            }
            if (!cancelled) setStatus("ready");
          } catch {
            if (!cancelled) setStatus("failed");
          }
        })();
      },
      { rootMargin: "240px" },
    );
    observer.observe(element);
    return () => {
      cancelled = true;
      // Switching files while a render is in flight used to leave the task
      // writing to a canvas that a new render already claimed.
      try { task?.cancel(); } catch { /* already settled */ }
      observer.disconnect();
    };
  }, [doc, pageIndex, rotationDelta]);

  const turned = rotationDelta % 180 !== 0;
  // Use the exporter's geometry for the live overlay too, so the page wall,
  // the placement editor and the downloaded PDF all show the same position.
  const overlayStyle = signaturePreview ? (() => {
    const box = computeSignatureRect(
      pageSize.width, pageSize.height, signaturePreview, signaturePreview.ratio,
    );
    return {
      left: `${(box.x / pageSize.width) * 100}%`,
      top: `${((pageSize.height - box.y - box.height) / pageSize.height) * 100}%`,
      width: `${(box.width / pageSize.width) * 100}%`,
      height: `${(box.height / pageSize.height) * 100}%`,
    };
  })() : undefined;

  return (
    <div
      ref={host}
      data-page-index={pageIndex}
      className={`${styles.thumb} ${selected ? styles.thumbSelected : ""}`}
      draggable={draggable}
      onDragStart={onDragStart}
      onDragOver={onDragOver}
      onDrop={onDrop}
      role="button"
      tabIndex={0}
      aria-pressed={selected}
      aria-label={pageLabel}
      onClick={onActivate}
      onKeyDown={(event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          onActivate();
        }
      }}
    >
      <div className={styles.thumbHead}>
        <span>{position}</span>
        <span className={styles.thumbMarks}>
          {signed && <span className={styles.thumbSigned} aria-hidden="true">✎</span>}
          {rotationDelta !== 0 && <span className={styles.thumbRotation} aria-hidden="true">{rotationDelta}°</span>}
          {selected && action !== "delete" && <span className={styles.thumbCheck}><Check size={11} /></span>}
        </span>
      </div>
      <div className={styles.thumbCanvas}>
        <div className={styles.thumbPage}>
          <canvas ref={canvas} className={`${status === "ready" ? styles.canvasReady : styles.canvasHidden} ${turned ? styles.canvasTurned : ""}`} />
          {/* A turned canvas no longer shares the media-box coordinates the
              signature overlay uses; hide it rather than draw it in the wrong place. */}
          {signaturePreview && status === "ready" && !turned && (
            <img
              className={styles.signatureLayer}
              src={signaturePreview.url}
              alt=""
              style={overlayStyle}
            />
          )}
        </div>
        {status === "loading" && <LoaderCircle size={16} className={`${styles.spin} ${styles.thumbStatus}`} aria-hidden="true" />}
        {status === "failed" && <span className={`${styles.thumbFailed} ${styles.thumbStatus}`}>PDF</span>}
      </div>
      {action && (
        <button
          type="button"
          className={action === "delete" ? styles.thumbDelete : styles.thumbRotate}
          aria-label={actionLabel}
          data-tip={actionLabel}
          {...tip.triggerProps}
          aria-pressed={selected}
          onClick={(event) => {
            event.stopPropagation();
            onAction?.();
          }}
        >
          {action === "delete" ? <X size={13} /> : <RotateCw size={13} />}
        </button>
      )}
      {tip.tooltip}
    </div>
  );
}
