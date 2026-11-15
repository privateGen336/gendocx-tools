import { useEffect, useMemo, useRef, useState } from "react";
import { Check, RotateCcw, Undo2, X } from "lucide-react";
import type { PDFDocumentProxy } from "pdfjs-dist";
import { renderPage, type RenderTaskHandle } from "../lib/pdf-preview";
import { computeSignatureRect } from "../lib/signature-rect";
import { useHoverTooltip } from "./use-hover-tooltip";
import type { SignaturePlacement } from "../lib/pdf-ops";
import styles from "../styles/pdf-tools.module.css";

const PREVIEW_WIDTH = 520;
const INK_COLOR = "#1b2430";
const INK_WIDTH = 2.2;
const FALLBACK_INK_RATIO = 280 / 96;
const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));

function dataUrlToBytes(dataUrl: string): Uint8Array {
  const binary = atob(dataUrl.split(",")[1] ?? "");
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

type Point = { x: number; y: number };

interface Props {
  doc: PDFDocumentProxy;
  pageIndex: number;
  initial?: SignaturePlacement | null;
  onConfirm: (placement: SignaturePlacement) => void;
  onCancel: () => void;
  labels: {
    title: string;
    draw: string;
    hint: string;
    clear: string;
    undo: string;
    size: string;
    preview: string;
    cancel: string;
    confirm: string;
    empty: string;
    disclaimer: string;
  };
}

export default function SignatureEditor({
  doc,
  pageIndex,
  initial,
  onConfirm,
  onCancel,
  labels,
}: Props) {
  const tip = useHoverTooltip();
  const ink = useRef<HTMLCanvasElement>(null);
  const preview = useRef<HTMLCanvasElement>(null);
  const previewHost = useRef<HTMLDivElement>(null);
  const drawing = useRef(false);
  const drag = useRef<{ x: number; y: number; u: number; v: number } | null>(null);
  // Strokes are kept as points so "undo" can repaint without pixel snapshots.
  const strokes = useRef<Point[][]>([]);
  const current = useRef<Point[]>([]);
  // An existing signature is painted back onto the ink canvas, otherwise
  // re-opening the editor would export a fully transparent PNG over it.
  const baseImage = useRef<HTMLImageElement | null>(null);

  const [hasInk, setHasInk] = useState(Boolean(initial));
  const [baseReady, setBaseReady] = useState(!initial);
  const [strokeCount, setStrokeCount] = useState(0);
  const [inkImage, setInkImage] = useState<string | null>(null);
  const [inkRatio, setInkRatio] = useState(FALLBACK_INK_RATIO);
  // 0.6 keeps the placeholder inside the visible part of the page preview.
  const [u, setU] = useState(initial?.u ?? 0.6);
  const [v, setV] = useState(initial?.v ?? 0.6);
  const [widthRatio, setWidthRatio] = useState(initial?.widthRatio ?? 0.3);
  const [error, setError] = useState("");
  const [previewReady, setPreviewReady] = useState(false);
  const [pageSize, setPageSize] = useState({ width: 595, height: 842 });

  useEffect(() => {
    const canvas = ink.current;
    if (!canvas) return;
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    const rect = canvas.getBoundingClientRect();
    canvas.width = Math.round(rect.width * dpr);
    canvas.height = Math.round(rect.height * dpr);
    // The exporter derives the signature's aspect ratio from the actual PNG,
    // so the preview must use the same ratio (not a hard-coded 280/96).
    setInkRatio(canvas.width / canvas.height);
    const ctx = canvas.getContext("2d");
    if (ctx) {
      ctx.scale(dpr, dpr);
      ctx.lineWidth = INK_WIDTH;
      ctx.lineCap = "round";
      ctx.lineJoin = "round";
      ctx.strokeStyle = INK_COLOR;
      ctx.fillStyle = INK_COLOR;
    }
  }, []);

  useEffect(() => {
    // baseReady already starts as true when there is no signature to restore.
    if (!initial) return;
    const url = URL.createObjectURL(new Blob([new Uint8Array(initial.png)], { type: "image/png" }));
    const image = new Image();
    let cancelled = false;
    image.onload = () => {
      if (cancelled) return;
      baseImage.current = image;
      const canvas = ink.current;
      const ctx = canvas?.getContext("2d");
      if (canvas && ctx) {
        ctx.save();
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
        ctx.restore();
        setInkImage(canvas.toDataURL("image/png"));
      }
      setHasInk(true);
      setBaseReady(true);
    };
    image.onerror = () => {
      if (cancelled) return;
      // Cannot restore the strokes: force the user to draw again instead of
      // silently replacing the signature with a blank image.
      setHasInk(false);
      setBaseReady(true);
    };
    image.src = url;
    return () => {
      cancelled = true;
      baseImage.current = null;
      URL.revokeObjectURL(url);
    };
  }, [initial]);

  useEffect(() => {
    let cancelled = false;
    let task: RenderTaskHandle | null = null;
    void (async () => {
      const canvas = preview.current;
      if (!canvas) return;
      try {
        const page = await doc.getPage(pageIndex + 1);
        const full = page.getViewport({ scale: 1 });
        if (cancelled) return;
        setPageSize({ width: full.width, height: full.height });
        // StrictMode runs effects twice. Without cancelling the first task,
        // pdf.js rejects the second render on the same canvas.
        await renderPage(doc, pageIndex + 1, canvas, PREVIEW_WIDTH / full.width, (started) => { task = started; });
        canvas.style.width = "100%";
        canvas.style.height = "auto";
        if (!cancelled) setPreviewReady(true);
      } catch {
        if (!cancelled) setPreviewReady(false);
      }
    })();
    return () => {
      cancelled = true;
      try { task?.cancel(); } catch { /* already settled */ }
    };
  }, [doc, pageIndex]);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") onCancel();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onCancel]);

  // The same geometry the exporter uses. Dragging therefore cannot place the
  // box at a position that computeSignatureRect later clamps to somewhere else.
  const rect = useMemo(
    () => computeSignatureRect(pageSize.width, pageSize.height, { u, v, widthRatio }, inkRatio),
    [pageSize, u, v, widthRatio, inkRatio],
  );
  const boxStyle = {
    left: `${(rect.x / pageSize.width) * 100}%`,
    top: `${((pageSize.height - rect.y - rect.height) / pageSize.height) * 100}%`,
    width: `${(rect.width / pageSize.width) * 100}%`,
    height: `${(rect.height / pageSize.height) * 100}%`,
  };

  function inkPoint(event: React.PointerEvent<HTMLCanvasElement>) {
    const bounds = event.currentTarget.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  }

  function redraw() {
    const canvas = ink.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    if (baseImage.current) {
      ctx.drawImage(baseImage.current, 0, 0, canvas.width, canvas.height);
    }
    ctx.restore();
    ctx.lineWidth = INK_WIDTH;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.strokeStyle = INK_COLOR;
    ctx.fillStyle = INK_COLOR;
    for (const stroke of strokes.current) {
      if (stroke.length === 1) {
        ctx.beginPath();
        ctx.arc(stroke[0].x, stroke[0].y, INK_WIDTH / 2, 0, Math.PI * 2);
        ctx.fill();
        continue;
      }
      ctx.beginPath();
      ctx.moveTo(stroke[0].x, stroke[0].y);
      for (let index = 1; index < stroke.length; index += 1) ctx.lineTo(stroke[index].x, stroke[index].y);
      ctx.stroke();
    }
    setInkImage(canvas.toDataURL("image/png"));
  }

  function onInkDown(event: React.PointerEvent<HTMLCanvasElement>) {
    const ctx = event.currentTarget.getContext("2d");
    if (!ctx) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    drawing.current = true;
    const { x, y } = inkPoint(event);
    current.current = [{ x, y }];
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + 0.1, y + 0.1);
    ctx.stroke();
    setHasInk(true);
    setError("");
  }

  function onInkMove(event: React.PointerEvent<HTMLCanvasElement>) {
    if (!drawing.current) return;
    const ctx = event.currentTarget.getContext("2d");
    if (!ctx) return;
    const { x, y } = inkPoint(event);
    current.current.push({ x, y });
    ctx.lineTo(x, y);
    ctx.stroke();
  }

  function endStroke() {
    if (!drawing.current) return;
    drawing.current = false;
    if (current.current.length > 0) {
      strokes.current.push(current.current);
      current.current = [];
      redraw();
      setStrokeCount(strokes.current.length);
    }
  }

  function undoStroke() {
    if (strokes.current.length === 0) return;
    strokes.current.pop();
    redraw();
    setStrokeCount(strokes.current.length);
    setHasInk(strokes.current.length > 0 || !!baseImage.current);
  }

  function clearInk() {
    strokes.current = [];
    current.current = [];
    baseImage.current = null;
    setStrokeCount(0);
    redraw();
    setInkImage(null);
    setHasInk(false);
  }

  function onBoxDown(event: React.PointerEvent<HTMLDivElement>) {
    event.currentTarget.setPointerCapture(event.pointerId);
    drag.current = { x: event.clientX, y: event.clientY, u, v };
  }

  function onBoxMove(event: React.PointerEvent<HTMLDivElement>) {
    if (!drag.current || !previewHost.current) return;
    const bounds = previewHost.current.getBoundingClientRect();
    const maxU = Math.max(0, 1 - rect.width / pageSize.width);
    const maxV = Math.max(0, 1 - rect.height / pageSize.height);
    setU(clamp(drag.current.u + (event.clientX - drag.current.x) / bounds.width, 0, maxU));
    setV(clamp(drag.current.v + (event.clientY - drag.current.y) / bounds.height, 0, maxV));
  }

  function confirm() {
    if (!hasInk || !ink.current || !baseReady) {
      setError(labels.empty);
      return;
    }
    onConfirm({
      pageIndex,
      u,
      v,
      widthRatio,
      png: dataUrlToBytes(ink.current.toDataURL("image/png")),
    });
  }

  return (
    <div className={styles.overlay} role="dialog" aria-modal="true" aria-label={labels.title}>
      {tip.tooltip}
      <div className={styles.overlayPanel}>
        <div className={styles.overlayHead}>
          <h2>{labels.title}</h2>
          <div className={styles.headActions}>
            <button type="button" className={styles.ghostButton} onClick={onCancel}>{labels.cancel}</button>
            <button type="button" className={`${styles.primary} ${styles.headPrimary}`} onClick={confirm}>
              <Check size={16} />{labels.confirm}
            </button>
            <button type="button" className={styles.overlayClose} aria-label={labels.cancel} data-tip={labels.cancel} {...tip.triggerProps} onClick={onCancel}>
              <X size={18} />
            </button>
          </div>
        </div>

        <div className={styles.overlayBody}>
          <div className={styles.overlayInk}>
            <span className={styles.slotLabel}>{labels.draw}</span>
            <canvas
              ref={ink}
              className={styles.signPad}
              aria-label={labels.draw}
              onPointerDown={onInkDown}
              onPointerMove={onInkMove}
              onPointerUp={endStroke}
              onPointerCancel={endStroke}
            />
            <div className={styles.signRow}>
              <span>{labels.hint}</span>
              <button type="button" onClick={undoStroke} disabled={strokeCount === 0}>
                <Undo2 size={13} />{labels.undo}
              </button>
              <button type="button" onClick={clearInk} disabled={!hasInk}>
                <RotateCcw size={13} />{labels.clear}
              </button>
            </div>
            <div className={styles.field}>
              <label htmlFor="sign-size">{labels.size}</label>
              <input
                id="sign-size"
                type="range"
                min={0.12}
                max={0.6}
                step={0.01}
                value={widthRatio}
                onChange={(event) => setWidthRatio(Number(event.target.value))}
              />
            </div>
            <p className={styles.hint}>{labels.disclaimer}</p>
          </div>

          <div className={styles.overlayPreview}>
            <span className={styles.slotLabel}>{labels.preview}</span>
            <div ref={previewHost} className={styles.signPreview}>
              <canvas ref={preview} />
              {previewReady && (
                <div
                  className={styles.signBox}
                  style={boxStyle}
                  onPointerDown={onBoxDown}
                  onPointerMove={onBoxMove}
                  onPointerUp={() => { drag.current = null; }}
                  onPointerCancel={() => { drag.current = null; }}
                >
                  {inkImage
                    ? <img className={styles.signBoxInk} src={inkImage} alt="" />
                    : <span>✎</span>}
                </div>
              )}
            </div>
          </div>
        </div>

        {error && <p role="alert" className={styles.overlayError}>{error}</p>}
      </div>
    </div>
  );
}
