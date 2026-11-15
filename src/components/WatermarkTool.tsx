import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useT } from "../i18n";
import {
  ArrowDownToLine,
  Check,
  Eraser,
  FileCheck2,
  LoaderCircle,
  LockKeyhole,
  Plus,
  RotateCcw,
  ScanSearch,
  ShieldCheck,
  X,
} from "lucide-react";
import hero from "../styles/tool-hero.module.css";
import { addDocxWatermark, DocxWatermarkError, removeDocxWatermarks, scanDocx } from "../lib/watermark/docx-watermark";
import { PdfWatermarkError, removePdfWatermarks, scanPdf } from "../lib/watermark/pdf-watermark";
import { addPdfWatermark } from "../lib/watermark/pdf-watermark-image";
import {
  DEFAULT_WATERMARK_OPTIONS,
  VARIABLE_TOKENS,
  expandWatermarkVariables,
  isValidPageRange,
  type HorizontalPosition,
  type VerticalPosition,
  type WatermarkImage,
  type WatermarkKind,
  type WatermarkLayer,
  type WatermarkPosition,
} from "../lib/watermark/watermark-options";
import { renderPagePreview, type PagePreview } from "../lib/watermark/page-preview";
import styles from "../styles/watermark.module.css";

type Format = "docx" | "pdf";
type Mode = "add" | "remove";
type Stage = "idle" | "scanning" | "ready" | "working" | "done";
type MarkTone = "solid" | "outline";

/** 水印颜色：前四个是常用色，最后一个走系统取色器。 */
const MARK_COLORS = ["#7a7a7a", "#2563eb", "#dc2626", "#16a34a", "#111827"];
/** 每页枚数：1 居中，2 上下两枚，其余按 n×n 平铺。 */
const MARK_COUNTS = [1, 2, 4, 9, 16];
/** 粗细四档：实心映射到字重，描边映射到线宽。 */
const MARK_WEIGHTS = [300, 500, 700, 900];
const PRESET_KEY = "gendocx-watermark-preset-v1";
const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

/** 只把可调参数存进 localStorage，图片二进制太大，不做持久化。 */
type MarkPreset = {
  text: string;
  opacity: number;
  angle: number;
  color: string;
  tone: MarkTone;
  weight: number;
  count: number;
  scale: number;
  gap: number;
  position: WatermarkPosition;
  layer: WatermarkLayer;
  pages: string;
};

const DEFAULT_PRESET: MarkPreset = {
  text: "",
  opacity: DEFAULT_WATERMARK_OPTIONS.opacity,
  angle: DEFAULT_WATERMARK_OPTIONS.angle,
  color: DEFAULT_WATERMARK_OPTIONS.color,
  tone: DEFAULT_WATERMARK_OPTIONS.tone,
  weight: DEFAULT_WATERMARK_OPTIONS.weight,
  count: DEFAULT_WATERMARK_OPTIONS.count,
  scale: DEFAULT_WATERMARK_OPTIONS.scale,
  gap: DEFAULT_WATERMARK_OPTIONS.gap,
  position: DEFAULT_WATERMARK_OPTIONS.position,
  layer: DEFAULT_WATERMARK_OPTIONS.layer,
  pages: DEFAULT_WATERMARK_OPTIONS.pages,
};

function readPreset(): Partial<MarkPreset> {
  if (typeof window === "undefined") return {};
  try {
    const raw = window.localStorage.getItem(PRESET_KEY);
    return raw ? (JSON.parse(raw) as Partial<MarkPreset>) : {};
  } catch {
    return {};
  }
}

type LoadedImage = WatermarkImage & { dataUrl: string };

const H_POSITIONS: HorizontalPosition[] = ["left", "center", "right"];
const V_POSITIONS: VerticalPosition[] = ["top", "middle", "bottom"];
const POSITION_CELLS = V_POSITIONS.flatMap((vertical) =>
  H_POSITIONS.map((horizontal) => ({ horizontal, vertical })),
);

type Item = {
  id: string;
  kind: "text" | "image" | "floating" | "form" | "annot";
  label: string;
  meta: string;
  /** 手动输入的条目在重新检测后不存在，界面上标记一下。 */
  custom?: boolean;
};

type Result = {
  bytes: Uint8Array;
  name: string;
  mode: Mode;
  removed?: number;
  pages?: number;
  /** 添加时每页的水印枚数，结果行据此说明是单枚还是平铺。 */
  count?: number;
};

const MB = 1024 * 1024;
const MAX_MARK_CHARS = 40;

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < MB) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / MB).toFixed(1)} MB`;
}

function extensionOf(name: string): Format | null {
  if (/\.docx$/i.test(name)) return "docx";
  if (/\.pdf$/i.test(name)) return "pdf";
  return null;
}

function baseName(name: string): string {
  return name.replace(/\.[^.]+$/, "");
}

function saveFile(bytes: Uint8Array, name: string, format: Format) {
  const blob = new Blob([bytes as unknown as BlobPart], {
    type:
      format === "pdf"
        ? "application/pdf"
        : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  // 立刻释放会让部分浏览器取消下载，给一个短延迟。
  window.setTimeout(() => URL.revokeObjectURL(url), 4000);
}

export default function WatermarkTool() {
  const t = useT("watermark");

  const [mode, setMode] = useState<Mode>("remove");
  const [stage, setStage] = useState<Stage>("idle");
  const [format, setFormat] = useState<Format | null>(null);
  const [fileName, setFileName] = useState("");
  const [fileSize, setFileSize] = useState(0);
  const [pageCount, setPageCount] = useState(0);
  const [findings, setFindings] = useState<Item[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [custom, setCustom] = useState("");
  const [manualNote, setManualNote] = useState("");
  const [error, setError] = useState("");
  const [dropping, setDropping] = useState(false);
  const [result, setResult] = useState<Result | null>(null);
  const [markText, setMarkText] = useState(DEFAULT_PRESET.text);
  const [markOpacity, setMarkOpacity] = useState(DEFAULT_PRESET.opacity);
  const [markAngle, setMarkAngle] = useState(DEFAULT_PRESET.angle);
  const [markColor, setMarkColor] = useState(DEFAULT_PRESET.color);
  const [markTone, setMarkTone] = useState<MarkTone>(DEFAULT_PRESET.tone);
  const [markCount, setMarkCount] = useState(DEFAULT_PRESET.count);
  const [markWeight, setMarkWeight] = useState(DEFAULT_PRESET.weight);
  const [markScale, setMarkScale] = useState(DEFAULT_PRESET.scale);
  const [markGap, setMarkGap] = useState(DEFAULT_PRESET.gap);
  const [markPosition, setMarkPosition] = useState<WatermarkPosition>(DEFAULT_PRESET.position);
  const [markLayer, setMarkLayer] = useState<WatermarkLayer>(DEFAULT_PRESET.layer);
  const [markPages, setMarkPages] = useState(DEFAULT_PRESET.pages);
  const [markKind, setMarkKind] = useState<WatermarkKind>("text");
  const [markImage, setMarkImage] = useState<LoadedImage | null>(null);
  const [pagePreview, setPagePreview] = useState<PagePreview | null>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);

  const inputRef = useRef<HTMLInputElement>(null);
  const workspaceRef = useRef<HTMLElement>(null);
  const mainRef = useRef<HTMLElement>(null);
  const sourceRef = useRef<Uint8Array | null>(null);
  /** 连续换文件时，只让最后一次预览的结果落地。 */
  const previewTokenRef = useRef(0);

  // 参数自动存在本机，下次打开工具直接接着上次的设置；SSR 首帧仍用默认值，避免 hydration 不一致。
  /* eslint-disable react-hooks/set-state-in-effect -- 预设必须在挂载后从 localStorage 读取，SSR 首帧要用默认值 */
  useEffect(() => {
    const preset = readPreset();
    if (typeof preset.text === "string") setMarkText(preset.text);
    if (typeof preset.opacity === "number") setMarkOpacity(preset.opacity);
    if (typeof preset.angle === "number") setMarkAngle(preset.angle);
    if (typeof preset.color === "string") setMarkColor(preset.color);
    if (preset.tone === "solid" || preset.tone === "outline") setMarkTone(preset.tone);
    if (typeof preset.weight === "number") setMarkWeight(preset.weight);
    if (typeof preset.count === "number") setMarkCount(preset.count);
    if (typeof preset.scale === "number") setMarkScale(preset.scale);
    if (typeof preset.gap === "number") setMarkGap(preset.gap);
    if (preset.position) setMarkPosition(preset.position);
    if (preset.layer === "above" || preset.layer === "below") setMarkLayer(preset.layer);
    if (typeof preset.pages === "string") setMarkPages(preset.pages);
  }, []);
  /* eslint-enable react-hooks/set-state-in-effect */

  useEffect(() => {
    if (typeof window === "undefined") return;
    const preset: MarkPreset = {
      text: markText,
      opacity: markOpacity,
      angle: markAngle,
      color: markColor,
      tone: markTone,
      weight: markWeight,
      count: markCount,
      scale: markScale,
      gap: markGap,
      position: markPosition,
      layer: markLayer,
      pages: markPages,
    };
    try {
      window.localStorage.setItem(PRESET_KEY, JSON.stringify(preset));
    } catch {
      // 隐私模式或配额满时静默跳过，不影响加水印。
    }
  }, [markText, markOpacity, markAngle, markColor, markTone, markWeight, markCount, markScale, markGap, markPosition, markLayer, markPages]);

  const resetSettings = useCallback(() => {
    setMarkText(DEFAULT_PRESET.text);
    setMarkOpacity(DEFAULT_PRESET.opacity);
    setMarkAngle(DEFAULT_PRESET.angle);
    setMarkColor(DEFAULT_PRESET.color);
    setMarkTone(DEFAULT_PRESET.tone);
    setMarkWeight(DEFAULT_PRESET.weight);
    setMarkCount(DEFAULT_PRESET.count);
    setMarkScale(DEFAULT_PRESET.scale);
    setMarkGap(DEFAULT_PRESET.gap);
    setMarkPosition(DEFAULT_PRESET.position);
    setMarkLayer(DEFAULT_PRESET.layer);
    setMarkPages(DEFAULT_PRESET.pages);
    setMarkKind("text");
    setMarkImage(null);
    setResult(null);
    setError("");
    if (imageInputRef.current) imageInputRef.current.value = "";
    try {
      window.localStorage.removeItem(PRESET_KEY);
    } catch {
      // 忽略
    }
  }, []);

  /** 图片水印：读成字节 + dataURL，再用浏览器解码出真实宽高。 */
  const loadMarkImage = useCallback(
    async (file: File) => {
      if (file.size > MAX_IMAGE_BYTES) {
        setError(t("imageTooLarge"));
        return;
      }
      const mime: WatermarkImage["mime"] | null =
        file.type === "image/jpeg" ? "image/jpeg" : file.type === "image/png" ? "image/png" : null;
      if (!mime) {
        setError(t("imageType"));
        return;
      }
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result));
          reader.onerror = () => reject(new Error("read-failed"));
          reader.readAsDataURL(file);
        });
        const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
          image.onerror = () => reject(new Error("decode-failed"));
          image.src = dataUrl;
        });
        setMarkImage({ bytes, mime, width: size.width, height: size.height, dataUrl });
        setError("");
        setResult(null);
      } catch {
        setError(t("imageType"));
      }
    },
    [t],
  );

  /** 首屏缩略图是异步的（PDF 要过一遍 pdf.js），失败就退回纯文字预览。 */
  const loadPreview = useCallback((bytes: Uint8Array, kind: Format) => {
    const token = previewTokenRef.current + 1;
    previewTokenRef.current = token;
    setPagePreview(null);
    void renderPagePreview(bytes, kind)
      .then((next) => {
        if (previewTokenRef.current === token) setPagePreview(next);
      })
      .catch(() => {
        if (previewTokenRef.current === token) setPagePreview(null);
      });
  }, []);

  const clearOutcome = useCallback(() => {
    setResult(null);
    setError("");
    setManualNote("");
  }, []);

  const revealWorkspace = useCallback(() => {
    window.setTimeout(() => {
      workspaceRef.current?.scrollIntoView({
        behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth",
        block: "start",
      });
    }, 60);
  }, []);

  const reset = useCallback(() => {
    sourceRef.current = null;
    previewTokenRef.current += 1;
    setPagePreview(null);
    setStage("idle");
    setFormat(null);
    setFileName("");
    setFileSize(0);
    setPageCount(0);
    setFindings([]);
    setSelected(new Set());
    setCustom("");
    setManualNote("");
    setError("");
    setResult(null);
    if (inputRef.current) inputRef.current.value = "";
    window.setTimeout(() => {
      mainRef.current?.scrollIntoView({
        behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth",
        block: "start",
      });
    }, 60);
  }, []);

  const toItems = useCallback(
    (scan: Awaited<ReturnType<typeof scanPdf>> | ReturnType<typeof scanDocx>, kind: Format): Item[] => {
      if (kind === "pdf") {
        const pdf = scan as Awaited<ReturnType<typeof scanPdf>>;
        return pdf.findings.map((finding) => ({
          id: finding.id,
          kind: finding.kind,
          // 注释水印没有文字可显示，只留类型徽标，避免和标签重复一遍。
          label: finding.kind === "annot" ? "" : finding.text,
          meta:
            finding.kind === "annot"
              ? t("meta.annot", { count: finding.blocks })
              : t("meta.pdfPages", { pages: finding.pages, total: pdf.pageCount, count: finding.blocks }),
        }));
      }
      const docx = scan as ReturnType<typeof scanDocx>;
      return docx.findings.map((finding) => {
        const part = t(`part.${finding.part}`);
        return {
          id: finding.id,
          kind: finding.kind,
          label: finding.text || t(`kind.${finding.kind}`),
          meta:
            (finding.count ?? 1) > 1
              ? t("meta.docxMultiple", { part, file: finding.partName, count: finding.count ?? 1 })
              : t("meta.docx", { part, file: finding.partName }),
        };
      });
    },
    [t],
  );

  const scanBytes = useCallback(
    async (bytes: Uint8Array, name: string, kind: Format) => {
      setStage("scanning");
      clearOutcome();
      sourceRef.current = bytes;
      setFormat(kind);
      setFileName(name);
      setFileSize(bytes.byteLength);

      try {
        if (kind === "docx") {
          const scan = scanDocx(bytes);
          const items = toItems(scan, kind);
          setFindings(items);
          setSelected(new Set(items.map((item) => item.id)));
          setPageCount(0);
        } else {
          const scan = await scanPdf(bytes);
          const items = toItems(scan, kind);
          setFindings(items);
          setSelected(new Set(items.map((item) => item.id)));
          setPageCount(scan.pageCount);
        }
        setStage("ready");
        revealWorkspace();
        // 移除模式也要一份原始渲染图，放在水印列表左侧作对照。
        loadPreview(bytes, kind);
      } catch (cause) {
        setFindings([]);
        setSelected(new Set());
        setStage("idle");
        setError(t(`errors.${errorCode(cause)}`));
      }
    },
    [clearOutcome, loadPreview, revealWorkspace, t, toItems],
  );

  /** 添加模式不扫描，直接进设置面板；切回移除模式时再按需检测。 */
  const prepareAdd = useCallback(
    (bytes: Uint8Array, name: string, kind: Format) => {
      sourceRef.current = bytes;
      setFormat(kind);
      setFileName(name);
      setFileSize(bytes.byteLength);
      setPageCount(0);
      setFindings([]);
      setSelected(new Set());
      setCustom("");
      setManualNote("");
      setError("");
      setResult(null);
      setStage("ready");
      revealWorkspace();
      loadPreview(bytes, kind);
    },
    [loadPreview, revealWorkspace],
  );

  const addFile = useCallback(
    async (file: File) => {
      const kind = extensionOf(file.name);
      if (!kind) {
        setError(t("errors.file"));
        return;
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (mode === "add") prepareAdd(bytes, file.name, kind);
      else await scanBytes(bytes, file.name, kind);
    },
    [mode, prepareAdd, scanBytes, t],
  );

  const changeMode = useCallback(
    (next: Mode) => {
      if (next === mode) return;
      setMode(next);
      setResult(null);
      setError("");
      setManualNote("");

      const source = sourceRef.current;
      if (!source || !format) {
        setStage("idle");
        return;
      }
      if (next === "remove") {
        void scanBytes(source, fileName, format);
        return;
      }
      setFindings([]);
      setSelected(new Set());
      setStage("ready");
      // 先在移除模式里选过文件时，切回添加要把同一份文件的首屏补上。
      if (!pagePreview) loadPreview(source, format);
    },
    [fileName, format, loadPreview, mode, pagePreview, scanBytes],
  );

  const handleRemove = useCallback(async () => {
    const source = sourceRef.current;
    if (!source || !format) return;
    const ids = findings.filter((item) => selected.has(item.id)).map((item) => item.id);
    if (ids.length === 0) {
      setError(t("selectHint"));
      return;
    }

    setStage("working");
    clearOutcome();

    try {
      if (format === "docx") {
        const { bytes, removed } = removeDocxWatermarks(source, ids);
        setResult({ bytes, name: `${baseName(fileName)}-clean.docx`, mode: "remove", removed });
      } else {
        const { bytes, removed } = await removePdfWatermarks(source, ids);
        setResult({
          bytes,
          name: `${baseName(fileName)}-clean.pdf`,
          mode: "remove",
          removed: removed.text + removed.form + removed.annot,
        });
      }
      setStage("done");
    } catch (cause) {
      setStage("ready");
      setError(t(`errors.${errorCode(cause)}`));
    }
  }, [clearOutcome, fileName, findings, format, selected, t]);

  const handleAdd = useCallback(async () => {
    const source = sourceRef.current;
    if (!source || !format) return;
    if (markKind === "text" && !markText.trim()) {
      setError(t("textRequired"));
      return;
    }
    if (markKind === "image" && !markImage) {
      setError(t("imageRequired"));
      return;
    }
    if (format === "pdf" && !isValidPageRange(markPages, pageCount || 1)) {
      setError(t("pagesInvalid"));
      return;
    }

    const options = {
      text: expandWatermarkVariables(markText.trim(), {
        fileName: baseName(fileName),
        pageCount: pageCount || 1,
        now: new Date(),
      }),
      opacity: markOpacity,
      angle: markAngle,
      color: markColor,
      tone: markTone,
      weight: markWeight,
      count: markCount,
      scale: markScale,
      gap: markGap,
      position: markPosition,
      layer: markLayer,
      pages: markPages,
      kind: markKind,
      image: markImage
        ? {
            bytes: markImage.bytes,
            mime: markImage.mime,
            width: markImage.width,
            height: markImage.height,
          }
        : undefined,
    };

    setStage("working");
    clearOutcome();

    try {
      if (format === "docx") {
        const { bytes } = addDocxWatermark(source, options);
        setResult({ bytes, name: `${baseName(fileName)}-watermarked.docx`, mode: "add", count: markCount });
      } else {
        const { bytes, pages } = await addPdfWatermark(source, options);
        setResult({ bytes, name: `${baseName(fileName)}-watermarked.pdf`, mode: "add", pages, count: markCount });
      }
      setStage("done");
    } catch (cause) {
      setStage("ready");
      setError(t(`errors.${errorCode(cause)}`));
    }
  }, [
    clearOutcome, fileName, format, markAngle, markColor, markCount, markGap, markImage, markKind,
    markLayer, markOpacity, markPages, markPosition, markScale, markText, markTone, markWeight, pageCount, t,
  ]);

  const toggle = useCallback((id: string) => {
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
    // 改了选择，上一次的结果就不再对应当前勾选；退回可执行状态并收起下载卡。
    setResult(null);
    setStage("ready");
  }, []);

  const addCustom = useCallback(() => {
    const text = custom.trim();
    if (!text) {
      setManualNote(t("manualEmpty"));
      return;
    }
    if (format !== "pdf") return;
    const key = text.replace(/\s+/g, "");
    if (key.length < 2) {
      setManualNote(t("manualEmpty"));
      return;
    }
    const id = `text:${encodeURIComponent(key)}`;
    if (findings.some((item) => item.id === id)) {
      setManualNote(t("manualExists"));
      return;
    }
    const next: Item = { id, kind: "text", label: text, meta: t("meta.custom"), custom: true };
    setFindings((current) => [...current, next]);
    setSelected((current) => new Set([...current, id]));
    setCustom("");
    setManualNote("");
    setResult(null);
    setStage("ready");
  }, [custom, findings, format, t]);

  const enabled = selected.size;
  const allSelected = findings.length > 0 && enabled === findings.length;
  const busy = stage === "scanning" || stage === "working";
  const isImage = markKind === "image";
  // 没填字时预览用样例顶着，但字号按真实字数缩放，长句也能整行放下。
  const previewIsSample = !isImage && markText.trim().length === 0;
  const resolvedText = useMemo(
    () =>
      expandWatermarkVariables(previewIsSample ? t("markPreviewSample") : markText.trim(), {
        fileName: baseName(fileName),
        pageCount: pageCount || 1,
        now: new Date(),
      }),
    [previewIsSample, markText, fileName, pageCount, t],
  );
  // 2 枚上下两行，其余按 n×n；字号再按列数、缩放与密度缩，平铺后不会互相压住。
  const previewColumns = markCount <= 2 ? 1 : Math.max(1, Math.min(4, Math.round(Math.sqrt(markCount))));
  const density = markCount > 1 ? markGap : 1;
  // 用 cqw：字号跟着纸宽走，单枚落在纸宽的 68% 上下，字多字少都不会溢出纸边。
  const previewFont = `${((Math.min(13, 68 / Math.max(1, resolvedText.length)) / previewColumns) * markScale * density).toFixed(2)}cqw`;
  const previewImageWidth = `${((60 / previewColumns) * markScale * density).toFixed(2)}cqw`;
  const previewTiles = Array.from({ length: markCount }, (_, index) => index);
  const previewWeight = MARK_WEIGHTS[markWeight];
  const previewStroke = `${(0.4 + markWeight * 0.5).toFixed(1)}px`;
  const previewOpacity = Math.min(0.86, 0.3 + markOpacity * 0.9);
  const posColumn = markPosition.horizontal === "left" ? 1 : markPosition.horizontal === "right" ? 3 : 2;
  const posRow = markPosition.vertical === "top" ? 1 : markPosition.vertical === "bottom" ? 3 : 2;
  const weightNames = [t("weightThin"), t("weightRegular"), t("weightBold"), t("weightHeavy")];
  const customColor = !MARK_COLORS.includes(markColor);

  /** 结果行的一句话说明：添加时报「每页几枚」，移除时报「删了几处」。 */
  const describeResult = (value: Result): string => {
    const size = formatBytes(value.bytes.byteLength);
    if (value.mode === "remove") return t("resultMeta", { count: value.removed ?? 0, size });

    const count = value.count ?? 1;
    const tiled = count > 1;
    if (format === "pdf") {
      const pages = value.pages ?? 0;
      return tiled
        ? t("addResultMetaPdfTiled", { pages, count, size })
        : t("addResultMetaPdf", { pages, size });
    }
    return tiled ? t("addResultMetaDocxTiled", { count, size }) : t("addResultMetaDocx", { size });
  };

  return (
    <div className={styles.page}>
      <input
        ref={inputRef}
        hidden
        type="file"
        accept=".docx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        aria-label={t("browse")}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) void addFile(file);
          event.target.value = "";
        }}
      />

      <main ref={mainRef}>
        <section className={hero.layout}>
          <div className={hero.copy}>
            <h1>
              {t("title")}
            </h1>
            <p>{t("intro")}</p>
          </div>

          <div
            className={`${hero.panel} ${dropping ? styles.dragging : ""}`}
            onDragOver={(event) => {
              event.preventDefault();
              setDropping(true);
            }}
            onDragLeave={() => setDropping(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDropping(false);
              const file = event.dataTransfer.files?.[0];
              if (file) void addFile(file);
            }}
          >
            <div className={hero.panelHeading}>
              <h2>{t("choose")}</h2>
              <span>.docx / .pdf</span>
            </div>

            <div className={styles.modes} role="group" aria-label={t("modeLabel")}>
              <button
                type="button"
                className={mode === "remove" ? styles.modeOn : undefined}
                aria-pressed={mode === "remove"}
                onClick={() => changeMode("remove")}
              >
                {t("modeRemove")}
              </button>
              <button
                type="button"
                className={mode === "add" ? styles.modeOn : undefined}
                aria-pressed={mode === "add"}
                onClick={() => changeMode("add")}
              >
                {t("modeAdd")}
              </button>
            </div>

            <button
              type="button"
              className={styles.dropzone}
              onClick={() => inputRef.current?.click()}
              disabled={busy}
            >
              <span className={styles.fileIcons} aria-hidden="true">
                {busy ? (
                  <LoaderCircle size={24} className={styles.spin} />
                ) : (
                  <>
                    <span
                      className={`${styles.fileIcon} ${format && format !== "docx" ? styles.iconDim : ""}`}
                    >
                      <i className="ri-file-word-2-line" />
                    </span>
                    <span
                      className={`${styles.fileIcon} ${styles.iconPdf} ${format && format !== "pdf" ? styles.iconDim : ""}`}
                    >
                      <i className="ri-file-pdf-2-line" />
                    </span>
                  </>
                )}
              </span>
              <span className={styles.dropText}>
                <strong>{busy ? t("scanning") : t("dropTitle")}</strong>
                <small>{t("dropHint")}</small>
              </span>
            </button>

            {fileName && (
              <div className={styles.fileRow}>
                <i className={format === "pdf" ? "ri-file-pdf-2-line" : "ri-file-word-2-line"} aria-hidden="true" />
                <span className={styles.fileRowName}>
                  {fileName}
                  <small>{formatBytes(fileSize)}</small>
                </span>
                <button
                  type="button"
                  className={styles.iconButton}
                  aria-label={t("clearFile")}
                  onClick={reset}
                >
                  <X size={15} aria-hidden="true" />
                </button>
              </div>
            )}

            <p className={hero.privacy}>
              <LockKeyhole size={14} aria-hidden="true" />
              {t("privacy")}
            </p>
          </div>
        </section>

        {error && (
          <div role="alert" className={styles.error}>
            <span>{error}</span>
            <button type="button" aria-label={t("dismiss")} onClick={() => setError("")}>
              <X size={17} aria-hidden="true" />
            </button>
          </div>
        )}

        {(stage === "ready" || stage === "working" || stage === "done") && (
          <section className={styles.workspace} ref={workspaceRef} aria-labelledby="watermark-results">
            <div className={styles.workspaceHead}>
              <h2 id="watermark-results">
                {mode === "add"
                  ? t("addHeadTitle")
                  : findings.length > 0
                    ? t("found", { count: findings.length })
                    : t("noFindings")}
                {mode === "remove" && format === "pdf" && pageCount > 0 && (
                  <small>{t("pageCount", { count: pageCount })}</small>
                )}
              </h2>
              <div className={styles.headActions}>
                {mode === "remove" && findings.length > 0 && (
                  <>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={() => {
                        setSelected(allSelected ? new Set() : new Set(findings.map((item) => item.id)));
                        setResult(null);
                        setStage("ready");
                      }}
                    >
                      {allSelected ? t("selectNone") : t("selectAll")}
                    </button>
                    <span>{t("selectedCount", { count: enabled })}</span>
                  </>
                )}
                {/* 标题行右侧的模式切换：和上传区那组按钮共用 changeMode，状态会一起重置。 */}
                <button
                  type="button"
                  className={styles.modeSwitch}
                  onClick={() => changeMode(mode === "add" ? "remove" : "add")}
                >
                  {mode === "add" ? (
                    <Eraser size={14} aria-hidden="true" />
                  ) : (
                    <Plus size={14} aria-hidden="true" />
                  )}
                  {mode === "add" ? t("modeRemove") : t("modeAdd")}
                </button>
              </div>
            </div>

            {mode === "add" ? (
              <>
                <div className={styles.markPanel}>
                  <div className={styles.markPreview}>
                    {/* 纸张一直在：缩略图没渲染好之前是白纸，换文件时不会先塌一块灰底再弹回来。 */}
                    <div
                      className={styles.markPage}
                      aria-hidden="true"
                      style={{ ["--page-ratio" as string]: pagePreview?.ratio ?? 0.707 }}
                    >
                      {pagePreview?.kind === "pdf" && (
                        // 首屏是本地生成的 data URL，next/image 的优化链在这里没有意义。
                        // eslint-disable-next-line @next/next/no-img-element
                        <img className={styles.markPageImage} src={pagePreview.image} alt="" />
                      )}
                      {pagePreview?.kind === "docx" && (
                        <div className={styles.markPageText}>
                          {pagePreview.paragraphs.map((line, index) => (
                            <p
                              key={index}
                              className={index === 0 && line.length < 32 ? styles.markPageLead : undefined}
                            >
                              {line}
                            </p>
                          ))}
                        </div>
                      )}
                      <div
                        className={styles.markGrid}
                        data-count={previewTiles.length}
                        style={{
                          // 与导出文件一致：below 沉在正文下，above 浮在正文上。
                          zIndex: markLayer === "above" ? 2 : 0,
                          // 单枚时切 3×3，把水印放到九宫格选中的那一格。
                          ...(markCount === 1
                            ? {
                                gridTemplateColumns: "repeat(3,minmax(0,1fr))",
                                gridTemplateRows: "repeat(3,minmax(0,1fr))",
                              }
                            : {}),
                        }}
                      >
                        {previewTiles.map((tile) => {
                          const cell =
                            markCount === 1 ? { gridColumn: posColumn, gridRow: posRow } : undefined;
                          if (isImage) {
                            return markImage ? (
                              // 用户选的本地图片，data URL 不需要 next/image 优化。
                              // eslint-disable-next-line @next/next/no-img-element
                              <img
                                key={tile}
                                className={styles.markImageTile}
                                src={markImage.dataUrl}
                                alt=""
                                style={{
                                  width: previewImageWidth,
                                  opacity: previewOpacity,
                                  transform: `rotate(${-markAngle}deg)`,
                                  ...cell,
                                }}
                              />
                            ) : (
                              <span key={tile} className={styles.markImagePlaceholder} style={cell}>
                                {t("imageRequired")}
                              </span>
                            );
                          }
                          return (
                            <span
                              key={tile}
                              className={styles.markOverlay}
                              style={{
                                // 缩略图上纯按 18% 画会淡到看不见，抬一个下限，方向仍随滑块变。
                                opacity: previewOpacity,
                                transform: `rotate(${-markAngle}deg)`,
                                color: markTone === "solid" ? markColor : "transparent",
                                fontWeight: previewWeight,
                                WebkitTextStroke:
                                  markTone === "outline" ? `${previewStroke} ${markColor}` : undefined,
                                ["--mark-fs" as string]: previewFont,
                                ...cell,
                              }}
                            >
                              {resolvedText}
                            </span>
                          );
                        })}
                      </div>
                    </div>
                    {previewIsSample && <em className={styles.markTag}>{t("markPreviewTag")}</em>}

                    {/* 操作按钮放在渲染图正下方：未完成是「添加水印」，完成后是「下载」。 */}
                    {stage === "done" && result?.mode === "add" ? (
                      <div className={styles.markActions} role="status">
                        <span className={styles.markActionsHint}>
                          <strong>{t("addResultTitle")}</strong> · {describeResult(result)}
                        </span>
                        <div className={styles.markActionsRow}>
                          <button
                            type="button"
                            className={styles.primary}
                            onClick={() => saveFile(result.bytes, result.name, format ?? "pdf")}
                          >
                            <ArrowDownToLine size={16} aria-hidden="true" />
                            {t("download")}
                          </button>
                          <button type="button" className={styles.ghostAction} onClick={reset}>
                            <RotateCcw size={16} aria-hidden="true" />
                            {t("another")}
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className={styles.markActions}>
                        <span className={styles.markActionsHint}>
                          {isImage
                            ? markImage
                              ? t("addActionHint")
                              : t("imageRequired")
                            : markText.trim()
                              ? t("addActionHint")
                              : t("textRequired")}
                        </span>
                        <div className={styles.markActionsRow}>
                          <button
                            type="button"
                            className={styles.primary}
                            disabled={busy || (isImage ? !markImage : !markText.trim())}
                            onClick={() => void handleAdd()}
                          >
                            {stage === "working" ? (
                              <>
                                <LoaderCircle size={17} className={styles.spin} aria-hidden="true" />
                                {t("adding")}
                              </>
                            ) : (
                              <>
                                {t("addAction")}
                                <ArrowDownToLine size={16} aria-hidden="true" />
                              </>
                            )}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  <div className={styles.markFields}>
                    <div className={`${styles.field} ${styles.fieldWide}`}>
                      <span>{t("markKindLabel")}</span>
                      <div className={styles.segmented}>
                        <button
                          type="button"
                          aria-pressed={markKind === "text"}
                          onClick={() => {
                            setMarkKind("text");
                            setResult(null);
                          }}
                        >
                          {t("kindText")}
                        </button>
                        <button
                          type="button"
                          aria-pressed={markKind === "image"}
                          onClick={() => {
                            setMarkKind("image");
                            setResult(null);
                          }}
                        >
                          {t("kindImage")}
                        </button>
                      </div>
                    </div>

                    {isImage && (
                      <div className={`${styles.field} ${styles.fieldWide}`}>
                        <span>
                          {t("markImageLabel")}
                          <em>{markImage ? `${markImage.width}×${markImage.height}` : ""}</em>
                        </span>
                        <div className={styles.imagePicker}>
                          <button
                            type="button"
                            className={styles.ghostButton}
                            onClick={() => imageInputRef.current?.click()}
                          >
                            {t("markImagePick")}
                          </button>
                          {markImage && (
                            // 本地图片缩略图，data URL 不需要 next/image。
                            // eslint-disable-next-line @next/next/no-img-element
                            <img className={styles.imageThumb} src={markImage.dataUrl} alt="" />
                          )}
                          <small className={styles.fieldHint}>{t("markImageHint")}</small>
                        </div>
                        <input
                          ref={imageInputRef}
                          hidden
                          type="file"
                          accept="image/png,image/jpeg"
                          aria-label={t("markImagePick")}
                          onChange={(event) => {
                            const file = event.target.files?.[0];
                            if (file) void loadMarkImage(file);
                            event.target.value = "";
                          }}
                        />
                      </div>
                    )}

                    <label
                      className={`${styles.field} ${styles.fieldWide}${isImage ? ` ${styles.fieldHidden}` : ""}`}
                      htmlFor="mark-text"
                    >
                      <span>
                        {t("markTextLabel")}
                        <em>
                          {markText.length}/{MAX_MARK_CHARS}
                        </em>
                      </span>
                      <input
                        id="mark-text"
                        value={markText}
                        maxLength={MAX_MARK_CHARS}
                        placeholder={t("markTextPlaceholder")}
                        onChange={(event) => {
                          setMarkText(event.target.value);
                          setError("");
                          setResult(null);
                        }}
                      />
                    </label>

                    {!isImage && (
                      <div className={`${styles.field} ${styles.fieldWide}`}>
                        <span>{t("variableLabel")}</span>
                        <div className={styles.tokenRow}>
                          {VARIABLE_TOKENS.map((token) => (
                            <button
                              key={token}
                              type="button"
                              className={styles.tokenButton}
                              title={t("variableHint")}
                              onClick={() => {
                                setMarkText((current) => (current + token).slice(0, MAX_MARK_CHARS));
                                setResult(null);
                              }}
                            >
                              {token}
                            </button>
                          ))}
                        </div>
                        <small className={styles.fieldHint}>{t("variableHint")}</small>
                      </div>
                    )}

                    <div className={`${styles.field}${isImage ? ` ${styles.fieldDisabled}` : ""}`}>
                      <span>{t("markColorLabel")}</span>
                      <div className={styles.swatches}>
                        {MARK_COLORS.map((color) => (
                          <button
                            key={color}
                            type="button"
                            className={styles.swatch}
                            style={{ ["--swatch" as string]: color }}
                            aria-pressed={markColor === color}
                            aria-label={`${t("markColorLabel")} ${color}`}
                            title={color}
                            onClick={() => {
                              setMarkColor(color);
                              setResult(null);
                            }}
                          />
                        ))}
                        {/* 自定义色：原生取色器藏在色块底下，点色块就弹出调色盘。 */}
                        <label
                          className={`${styles.swatch} ${styles.swatchCustom}${customColor ? ` ${styles.swatchOn}` : ""}`}
                          style={{ ["--swatch" as string]: markColor }}
                          title={t("markColorCustom")}
                        >
                          <input
                            type="color"
                            value={markColor}
                            aria-label={t("markColorCustom")}
                            onChange={(event) => {
                              setMarkColor(event.target.value);
                              setResult(null);
                            }}
                          />
                        </label>
                      </div>
                    </div>

                    <div className={`${styles.field}${isImage ? ` ${styles.fieldDisabled}` : ""}`}>
                      <span>{t("markStyleLabel")}</span>
                      <div className={styles.segmented}>
                        <button
                          type="button"
                          aria-pressed={markTone === "solid"}
                          onClick={() => {
                            setMarkTone("solid");
                            setResult(null);
                          }}
                        >
                          {t("styleSolid")}
                        </button>
                        <button
                          type="button"
                          aria-pressed={markTone === "outline"}
                          onClick={() => {
                            setMarkTone("outline");
                            setResult(null);
                          }}
                        >
                          {t("styleOutline")}
                        </button>
                      </div>
                    </div>

                    <div className={styles.field}>
                      <span>{t("markCountLabel")}</span>
                      <div className={styles.segmented}>
                        {MARK_COUNTS.map((count) => (
                          <button
                            key={count}
                            type="button"
                            aria-pressed={markCount === count}
                            aria-label={`${t("markCountLabel")} ${count}`}
                            onClick={() => {
                              setMarkCount(count);
                              setResult(null);
                            }}
                          >
                            {count}
                          </button>
                        ))}
                      </div>
                    </div>

                    <label className={styles.field} htmlFor="mark-angle">
                      <span>
                        {t("markAngleLabel")}
                        <em>{markAngle}°</em>
                      </span>
                      <input
                        id="mark-angle"
                        type="range"
                        min={0}
                        max={360}
                        step={15}
                        value={markAngle}
                        onChange={(event) => {
                          setMarkAngle(Number(event.target.value));
                          setResult(null);
                        }}
                      />
                    </label>

                    <label className={`${styles.field}${isImage ? ` ${styles.fieldDisabled}` : ""}`} htmlFor="mark-weight">
                      <span>
                        {t("markWeightLabel")}
                        <em>{weightNames[markWeight]}</em>
                      </span>
                      <input
                        id="mark-weight"
                        type="range"
                        min={0}
                        max={3}
                        step={1}
                        value={markWeight}
                        onChange={(event) => {
                          setMarkWeight(Number(event.target.value));
                          setResult(null);
                        }}
                      />
                    </label>

                    <label className={styles.field} htmlFor="mark-opacity">
                      <span>
                        {t("markOpacityLabel")}
                        <em>{Math.round(markOpacity * 100)}%</em>
                      </span>
                      <input
                        id="mark-opacity"
                        type="range"
                        min={5}
                        max={60}
                        step={5}
                        value={Math.round(markOpacity * 100)}
                        onChange={(event) => {
                          setMarkOpacity(Number(event.target.value) / 100);
                          setResult(null);
                        }}
                      />
                    </label>

                    <details open className={styles.advanced}>
                      <summary className={styles.advancedSummary}>{t("advancedTitle")}</summary>
                      <div className={styles.markFields}>
                        <label className={styles.field} htmlFor="mark-scale">
                          <span>
                            {t("markScaleLabel")}
                            <em>{Math.round(markScale * 100)}%</em>
                          </span>
                          <input
                            id="mark-scale"
                            type="range"
                            min={50}
                            max={160}
                            step={10}
                            value={Math.round(markScale * 100)}
                            onChange={(event) => {
                              setMarkScale(Number(event.target.value) / 100);
                              setResult(null);
                            }}
                          />
                        </label>

                        <label
                          className={`${styles.field}${markCount === 1 ? ` ${styles.fieldDisabled}` : ""}`}
                          htmlFor="mark-gap"
                        >
                          <span>
                            {t("markGapLabel")}
                            <em>{Math.round(markGap * 100)}%</em>
                          </span>
                          <input
                            id="mark-gap"
                            type="range"
                            min={60}
                            max={140}
                            step={10}
                            value={Math.round(markGap * 100)}
                            onChange={(event) => {
                              setMarkGap(Number(event.target.value) / 100);
                              setResult(null);
                            }}
                          />
                        </label>

                        <div className={styles.field}>
                          <span>{t("markPositionLabel")}</span>
                          <div className={styles.positionGrid} role="group" aria-label={t("markPositionLabel")}>
                            {POSITION_CELLS.map(({ horizontal, vertical }) => (
                              <button
                                key={`${vertical}-${horizontal}`}
                                type="button"
                                aria-pressed={
                                  markPosition.horizontal === horizontal && markPosition.vertical === vertical
                                }
                                aria-label={`${vertical}-${horizontal}`}
                                onClick={() => {
                                  setMarkPosition({ horizontal, vertical });
                                  setResult(null);
                                }}
                              />
                            ))}
                          </div>
                          <small className={styles.fieldHint}>
                            {markCount > 1 ? t("positionTiledNote") : ""}
                          </small>
                        </div>

                        <div className={styles.field}>
                          <span>{t("markLayerLabel")}</span>
                          <div className={styles.segmented}>
                            <button
                              type="button"
                              aria-pressed={markLayer === "below"}
                              onClick={() => {
                                setMarkLayer("below");
                                setResult(null);
                              }}
                            >
                              {t("layerBelow")}
                            </button>
                            <button
                              type="button"
                              aria-pressed={markLayer === "above"}
                              onClick={() => {
                                setMarkLayer("above");
                                setResult(null);
                              }}
                            >
                              {t("layerAbove")}
                            </button>
                          </div>
                        </div>

                        <label className={`${styles.field} ${styles.fieldWide}`} htmlFor="mark-pages">
                          <span>
                            {t("markPagesLabel")}
                            <em>{format === "docx" ? t("pagesDocxNote") : ""}</em>
                          </span>
                          <input
                            id="mark-pages"
                            value={markPages}
                            disabled={format === "docx"}
                            placeholder="1,3,5-8"
                            onChange={(event) => {
                              setMarkPages(event.target.value);
                              setResult(null);
                            }}
                          />
                          <small className={styles.fieldHint}>
                            {format === "docx" ? t("pagesDocxHint") : t("markPagesHint")}
                          </small>
                        </label>

                        <div className={`${styles.field} ${styles.fieldWide} ${styles.presetRow}`}>
                          <button type="button" className={styles.ghostButton} onClick={resetSettings}>
                            {t("presetReset")}
                          </button>
                          <small className={styles.fieldHint}>{t("presetHint")}</small>
                        </div>
                      </div>
                    </details>
                  </div>
                </div>

                <p className={styles.notice}>
                  <strong>{t("addNoticeTitle")}</strong>
                  {t("addNoticeBody")}
                </p>
              </>
            ) : (
              <div className={styles.removePanel}>
                <div className={styles.removePreview}>
                  {/* 原始文档第一页：左边这份永远是不带水印的原文，供对照勾选。 */}
                  <div
                    className={styles.markPage}
                    style={{ ["--page-ratio" as string]: pagePreview?.ratio ?? 0.707 }}
                  >
                    {pagePreview?.kind === "pdf" && (
                      // 首屏是本地生成的 data URL，next/image 的优化链在这里没有意义。
                      // eslint-disable-next-line @next/next/no-img-element
                      <img className={styles.markPageImage} src={pagePreview.image} alt="" />
                    )}
                    {pagePreview?.kind === "docx" && (
                      <div className={styles.markPageText}>
                        {pagePreview.paragraphs.map((line, index) => (
                          <p
                            key={index}
                            className={index === 0 && line.length < 32 ? styles.markPageLead : undefined}
                          >
                            {line}
                          </p>
                        ))}
                      </div>
                    )}
                  </div>
                  <em className={styles.markTag}>{t("removePreviewTag")}</em>
                </div>

                <div className={styles.removeList}>
                  {findings.length > 0 && (
                    <ul className={styles.findings}>
                      {findings.map((item) => {
                        const on = selected.has(item.id);
                        return (
                          <li key={item.id}>
                            <label className={`${styles.finding} ${on ? styles.findingOn : ""}`}>
                              <input type="checkbox" checked={on} onChange={() => toggle(item.id)} />
                              <span className={styles.findingBody}>
                                <span className={styles.findingTop}>
                                  <span className={styles.badge}>{t(`kind.${item.kind}`)}</span>
                                  {item.label && <span className={styles.findingText}>{item.label}</span>}
                                </span>
                                <span className={styles.findingMeta}>{item.meta}</span>
                              </span>
                            </label>
                          </li>
                        );
                      })}
                    </ul>
                  )}

                  {format === "pdf" && stage !== "done" && (
                    <div className={styles.manual}>
                      <label htmlFor="watermark-manual">{t("manualLabel")}</label>
                      <div className={styles.manualRow}>
                        <input
                          id="watermark-manual"
                          value={custom}
                          placeholder={t("manualPlaceholder")}
                          onChange={(event) => {
                            setCustom(event.target.value);
                            setManualNote("");
                          }}
                          onKeyDown={(event) => {
                            if (event.key === "Enter") {
                              event.preventDefault();
                              addCustom();
                            }
                          }}
                        />
                        <button type="button" className={styles.ghostButton} onClick={addCustom}>
                          <Plus size={14} aria-hidden="true" />
                          {t("manualAdd")}
                        </button>
                      </div>
                      <p className={styles.manualHint}>{manualNote || t("manualHint")}</p>
                    </div>
                  )}

                  {findings.length === 0 && (
                    <p className={styles.notice}>
                      <strong>{t("emptyTitle")}</strong>
                      {t("emptyBody")}
                    </p>
                  )}

                  <p className={styles.notice}>
                    <strong>{t("limitsTitle")}</strong>
                    {t("limitsBody")}
                  </p>
                </div>
              </div>
            )}

            {mode === "remove" &&
              (stage === "done" && result ? (
                <div className={styles.result} role="status">
                  <span className={styles.resultIcon}>
                    <Check size={18} aria-hidden="true" />
                  </span>
                  <span className={styles.resultText}>
                    <strong>{result.mode === "add" ? t("addResultTitle") : t("resultTitle")}</strong>
                    <small>{describeResult(result)}</small>
                  </span>
                  <span className={styles.resultActions}>
                    <button
                      type="button"
                      className={styles.primary}
                      onClick={() => saveFile(result.bytes, result.name, format ?? "pdf")}
                    >
                      <ArrowDownToLine size={16} aria-hidden="true" />
                      {t("download")}
                    </button>
                    <button type="button" className={styles.ghostAction} onClick={reset}>
                      <RotateCcw size={16} aria-hidden="true" />
                      {t("another")}
                    </button>
                  </span>
                </div>
              ) : (
                <div className={styles.actions}>
                  <span className={styles.actionsHint}>
                    {enabled > 0 ? t("actionsHint") : t("selectHint")}
                  </span>
                  <button
                    type="button"
                    className={styles.primary}
                    disabled={busy || enabled === 0}
                    onClick={() => void handleRemove()}
                  >
                    {stage === "working" ? (
                      <>
                        <LoaderCircle size={17} className={styles.spin} aria-hidden="true" />
                        {t("removing")}
                      </>
                    ) : (
                      <>
                        {t("remove", { count: enabled })}
                        <ArrowDownToLine size={16} aria-hidden="true" />
                      </>
                    )}
                  </button>
                </div>
              ))}
          </section>
        )}

        {stage !== "ready" && stage !== "working" && stage !== "done" && (
          <section className={styles.features}>
            {[ScanSearch, ShieldCheck, FileCheck2].map((Icon, index) => (
              <article key={index}>
                <Icon size={22} aria-hidden="true" />
                <h2>{t(`feature${index + 1}`)}</h2>
                <p>{t(`body${index + 1}`)}</p>
              </article>
            ))}
          </section>
        )}

        <div className={styles.seo}></div>
      </main>
    </div>
  );
}

/** 把引擎错误码收敛成 messages 里的键，组件负责翻译。 */
function errorCode(
  cause: unknown,
): "tooLarge" | "notDocx" | "encrypted" | "empty" | "invalidPdf" | "noText" | "canvas" | "image" | "unknown" {
  if (cause instanceof DocxWatermarkError) {
    if (cause.code === "too-large") return "tooLarge";
    if (cause.code === "empty-text") return "noText";
    if (cause.code === "invalid-image") return "image";
    return "notDocx";
  }
  if (cause instanceof PdfWatermarkError) {
    if (cause.code === "encrypted") return "encrypted";
    if (cause.code === "too-large") return "tooLarge";
    if (cause.code === "empty") return "empty";
    return "invalidPdf";
  }
  if (cause instanceof Error && cause.message === "canvas-unavailable") return "canvas";
  return "unknown";
}
