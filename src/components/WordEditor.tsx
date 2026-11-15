import React, { Suspense, useEffect, useRef, useState } from "react";
import { useLang, useT } from "../i18n";
import type { DocxEditorRef } from "@docx-editor.dev/react";
import { ArrowDownToLine, ArrowRight, FilePlus2, FolderOpen, LoaderCircle, LockKeyhole, PencilLine, Printer, Table2 } from "lucide-react";
import { useTheme } from "../lib/use-theme";
import { ConfirmDialog } from "./ConfirmDialog";
import { useHoverTooltip } from "./use-hover-tooltip";
import { readEditorDocument } from "../lib/local-document-transfer";
import "@docx-editor.dev/react/styles.css";
import { defaultEditorLocale, editorLocales } from "../lib/editor-locales";
import hero from "../styles/tool-hero.module.css";
import styles from "../styles/word-editor.module.css";

const Editor = React.lazy(() => import("@docx-editor.dev/react").then(module => ({ default: module.DocxEditor })));
// Help only links out to the upstream issue tracker, so drop it together with its row.
const editorMenu = { reportIssue: false };

class EditorBoundary extends React.Component<{ children: React.ReactNode; message: string }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  render() { return this.state.failed ? <p role="alert" className={styles.error}>{this.props.message}</p> : this.props.children; }
}

export default function WordEditor() {
  const t = useT("wordEditor");
  const locale = useLang().lang;
  const { isDark } = useTheme();
  const tip = useHoverTooltip();
  const [source, setSource] = useState<ArrayBuffer | "blank" | null>(null);
  const [name, setName] = useState("");
  const [documentId, setDocumentId] = useState(0);
  const [ready, setReady] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [printing, setPrinting] = useState(false);
  const [opening, setOpening] = useState(false);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [pending, setPending] = useState<{ source: ArrayBuffer | "blank"; name: string } | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const editor = useRef<DocxEditorRef | null>(null);
  const revision = useRef<number | null>(null);
  const saveLock = useRef(false);
  const currentRevision = useRef<number | null>(null);

  useEffect(() => {
    const id = new URL(window.location.href).searchParams.get("document");
    if (!id) return;
    let cancelled = false;
    // Keep file controls disabled while the cross-tab document is loading.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setOpening(true);
    readEditorDocument(id).then(async file => ({ bytes: await file.arrayBuffer(), name: file.name })).then(document => {
      if (cancelled) return;
      setSource(document.bytes);
      setName(document.name);
    }).catch(() => {
      if (!cancelled) setError(t("error"));
    }).finally(() => {
      if (!cancelled) setOpening(false);
    });
    return () => { cancelled = true; };
  }, [t]);

  function install(next: { source: ArrayBuffer | "blank"; name: string }) {
    setSource(next.source); setName(next.name); setDocumentId(id => id + 1);
    setReady(false); setDirty(false); setError(""); setPending(null);
    revision.current = null;
    currentRevision.current = null;
  }

  function replace(next: { source: ArrayBuffer | "blank"; name: string }) {
    if (dirty) setPending(next); else install(next);
  }

  async function openFile(file: File) {
    if (saving || opening) return;
    if (!file.name.toLowerCase().endsWith(".docx") || file.size > 20 * 1024 * 1024) {
      setError(t("fileError")); return;
    }
    setOpening(true);
    try { replace({ source: await file.arrayBuffer(), name: file.name }); }
    catch { setError(t("error")); }
    finally { setOpening(false); }
  }

  async function openDemo() {
    setOpening(true); setError("");
    try {
      const response = await fetch(`/demo/compare/${locale}/before.docx`);
      if (!response.ok) throw new Error("Sample unavailable");
      replace({ source: await response.arrayBuffer(), name: t("sampleName") });
    } catch { setError(t("error")); }
    finally { setOpening(false); }
  }

  async function download() {
    if (!ready || saveLock.current) return;
    saveLock.current = true; setSaving(true); setError("");
    const savedRevision = currentRevision.current;
    try {
      const bytes = await editor.current?.save();
      if (!bytes) throw new Error("No document");
      const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
      const link = document.createElement("a"); link.href = url;
      const filename = name.trim().replace(/[\\/:*?"<>|]/g, "_") || t("untitled");
      link.download = filename.toLowerCase().endsWith(".docx") ? filename : `${filename}.docx`;
      link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
      if (currentRevision.current === savedRevision) { setDirty(false); }
    } catch { setError(t("saveError")); }
    finally { saveLock.current = false; setSaving(false); }
  }

  // Print and "Save as PDF" are the same browser dialog. The DOCX is converted
  // and paginated locally (no upload); save() supplies the current edits, so
  // printing does not require downloading first.
  async function printDocument() {
    if (!ready || saveLock.current) return;
    saveLock.current = true; setPrinting(true); setError("");
    try {
      const bytes = await editor.current?.save();
      if (!bytes) throw new Error("No document");
      const { printDocx } = await import("../lib/print-document");
      await printDocx(bytes, name.trim() || t("untitled"));
    } catch { setError(t("printError")); }
    finally { saveLock.current = false; setPrinting(false); }
  }

  return <div className={styles.page}>
    <input hidden ref={fileInput} type="file" accept=".docx" aria-label={t("browse")} onChange={event => {
      const file = event.target.files?.[0]; if (file) void openFile(file); event.target.value = "";
    }} />
    {!source ? <main>
      <section className={hero.layout}>
        <div className={hero.copy}>
          <h1>{t("title")}</h1>
          <p>{t("intro")}</p>
          <button className={hero.demo} disabled={opening} onClick={openDemo}><i className="ri-file-word-line" aria-hidden="true" />{t("demo")}<ArrowRight size={16} /></button>
        </div>
        <div className={`${hero.panel} ${styles.upload} ${dragging ? styles.dragging : ""}`}
          onDragOver={event => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)}
          onDrop={event => { event.preventDefault(); setDragging(false); if (event.dataTransfer.files[0]) void openFile(event.dataTransfer.files[0]); }}>
          <div className={styles.uploadIcon}><i className="ri-file-word-line" aria-hidden="true" /><span><PencilLine size={15} /></span></div>
          <h2>{t("upload")}</h2><p>{t("limit")}</p>
          <button className={styles.primary} disabled={opening} onClick={() => fileInput.current?.click()}>{opening ? <LoaderCircle size={17} className={styles.spin} /> : <FolderOpen size={17} />}{t("browse")}</button>
          <button className={styles.blank} disabled={opening} onClick={() => replace({ source: "blank", name: t("untitled") })}><FilePlus2 size={16} />{t("blank")}</button>
          <p className={hero.privacy}><LockKeyhole size={14} aria-hidden="true" />{t("privacy")}</p>
        </div>
      </section>
      <div className={styles.landing}>
        {error && <p role="alert" className={styles.error}>{error}</p>}
        <section className={styles.features}>{[PencilLine, Table2, ArrowDownToLine].map((Icon, index) => <article key={index}><Icon size={22} strokeWidth={1.5} /><h2>{t(`feature${index + 1}`)}</h2><p>{t(`body${index + 1}`)}</p></article>)}</section></div>
    </main> : <main className={styles.workspace}>
      {error && <p role="alert" className={styles.error}>{error}</p>}
      <div className={styles.editor} inert={saving}>
        <EditorBoundary key={documentId} message={t("error")}>
          <Suspense fallback={<div role="status" className={styles.loading}><LoaderCircle size={24} className={styles.spin} />{t("loading")}</div>}>
            <Editor ref={editor} document={source} mode="edit" title={name} onTitleChange={setName}
              menu={editorMenu}
              colorMode={isDark ? "dark" : "light"} i18n={editorLocales[locale] ?? defaultEditorLocale}
              locale={locale === "zh" ? "zh-CN" : locale} onOpen={() => fileInput.current?.click()} onSave={download}
              renderTitleBarLeft={() => (
                <i className={`ri-file-word-line ${styles.titleIcon}`} aria-hidden="true" />
              )}
              renderTitleBarRight={() => (
                <div className={styles.titleActions}>
                  <button type="button" data-tip={t("browse")} {...tip.triggerProps} disabled={saving || opening} onClick={() => fileInput.current?.click()} aria-label={t("browse")}>
                    <FolderOpen size={16} />
                  </button>
                  <button type="button" data-tip={t("blank")} {...tip.triggerProps} disabled={saving} onClick={() => replace({ source: "blank", name: t("untitled") })} aria-label={t("blank")}>
                    <FilePlus2 size={16} />
                  </button>
                  <button type="button" data-tip={t("print")} {...tip.triggerProps} onClick={printDocument} disabled={!ready || saving || printing} aria-label={t("print")}>
                    {printing ? <LoaderCircle size={16} className={styles.spin} /> : <Printer size={16} />}
                  </button>
                  <button type="button" data-tip={t("download")} {...tip.triggerProps} className={`${styles.download} ${dirty ? styles.downloadDirty : ""}`} onClick={download} disabled={!ready || saving || printing} aria-label={t("download")}>
                    {saving ? <LoaderCircle size={16} className={styles.spin} /> : <ArrowDownToLine size={16} />}
                  </button>
                </div>
              )}
              onReady={instance => {
                revision.current = instance.getDocumentHandle().revision;
                currentRevision.current = revision.current;
                setReady(true);
              }}
              onChange={change => { currentRevision.current = change.revision; if (revision.current !== null && change.revision !== revision.current) { setDirty(true); } }}
            />
          </Suspense>
        </EditorBoundary>
      </div>
    </main>}
    {tip.tooltip}
    {pending && <ConfirmDialog title={t("replaceTitle")} message={t("replaceBody")} confirmLabel={t("replace")} cancelLabel={t("cancel")} variant="danger" onConfirm={() => install(pending)} onCancel={() => setPending(null)} />}
  </div>;
}
