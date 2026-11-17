import { useCallback, useEffect, useRef, useState } from 'react'
import { ArrowDownToLine, FileCode2, LockKeyhole, Upload } from 'lucide-react'
import { useT } from '../i18n'
import { markdownToDocxBlob, TEMPLATE_IDS, type TemplateId } from '../lib/markdown-to-docx'
import styles from '../styles/markdown-to-word.module.css'

const MAX_BYTES = 1024 * 1024
/** Long enough that typing does not queue a conversion per keystroke. */
const RENDER_DELAY = 700

/** Strip characters Windows and macOS refuse in file names. */
function safeName(markdown: string): string {
  const heading = markdown.match(/^#\s+(.{1,60})$/m)?.[1]?.trim()
  const base = heading || 'document'
  return base.replace(/[\\/:*?"<>|]/g, '').trim() || 'document'
}

export default function MarkdownToWord() {
  const t = useT('markdownToWord')
  const [value, setValue] = useState('')
  const [template, setTemplate] = useState<TemplateId>('standard')
  const [busy, setBusy] = useState(false)
  const [rendering, setRendering] = useState(false)
  const [error, setError] = useState('')

  const preview = useRef<HTMLDivElement>(null)
  /** Bumped per render so a slow conversion cannot paint over a newer one. */
  const seq = useRef(0)

  /**
   * The preview is the generated .docx rendered back into the page, not a second
   * Markdown renderer: what you see is literally what downloads, and there is
   * only one styling implementation to keep in sync.
   */
  const renderPreview = useCallback(
    async (markdown: string, tpl: TemplateId) => {
      const host = preview.current
      if (!host) return

      if (!markdown.trim()) {
        host.innerHTML = ''
        setRendering(false)
        return
      }

      const current = ++seq.current
      setRendering(true)
      try {
        const blob = await markdownToDocxBlob(markdown, tpl)
        if (current !== seq.current) return
        const { renderAsync } = await import('docx-preview')
        if (current !== seq.current) return
        host.innerHTML = ''
        await renderAsync(blob, host, undefined, {
          className: 'docx',
          inWrapper: true,
          breakPages: true,
          ignoreWidth: false,
          ignoreHeight: false,
        })
        if (current === seq.current) setError('')
      } catch {
        if (current === seq.current) setError(t('failed'))
      } finally {
        if (current === seq.current) setRendering(false)
      }
    },
    [t],
  )

  useEffect(() => {
    const timer = window.setTimeout(() => void renderPreview(value, template), RENDER_DELAY)
    return () => window.clearTimeout(timer)
  }, [value, template, renderPreview])

  async function load(file: File) {
    setError('')
    if (file.size > MAX_BYTES) {
      setError(t('tooLarge'))
      return
    }
    try {
      setValue(await file.text())
    } catch {
      setError(t('failed'))
    }
  }

  async function download() {
    if (busy || !value.trim()) return
    setError('')
    if (new Blob([value]).size > MAX_BYTES) {
      setError(t('tooLarge'))
      return
    }
    setBusy(true)
    try {
      const blob = await markdownToDocxBlob(value, template)
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = `${safeName(value)}.docx`
      link.click()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
    } catch {
      setError(t('failed'))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className={styles.page}>
      <div className={styles.head}>
        <div>
          <h1 className={styles.title}>
            <FileCode2 size={22} style={{ verticalAlign: '-3px', marginRight: 8, color: 'var(--color-primary)' }} aria-hidden="true" />
            {t('title')}
          </h1>
          <p className={styles.lead}>{t('description')}</p>
        </div>
        <div className={styles.actions}>
          <button type="button" className={styles.primary} onClick={() => void download()} disabled={busy || !value.trim()}>
            <ArrowDownToLine size={16} aria-hidden="true" />
            {busy ? t('working') : t('download')}
          </button>
        </div>
      </div>

      <div className={styles.bar}>
        <label className={styles.ghost}>
          <Upload size={14} aria-hidden="true" />
          {t('open')}
          <input
            type="file"
            accept=".md,.markdown,.txt"
            hidden
            onChange={(event) => {
              const file = event.currentTarget.files?.[0]
              event.currentTarget.value = ''
              if (file) void load(file)
            }}
          />
        </label>
        <button type="button" className={styles.ghost} onClick={() => setValue(t('sample'))}>
          {t('trySample')}
        </button>

        {/* Four presets, all visible: a <select> would hide three of them behind
            a click for a choice this small. */}
        <div className={styles.segment} role="group" aria-label={t('style')}>
          {TEMPLATE_IDS.map((id) => (
            <button
              key={id}
              type="button"
              className={`${styles.segmentButton} ${template === id ? styles.isActive : ''}`}
              aria-pressed={template === id}
              onClick={() => setTemplate(id)}
            >
              {t(id)}
            </button>
          ))}
        </div>

        <span className={styles.privacy}>
          <LockKeyhole size={13} aria-hidden="true" />
          {t('privacy')}
        </span>
      </div>

      {error && (
        <p className={styles.error} role="alert">
          {error}
        </p>
      )}

      <div className={styles.cols}>
        <section className={styles.pane}>
          <div className={styles.paneHead}>
            <span>Markdown</span>
            {rendering && <span className={styles.placeholder}>{t('rendering')}</span>}
          </div>
          <textarea
            className={styles.editor}
            value={value}
            maxLength={MAX_BYTES}
            spellCheck={false}
            placeholder={t('placeholder')}
            aria-label="Markdown"
            onChange={(event) => {
              setValue(event.target.value)
              setError('')
            }}
          />
        </section>

        <section className={styles.pane}>
          <div className={styles.paneHead}>
            <span>{t('preview')}</span>
          </div>
          {!value.trim() && !rendering && <p className={styles.placeholder}>{t('placeholder')}</p>}
          <div ref={preview} className={styles.preview} aria-live="polite" />
        </section>
      </div>
    </div>
  )
}
