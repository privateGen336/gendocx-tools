import { useEffect, useState } from 'react'
import { Code2, FileImage, FileText, ShieldCheck } from 'lucide-react'
import ImageToWord from './components/ImageToWord'
import PdfToolbox from './components/PdfToolbox'
import { useLang } from './i18n'

type Tool = 'images-to-docx' | 'pdf'

const TOOLS: Tool[] = ['images-to-docx', 'pdf']

const UI = {
  zh: {
    images: '图片转 Word',
    pdf: 'PDF 工具箱',
    online: '在线版',
    product: '完整产品',
    privacy: '文件不上传 · 全部在浏览器内处理',
    source: '源码',
    switchTo: 'Switch to English',
  },
  en: {
    images: 'Images to Word',
    pdf: 'PDF tools',
    online: 'Hosted version',
    product: 'Full product',
    privacy: 'No upload · every byte stays in this browser',
    source: 'Source',
    switchTo: '切换到中文',
  },
} as const

function readTool(): Tool {
  const hash = window.location.hash.replace('#', '')
  return hash === 'pdf' ? 'pdf' : 'images-to-docx'
}

export function App() {
  const { lang, setLang } = useLang()
  const [tool, setTool] = useState<Tool>(readTool)
  const copy = UI[lang]

  useEffect(() => {
    const onHash = () => setTool(readTool())
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  const select = (next: Tool) => {
    setTool(next)
    // Keep the URL shareable without triggering a re-render loop.
    if (readTool() !== next) window.location.hash = next
  }

  return (
    <>
      <header className="app-header">
        <div className="app-header-inner">
          <span className="brand">
            <span className="brand-mark" aria-hidden="true">G</span>
            <span className="brand-name">
              GenDocX <em>Tools</em>
            </span>
          </span>

          <nav className="tool-tabs" aria-label="Tools">
            <button
              type="button"
              className={tool === 'images-to-docx' ? 'tool-tab is-active' : 'tool-tab'}
              aria-current={tool === 'images-to-docx' ? 'page' : undefined}
              aria-label={copy.images}
              title={copy.images}
              onClick={() => select('images-to-docx')}
            >
              <FileImage size={16} aria-hidden="true" />
              <span>{copy.images}</span>
            </button>
            <button
              type="button"
              className={tool === 'pdf' ? 'tool-tab is-active' : 'tool-tab'}
              aria-current={tool === 'pdf' ? 'page' : undefined}
              aria-label={copy.pdf}
              title={copy.pdf}
              onClick={() => select('pdf')}
            >
              <FileText size={16} aria-hidden="true" />
              <span>{copy.pdf}</span>
            </button>
          </nav>

          <div className="header-actions">
            <button
              type="button"
              className="ghost-button"
              onClick={() => setLang(lang === 'zh' ? 'en' : 'zh')}
              aria-label={copy.switchTo}
              title={copy.switchTo}
            >
              {lang === 'zh' ? 'EN' : '中文'}
            </button>
            <a
              className="ghost-button source-link"
              href="https://github.com/gendocx/gendocx-tools"
              target="_blank"
              rel="noreferrer noopener"
            >
              <Code2 size={15} aria-hidden="true" />
              <span>{copy.source}</span>
            </a>
            <a className="online-link" href="https://gendocx.com" target="_blank" rel="noreferrer noopener">
              {copy.online}
            </a>
          </div>
        </div>
      </header>

      {tool === 'images-to-docx' ? <ImageToWord /> : <PdfToolbox />}

      <footer className="app-footer">
        <p>
          <ShieldCheck size={14} aria-hidden="true" />
          {copy.privacy}
        </p>
        <a href="https://gendocx.com" target="_blank" rel="noreferrer noopener">
          {copy.product} · gendocx.com
        </a>
      </footer>
    </>
  )
}
