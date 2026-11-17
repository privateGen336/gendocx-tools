import { useEffect, useState } from 'react'
import { Code2, Droplets, FileCode2, FileImage, FileText, GitCompare, PenLine, ShieldCheck } from 'lucide-react'
import ImageToWord from './components/ImageToWord'
import MarkdownToWord from './components/MarkdownToWord'
import PdfToolbox from './components/PdfToolbox'
import WatermarkTool from './components/WatermarkTool'
import WordCompare from './components/WordCompare'
import WordEditor from './components/WordEditor'
import { useLang } from './i18n'

type Tool = 'images-to-docx' | 'markdown-to-docx' | 'pdf' | 'watermark' | 'compare' | 'editor'

const TOOLS: Tool[] = ['images-to-docx', 'markdown-to-docx', 'pdf', 'watermark', 'compare', 'editor']

const UI = {
  zh: {
    'images-to-docx': '图片转 Word',
    'markdown-to-docx': 'Markdown 转 Word',
    pdf: 'PDF 工具箱',
    watermark: 'PDF 水印',
    compare: '文档对比',
    editor: 'Word 编辑器',
    online: '在线版',
    product: '完整产品',
    privacy: '文件不上传 · 全部在浏览器内处理',
    source: '源码',
    switchTo: 'Switch to English',
  },
  en: {
    'images-to-docx': 'Images to Word',
    'markdown-to-docx': 'Markdown to Word',
    pdf: 'PDF tools',
    watermark: 'PDF watermark',
    compare: 'Compare docs',
    editor: 'Word editor',
    online: 'Hosted version',
    product: 'Full product',
    privacy: 'No upload · every byte stays in this browser',
    source: 'Source',
    switchTo: '切换到中文',
  },
} as const

const ICONS = {
  'images-to-docx': FileImage,
  'markdown-to-docx': FileCode2,
  pdf: FileText,
  watermark: Droplets,
  compare: GitCompare,
  editor: PenLine,
} as const

function readTool(): Tool {
  const hash = window.location.hash.replace('#', '')
  return (TOOLS as string[]).includes(hash) ? (hash as Tool) : 'images-to-docx'
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

  const panels: Record<Tool, React.ReactNode> = {
    'images-to-docx': <ImageToWord />,
    'markdown-to-docx': <MarkdownToWord />,
    pdf: <PdfToolbox />,
    watermark: <WatermarkTool />,
    compare: <WordCompare />,
    editor: <WordEditor />,
  }

  return (
    <>
      <header className="app-header">
        <div className="app-header-inner">
          <span className="brand">
            <img className="brand-mark" src="./favicon.svg" alt="" width={26} height={26} aria-hidden="true" />
            <span className="brand-name">
              GenDocX <em>Tools</em>
            </span>
          </span>

          <nav className="tool-tabs" aria-label="Tools">
            {TOOLS.map((id) => {
              const Icon = ICONS[id]
              return (
                <button
                  key={id}
                  type="button"
                  className={tool === id ? 'tool-tab is-active' : 'tool-tab'}
                  aria-current={tool === id ? 'page' : undefined}
                  aria-label={copy[id]}
                  title={copy[id]}
                  onClick={() => select(id)}
                >
                  <Icon size={16} aria-hidden="true" />
                  <span>{copy[id]}</span>
                </button>
              )
            })}
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
              href="https://github.com/privateGen336/gendocx-tools"
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

      {panels[tool]}

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
