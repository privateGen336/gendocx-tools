import markdownDocx, { Packer, type IMarkdownTheme } from 'markdown-docx'
import JSZip from 'jszip'

export type TemplateId = 'standard' | 'compact' | 'report' | 'official'

/**
 * Shared typography. The four presets differ in size, leading and margins, not
 * in colour or typeface lineage, so those live here once.
 *
 * Sizes are in points — the library converts them to docx half-points. Fonts
 * are written as ascii + hAnsi + eastAsia families; 宋体 / 黑体 / 仿宋 ship with
 * every Chinese-capable Windows and macOS install, so Word does not fall back
 * to Calibri mid-paragraph.
 */
const BASE: Partial<IMarkdownTheme> = {
  bodyFont: '宋体',
  headingFont: '黑体',
  codeFont: 'Consolas',
  heading1: '000000',
  heading2: '000000',
  heading3: '000000',
  heading4: '000000',
  heading5: '000000',
  heading6: '000000',
  link: '0563C1',
  linkUnderline: true,
  code: '1F2328',
  codespan: 'A31515',
  blockquote: '595959',
  blockquoteBackground: 'F7F7F7',
  codeBackground: 'F5F5F5',
  border: 'BFBFBF',
  hr: 'D0D0D0',
  del: 'C00000',
  html: '1F2328',
  tableHeaderBackground: 'F2F2F2',
  imageHorizontalAlign: 'center',
}

const TEMPLATES: Record<TemplateId, Partial<IMarkdownTheme>> = {
  /** Everyday business document: 12pt body, 1.5 leading, Word's 1-inch margins. */
  standard: {
    ...BASE,
    bodySize: 12,
    heading1Size: 22,
    heading2Size: 16,
    heading3Size: 14,
    heading4Size: 13,
    heading5Size: 12,
    heading6Size: 12,
    codeSize: 10,
    lineSpacing: 1.5,
    spaceSize: 8,
    margin: '2.54cm',
  },
  /**
   * Long internal material — specs, meeting notes, anything read on screen and
   * scrolled. Smaller type, tighter leading, narrow margins, and blank spacer
   * paragraphs dropped so the page holds more without looking cramped.
   */
  compact: {
    ...BASE,
    bodySize: 11,
    heading1Size: 18,
    heading2Size: 15,
    heading3Size: 13,
    heading4Size: 12,
    heading5Size: 11,
    heading6Size: 11,
    codeSize: 9,
    lineSpacing: 1.3,
    spaceSize: 4,
    margin: '2cm',
    collapseEmptyLines: true,
  },
  /** Proposals and reports: larger headings and more air around paragraphs. */
  report: {
    ...BASE,
    bodySize: 12,
    heading1Size: 24,
    heading2Size: 18,
    heading3Size: 15,
    heading4Size: 13,
    heading5Size: 12,
    heading6Size: 12,
    codeSize: 10,
    lineSpacing: 1.75,
    spaceSize: 12,
    margin: '2.5cm',
  },
  /**
   * Chinese official-document layout, following GB/T 9704 conventions: 仿宋 body
   * at 三号 (16pt), 黑体 headings, 28pt-equivalent leading, no space between
   * paragraphs (the indent does that job in print), and the wide 37/26/35/28 mm
   * margins that leave room for the red header and the folio.
   *
   * Table headers lose their fill here because printed official documents use
   * ruled tables, not shaded ones.
   */
  official: {
    ...BASE,
    bodyFont: '仿宋',
    bodySize: 16,
    heading1Size: 22,
    heading2Size: 16,
    heading3Size: 16,
    heading4Size: 16,
    heading5Size: 16,
    heading6Size: 16,
    codeSize: 12,
    lineSpacing: 1.75,
    spaceSize: 0,
    margin: '3.7cm 2.6cm 3.5cm 2.8cm',
    tableHeaderBackground: 'FFFFFF',
    collapseEmptyLines: true,
  },
}

export const TEMPLATE_IDS = ['standard', 'compact', 'report', 'official'] as const

/**
 * Markdown → .docx, straight from the AST.
 *
 * This replaces a hand-written HTML→docx walker: that one rendered Markdown to
 * DOM first and then guessed at semantics from tag names, so code-fence
 * languages, table header cells and nested list levels were lost on the way.
 * Here `marked` produces tokens and the library maps them directly.
 *
 * Kept as a dynamic import at the call site — the converter and KaTeX together
 * are the largest thing this page ships, and a visitor who only reads the
 * placeholder never needs them.
 */
export async function markdownToDocxBlob(markdown: string, template: TemplateId): Promise<Blob> {
  const theme = TEMPLATES[template]
  const doc = await markdownDocx(markdown, {
    theme,
    gfm: true,
    math: { engine: 'katex' },
    // Images are fetched by the library's browser adapter. A broken image URL
    // should not abort the whole conversion, so failures are left to the
    // adapter rather than thrown here.
    ignoreImage: false,
  })
  return finalize(await Packer.toBlob(doc), theme, template)
}

/**
 * Patch the packed package in one pass: give tables visible borders, and
 * rewrite the `Md*` styles so Word and the browser preview agree.
 *
 * Both fixes work on XML because the objects that need changing are constructed
 * inside the library and never exposed.
 */
async function finalize(blob: Blob, theme: Partial<IMarkdownTheme>, template: TemplateId): Promise<Blob> {
  const zip = await JSZip.loadAsync(blob)

  const docPart = zip.file('word/document.xml')
  if (docPart) {
    let docXml = await docPart.async('string')
    docXml = addTableBorders(docXml, String(theme.border ?? 'BFBFBF'))
    docXml = pushRunStyles(docXml, theme)
    if (template === 'official') docXml = addFirstLineIndent(docXml)
  zip.file('word/document.xml', docXml)
  }

  const stylesPart = zip.file('word/styles.xml')
  if (stylesPart) {
    const stylesXml = await stylesPart.async('string')
    zip.file('word/styles.xml', applyTheme(stylesXml, theme))
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })
}

/**
 * `markdown-docx` builds its tables with a `tblStyle` reference and no
 * `tblBorders`, so the grid only exists if the reader honours the table style —
 * Word does, in-browser renderers often do not. Writing the borders explicitly
 * makes the export and the preview agree.
 */
function addTableBorders(xml: string, color: string): string {
  if (!xml.includes('<w:tblPr>')) return xml
  if (xml.includes('<w:tblBorders>')) return xml

  const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
  const borders = `<w:tblBorders>${sides
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${color}"/>`)
    .join('')}</w:tblBorders>`

  // `<w:tblPr>` never nests inside itself, so a lazy match per table is safe.
  return xml.replace(/<w:tblPr>([\s\S]*?)<\/w:tblPr>/g, (match, inner: string) =>
    inner.includes('<w:tblBorders>') ? match : `<w:tblPr>${inner}${borders}</w:tblPr>`
  )
}

/**
 * Copy each paragraph style's type settings onto its runs.
 *
 * docx-preview — the renderer behind the live preview — resolves a paragraph's
 * `pStyle` for layout but takes font family and size from the run, falling back
 * to document defaults. Word reads the style table; the preview does not. A
 * document styled only through `styles.xml` therefore looks different in the
 * two, which defeats the point of previewing the real file. Writing family and
 * size onto every run removes the discrepancy however the file is opened.
 *
 * Only family and size are pushed. Bold, italic and colour that a run already
 * carries stay as they are, so emphasis and links keep their own weight.
 */
function pushRunStyles(xml: string, theme: Partial<IMarkdownTheme>): string {
  const half = (pt: number) => Math.round(pt * 2)
  const fonts = (name?: string) =>
    name ? `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}"/>` : ''
  const sizes = (pt: number) => {
    const value = half(pt)
    return `<w:sz w:val="${value}"/><w:szCs w:val="${value}"/>`
  }
  const asRPr = (font: string, size: string) => `<w:rPr>${font}${size}</w:rPr>`

  const body = { font: fonts(theme.bodyFont), size: sizes(theme.bodySize ?? 12) }
  const mono = { font: fonts(theme.codeFont), size: sizes(theme.codeSize ?? 10) }
  const headingSizes = [
    theme.heading1Size ?? 22,
    theme.heading2Size ?? 16,
    theme.heading3Size ?? 14,
    theme.heading4Size ?? 13,
    theme.heading5Size ?? 12,
    theme.heading6Size ?? 12,
  ]

  const byStyle: Record<string, { font: string; size: string }> = {
    MdParagraph: body,
    MdText: body,
    MdListItem: body,
    MdTableCell: body,
    MdTable: body,
    MdDef: body,
    MdFootnote: body,
    MdTableHeader: body,
    MdBlockquote: body,
    MdCode: mono,
    MdCodespan: mono,
    MdTag: mono,
    MdHtml: mono,
  }
  headingSizes.forEach((size, index) => {
    byStyle[`MdHeading${index + 1}`] = { font: fonts(theme.headingFont), size: sizes(size) }
  })

  return xml.replace(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g, (paragraph, inner: string) => {
    const styleId = inner.match(/<w:pStyle w:val="([^"]+)"\/>/)?.[1]
    const style = styleId ? byStyle[styleId] : undefined
    if (!style) return paragraph

    const patched = inner.replace(/<w:r>([\s\S]*?)<\/w:r>/g, (run, runInner: string) => {
      const existing = runInner.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/)
      if (!existing) return `<w:r>${asRPr(style.font, style.size)}${runInner}</w:r>`

      let props = existing[1]
      if (!props.includes('<w:rFonts')) props += style.font
      if (!props.includes('<w:sz ')) props += style.size
      return `<w:r><w:rPr>${props}</w:rPr>${runInner.replace(existing[0], '')}</w:r>`
    })

    return paragraph.replace(inner, patched)
  })
}

/**
 * Two-character first-line indent, official template only.
 *
 * GB/T 9704 sets the body indent by character rather than by a gap between
 * paragraphs — which is why that template zeroes `spaceSize`. docx has no
 * character-relative indent, so the value is pinned to two 16pt characters
 * (640 twips), matching that template's body size.
 */
function addFirstLineIndent(xml: string): string {
  return xml.replace(
    /<w:pPr>([\s\S]*?<w:pStyle w:val="MdParagraph"\/>[\s\S]*?)<\/w:pPr>/g,
    (match, inner: string) =>
      inner.includes('<w:ind ') ? match : `<w:pPr>${inner}<w:ind w:firstLine="640"/></w:pPr>`
  )
}

/**
 * Rewrite the generated style sheet with our values.
 *
 * The library writes its own palette into the `Md*` styles — headings ship as
 * 2F5597 / 5B9BD5 / 44546A, code as 11pt Courier — and its theme options only
 * reach `docDefaults`. Two consequences: Word renders blue headings whatever
 * theme you pass, and a renderer that ignores the style table (docx-preview
 * among them) shows body defaults instead. Rewriting the styles fixes both at
 * once, because every consumer then reads the same numbers.
 */
function applyTheme(styles: string, theme: Partial<IMarkdownTheme>): string {
  const sz = (pt?: number, fallback = 12) => Math.round((pt ?? fallback) * 2)

  const rPr = (o: {
    font?: string
    size?: number
    color?: string
    bold?: boolean
    underline?: boolean
  }) => {
    const parts: string[] = []
    if (o.font) {
      parts.push(`<w:rFonts w:ascii="${o.font}" w:hAnsi="${o.font}" w:eastAsia="${o.font}" w:cs="${o.font}"/>`)
    }
    if (o.bold) parts.push('<w:b/><w:bCs/>')
    if (o.underline) parts.push('<w:u w:val="single"/>')
    if (o.color) parts.push(`<w:color w:val="${o.color}"/>`)
    if (o.size) parts.push(`<w:sz w:val="${sz(o.size)}"/><w:szCs w:val="${sz(o.size)}"/>`)
    return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
  }

  const body = { font: theme.bodyFont, size: theme.bodySize ?? 12, color: '000000' }
  const headingSizes = [
    theme.heading1Size ?? 22,
    theme.heading2Size ?? 16,
    theme.heading3Size ?? 14,
    theme.heading4Size ?? 13,
    theme.heading5Size ?? 12,
    theme.heading6Size ?? 12,
  ]

  const overrides: Record<string, string> = {
    // Body-family styles: no colour override beyond black, so emphasis runs
    // (strong/em) keep the weight they were given.
    MdParagraph: rPr(body),
    MdText: rPr(body),
    MdListItem: rPr(body),
    MdTableCell: rPr(body),
    MdTable: rPr(body),
    MdDef: rPr(body),
    MdFootnote: rPr(body),
    MdTableHeader: rPr({ ...body, bold: true }),

    MdCode: rPr({ font: theme.codeFont, size: theme.codeSize ?? 10, color: theme.code ?? '1F2328' }),
    MdCodespan: rPr({ font: theme.codeFont, color: theme.codespan ?? theme.code ?? '1F2328' }),
    MdTag: rPr({ font: theme.codeFont, color: theme.code ?? '1F2328' }),
    MdHtml: rPr({ font: theme.codeFont, color: theme.html ?? '1F2328' }),

    MdLink: rPr({ color: theme.link ?? '0563C1', underline: theme.linkUnderline !== false }),
    MdBlockquote: rPr({ color: theme.blockquote ?? '595959' }),
    MdDel: rPr({ color: theme.del ?? 'C00000' }),
  }

  headingSizes.forEach((size, index) => {
    overrides[`MdHeading${index + 1}`] = rPr({
      font: theme.headingFont,
      size,
      // Every level reads from heading1 so a template cannot end up with a
      // gradient of six different colours.
      color: theme.heading1 ?? '000000',
      bold: true,
    })
  })

  let out = styles
  for (const [id, xml] of Object.entries(overrides)) {
    if (!xml) continue
    const re = new RegExp(`(<w:style\\b[^>]*w:styleId="${id}"[^>]*>)([\\s\\S]*?)(</w:style>)`)
    out = out.replace(re, (_match, open: string, inner: string, close: string) => {
      const stripped = inner.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '')
      return `${open}${stripped}${xml}${close}`
    })
  }
  return out
}


