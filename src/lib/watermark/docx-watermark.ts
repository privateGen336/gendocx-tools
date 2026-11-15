import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { normalizeOptions, type WatermarkImage, type WatermarkOptions, type WatermarkPosition } from './watermark-options'

/**
 * DOCX 水印识别与移除。
 *
 * Word 的水印不是文档正文里的一段文字，而是塞在页眉里的一枚「绘图对象」：
 * 传统水印是 VML 的 WordArt（`<v:shape><v:textpath string="…"/>`），图片水印是
 * `<v:imagedata>` 加负 z-index，新版「浮层」则是 DrawingML 的 `<wp:anchor
 * behindDoc="1">`。它们在页眉里不可见地存在，所以只能按结构识别，不能按文字找。
 *
 * 这里只删「确定是水印」的节点，宁可漏也不误删：
 *   - WordArt 文本框（v:textpath）：页眉页脚里直接判定，正文里要求负 z-index
 *     或 id 带 watermark（正文里的艺术字可能只是排版）；
 *   - 图片水印：o:title 写明 watermark、id 带 watermark，或负 z-index；
 *   - DrawingML 浮层：必须 behindDoc="1"，且带文本框 / WordArt / 名为 watermark。
 *
 * 解析用 DOMParser 而不是正则：节点删除、命名空间前缀、实体转义都交给解析器，
 * 序列化后仍是 Word 认得的 XML。
 */

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const

/** 只扫水印可能出现的位置：正文、页眉、页脚。脚注里不会有水印。 */
const PART_PATTERN = /^word\/(document|header\d*|footer\d*)\.xml$/

export type DocxFindingKind = 'text' | 'image' | 'floating'
export type DocxPart = 'body' | 'header' | 'footer'

export interface DocxFinding {
  /** `part#序号`，同一份文件内稳定，扫描结果与移除操作共用。 */
  id: string
  part: DocxPart
  /** 具体分件名，如 `header1.xml`，用于说明「在哪」。 */
  partName: string
  kind: DocxFindingKind
  /** WordArt 的文字；图片水印可能是 o:title；没有就是空串。 */
  text: string
  /** 同一段文字在同一个分件里出现多枚（平铺水印）时合并成一条，这里记枚数。 */
  count?: number
}

export interface DocxScan {
  findings: DocxFinding[]
  /** 实际检查过的 XML 分件数，0 说明不是 .docx。 */
  parts: number
}

export const DOCX_EXTENSIONS = ['.docx'] as const
export const MAX_DOCX_BYTES = 60 * 1024 * 1024

export class DocxWatermarkError extends Error {
  constructor(public code: 'not-docx' | 'invalid-docx' | 'too-large' | 'empty-text' | 'invalid-image') {
    super(code)
  }
}

type Candidate = {
  node: Element
  kind: DocxFindingKind
  text: string
}

function parseXml(xml: string): Document | null {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) return null
  return doc
}

function serializeXml(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc)
  if (xml.startsWith('<?xml')) return xml
  // Word 期望这个声明；漏了它会拿默认编码重新解释整份 XML。
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${xml}`
}

function partOf(name: string): DocxPart {
  if (name.includes('header')) return 'header'
  if (name.includes('footer')) return 'footer'
  return 'body'
}

function attr(el: Element, ns: string, local: string): string {
  return el.getAttributeNS(ns, local) ?? ''
}

/** VML style 里的负 z-index 是「沉到文字下面」的信号，Logo 一般不会有。 */
function isBehind(style: string): boolean {
  return /z-index\s*:\s*-\s*\d/i.test(style)
}

function isWatermarkName(value: string): boolean {
  return /watermark|水印/i.test(value)
}

function matchVml(shape: Element, part: DocxPart): { kind: DocxFindingKind; text: string } | null {
  const textpath = shape.getElementsByTagNameNS(NS.v, 'textpath')[0]
  const imagedata = shape.getElementsByTagNameNS(NS.v, 'imagedata')[0]
  const id = shape.getAttribute('id') ?? ''
  const style = shape.getAttribute('style') ?? ''
  const behind = isBehind(style) || shape.getAttribute('behindDoc') === '1'
  const named = isWatermarkName(id)

  if (textpath) {
    // 正文里的艺术字可能只是标题装饰，必须有「沉底」或显式命名才动。
    if (part === 'body' && !behind && !named) return null
    const text = (textpath.getAttribute('string') ?? textpath.textContent ?? '').trim()
    return { kind: 'text', text }
  }

  if (imagedata) {
    const title = attr(imagedata, NS.o, 'title') || imagedata.getAttribute('o:title') || ''
    if (!behind && !named && !isWatermarkName(title)) return null
    return { kind: 'image', text: title.trim() }
  }

  return null
}

function matchDrawing(el: Element): { kind: DocxFindingKind; text: string } | null {
  const anchors = Array.from(el.getElementsByTagNameNS(NS.wp, 'anchor'))
  for (const anchor of anchors) {
    // 只有「浮在文字下面」的锚定对象才可能是水印；inline 图片是从正文插入的。
    if (anchor.getAttribute('behindDoc') !== '1') continue
    const docPr = anchor.getElementsByTagNameNS(NS.wp, 'docPr')[0]
    const name = `${docPr?.getAttribute('name') ?? ''} ${docPr?.getAttribute('descr') ?? ''}`
    const txbx = anchor.getElementsByTagNameNS(NS.wps, 'txbx')[0]
    const wordart = Array.from(anchor.getElementsByTagNameNS(NS.a, 'graphicData')).some((node) =>
      /wordart/i.test(node.getAttribute('uri') ?? ''),
    )
    if (!txbx && !wordart && !isWatermarkName(name)) continue
    const text = txbx ? (txbx.textContent ?? '').replace(/\s+/g, ' ').trim() : ''
    return { kind: 'text', text }
  }
  return null
}

function isContainer(el: Element): boolean {
  if (el.namespaceURI === NS.w && el.localName === 'pict') return true
  if (el.namespaceURI === NS.w && el.localName === 'drawing') return true
  // mc:AlternateContent 同时包着新版 drawing 与旧版 pict；整体删掉才不会留下另一份。
  if (el.namespaceURI === NS.mc && el.localName === 'AlternateContent') return true
  return false
}

/** 按文档顺序收集水印容器；被外层容器罩住的不重复收集。 */
function collectNodes(xml: string, part: DocxPart): { doc: Document; nodes: Element[]; matches: Candidate[] } | null {
  const doc = parseXml(xml)
  if (!doc) return null

  const nodes: Element[] = []
  const matches: Candidate[] = []
  const taken: Element[] = []

  for (const el of Array.from(doc.getElementsByTagName('*'))) {
    if (!isContainer(el)) continue
    if (taken.some((root) => root.contains(el))) continue

    const match = el.localName === 'drawing' ? matchDrawing(el) : matchVmlContainer(el, part)
    if (!match) continue

    taken.push(el)
    nodes.push(el)
    matches.push({ node: el, ...match })
  }

  return { doc, nodes, matches }
}

function collect(xml: string, part: DocxPart): Candidate[] {
  return collectNodes(xml, part)?.matches ?? []
}

/** w:pict 直接匹配；mc:AlternateContent 里的 pict / drawing 都交给子节点规则。 */
function matchVmlContainer(el: Element, part: DocxPart): { kind: DocxFindingKind; text: string } | null {
  if (el.localName === 'pict') {
    const shape = el.getElementsByTagNameNS(NS.v, 'shape')[0]
    return shape ? matchVml(shape, part) : null
  }

  if (el.localName === 'AlternateContent') {
    const direct = el.getElementsByTagNameNS(NS.v, 'shape')[0]
    if (direct) {
      const match = matchVml(direct, part)
      if (match) return match
    }
    return matchDrawing(el)
  }

  return null
}

function readZip(bytes: Uint8Array): Record<string, Uint8Array> {
  if (bytes.byteLength > MAX_DOCX_BYTES) throw new DocxWatermarkError('too-large')
  try {
    return unzipSync(bytes)
  } catch {
    // 旧版 .doc 是二进制，不是 zip；这里统一报「不是 docx」。
    throw new DocxWatermarkError('not-docx')
  }
}

export function scanDocx(bytes: Uint8Array): DocxScan {
  const zip = readZip(bytes)
  if (!zip['word/document.xml']) throw new DocxWatermarkError('not-docx')

  const findings: DocxFinding[] = []
  let parts = 0

  for (const [name, data] of Object.entries(zip)) {
    if (!PART_PATTERN.test(name)) continue
    parts += 1
    const xml = strFromU8(data)
    const part = partOf(name)
    // 平铺水印是同一段文字的多枚形状，列表里合成一条，只标枚数。
    const groups = new Map<string, { first: Candidate; indexes: number[] }>()
    collect(xml, part).forEach((item, index) => {
      const key = `${item.kind}\u0000${item.text}`
      const group = groups.get(key)
      if (group) group.indexes.push(index)
      else groups.set(key, { first: item, indexes: [index] })
    })

    for (const { first, indexes } of groups.values()) {
      findings.push({
        // 一条 finding 可能对应多枚形状，删除时按逗号列表一起回收。
        id: `${name}#${indexes.join(',')}`,
        part,
        partName: name.replace(/^word\//, ''),
        kind: first.kind,
        text: first.text,
        count: indexes.length,
      })
    }
  }

  return { findings, parts }
}

export interface DocxRemoval {
  bytes: Uint8Array
  removed: number
}

export function removeDocxWatermarks(bytes: Uint8Array, ids: string[]): DocxRemoval {
  if (ids.length === 0) return { bytes, removed: 0 }
  const zip = readZip(bytes)
  const wanted = new Map<string, Set<number>>()

  for (const id of ids) {
    const hash = id.lastIndexOf('#')
    if (hash < 0) continue
    const part = id.slice(0, hash)
    // 平铺水印一条 finding 对应多枚，id 里是逗号分隔的下标列表。
    const indexes = id
      .slice(hash + 1)
      .split(',')
      .map((piece) => Number(piece))
      .filter((index) => Number.isInteger(index))
    if (indexes.length === 0) continue
    const set = wanted.get(part) ?? new Set<number>()
    for (const index of indexes) set.add(index)
    wanted.set(part, set)
  }

  let removed = 0

  for (const [part, indexes] of wanted) {
    const data = zip[part]
    if (!data) continue
    const parsed = collectNodes(strFromU8(data), partOf(part))
    if (!parsed) continue

    let removedHere = 0
    for (const index of indexes) {
      const node = parsed.nodes[index]
      if (!node?.parentNode) continue
      node.parentNode.removeChild(node)
      removedHere += 1
    }

    if (removedHere === 0) continue
    zip[part] = strToU8(serializeXml(parsed.doc))
    removed += removedHere
  }

  if (removed === 0) return { bytes, removed: 0 }
  return { bytes: zipSync(zip, { level: 6 }), removed }
}

/* ────────────────────────── 添加水印 ──────────────────────────
 *
 * 和水印移除共用一套结构认知：Word 的水印就是页眉里的一枚 VML WordArt，
 * 所以「添加」= 往页眉追加一枚 `v:shape[type=_x0000_t136]`，和 Word 自带
 * 水印生成的结构一致，打开后仍能在「设计 → 水印」里看到并修改。
 *
 * 麻烦的是文件可能根本没有页眉：这时要新建 header 分件、注册 Content_Types
 * 与关系，并在每个还没有默认页眉的 sectPr 上挂 headerReference。少任何一步
 * Word 都会判定文件损坏。已有默认页眉则直接追加，重复添加先删掉自己上次写的
 * 那枚，保证可反复调用。
 */

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/'
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
const HEADER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

/** Word 生成艺术字时写的标准 shapetype；缺了它 Word 不肯渲染。 */
const WORDART_SHAPETYPE = `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles></v:shapetype>`

export type DocxWatermarkTone = 'solid' | 'outline'

/** UI 只传改动过的字段，缺省值由 normalizeOptions 补齐。 */
export type DocxAddOptions = Partial<WatermarkOptions> & { text: string }

export interface DocxAddition {
  bytes: Uint8Array
  /** 水印写进了哪个页眉分件，便于说明「加在哪」。 */
  part: string
}

const DEFAULT_COLOR = '#7a7a7a'
/** 文档里读不到 w:pgSz 时按 A4 铺版，单位 pt。 */
const A4_PAGE = { w: 595.3, h: 841.9 }

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function parseHexColor(value: string): [number, number, number] | null {
  const match = /^#?([0-9a-f]{6})$/i.exec(value.trim())
  if (!match) return null
  const int = Number.parseInt(match[1], 16)
  return [(int >> 16) & 255, (int >> 8) & 255, int & 255]
}

/** 把颜色按不透明度混到白底：得到的实色在各家 Office 里表现一致。 */
function fadeColor(color: string, opacity: number): string {
  const rgb = parseHexColor(color) ?? parseHexColor(DEFAULT_COLOR)!
  const alpha = Math.min(1, Math.max(0.02, opacity))
  const channel = (value: number) =>
    Math.round(255 - (255 - value) * alpha)
      .toString(16)
      .padStart(2, '0')
  return `#${rgb.map(channel).join('')}`
}

type WatermarkBox = {
  /** 相对页面左上角的偏移，单位 pt。 */
  x: number
  y: number
  w: number
  h: number
  /** 单枚水印仍用「页面居中」定位，和 Word 自带水印一致。 */
  centered?: boolean
}

/** 每页枚数 → 网格：2 枚上下两行，4 / 9 / 16 走 n×n，其余按近方形补齐。 */
function watermarkGrid(count: number): { rows: number; columns: number } {
  if (count <= 2) return { rows: Math.max(1, count), columns: 1 }
  const columns = Math.max(1, Math.min(4, Math.round(Math.sqrt(count))))
  return { rows: Math.ceil(count / columns), columns }
}

/** 图片水印按原图比例定高，文字水印保持 468×117 的横幅比例。 */
function boxAspect(options: WatermarkOptions): number {
  if (options.kind === 'image' && options.image?.width) {
    return Math.min(1.2, options.image.height / options.image.width)
  }
  return 0.26
}

/** 单枚居中/按九宫格定位；多枚把页面按网格均分，每个格子中间放一枚。 */
function watermarkBoxes(options: WatermarkOptions, page: { w: number; h: number }): WatermarkBox[] {
  if (options.count <= 1) {
    const w = 468 * options.scale
    const h = Math.min(w * boxAspect(options), page.h * 0.6)
    return [{ x: 0, y: 0, w, h, centered: true }]
  }

  const { rows, columns } = watermarkGrid(options.count)
  const cellW = page.w / columns
  const cellH = page.h / rows
  // 2 枚上下排时一行只有一枚，横幅不能按整页宽撑满；gap 越小越稀疏。
  const w = (columns === 1 ? Math.min(468, page.w * 0.78) : cellW * 0.9) * options.gap * options.scale
  const h = Math.min(cellH * 0.5, w * boxAspect(options))
  const boxes: WatermarkBox[] = []
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      boxes.push({
        x: column * cellW + (cellW - w) / 2,
        y: row * cellH + (cellH - h) / 2,
        w,
        h,
      })
    }
  }
  return boxes
}

function round(value: number): string {
  return Number(value.toFixed(2)).toString()
}

/** Word 的 mso-position 用 center 表示「中」，不是 middle。 */
function positionStyle(position: WatermarkPosition): string {
  const horizontal = position.horizontal
  const vertical = position.vertical === 'middle' ? 'center' : position.vertical
  return `margin-left:0;margin-top:0;mso-position-horizontal:${horizontal};mso-position-horizontal-relative:margin;mso-position-vertical:${vertical};mso-position-vertical-relative:margin`
}

function watermarkShape(
  text: string,
  options: WatermarkOptions,
  box: WatermarkBox,
  index: number,
  imageRelId: string,
): string {
  const rotation = options.angle === 0 ? 0 : 360 - options.angle
  const paint = fadeColor(options.color, options.opacity)
  const outline = options.tone === 'outline'
  const level = options.weight
  const fontWeight = [300, 500, 700, 900][level]
  const strokeWidth = (box.centered ? 1 : 0.75) + (level - 1) * 0.5
  const isImage = options.kind === 'image' && Boolean(imageRelId)

  const position = box.centered
    ? positionStyle(options.position)
    : `margin-left:${round(box.x)}pt;margin-top:${round(box.y)}pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:page;mso-position-vertical:absolute;mso-position-vertical-relative:page`

  const paintAttrs = isImage
    ? 'stroked="f"'
    : outline
      ? `fillcolor="none" stroked="t" strokecolor="${paint}" strokeweight="${strokeWidth}pt"`
      : `fillcolor="${paint}" stroked="f"`

  // 同一份文档里 shape id 必须唯一，否则 Word 只认第一枚。
  const id = index === 0 ? 'GDXWatermark' : `GDXWatermark${index + 1}`
  // 图片水印用 Office 内置的矩形图片 shape 类型，Word 才肯渲染 v:imagedata。
  const wordArtType = isImage ? ' type="#_x0000_t75"' : ' type="#_x0000_t136"'
  const zIndex = options.layer === 'below' ? -251655168 : 251658240
  const behindDoc = options.layer === 'below' ? ' behindDoc="1"' : ''
  const inner = isImage
    ? `<v:imagedata r:id="${imageRelId}" o:title="GDXWatermark"/>`
    : `<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt;font-weight:${fontWeight}" string="${escapeXml(text)}"/>`

  return `<v:shape id="${id}" o:spid="_x0000_s${2049 + index}"${wordArtType} style="position:absolute;${position};width:${round(box.w)}pt;height:${round(box.h)}pt;rotation:${rotation};z-index:${zIndex}" o:allowincell="f"${behindDoc} ${paintAttrs}>${inner}</v:shape>`
}

function watermarkParagraph(
  text: string,
  options: WatermarkOptions,
  page: { w: number; h: number },
  withShapetype: boolean,
  imageRelId: string,
): string {
  // 多枚水印放在同一个段落里：删除时按段落整段回收，不会留下空行。
  const runs = watermarkBoxes(options, page).map((box, index) => {
    const shapetype = index === 0 && withShapetype && options.kind !== 'image' ? WORDART_SHAPETYPE : ''
    return `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>${shapetype}${watermarkShape(text, options, box, index, imageRelId)}</w:pict></w:r>`
  })
  return `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>${runs.join('')}</w:p>`
}

/** 平铺时要知道页面多大；读第一个 sectPr 的 w:pgSz（单位 twip，/20 得 pt）。 */
function documentPageSize(xml: string): { w: number; h: number } {
  const doc = parseXml(xml)
  const size = doc?.getElementsByTagNameNS(NS.w, 'pgSz')[0]
  if (!size) return A4_PAGE
  const w = Number(size.getAttributeNS(NS.w, 'w') || size.getAttribute('w:w') || 0)
  const h = Number(size.getAttributeNS(NS.w, 'h') || size.getAttribute('w:h') || 0)
  if (!w || !h) return A4_PAGE
  return { w: w / 20, h: h / 20 }
}

function ensureNamespace(root: Element, prefix: string, uri: string) {
  if (root.getAttributeNS(XMLNS_NS, prefix)) return
  root.setAttributeNS(XMLNS_NS, `xmlns:${prefix}`, uri)
}

/** 页眉里已有艺术字 shapetype 时不能再插一份，否则重复定义。 */
function hasWordArtShapetype(xml: string): boolean {
  return xml.includes('id="_x0000_t136"')
}

function removeOwnWatermarks(root: Element) {
  const parentElement = (node: Element | null): Element | null =>
    node && node.parentNode && node.parentNode.nodeType === 1 ? (node.parentNode as Element) : null

  for (const shape of Array.from(root.getElementsByTagNameNS(NS.v, 'shape'))) {
    if (!(shape.getAttribute('id') ?? '').startsWith('GDXWatermark')) continue
    // 逐层往上收到 w:p：留下空的 pict/run 会在页眉里多出一个空行。
    const pict = parentElement(shape)
    const run = parentElement(pict)
    const paragraph = parentElement(run)
    const target = paragraph?.localName === 'p' ? paragraph : pict ?? shape
    target.parentNode?.removeChild(target)
  }
}

function appendWatermarkToHeader(xml: string, paragraph: string): string {
  const doc = parseXml(xml)
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const root = doc.documentElement

  ensureNamespace(root, 'v', NS.v)
  ensureNamespace(root, 'o', NS.o)
  ensureNamespace(root, 'r', NS.r)
  ensureNamespace(root, 'w10', 'urn:schemas-microsoft-com:office:word')

  removeOwnWatermarks(root)

  const wrapper = parseXml(
    `<w:body xmlns:w="${NS.w}" xmlns:v="${NS.v}" xmlns:o="${NS.o}">${paragraph}</w:body>`,
  )
  const para = wrapper?.documentElement.firstElementChild
  if (!para) throw new DocxWatermarkError('invalid-docx')
  root.appendChild(doc.importNode(para, true))
  return serializeXml(doc)
}

function newHeaderXml(paragraph: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr xmlns:w="${NS.w}" xmlns:v="${NS.v}" xmlns:o="${NS.o}" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:r="${NS.r}">${paragraph}</w:hdr>`
}

function headerTargets(zip: Record<string, Uint8Array>): Map<string, string> {
  const relPath = 'word/_rels/document.xml.rels'
  const data = zip[relPath]
  if (!data) return new Map()
  const doc = parseXml(strFromU8(data))
  if (!doc) return new Map()

  const byId = new Map<string, string>()
  for (const rel of Array.from(doc.documentElement.getElementsByTagNameNS(REL_NS, 'Relationship'))) {
    const id = rel.getAttribute('Id')
    const target = rel.getAttribute('Target')
    const type = rel.getAttribute('Type') ?? ''
    if (id && target && type.endsWith('/header')) byId.set(id, target)
  }
  return byId
}

/** 找当前文档「默认页眉」实际指向的分件；没有就返回 null。 */
function defaultHeaderPart(zip: Record<string, Uint8Array>): string | null {
  const docData = zip['word/document.xml']
  if (!docData) return null
  const doc = parseXml(strFromU8(docData))
  if (!doc) return null
  const rels = headerTargets(zip)

  for (const ref of Array.from(doc.getElementsByTagNameNS(NS.w, 'headerReference'))) {
    const type = ref.getAttributeNS(NS.w, 'type') || ref.getAttribute('w:type') || 'default'
    if (type !== 'default') continue
    const id = ref.getAttributeNS(NS.r, 'id') || ref.getAttribute('r:id') || ''
    const target = rels.get(id)
    if (!target) continue
    const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`
    if (zip[path]) return path
  }
  return null
}

function nextHeaderPart(zip: Record<string, Uint8Array>): string {
  let max = 0
  for (const name of Object.keys(zip)) {
    const match = /^word\/header(\d+)\.xml$/.exec(name)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return `word/header${max + 1}.xml`
}

function addContentTypeOverride(zip: Record<string, Uint8Array>, partName: string) {
  const data = zip['[Content_Types].xml']
  if (!data) throw new DocxWatermarkError('invalid-docx')
  const doc = parseXml(strFromU8(data))
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const root = doc.documentElement
  const part = `/${partName}`
  const exists = Array.from(root.getElementsByTagNameNS(CT_NS, 'Override')).some(
    (node) => node.getAttribute('PartName') === part,
  )
  if (exists) return

  const override = doc.createElementNS(CT_NS, 'Override')
  override.setAttribute('PartName', part)
  override.setAttribute('ContentType', HEADER_CONTENT_TYPE)
  root.appendChild(override)
  zip['[Content_Types].xml'] = strToU8(serializeXml(doc))
}

function addHeaderRelationship(zip: Record<string, Uint8Array>, target: string): string {
  const relPath = 'word/_rels/document.xml.rels'
  const data = zip[relPath]
  const doc = data
    ? parseXml(strFromU8(data))
    : parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"/>`)
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const root = doc.documentElement

  const used = new Set(
    Array.from(root.getElementsByTagNameNS(REL_NS, 'Relationship')).map((rel) => rel.getAttribute('Id') ?? ''),
  )
  let index = 1
  while (used.has(`rId${index}`)) index += 1
  const id = `rId${index}`

  const rel = doc.createElementNS(REL_NS, 'Relationship')
  rel.setAttribute('Id', id)
  rel.setAttribute('Type', HEADER_REL)
  rel.setAttribute('Target', target)
  root.appendChild(rel)
  zip[relPath] = strToU8(serializeXml(doc))
  return id
}

/** 把 headerReference 挂到每个还没有默认页眉的节上；没有节就补一个。 */
function attachHeaderReference(xml: string, relationshipId: string): string {
  const doc = parseXml(xml)
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const root = doc.documentElement
  ensureNamespace(root, 'r', NS.r)

  const sections = Array.from(doc.getElementsByTagNameNS(NS.w, 'sectPr'))
  const targets: Element[] = []

  if (sections.length === 0) {
    const body = doc.getElementsByTagNameNS(NS.w, 'body')[0]
    if (!body) throw new DocxWatermarkError('invalid-docx')
    const section = doc.createElementNS(NS.w, 'w:sectPr')
    body.appendChild(section)
    targets.push(section)
  } else {
    for (const section of sections) {
      const hasDefault = Array.from(section.getElementsByTagNameNS(NS.w, 'headerReference')).some(
        (ref) => (ref.getAttributeNS(NS.w, 'type') || ref.getAttribute('w:type') || 'default') === 'default',
      )
      if (!hasDefault) targets.push(section)
    }
  }

  for (const section of targets) {
    const ref = doc.createElementNS(NS.w, 'w:headerReference')
    // 序列化只保留前缀，命名空间由根节点统一声明，因此带前缀的属性用 setAttribute 即可。
    ref.setAttribute('w:type', 'default')
    ref.setAttribute('r:id', relationshipId)
    section.insertBefore(ref, section.firstChild)
  }

  return serializeXml(doc)
}

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** 让 [Content_Types].xml 认得图片扩展名；已有的 Default 不重复写。 */
function ensureDefaultExtension(zip: Record<string, Uint8Array>, extension: string, contentType: string) {
  const data = zip['[Content_Types].xml']
  if (!data) throw new DocxWatermarkError('invalid-docx')
  const doc = parseXml(strFromU8(data))
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const exists = Array.from(doc.documentElement.getElementsByTagNameNS(CT_NS, 'Default')).some(
    (node) => (node.getAttribute('Extension') ?? '').toLowerCase() === extension,
  )
  if (exists) return
  const entry = doc.createElementNS(CT_NS, 'Default')
  entry.setAttribute('Extension', extension)
  entry.setAttribute('ContentType', contentType)
  doc.documentElement.appendChild(entry)
  zip['[Content_Types].xml'] = strToU8(serializeXml(doc))
}

/** 图片水印：写 media 分件 + 注册扩展名 + 在页眉 rels 里挂 image 关系，返回 r:id。 */
function addImageRelationship(zip: Record<string, Uint8Array>, headerPart: string, image: WatermarkImage): string {
  const extension = image.mime === 'image/jpeg' ? 'jpeg' : 'png'
  let index = 1
  let mediaName = `gdx-watermark${index}.${extension}`
  while (zip[`word/media/${mediaName}`]) {
    index += 1
    mediaName = `gdx-watermark${index}.${extension}`
  }
  zip[`word/media/${mediaName}`] = image.bytes
  ensureDefaultExtension(zip, extension, image.mime)

  const relPath = `word/_rels/${headerPart.replace(/^word\//, '')}.rels`
  const data = zip[relPath]
  const doc = data
    ? parseXml(strFromU8(data))
    : parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"/>`)
  if (!doc) throw new DocxWatermarkError('invalid-docx')
  const root = doc.documentElement
  const used = new Set(
    Array.from(root.getElementsByTagNameNS(REL_NS, 'Relationship')).map((rel) => rel.getAttribute('Id') ?? ''),
  )
  let n = 1
  while (used.has(`rId${n}`)) n += 1
  const id = `rId${n}`
  const rel = doc.createElementNS(REL_NS, 'Relationship')
  rel.setAttribute('Id', id)
  rel.setAttribute('Type', IMAGE_REL)
  rel.setAttribute('Target', `media/${mediaName}`)
  root.appendChild(rel)
  zip[relPath] = strToU8(serializeXml(doc))
  return id
}

export function addDocxWatermark(bytes: Uint8Array, input: DocxAddOptions): DocxAddition {
  const zip = readZip(bytes)
  const documentXml = zip['word/document.xml']
  if (!documentXml) throw new DocxWatermarkError('not-docx')

  // 明确要求图片水印却没给图片时直接报错，别静默退回文字模式再报「文字为空」。
  if (input.kind === 'image' && !input.image?.bytes.length) throw new DocxWatermarkError('invalid-image')
  const options = normalizeOptions(input)
  if (options.kind === 'text' && !options.text) throw new DocxWatermarkError('empty-text')

  const page = documentPageSize(strFromU8(documentXml))
  const existing = defaultHeaderPart(zip)
  const part = existing ?? nextHeaderPart(zip)
  const imageRelId = options.kind === 'image' && options.image ? addImageRelationship(zip, part, options.image) : ''
  const withShapetype = !existing || !hasWordArtShapetype(strFromU8(zip[existing]))
  const paragraph = watermarkParagraph(options.text, options, page, withShapetype, imageRelId)

  if (existing) {
    zip[existing] = strToU8(appendWatermarkToHeader(strFromU8(zip[existing]), paragraph))
  } else {
    zip[part] = strToU8(newHeaderXml(paragraph))
    addContentTypeOverride(zip, part)
    const relationshipId = addHeaderRelationship(zip, part.replace(/^word\//, ''))
    zip['word/document.xml'] = strToU8(attachHeaderReference(strFromU8(documentXml), relationshipId))
  }

  return { bytes: zipSync(zip, { level: 6 }), part: part.replace(/^word\//, '') }
}
