import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFPage,
  PDFObject,
  PDFRawStream,
  PDFStream,
  decodePDFRawStream,
  degrees,
} from 'pdf-lib'
import { zlibSync } from 'fflate'
import { normalizeOptions, parsePageRange, positionRatio, type WatermarkOptions } from './watermark-options'

/**
 * PDF 水印识别与移除。
 *
 * PDF 里没有「水印」这个对象，水印只是「画在页面上的东西」的一种意图。所以识别
 * 分三条线，全部要求用户确认后才动手：
 *
 *   1. 文字水印 —— 页面内容流（以及页面引用的 Form XObject）里 BT…ET 文本块，
 *      按字体 ToUnicode 解码后与候选文字比对；
 *   2. 表单水印 —— 重复出现的 Form XObject 里的文字即水印（Word 导出的 PDF
 *      常把艺术字水印放进 Form），移除页面上对它的 `/Name Do` 调用；
 *   3. 注释水印 —— /Watermark 与 /Stamp 注释，直接从页面的 /Annots 里摘掉。
 *
 * 三条线都不重建页面，只改内容流，因此书签、元数据、表单域都原样保留。
 *
 * 局限（UI 里也照实说）：扫描件或已轮廓化的文字是图像，没有可删的文本节点；
 * 用 Identity-H 又没有 ToUnicode 的字体解不出字，只能按「重复出现的原始字符串」
 * 统计，不会列进候选。
 */

export type PdfFindingKind = 'text' | 'form' | 'annot'

export interface PdfFinding {
  /** `text:…` / `form:…` / `annot:watermark`，扫描与移除共用。 */
  id: string
  kind: PdfFindingKind
  text: string
  /** 出现的页数。 */
  pages: number
  /** 出现的次数（同一页可能出现多次）。 */
  blocks: number
}

export interface PdfScan {
  pageCount: number
  findings: PdfFinding[]
  /** 解不出文字的文本块数量，用于提示「这份 PDF 有无法识别的字体」。 */
  undecoded: number
}

export class PdfWatermarkError extends Error {
  constructor(public code: 'encrypted' | 'invalid' | 'too-large' | 'empty') {
    super(code)
  }
}

export const PDF_EXTENSIONS = ['.pdf'] as const
export const MAX_PDF_BYTES = 80 * 1024 * 1024

/** 单页里最多扫这么多字符，防止异常内容流把浏览器拖死。 */
const MAX_CONTENT_CHARS = 12 * 1024 * 1024
/** 候选文字长度上限：水印通常很短，长段落不该出现在候选里。 */
const MAX_CANDIDATE_CHARS = 60

type FontInfo = {
  /** ToUnicode 映射；有它就能把自定义编码还原成文字。 */
  map?: Map<number, string>
  /** 每个字符码占几个字节。 */
  codeBytes: 1 | 2
  type0: boolean
  encoding?: string
}

type Fonts = Map<string, FontInfo>

/**
 * pdf-lib 的 `dict.lookup(key, Type)` 在键不存在时会抛错，而 PDF 里几乎所有
 * 可选键都可能不存在。这里统一用「先取、再判类型」的写法。
 */
function valueOf(dict: PDFDict | undefined, name: string) {
  if (!dict) return undefined
  const raw = dict.get(PDFName.of(name))
  if (raw === undefined) return undefined
  return dict.context.lookup(raw)
}

function dictOf(dict: PDFDict | undefined, name: string): PDFDict | undefined {
  const value = valueOf(dict, name)
  return value instanceof PDFDict ? value : undefined
}

function numberOf(dict: PDFDict | undefined, name: string): number | undefined {
  const value = valueOf(dict, name)
  return value instanceof PDFNumber ? value.asNumber() : undefined
}

function nameOf(dict: PDFDict | undefined, name: string): string | undefined {
  const value = valueOf(dict, name)
  return value instanceof PDFName ? value.decodeText() : undefined
}

type TextBlock = {
  /** 内容流里的 [start, end)，整段替换成空即删除。 */
  start: number
  end: number
  text: string
  /** 文本矩阵带旋转（水印常见，正文几乎没有）。 */
  rotated: boolean
  /** 半透明绘制（水印常见）。 */
  faint: boolean
  /** 解不出文字但确实画了字符串。 */
  undecoded: boolean
}

type FormUse = {
  name: string
  start: number
  end: number
}

type Token = {
  type: 'op' | 'num' | 'name' | 'str' | 'hex' | 'punct'
  text: string
  bytes?: Uint8Array
  start: number
  end: number
}

const WHITESPACE = new Set(['\u0000', '\t', '\n', '\f', '\r', ' '])
const DELIMITERS = new Set(['(', ')', '<', '>', '[', ']', '{', '}', '/', '%'])

const isDigit = (ch: string) => ch >= '0' && ch <= '9'
const isRegular = (ch: string) => !WHITESPACE.has(ch) && !DELIMITERS.has(ch)

function latin1(bytes: Uint8Array): string {
  let out = ''
  for (let i = 0; i < bytes.length; i += 1) out += String.fromCharCode(bytes[i])
  return out
}

function utf16be(bytes: Uint8Array): string {
  let out = ''
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    out += String.fromCharCode((bytes[i] << 8) | bytes[i + 1])
  }
  return out
}

/** 内容流是字节流，用 latin1 保证 1:1 往返；用 UTF-8 会改坏二进制片段。 */
function bytesToLatin1(bytes: Uint8Array): string {
  return latin1(bytes)
}

function latin1ToBytes(text: string): Uint8Array {
  const out = new Uint8Array(text.length)
  for (let i = 0; i < text.length; i += 1) out[i] = text.charCodeAt(i) & 0xff
  return out
}

/** 归一化：去掉所有空白再比对，避免字距调整把水印拆成几个字符串。 */
function normalize(text: string): string {
  return text.replace(/\s+/g, '')
}

function hexToNumber(hex: string): number {
  return parseInt(hex, 16) || 0
}

function hexToUnicode(hex: string): string {
  const clean = hex.replace(/[^0-9a-fA-F]/g, '')
  const bytes = new Uint8Array(clean.length >> 1)
  for (let i = 0; i < bytes.length; i += 1) bytes[i] = parseInt(clean.slice(i * 2, i * 2 + 2), 16)
  return utf16be(bytes)
}

type CMap = { map: Map<number, string>; codeBytes: 1 | 2 }

/** 解析 ToUnicode CMap 的 bfchar / bfrange / codespacerange 三段。 */
export function parseToUnicode(source: string): CMap | null {
  const map = new Map<number, string>()
  let codeBytes: 1 | 2 = 2

  const space = source.match(/begincodespacerange([\s\S]*?)endcodespacerange/)
  if (space) {
    const first = space[1].match(/<([0-9a-fA-F]+)>/)
    if (first) codeBytes = first[1].length / 2 >= 2 ? 2 : 1
  }

  for (const section of source.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const pair of section[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]*)>/g)) {
      map.set(hexToNumber(pair[1]), hexToUnicode(pair[2]))
    }
  }

  for (const section of source.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    const body = section[1]
    // 三种写法：<lo> <hi> <dst>、<lo> <hi> [<d1> <d2> …]、以及跨行排列。
    for (const line of body.matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(\[[\s\S]*?\]|<[0-9a-fA-F]*>)/g)) {
      const lo = hexToNumber(line[1])
      const hi = hexToNumber(line[2])
      const dst = line[3]
      if (dst.startsWith('[')) {
        const items = Array.from(dst.matchAll(/<([0-9a-fA-F]*)>/g))
        items.forEach((item, index) => map.set(lo + index, hexToUnicode(item[1])))
      } else {
        const base = dst.slice(1, -1)
        const baseBytes = new Uint8Array(base.length >> 1)
        for (let i = 0; i < baseBytes.length; i += 1) baseBytes[i] = parseInt(base.slice(i * 2, i * 2 + 2), 16)
        for (let code = lo; code <= hi && code - lo < 65536; code += 1) {
          const next = baseBytes.slice()
          // 递增落在最后一个 UTF-16 码元上，和规范一致。
          const last = next.length - 2
          if (last >= 0) {
            const value = (next[last] << 8) | next[last + 1]
            next[last] = ((value + (code - lo)) >> 8) & 0xff
            next[last + 1] = (value + (code - lo)) & 0xff
          }
          map.set(code, utf16be(next))
        }
      }
    }
  }

  return map.size > 0 ? { map, codeBytes } : { map, codeBytes }
}

function streamText(stream: PDFStream): string | null {
  try {
    if (stream instanceof PDFRawStream) {
      return bytesToLatin1(decodePDFRawStream(stream).decode())
    }
    const maybe = stream as PDFStream & { getContentsString?: () => string }
    if (typeof maybe.getContentsString === 'function') return maybe.getContentsString()
  } catch {
    return null
  }
  return null
}

function decodeWithFont(bytes: Uint8Array, font: FontInfo | undefined): string | null {
  if (font?.map && font.map.size > 0) {
    let out = ''
    const size = font.codeBytes
    for (let i = 0; i + size - 1 < bytes.length; i += size) {
      const code = size === 2 ? (bytes[i] << 8) | bytes[i + 1] : bytes[i]
      out += font.map.get(code) ?? ''
    }
    return out
  }

  if (font?.type0) {
    const encoding = font.encoding ?? ''
    if (/ucs2|utf16/i.test(encoding)) return utf16be(bytes)
    if (/gbk|gb-euc|gbpc|gb18030/i.test(encoding)) {
      try {
        return new TextDecoder('gbk').decode(bytes)
      } catch {
        return null
      }
    }
    // Identity-H 且没有 ToUnicode：里面是字形编号，硬解只会得到乱码。
    return null
  }

  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return utf16be(bytes.subarray(2))
  return latin1(bytes)
}

/* ── 内容流词法 ─────────────────────────────────────────────── */

function readLiteralString(source: string, start: number): { bytes: Uint8Array; end: number } {
  const out: number[] = []
  let depth = 1
  let i = start + 1
  while (i < source.length) {
    const ch = source[i]
    if (ch === '\\') {
      const next = source[i + 1]
      i += 2
      if (next === 'n') out.push(10)
      else if (next === 'r') out.push(13)
      else if (next === 't') out.push(9)
      else if (next === 'b') out.push(8)
      else if (next === 'f') out.push(12)
      else if (next === '\n') {
        /* 续行：反斜杠 + 换行什么都不输出 */
      } else if (next === '\r') {
        if (source[i] === '\n') i += 1
      } else if (next >= '0' && next <= '7') {
        let value = next.charCodeAt(0) - 48
        for (let k = 0; k < 2 && source[i] >= '0' && source[i] <= '7'; k += 1) {
          value = value * 8 + (source.charCodeAt(i) - 48)
          i += 1
        }
        out.push(value & 0xff)
      } else {
        out.push(next.charCodeAt(0) & 0xff)
      }
      continue
    }
    if (ch === '(') {
      depth += 1
      out.push(40)
      i += 1
      continue
    }
    if (ch === ')') {
      depth -= 1
      if (depth === 0) return { bytes: Uint8Array.from(out), end: i + 1 }
      out.push(41)
      i += 1
      continue
    }
    out.push(ch.charCodeAt(0) & 0xff)
    i += 1
  }
  return { bytes: Uint8Array.from(out), end: source.length }
}

/** 跳过内联图片（BI … ID <二进制> EI），否则二进制里的括号会打乱词法。 */
function skipInlineImage(source: string, from: number): number {
  const id = source.indexOf('ID', from)
  if (id < 0) return source.length
  let i = id + 3
  while (i < source.length - 1) {
    if (source[i] === 'E' && source[i + 1] === 'I') {
      const before = source[i - 1]
      const after = source[i + 2]
      if ((i === 0 || WHITESPACE.has(before)) && (after === undefined || WHITESPACE.has(after) || DELIMITERS.has(after))) {
        return i + 2
      }
    }
    i += 1
  }
  return source.length
}

function tokenize(source: string): Token[] {
  const tokens: Token[] = []
  let i = 0

  while (i < source.length) {
    const ch = source[i]

    if (WHITESPACE.has(ch)) {
      i += 1
      continue
    }
    if (ch === '%') {
      while (i < source.length && source[i] !== '\n' && source[i] !== '\r') i += 1
      continue
    }

    const start = i

    if (ch === '(') {
      const { bytes, end } = readLiteralString(source, i)
      tokens.push({ type: 'str', text: '', bytes, start, end })
      i = end
      continue
    }

    if (ch === '<' && source[i + 1] === '<') {
      tokens.push({ type: 'punct', text: '<<', start, end: i + 2 })
      i += 2
      continue
    }

    if (ch === '>') {
      if (source[i + 1] === '>') {
        tokens.push({ type: 'punct', text: '>>', start, end: i + 2 })
        i += 2
      } else {
        tokens.push({ type: 'punct', text: '>', start, end: i + 1 })
        i += 1
      }
      continue
    }

    if (ch === '<') {
      i += 1
      let hex = ''
      while (i < source.length && source[i] !== '>') {
        if (!WHITESPACE.has(source[i])) hex += source[i]
        i += 1
      }
      i += 1
      const bytes = new Uint8Array(hex.length >> 1)
      for (let k = 0; k < bytes.length; k += 1) bytes[k] = parseInt(hex.slice(k * 2, k * 2 + 2), 16) || 0
      tokens.push({ type: 'hex', text: hex, bytes, start, end: i })
      continue
    }

    if (ch === '[' || ch === ']' || ch === '{' || ch === '}') {
      tokens.push({ type: 'punct', text: ch, start, end: i + 1 })
      i += 1
      continue
    }

    if (ch === '/') {
      i += 1
      let name = ''
      while (i < source.length && isRegular(source[i])) {
        if (source[i] === '#' && i + 2 < source.length) {
          name += String.fromCharCode(parseInt(source.slice(i + 1, i + 3), 16) || 0)
          i += 3
          continue
        }
        name += source[i]
        i += 1
      }
      tokens.push({ type: 'name', text: name, start, end: i })
      continue
    }

    if (isDigit(ch) || ch === '+' || ch === '-' || ch === '.') {
      while (i < source.length && (isDigit(source[i]) || source[i] === '.' || source[i] === '-' || source[i] === '+')) i += 1
      tokens.push({ type: 'num', text: source.slice(start, i), start, end: i })
      continue
    }

    let op = ''
    while (i < source.length && isRegular(source[i])) {
      op += source[i]
      i += 1
    }
    if (op === '') {
      i += 1
      continue
    }
    tokens.push({ type: 'op', text: op, start, end: i })

    if (op === 'BI') {
      // 内联图片的数据不是内容流语法，整段跳过后再继续。
      i = skipInlineImage(source, i)
    }
  }

  return tokens
}

type Operator = { op: string; args: Token[]; start: number; end: number }

function* operators(tokens: Token[]): Generator<Operator> {
  let args: Token[] = []
  for (const token of tokens) {
    if (token.type !== 'op') {
      args.push(token)
      continue
    }
    yield { op: token.text, args, start: args[0]?.start ?? token.start, end: token.end }
    args = []
  }
}

/* ── 资源解析 ───────────────────────────────────────────────── */

function readFonts(doc: PDFDocument, resources: PDFDict | undefined): Fonts {
  const fonts: Fonts = new Map()
  const fontDict = dictOf(resources, 'Font')
  if (!fontDict) return fonts

  for (const [key, value] of fontDict.entries()) {
    const resolved = doc.context.lookup(value)
    if (!(resolved instanceof PDFDict)) continue
    const dict = resolved
    const type0 = nameOf(dict, 'Subtype') === 'Type0'
    const encodingValue = valueOf(dict, 'Encoding')
    const encoding =
      encodingValue instanceof PDFName
        ? encodingValue.decodeText()
        : encodingValue instanceof PDFDict
          ? nameOf(encodingValue, 'CMapName')
          : undefined

    let map: Map<number, string> | undefined
    let codeBytes: 1 | 2 = type0 ? 2 : 1
    const toUnicode = valueOf(dict, 'ToUnicode')
    if (toUnicode instanceof PDFStream) {
      const text = streamText(toUnicode)
      if (text) {
        const cmap = parseToUnicode(text)
        if (cmap && cmap.map.size > 0) {
          map = cmap.map
          codeBytes = cmap.codeBytes
        }
      }
    }

    fonts.set(key.decodeText(), { map, codeBytes, type0, encoding })
  }

  return fonts
}

function readAlphas(resources: PDFDict | undefined): Map<string, number> {
  const alphas = new Map<string, number>()
  const dict = dictOf(resources, 'ExtGState')
  if (!dict) return alphas
  for (const [key, value] of dict.entries()) {
    const state = dict.context.lookup(value)
    if (!(state instanceof PDFDict)) continue
    const alpha = numberOf(state, 'ca') ?? numberOf(state, 'CA')
    if (typeof alpha === 'number') alphas.set(key.decodeText(), alpha)
  }
  return alphas
}

type FormInfo = { name: string; text: string; dict: PDFDict }

function readForms(doc: PDFDocument, resources: PDFDict | undefined, fallbackFonts: Fonts): FormInfo[] {
  const xobjects = dictOf(resources, 'XObject')
  if (!xobjects) return []
  const forms: FormInfo[] = []
  const context = doc.context

  for (const [key, value] of xobjects.entries()) {
    const stream = context.lookup(value)
    if (!(stream instanceof PDFRawStream)) continue
    if (nameOf(stream.dict, 'Subtype') !== 'Form') continue
    const content = streamText(stream)
    if (!content) continue
    const ownResources = dictOf(stream.dict, 'Resources')
    const fonts = ownResources ? readFonts(doc, ownResources) : fallbackFonts
    const blocks = scanBlocks(content, fonts, new Map())
    const text = blocks.map((block) => block.text).join('')
    forms.push({ name: key.decodeText(), text, dict: stream.dict })
  }

  return forms
}

/* ── 文本块扫描 ─────────────────────────────────────────────── */

type Matrix = [number, number, number, number, number, number]

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

/** PDF 的 cm 是「新矩阵 × 当前矩阵」，顺序不能反。 */
function concatMatrix(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ]
}

function isRotated(m: Matrix): boolean {
  return Math.abs(m[1]) > 0.001 || Math.abs(m[2]) > 0.001
}

function scanBlocks(content: string, fonts: Fonts, alphas: Map<string, number>): TextBlock[] {
  if (content.length > MAX_CONTENT_CHARS) return []

  const blocks: TextBlock[] = []
  let current: { start: number; end: number; parts: string[]; undecoded: boolean; rotated: boolean; faint: boolean; font?: FontInfo } | null = null
  let fontName = ''
  let alpha = 1
  let ctm = IDENTITY
  const stack: { ctm: Matrix; alpha: number }[] = []

  for (const item of operators(tokenize(content))) {
    switch (item.op) {
      case 'q':
        stack.push({ ctm, alpha })
        break
      case 'Q': {
        const previous = stack.pop()
        if (previous) {
          ctm = previous.ctm
          alpha = previous.alpha
        }
        break
      }
      case 'cm': {
        const n = item.args.filter((arg) => arg.type === 'num').map((arg) => Number(arg.text))
        if (n.length >= 6) ctm = concatMatrix([n[0], n[1], n[2], n[3], n[4], n[5]], ctm)
        break
      }
      case 'BT':
        current = {
          start: item.start,
          end: item.end,
          parts: [],
          undecoded: false,
          rotated: isRotated(ctm),
          faint: alpha < 0.9,
          font: fonts.get(fontName),
        }
        break
      case 'ET':
        if (current) {
          current.end = item.end
          blocks.push({
            start: current.start,
            end: current.end,
            text: current.parts.join(''),
            rotated: current.rotated,
            faint: current.faint,
            undecoded: current.undecoded,
          })
          current = null
        }
        break
      case 'Tf': {
        const name = item.args.find((arg) => arg.type === 'name')
        if (name) fontName = name.text
        if (current) current.font = fonts.get(fontName)
        break
      }
      case 'Tm': {
        if (!current) break
        const n = item.args.filter((arg) => arg.type === 'num').map((arg) => Number(arg.text))
        // 文本矩阵自带旋转，或外层 cm 已旋转，都算「斜着画」。
        if (n.length >= 6 && isRotated([n[0], n[1], n[2], n[3], n[4], n[5]])) current.rotated = true
        break
      }
      case 'gs': {
        const name = item.args.find((arg) => arg.type === 'name')
        if (name) alpha = alphas.get(name.text) ?? 1
        break
      }
      case 'Tj':
      case "'":
      case '"': {
        if (!current) break
        const str = item.args.find((arg) => arg.type === 'str' || arg.type === 'hex')
        if (!str?.bytes) break
        const decoded = decodeWithFont(str.bytes, current.font)
        if (decoded === null) current.undecoded = true
        else current.parts.push(decoded)
        break
      }
      case 'TJ': {
        if (!current) break
        for (const arg of item.args) {
          if ((arg.type !== 'str' && arg.type !== 'hex') || !arg.bytes) continue
          const decoded = decodeWithFont(arg.bytes, current.font)
          if (decoded === null) current.undecoded = true
          else current.parts.push(decoded)
        }
        break
      }
    }
  }

  return blocks
}

function scanFormUses(content: string): FormUse[] {
  const uses: FormUse[] = []
  let depth = 0
  for (const item of operators(tokenize(content))) {
    if (item.op === 'BT') depth += 1
    else if (item.op === 'ET') depth = Math.max(0, depth - 1)
    else if (item.op === 'Do' && depth === 0) {
      const name = item.args.find((arg) => arg.type === 'name')
      if (name) uses.push({ name: name.text, start: name.start, end: item.end })
    }
  }
  return uses
}

/* ── 页面内容流 ─────────────────────────────────────────────── */

type PageContent = { text: string }

function readPageContent(doc: PDFDocument, page: PDFPage): PageContent | null {
  const contents = page.node.Contents()
  if (!contents) return null
  // 规范允许 /Contents 是单个流，也允许是流数组；两种都要拼成一条逻辑内容流。
  const entries = contents instanceof PDFArray ? contents.asArray() : [contents]
  const parts: string[] = []

  for (const entry of entries) {
    const stream = doc.context.lookup(entry)
    if (!(stream instanceof PDFRawStream)) continue
    const text = streamText(stream)
    if (text === null) return null
    parts.push(text)
  }

  if (parts.length === 0) return null
  return { text: parts.join('\n') }
}

/* ── 扫描 ───────────────────────────────────────────────────── */

type Group = { text: string; pages: Set<number>; blocks: number; rotated: boolean; faint: boolean }

function addToGroup(groups: Map<string, Group>, text: string, page: number, rotated: boolean, faint: boolean) {
  const key = normalize(text)
  if (key.length < 2 || key.length > MAX_CANDIDATE_CHARS) return
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffd]/.test(key)) return
  const group = groups.get(key)
  if (group) {
    group.pages.add(page)
    group.blocks += 1
    group.rotated = group.rotated || rotated
    group.faint = group.faint || faint
  } else {
    groups.set(key, { text: text.trim(), pages: new Set([page]), blocks: 1, rotated, faint })
  }
}

/** 页码、日期这类「每页都出现但每页不同」的文字不该进候选。 */
function looksLikePagination(text: string): boolean {
  const compact = normalize(text)
  return /^(第)?\d{1,4}(页|\/\d{1,4})?$/.test(compact) || /^page\d+(of\d+)?$/i.test(compact)
}

function collectCandidates(groups: Map<string, Group>, pageCount: number): PdfFinding[] {
  const findings: PdfFinding[] = []
  const threshold = Math.max(2, Math.ceil(pageCount * 0.35))

  for (const [key, group] of groups) {
    if (looksLikePagination(group.text)) continue
    const repeated = group.pages.size >= threshold
    const singlePageSignal = pageCount === 1 && (group.rotated || group.faint) && key.length <= 32
    if (!repeated && !singlePageSignal) continue
    findings.push({
      id: `text:${encodeURIComponent(key)}`,
      kind: 'text',
      text: group.text,
      pages: group.pages.size,
      blocks: group.blocks,
    })
  }

  return findings.sort((a, b) => b.pages - a.pages || b.blocks - a.blocks)
}

export async function scanPdf(bytes: Uint8Array): Promise<PdfScan> {
  if (bytes.byteLength > MAX_PDF_BYTES) throw new PdfWatermarkError('too-large')
  if (bytes.byteLength === 0) throw new PdfWatermarkError('empty')

  let doc: PDFDocument
  try {
    doc = await PDFDocument.load(bytes, { updateMetadata: false })
  } catch (error) {
    const name = (error as Error)?.name ?? ''
    if (/encrypt/i.test(name) || /encrypt/i.test((error as Error)?.message ?? '')) {
      throw new PdfWatermarkError('encrypted')
    }
    throw new PdfWatermarkError('invalid')
  }

  const pages = doc.getPages()
  const groups = new Map<string, Group>()
  const formGroups = new Map<string, Group>()
  let undecoded = 0
  let annotationCount = 0

  pages.forEach((page, index) => {
    const resources = page.node.Resources()
    const fonts = readFonts(doc, resources)
    const alphas = readAlphas(resources)

    const content = readPageContent(doc, page)
    if (content) {
      for (const block of scanBlocks(content.text, fonts, alphas)) {
        if (block.undecoded && !block.text) {
          undecoded += 1
          continue
        }
        addToGroup(groups, block.text, index, block.rotated, block.faint)
      }

      const forms = readForms(doc, resources, fonts)
      const byName = new Map(forms.map((form) => [form.name, form]))
      for (const use of scanFormUses(content.text)) {
        const form = byName.get(use.name)
        if (!form) continue
        const key = normalize(form.text)
        if (key.length < 2 || key.length > MAX_CANDIDATE_CHARS) continue
        addToGroup(formGroups, form.text, index, false, false)
      }
    }

    const annots = page.node.Annots()
    if (annots) {
      annots.asArray().forEach((entry) => {
        const dict = doc.context.lookup(entry)
        if (!(dict instanceof PDFDict)) return
        const subtype = nameOf(dict, 'Subtype')
        if (subtype === 'Watermark' || subtype === 'Stamp') annotationCount += 1
      })
    }
  })

  const findings = collectCandidates(groups, pages.length)
  const formFindings = collectCandidates(formGroups, pages.length)
    .map((item) => ({ ...item, kind: 'form' as const, id: `form:${item.id.slice('text:'.length)}` }))
    // 同一个词同时出现在页面文本和 Form 里时，删 Form 更彻底，去重保留后者。
    .filter((item) => !findings.some((text) => text.text === item.text))

  if (annotationCount > 0) {
    findings.push({
      id: 'annot:watermark',
      kind: 'annot',
      text: '',
      pages: pages.length,
      blocks: annotationCount,
    })
  }

  return { pageCount: pages.length, findings: [...formFindings, ...findings], undecoded }
}

/* ── 移除 ───────────────────────────────────────────────────── */

function splice(content: string, spans: { start: number; end: number }[]): string {
  if (spans.length === 0) return content
  const sorted = [...spans].sort((a, b) => a.start - b.start)
  let out = ''
  let cursor = 0
  for (const span of sorted) {
    if (span.start < cursor) continue
    out += content.slice(cursor, span.start)
    cursor = span.end
  }
  out += content.slice(cursor)
  return out
}

export interface PdfRemoval {
  bytes: Uint8Array
  removed: { text: number; form: number; annot: number }
}

export async function removePdfWatermarks(bytes: Uint8Array, ids: string[]): Promise<PdfRemoval> {
  const removed = { text: 0, form: 0, annot: 0 }
  if (ids.length === 0) return { bytes, removed }

  const textKeys = new Set(
    ids.filter((id) => id.startsWith('text:')).map((id) => decodeURIComponent(id.slice(5))),
  )
  const formKeys = new Set(
    ids.filter((id) => id.startsWith('form:')).map((id) => decodeURIComponent(id.slice(5))),
  )
  const dropAnnotations = ids.includes('annot:watermark')

  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const context = doc.context

  for (const page of doc.getPages()) {
    const resources = page.node.Resources()
    const fonts = readFonts(doc, resources)
    const alphas = readAlphas(resources)
    const content = readPageContent(doc, page)
    if (!content) continue

    const spans: { start: number; end: number }[] = []

    if (textKeys.size > 0) {
      for (const block of scanBlocks(content.text, fonts, alphas)) {
        if (!block.text) continue
        if (textKeys.has(normalize(block.text))) {
          spans.push({ start: block.start, end: block.end })
          removed.text += 1
        }
      }
    }

    if (formKeys.size > 0) {
      const forms = readForms(doc, resources, fonts)
      const targets = new Set(forms.filter((form) => formKeys.has(normalize(form.text))).map((form) => form.name))
      if (targets.size > 0) {
        for (const use of scanFormUses(content.text)) {
          if (!targets.has(use.name)) continue
          spans.push({ start: use.start, end: use.end })
          removed.form += 1
        }
      }
    }

    if (spans.length > 0) {
      const next = splice(content.text, spans)
      const dict = context.obj({})
      dict.set(PDFName.of('Filter'), PDFName.of('FlateDecode'))
      const compressed = zlibSync(latin1ToBytes(next), { level: 6 })
      const stream = PDFRawStream.of(dict, compressed)
      const ref = context.register(stream)
      page.node.set(PDFName.of('Contents'), context.obj([ref]))
    }

    if (dropAnnotations) {
      const annots = page.node.Annots()
      if (annots) {
        const kept = annots.asArray().filter((entry) => {
          const dict = context.lookup(entry)
          const subtype = dict instanceof PDFDict ? nameOf(dict, 'Subtype') : undefined
          const isWatermark = subtype === 'Watermark' || subtype === 'Stamp'
          if (isWatermark) removed.annot += 1
          return !isWatermark
        })
        if (kept.length > 0) {
          page.node.set(PDFName.of('Annots'), context.obj(kept))
        } else {
          page.node.delete(PDFName.of('Annots'))
        }
      }
    }
  }

  if (removed.text + removed.form + removed.annot === 0) return { bytes, removed }

  const saved = await doc.save({ useObjectStreams: true, addDefaultPage: false, updateFieldAppearances: false })
  return { bytes: saved, removed }
}

/** 页面文字里是否还有某个水印——测试和「移除后再检测」共用。 */
export async function containsPdfText(bytes: Uint8Array, needle: string): Promise<boolean> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const target = normalize(needle)
  for (const page of doc.getPages()) {
    const content = readPageContent(doc, page)
    if (!content) continue
    const fonts = readFonts(doc, page.node.Resources())
    const blocks = scanBlocks(content.text, fonts, readAlphas(page.node.Resources()))
    if (blocks.some((block) => normalize(block.text).includes(target))) return true
  }
  return false
}

/* ────────────────────────── 添加水印 ──────────────────────────
 *
 * 文字水印先由 canvas 渲染成一张透明 PNG（见 pdf-watermark-image.ts），再按页
 * 居中贴上——这样中文、日文、符号都不用嵌字体，也不会因为标准字体缺字变方框。
 *
 * 贴图不重建页面：只往每页内容流末尾追加一条 Do 调用，书签、表单域、元数据
 * 都不受影响。旋转围绕页面中心计算，pdf-lib 的 rotate 以绘制原点为轴，所以
 * 还要把原点按旋转后的半宽/半高平移回去。
 */

export interface PdfWatermarkImage {
  /** PNG 或 JPEG 字节，mime 决定用 embedPng 还是 embedJpg。 */
  png: Uint8Array
  mime?: 'image/png' | 'image/jpeg'
  width: number
  height: number
}

/** 与 DOCX 共用一套参数定义，缺省值由 normalizeOptions 补齐。 */
export type PdfStampOptions = Partial<WatermarkOptions>

export interface PdfStamp {
  bytes: Uint8Array
  pages: number
}

/** 每页枚数 → 网格：2 枚上下两行，4 / 9 / 16 走 n×n。 */
function pdfGrid(count: number): { rows: number; columns: number } {
  if (count <= 2) return { rows: Math.max(1, count), columns: 1 }
  const columns = Math.max(1, Math.min(4, Math.round(Math.sqrt(count))))
  return { rows: Math.ceil(count / columns), columns }
}

/** 内容流可能是单段 ref，也可能是 ref 数组；统一取成数组。 */
function contentRefs(page: PDFPage): PDFObject[] {
  const contents = page.node.Contents()
  if (!contents) return []
  // 必须浅拷贝：pdf-lib 会往同一个 PDFArray 里 push 新段，直接持有引用会让 before 一起变长。
  return contents instanceof PDFArray ? [...contents.asArray()] : [contents]
}

/** layer=below：把本次新增的绘制段挪到原内容之前，水印就衬在正文下面。 */
function moveNewContentBehind(doc: PDFDocument, page: PDFPage, before: PDFObject[]) {
  const contents = page.node.Contents()
  if (!contents) return
  const after = contents instanceof PDFArray ? [...contents.asArray()] : [contents]
  const added = after.slice(before.length)
  if (added.length === 0) return
  const reordered = doc.context.obj([]) as PDFArray
  for (const ref of added) reordered.push(ref)
  for (const ref of before) reordered.push(ref)
  page.node.set(PDFName.of('Contents'), reordered)
}

export async function stampPdfWatermark(
  bytes: Uint8Array,
  image: PdfWatermarkImage,
  input: PdfStampOptions,
): Promise<PdfStamp> {
  if (bytes.byteLength > MAX_PDF_BYTES) throw new PdfWatermarkError('too-large')
  if (bytes.byteLength === 0) throw new PdfWatermarkError('empty')

  let doc: PDFDocument
  try {
    doc = await PDFDocument.load(bytes, { updateMetadata: false })
  } catch (cause) {
    throw new PdfWatermarkError(/encrypt/i.test(String(cause)) ? 'encrypted' : 'invalid')
  }

  const pageCount = doc.getPageCount()
  if (pageCount === 0) throw new PdfWatermarkError('empty')

  const options = normalizeOptions(input)
  const embedded = image.mime === 'image/jpeg' ? await doc.embedJpg(image.png) : await doc.embedPng(image.png)
  const radians = (options.angle * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  const ratio = image.height / Math.max(1, image.width)
  const count = options.count
  const { rows, columns } = pdfGrid(count)
  const selected = parsePageRange(options.pages, pageCount)
  const anchor = positionRatio(options.position)

  /** 以 (centerX, centerY) 为中心贴一枚；原点退到「旋转后仍居中」的位置。 */
  const place = (page: PDFPage, centerX: number, centerY: number, width: number, height: number) => {
    const x = centerX - (width / 2) * cos + (height / 2) * sin
    const y = centerY - (width / 2) * sin - (height / 2) * cos
    page.drawImage(embedded, {
      x,
      y,
      width,
      height,
      rotate: degrees(options.angle),
      opacity: options.opacity,
    })
  }

  doc.getPages().forEach((page, index) => {
    // 空集合说明用户填的范围一个有效页都没有，退回到「整篇」而不是生成无水印文件。
    if (selected.size > 0 && !selected.has(index)) return
    // 先触发 pdf-lib 的内容流规范化：否则首次 drawImage 会把单一流扩成数组并插入 q/Q，
    // 让「新增了几段」的计数错位，水印就挪不到正文前面。
    if (options.layer === 'below') page.node.normalizedEntries()
    const before = options.layer === 'below' ? contentRefs(page) : []
    const { width: pageWidth, height: pageHeight } = page.getSize()

    if (count === 1) {
      let width = pageWidth * 0.72 * options.scale
      let height = width * ratio
      // 竖版长水印可能比页面还高，按高度收一次，别让文字戳出页边。
      if (height > pageHeight * 0.7) {
        height = pageHeight * 0.7
        width = height / ratio
      }
      place(page, pageWidth * anchor.x, pageHeight * anchor.y, width, height)
    } else {
      // 平铺：页面切成 rows×columns 个格子，每格中间贴一枚，大小按格子内缩一圈。
      const cellW = pageWidth / columns
      const cellH = pageHeight / rows
      // 2 枚上下排时一行只有一枚，用单枚的宽度；gap 越小越稀疏。
      let width = (columns === 1 ? pageWidth * 0.72 : cellW * 0.88) * options.gap * options.scale
      let height = width * ratio
      if (height > cellH * 0.6) {
        height = cellH * 0.6
        width = height / ratio
      }
      for (let row = 0; row < rows; row += 1) {
        for (let column = 0; column < columns; column += 1) {
          place(
            page,
            column * cellW + cellW / 2,
            // pdf-lib 的 y 轴朝上，第一行要放在页面顶部。
            pageHeight - (row * cellH + cellH / 2),
            width,
            height,
          )
        }
      }
    }

    if (options.layer === 'below') moveNewContentBehind(doc, page, before)
  })

  const saved = await doc.save({ useObjectStreams: true, addDefaultPage: false, updateFieldAppearances: false })
  return { bytes: saved, pages: pageCount }
}
