/**
 * 加水印参数的共享定义与纯函数。
 *
 * DOCX 与 PDF 是两套完全不同的写入方式（VML 页眉 / 内容流贴图），但用户看到
 * 的控件只有一份，所以数值收敛、页面范围、动态变量这类逻辑集中放这里，避免
 * 两个引擎各写一套导致行为漂移。
 */

export type WatermarkTone = 'solid' | 'outline'
export type WatermarkLayer = 'above' | 'below'
export type WatermarkKind = 'text' | 'image'
export type HorizontalPosition = 'left' | 'center' | 'right'
export type VerticalPosition = 'top' | 'middle' | 'bottom'

export interface WatermarkPosition {
  horizontal: HorizontalPosition
  vertical: VerticalPosition
}

export interface WatermarkImage {
  bytes: Uint8Array
  mime: 'image/png' | 'image/jpeg'
  width: number
  height: number
}

/** 一次加水印操作的完整参数。两个引擎各取所需，缺省值保持一致。 */
export interface WatermarkOptions {
  text: string
  /** 0.02–1 */
  opacity: number
  /** 0–360，逆时针角度 */
  angle: number
  color: string
  tone: WatermarkTone
  /** 0–3：实心调字重，描边调线宽 */
  weight: number
  /** 1 / 2 / 4 / 9 / 16 */
  count: number
  /** 整体缩放 0.5–1.6 */
  scale: number
  /** 单枚时的九宫格位置 */
  position: WatermarkPosition
  /** 平铺间距 0.6–1.4，小于 1 更稀疏 */
  gap: number
  /** 图层：衬于文字下方 / 浮于上方 */
  layer: WatermarkLayer
  /** 页面范围，如 `1,3,5-8`；空或 all 表示全部。仅 PDF 生效。 */
  pages: string
  kind: WatermarkKind
  image?: WatermarkImage
}

export const DEFAULT_WATERMARK_OPTIONS: Omit<WatermarkOptions, 'text' | 'image'> = {
  opacity: 0.18,
  angle: 45,
  color: '#7a7a7a',
  tone: 'solid',
  weight: 1,
  count: 1,
  scale: 1,
  position: { horizontal: 'center', vertical: 'middle' },
  gap: 1,
  layer: 'above',
  pages: 'all',
  kind: 'text',
}

export function clamp(value: number, min: number, max: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback
  return Math.min(max, Math.max(min, value))
}

/** 把用户输入收敛到合法区间，避免异常值写进文档结构。 */
export function normalizeOptions(input: Partial<WatermarkOptions>): WatermarkOptions {
  const base = DEFAULT_WATERMARK_OPTIONS
  const position = input.position ?? base.position
  return {
    text: (input.text ?? '').trim(),
    opacity: clamp(input.opacity ?? base.opacity, 0.02, 1, base.opacity),
    angle: clamp(input.angle ?? base.angle, 0, 360, base.angle),
    color: input.color ?? base.color,
    tone: input.tone === 'outline' ? 'outline' : 'solid',
    weight: Math.round(clamp(input.weight ?? base.weight, 0, 3, base.weight)),
    count: Math.min(16, Math.max(1, Math.round(input.count ?? base.count))),
    scale: clamp(input.scale ?? base.scale, 0.5, 1.6, base.scale),
    position: {
      horizontal: (['left', 'center', 'right'] as const).includes(position.horizontal) ? position.horizontal : 'center',
      vertical: (['top', 'middle', 'bottom'] as const).includes(position.vertical) ? position.vertical : 'middle',
    },
    gap: clamp(input.gap ?? base.gap, 0.6, 1.4, base.gap),
    layer: input.layer === 'below' ? 'below' : 'above',
    pages: (input.pages ?? base.pages).trim() || 'all',
    kind: input.kind === 'image' && input.image ? 'image' : 'text',
    image: input.image,
  }
}

/** 解析 `1,3,5-8`；返回 0 基页索引集合，无法解析的片段直接忽略。 */
export function parsePageRange(input: string, total: number): Set<number> {
  const all = () => new Set(Array.from({ length: total }, (_, index) => index))
  const text = input.trim()
  if (!text || /^(all|\*|全部)$/i.test(text)) return all()

  const pages = new Set<number>()
  for (const piece of text.split(/[,，\s]+/)) {
    if (!piece) continue
    const range = /^(\d+)\s*[-–~]\s*(\d+)$/.exec(piece)
    if (range) {
      const start = Math.min(Number(range[1]), Number(range[2]))
      const end = Math.max(Number(range[1]), Number(range[2]))
      for (let page = start; page <= end && page <= total; page += 1) {
        if (page >= 1) pages.add(page - 1)
      }
      continue
    }
    if (/^\d+$/.test(piece)) {
      const page = Number(piece)
      if (page >= 1 && page <= total) pages.add(page - 1)
    }
  }
  return pages
}

/** 页面范围是否命中了任何有效页（空集合表示输入全无效，UI 要提示）。 */
export function isValidPageRange(input: string, total: number): boolean {
  const text = input.trim()
  if (!text || /^(all|\*|全部)$/i.test(text)) return true
  if (parsePageRange(text, total).size > 0) return true
  return false
}

export interface WatermarkVariables {
  fileName: string
  pageCount: number
  now: Date
}

const pad = (value: number) => String(value).padStart(2, '0')

/** 文字里的 {date} / {filename} 等占位符在生成时替换成实际值。 */
export function expandWatermarkVariables(text: string, vars: WatermarkVariables): string {
  const d = vars.now
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
  const time = `${pad(d.getHours())}:${pad(d.getMinutes())}`
  const values: Record<string, string> = {
    '{date}': date,
    '{time}': time,
    '{datetime}': `${date} ${time}`,
    '{year}': String(d.getFullYear()),
    '{filename}': vars.fileName,
    '{pages}': String(vars.pageCount),
  }
  return text.replace(/\{(date|time|datetime|year|filename|pages)\}/gi, (match) => values[match.toLowerCase()] ?? match)
}

export const VARIABLE_TOKENS = ['{date}', '{time}', '{datetime}', '{year}', '{filename}', '{pages}'] as const

/** 单枚水印在页面上的相对中心点（PDF 用 y 轴朝上的比例）。 */
export function positionRatio(position: WatermarkPosition): { x: number; y: number } {
  const x = position.horizontal === 'left' ? 0.28 : position.horizontal === 'right' ? 0.72 : 0.5
  // PDF 的 y 轴朝上：top 对应更大的 y。
  const y = position.vertical === 'top' ? 0.74 : position.vertical === 'bottom' ? 0.26 : 0.5
  return { x, y }
}
