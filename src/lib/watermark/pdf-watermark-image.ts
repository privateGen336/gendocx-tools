import { stampPdfWatermark, type PdfStamp } from './pdf-watermark'
import { normalizeOptions, type WatermarkOptions } from './watermark-options'

/**
 * 把水印文字画成一张透明 PNG。
 *
 * 不走 pdf-lib 内嵌字体：标准 14 款字体没有中文，嵌一套中文字体又要多几 MB。
 * 用 canvas 渲染则直接复用系统字体，中英文、数字和符号都能画，且浏览器本来
 * 就有；代价是位图——按目标宽度的 3 倍渲染，正常阅读与打印都看不出锯齿。
 */

const FONT_STACK =
  '"PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans SC", "Helvetica Neue", Arial, sans-serif'

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const out = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) out[index] = binary.charCodeAt(index)
  return out
}

export interface RenderedWatermark {
  png: Uint8Array
  width: number
  height: number
}

export type WatermarkTone = 'solid' | 'outline'

/** 粗细档位 0–3 对应的 canvas 字重，和 VML 里的 font-weight 保持一致。 */
const WEIGHTS = [300, 500, 700, 900]

/** 目标宽度以 pt 计；返回值里的宽高是位图像素，比例与实际尺寸一致。 */
export function renderWatermarkImage(
  text: string,
  targetWidth = 720,
  color = '#7a7a7a',
  tone: WatermarkTone = 'solid',
  weight = 1,
): RenderedWatermark {
  const source = text.trim()
  if (!source) throw new Error('empty-text')

  const probe = document.createElement('canvas').getContext('2d')
  if (!probe) throw new Error('canvas-unavailable')

  const level = Math.max(0, Math.min(3, Math.round(weight)))
  const fontWeight = WEIGHTS[level]

  // 先在 100px 下量宽度，再按目标宽度反推字号，长文字不会被裁掉。
  const base = 100
  probe.font = `${fontWeight} ${base}px ${FONT_STACK}`
  const measured = probe.measureText(source).width || base
  const fontPx = Math.min(260, Math.max(24, Math.round(((targetWidth * 3) * 0.94) / (measured / base))))

  const canvas = document.createElement('canvas')
  canvas.width = Math.max(32, Math.round(targetWidth * 3))
  // 描边会往外扩一圈，画布留出余量，别把字的边缘切掉。
  canvas.height = Math.round(fontPx * (tone === 'outline' ? 1.46 : 1.3))

  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('canvas-unavailable')
  ctx.clearRect(0, 0, canvas.width, canvas.height)
  ctx.font = `${fontWeight} ${fontPx}px ${FONT_STACK}`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'

  if (tone === 'outline') {
    ctx.strokeStyle = color
    // 线宽跟着字号与粗细档位走，放大缩小时比例一致。
    ctx.lineWidth = Math.max(2, fontPx * (0.02 + level * 0.018))
    ctx.strokeText(source, canvas.width / 2, canvas.height / 2)
  } else {
    ctx.fillStyle = color
    ctx.fillText(source, canvas.width / 2, canvas.height / 2)
  }

  const dataUrl = canvas.toDataURL('image/png')
  const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
  return { png: base64ToBytes(base64), width: canvas.width, height: canvas.height }
}

/** 与 DOCX 共用一套参数定义，缺省值由 normalizeOptions 补齐。 */
export type PdfAddOptions = Partial<WatermarkOptions> & { text: string }

/** 文字 → 透明 PNG（或用户图片）→ 每页按参数贴图。 */
export async function addPdfWatermark(bytes: Uint8Array, input: PdfAddOptions): Promise<PdfStamp> {
  const options = normalizeOptions(input)
  const image =
    options.kind === 'image' && options.image
      ? {
          png: options.image.bytes,
          mime: options.image.mime,
          width: options.image.width,
          height: options.image.height,
        }
      : renderWatermarkImage(options.text, 720, options.color, options.tone, options.weight)
  return stampPdfWatermark(bytes, image, options)
}
