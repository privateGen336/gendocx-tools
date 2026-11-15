import { useLayoutEffect, useRef } from 'react'
import { createPortal } from 'react-dom'

const TOOLTIP_OFFSET = 6
export type TooltipPlacement = 'top' | 'bottom' | 'left' | 'right'

/**
 * Portal-rendered hover tooltip anchored below a trigger element.
 * Styling comes from the global `.icon-tooltip` class.
 */
export function TooltipPortal({
  text,
  triggerRect,
  id,
  placement = 'bottom',
}: {
  text: string
  triggerRect: DOMRect
  id?: string
  placement?: TooltipPlacement
}) {
  const boxRef = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    const box = boxRef.current
    if (!box) return
    const fit = () => {
      box.style.translate = ''
      const bounds = box.getBoundingClientRect()
      const dx = Math.max(8 - bounds.left, Math.min(0, document.documentElement.clientWidth - 8 - bounds.right))
      const dy = Math.max(8 - bounds.top, Math.min(0, document.documentElement.clientHeight - 8 - bounds.bottom))
      if (dx || dy) box.style.translate = `${dx}px ${dy}px`
    }
    fit()
    window.addEventListener('resize', fit)
    return () => window.removeEventListener('resize', fit)
  }, [placement, text, triggerRect])
  if (typeof document === 'undefined') return null
  const centerX = triggerRect.left + triggerRect.width / 2
  const centerY = triggerRect.top + triggerRect.height / 2
  const position = {
    top: { top: triggerRect.top - TOOLTIP_OFFSET, left: centerX, transform: 'translate(-50%, -100%)' },
    bottom: { top: triggerRect.bottom + TOOLTIP_OFFSET, left: centerX, transform: 'translate(-50%, 0)' },
    left: { top: centerY, left: triggerRect.left - TOOLTIP_OFFSET, transform: 'translate(-100%, -50%)' },
    right: { top: centerY, left: triggerRect.right + TOOLTIP_OFFSET, transform: 'translate(0, -50%)' },
  }[placement]

  return createPortal(
    <div
      ref={boxRef}
      id={id}
      role="tooltip"
      className="icon-tooltip"
      style={{
        position: 'fixed',
        ...position,
        zIndex: 99999,
        pointerEvents: 'none',
      }}
    >
      {text}
    </div>,
    document.body
  )
}
