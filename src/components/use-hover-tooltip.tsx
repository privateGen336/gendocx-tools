import { useCallback, useId, useState } from 'react'
import { TooltipPortal, type TooltipPlacement } from './tooltip-portal'

/**
 * Hover/focus tooltip that matches the site header.
 *
 * Spread `triggerProps` on any element and give it a `data-tip` label; render
 * `tooltip` once anywhere in the subtree — it portals to <body>, so placement
 * in the tree does not matter. Styling comes from the shared `.icon-tooltip`
 * class, the same one `ThemeToggle` uses.
 */
export function useHoverTooltip(placement: TooltipPlacement = 'bottom') {
  const [tip, setTip] = useState<{ text: string; rect: DOMRect; trigger: HTMLElement } | null>(null)
  const id = useId()

  const show = useCallback((event: React.SyntheticEvent<HTMLElement>) => {
    const element = event.currentTarget
    const text = element.dataset.tip
    setTip(text ? { text, rect: element.getBoundingClientRect(), trigger: element } : null)
  }, [])

  const hide = useCallback(() => setTip(null), [])

  return {
    hide,
    triggerProps: {
      onMouseEnter: show,
      onMouseLeave: hide,
      onFocus: show,
      onBlur: hide,
      onKeyDown: (event: React.KeyboardEvent<HTMLElement>) => { if (event.key === 'Escape') hide() },
    },
    tooltip: tip?.trigger.isConnected ? <TooltipPortal id={id} text={tip.text} triggerRect={tip.rect} placement={placement} /> : null,
  }
}
