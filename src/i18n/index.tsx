import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from 'react'
import { en } from './en'
import { zh } from './zh'

export type Lang = 'zh' | 'en'

const dictionaries = { zh, en }
const STORAGE_KEY = 'gendocx-tools.lang'

interface LangContextValue {
  lang: Lang
  setLang: (lang: Lang) => void
}

const LangContext = createContext<LangContextValue>({ lang: 'en', setLang: () => {} })

function detectLang(): Lang {
  if (typeof window === 'undefined') return 'en'
  const saved = window.localStorage.getItem(STORAGE_KEY)
  if (saved === 'zh' || saved === 'en') return saved
  return navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en'
}

export function LangProvider({ children }: { children: ReactNode }) {
  const [lang, setLangState] = useState<Lang>(detectLang)

  const setLang = useCallback((next: Lang) => {
    setLangState(next)
    try {
      window.localStorage.setItem(STORAGE_KEY, next)
    } catch {
      // Private mode: keep the in-memory choice only.
    }
  }, [])

  useEffect(() => {
    document.documentElement.lang = lang === 'zh' ? 'zh-CN' : 'en'
  }, [lang])

  const value = useMemo(() => ({ lang, setLang }), [lang, setLang])

  return <LangContext.Provider value={value}>{children}</LangContext.Provider>
}

export function useLang() {
  return useContext(LangContext)
}

type Params = Record<string, string | number>

function resolve(dict: unknown, path: string): unknown {
  return path.split('.').reduce<unknown>((node, key) => {
    if (node && typeof node === 'object' && key in (node as Record<string, unknown>)) {
      return (node as Record<string, unknown>)[key]
    }
    return undefined
  }, dict)
}

const PLURAL = /\{(\w+),\s*plural,\s*((?:[a-z]+\s*\{[^{}]*\}\s*)+)\}/g

/**
 * Minimal ICU support: `{count, plural, one {# file} other {# files}}`.
 * The site copy uses plural only for file/page counts, so one/other is enough;
 * Chinese ships the `other` branch alone, which still resolves here.
 */
function interpolate(template: string, params?: Params): string {
  if (!params) return template

  const withPlural = template.replace(PLURAL, (_match, name: string, body: string) => {
    const value = Number(params[name] ?? 0)
    const branches: Record<string, string> = {}
    body.replace(/([a-z]+)\s*\{([^{}]*)\}/g, (_branch, key: string, text: string) => {
      branches[key] = text
      return ''
    })
    const chosen = value === 1 ? (branches.one ?? branches.other) : (branches.other ?? branches.one)
    return (chosen ?? '').replace(/#/g, String(value))
  })

  return withPlural.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in params ? String(params[key]) : match,
  )
}

/**
 * Same call shape as the site's next-intl `useTranslations`, so the ported
 * components keep their `t('seo.step1Title')` / `t('remove', { index })` calls.
 * A missing key returns the key itself instead of throwing — a demo page should
 * never blank out because one string is missing.
 */
export function useT(namespace: keyof typeof dictionaries.en) {
  const { lang } = useLang()

  return useCallback(
    (key: string, params?: Params): string => {
      const dict = dictionaries[lang][namespace] as unknown
      const value = resolve(dict, key)
      if (typeof value !== 'string') {
        if (import.meta.env.DEV) console.warn(`[i18n] missing ${lang}.${namespace}.${key}`)
        return key
      }
      return interpolate(value, params)
    },
    [lang, namespace],
  )
}
