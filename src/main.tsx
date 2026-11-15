import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { LangProvider } from './i18n'
import { App } from './App'
import './app.css'

const root = document.getElementById('root')
if (!root) throw new Error('#root missing')

// The ported tool stylesheets key their dark-theme tweaks off this attribute.
const dark = window.matchMedia('(prefers-color-scheme: dark)')
const syncTheme = () => {
  document.documentElement.dataset.theme = dark.matches ? 'gendocx-dark' : 'gendocx-light'
}
syncTheme()
dark.addEventListener('change', syncTheme)

createRoot(root).render(
  <StrictMode>
    <LangProvider>
      <App />
    </LangProvider>
  </StrictMode>,
)
