import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // Relative base so the built `dist/` works from a sub-path
  // (GitHub Pages project site, a sub-folder, or `file://`).
  base: './',
  build: {
    outDir: 'dist',
    sourcemap: false,
  },
})
