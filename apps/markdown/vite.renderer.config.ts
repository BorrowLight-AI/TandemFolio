import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// npm hoists some @tiptap packages to the repo root (shared with docs at a
// different version) and nests others under this app — dedupe forces every
// import onto this app's single copy so the bundle never carries two cores.
const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]

// Standalone browser development and self-contained MCP App packaging.
export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  resolve: { dedupe: TIPTAP_DEDUPE },
  base: './',
  build: {
    outDir: '../../dist/renderer',
    emptyOutDir: true,
    rollupOptions: {
      output: { inlineDynamicImports: true },
    },
  },
  server: {
    port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
    strictPort: true,
  },
})
