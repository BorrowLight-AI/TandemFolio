import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// Browser and MCP Apps renderer. Presentation semantics remain format-owned.
export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  base: './',
  build: {
    outDir: new URL('./dist/renderer', import.meta.url).pathname,
    emptyOutDir: true,
    assetsInlineLimit: Number.MAX_SAFE_INTEGER,
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
  server: {
    port: Number(process.env.SLIDES_DEV_PORT) || 5175,
    strictPort: true,
  },
})
