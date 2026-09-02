import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// Browser and MCP Apps renderer. Format behavior remains owned by apps/sheets.
export default defineConfig({
  root: 'src/renderer',
  plugins: [react()],
  base: './',
  resolve: {
    alias: {
      '@univerjs/telemetry': new URL(
        './src/renderer/vendor/univer-optional-metrics.ts',
        import.meta.url,
      ).pathname,
    },
  },
  build: {
    // MCP Apps hosts use a modern Chromium runtime; avoid compatibility
    // transforms that only add parse/evaluation work to the 19 MB Univer bundle.
    target: 'es2022',
    outDir: new URL('./dist/renderer', import.meta.url).pathname,
    emptyOutDir: true,
    assetsInlineLimit: Number.MAX_SAFE_INTEGER,
    // Preserve optional locale and hyphenation imports as real lazy chunks.
    // package-plugin embeds them into the single XLSX UI resource and serves
    // them from an in-memory module vault rather than compiling them at boot.
    rollupOptions: { output: { inlineDynamicImports: false } },
  },
  server: {
    port: Number(process.env.SHEETS_DEV_PORT) || 5174,
    strictPort: true,
  },
})
