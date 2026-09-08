// Keep the bundled Office fallback fonts outside the initial executable entry.
// The module still lives in the self-contained vault, and main awaits this
// promise before Univer performs its first canvas measurement.
export const cellFontAliasesReady = import('./cell-font-fallback').then(
  ({ installCanvasFontFallback, registerCellFontAliases }) => {
    installCanvasFontFallback()
    return registerCellFontAliases()
  },
)
