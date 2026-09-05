// Modified by TandemFolio contributors: retained browser-safe OLE display helpers.
/** Read-only rendering policy for worksheet OLE objects and their cached previews. */

export type OleRenderKind = 'preview' | 'placeholder'

export function oleRenderKind(
  visual: { readonly mediaPath?: string | undefined },
  previewFailed: boolean,
): OleRenderKind {
  return visual.mediaPath !== undefined && !previewFailed ? 'preview' : 'placeholder'
}

const OLE_PROG_ID_NAMES: readonly (readonly [RegExp, string])[] = [
  [/^Word\.Document\b/i, 'Microsoft Word Document'],
  [/^Word\.Template\b/i, 'Microsoft Word Template'],
  [/^Excel\.Sheet\b/i, 'Microsoft Excel Worksheet'],
  [/^Excel\.SheetMacroEnabled\b/i, 'Microsoft Excel Macro-Enabled Worksheet'],
  [/^Excel\.Chart\b/i, 'Microsoft Excel Chart'],
  [/^Worksheet$/i, 'Microsoft Excel Worksheet'],
  [/^PowerPoint\.Show\b/i, 'Microsoft PowerPoint Presentation'],
  [/^PowerPoint\.Slide\b/i, 'Microsoft PowerPoint Slide'],
  [/^Visio\.Drawing\b/i, 'Microsoft Visio Drawing'],
  [/^(Acrobat Document|AcroExch\.Document)\b/i, 'Adobe Acrobat Document'],
  [/^Equation\b/i, 'Microsoft Equation'],
  [/^MSGraph\.Chart\b/i, 'Microsoft Graph Chart'],
  [/^(Paint\.Picture|PBrush)\b/i, 'Bitmap Image'],
  [/^Package\b/i, 'Package'],
]

export function oleCaption(progId: string | undefined): string {
  const id = progId?.trim() ?? ''
  if (id.length === 0) return 'Embedded Object'
  for (const [pattern, name] of OLE_PROG_ID_NAMES) {
    if (pattern.test(id)) return name
  }
  return id.replace(/(\.\d+)+$/, '')
}

export function oleFrameStyle(visual: {
  readonly lineColor?: string | undefined
  readonly fillColor?: string | undefined
}): { readonly border?: string; readonly background?: string } {
  const style: { border?: string; background?: string } = {}
  if (visual.lineColor !== undefined && visual.lineColor !== 'none') {
    style.border = `1px solid ${visual.lineColor}`
  }
  if (visual.fillColor !== undefined && visual.fillColor !== 'none') {
    style.background = visual.fillColor
  }
  return style
}
