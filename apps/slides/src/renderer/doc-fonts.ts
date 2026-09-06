import type { EmbeddedFontFace } from '@genoffice/pptx-engine'

interface LoadableFontFace {
  load(): Promise<unknown>
}

interface EmbeddedFontRuntime<T extends LoadableFontFace> {
  readonly FontFaceCtor: new (
    family: string,
    source: ArrayBuffer,
    descriptors: { weight: string; style: string },
  ) => T
  readonly fontSet: { add(face: T): unknown }
}

type BundledFontAssetReader = (fileName: string) => Promise<ArrayBuffer>

const bundledFaces = [
  { fileName: 'Carlito-Regular.ttf', weight: '400', style: 'normal' },
  { fileName: 'Carlito-Bold.ttf', weight: '700', style: 'normal' },
  { fileName: 'Carlito-Italic.ttf', weight: '400', style: 'italic' },
  { fileName: 'Carlito-BoldItalic.ttf', weight: '700', style: 'italic' },
] as const

/** Load the metric-compatible Calibri fallback from the shared host asset store. */
export async function registerBundledFallbackFontFaces<T extends LoadableFontFace>(
  read: BundledFontAssetReader,
  runtime: EmbeddedFontRuntime<T>,
): Promise<number> {
  let loaded = 0
  const assets = new Map<string, Promise<ArrayBuffer>>()
  for (const item of bundledFaces) {
    const data = assets.get(item.fileName) ?? read(item.fileName)
    assets.set(item.fileName, data)
    for (const family of ['Carlito', 'Carlito GO']) {
      try {
        const face = new runtime.FontFaceCtor(family, await data, {
          weight: item.weight,
          style: item.style,
        })
        await face.load()
        runtime.fontSet.add(face)
        loaded += 1
      } catch {
        // Standalone hosts may not expose bundled assets; local font fallbacks remain usable.
      }
    }
  }
  return loaded
}

/** Register the usable font faces embedded in a PPTX before its first canvas render. */
export async function registerEmbeddedFontFaces<T extends LoadableFontFace>(
  faces: readonly EmbeddedFontFace[],
  runtime: EmbeddedFontRuntime<T>,
): Promise<number> {
  let loaded = 0
  for (const item of faces) {
    const bold = item.style === 'bold' || item.style === 'boldItalic'
    const italic = item.style === 'italic' || item.style === 'boldItalic'
    try {
      const face = new runtime.FontFaceCtor(item.typeface, item.sfnt.slice().buffer, {
        weight: bold ? '700' : '400',
        style: italic ? 'italic' : 'normal',
      })
      await face.load()
      runtime.fontSet.add(face)
      loaded += 1
    } catch {
      // A malformed or unsupported embedded face must not prevent the deck from opening.
    }
  }
  return loaded
}

export async function registerBrowserEmbeddedFontFaces(
  faces: readonly EmbeddedFontFace[],
): Promise<number> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined' || !document.fonts)
    return 0
  return registerEmbeddedFontFaces(faces, {
    FontFaceCtor: FontFace,
    fontSet: document.fonts,
  })
}

export async function registerBrowserBundledFallbackFontFaces(
  read: BundledFontAssetReader,
): Promise<number> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined' || !document.fonts)
    return 0
  return registerBundledFallbackFontFaces(read, {
    FontFaceCtor: FontFace,
    fontSet: document.fonts,
  })
}
