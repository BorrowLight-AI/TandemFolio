import { cssFontFamily } from './line-metrics'

export interface BundledFontFace {
  family: string
  fileName: string
  weight: '400' | '700'
  style: 'normal' | 'italic'
  unicodeRange?: string
  sizeAdjust?: string
}

export type BundledFontAssetReader = (fileName: string) => Promise<ArrayBuffer>
export type BundledFontRegistrar = (
  face: BundledFontFace,
  data: ArrayBuffer,
) => Promise<void>

const faces = (
  family: string,
  baseName: string,
  extension = 'ttf',
): BundledFontFace[] => [
  { family, fileName: `${baseName}-Regular.${extension}`, weight: '400', style: 'normal' },
  { family, fileName: `${baseName}-Bold.${extension}`, weight: '700', style: 'normal' },
  { family, fileName: `${baseName}-Italic.${extension}`, weight: '400', style: 'italic' },
  { family, fileName: `${baseName}-BoldItalic.${extension}`, weight: '700', style: 'italic' },
]

export const BUNDLED_FONT_FACES: readonly BundledFontFace[] = [
  ...faces('Carlito GO', 'Carlito'),
  ...faces('Caladea', 'Caladea'),
  ...faces('Liberation Serif', 'LiberationSerif'),
  ...faces('Liberation Sans', 'LiberationSans'),
  ...faces('Liberation Mono', 'LiberationMono'),
  {
    family: 'Noto Sans CJK SC',
    fileName: 'NotoSansCJKsc-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'Noto Serif CJK SC',
    fileName: 'NotoSerifCJKsc-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'GenOffice Sans KR',
    fileName: 'GenOfficeSansKR-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'GenOffice Serif KR',
    fileName: 'GenOfficeSerifKR-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'Noto Naskh Arabic',
    fileName: 'NotoNaskhArabic-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'Noto Naskh Arabic TA',
    fileName: 'NotoNaskhArabic-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
    sizeAdjust: '90%',
  },
  {
    family: 'Noto Sans Arabic',
    fileName: 'NotoSansArabic-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
  },
  {
    family: 'GenOffice Fullwidth TC',
    fileName: 'NotoSerifCJKsc-Regular-subset.woff2',
    weight: '400',
    style: 'normal',
    unicodeRange: 'U+FF0D,U+FF0F,U+FF3C,U+FF3F,U+FF5E',
  },
]

const descriptorsByFamily = new Map<string, readonly BundledFontFace[]>()
for (const face of BUNDLED_FONT_FACES) {
  descriptorsByFamily.set(face.family, [...(descriptorsByFamily.get(face.family) ?? []), face])
}

const hasCjk = (text: string) => /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(text)
const hasKorean = (text: string) => /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/.test(text)
const hasArabic = (text: string) => /[\u0600-\u077f\u08a0-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/.test(text)
const hasFullwidthTcShim = (text: string) => /[－／＼＿～]/.test(text)

function familyNeeded(family: string, declared: Set<string>, text: string): boolean {
  if (declared.has(family.toLowerCase())) return true
  if (family === 'Noto Sans CJK SC' || family === 'Noto Serif CJK SC') return hasCjk(text)
  if (family === 'GenOffice Sans KR' || family === 'GenOffice Serif KR') return hasKorean(text)
  if (family.includes('Arabic')) return hasArabic(text)
  if (family === 'GenOffice Fullwidth TC') return hasFullwidthTcShim(text)
  return true
}

export function requiredBundledFontFamilies(fontNames: string[], text: string): string[] {
  const declared = new Set(fontNames.map((name) => name.toLowerCase()))
  const required = new Set<string>()
  for (const fontName of fontNames) {
    for (const match of cssFontFamily(fontName).matchAll(/'([^']+)'/g)) {
      const family = match[1]
      if (descriptorsByFamily.has(family) && familyNeeded(family, declared, text)) {
        required.add(family)
      }
    }
  }
  return [...required]
}

const loadedFaces = new Set<string>()
const assetPromises = new Map<string, Promise<ArrayBuffer>>()

async function registerInDocument(face: BundledFontFace, data: ArrayBuffer): Promise<void> {
  const font = new FontFace(face.family, data, {
    weight: face.weight,
    style: face.style,
    ...(face.unicodeRange ? { unicodeRange: face.unicodeRange } : {}),
  })
  if (face.sizeAdjust) {
    Object.defineProperty(font, 'sizeAdjust', { configurable: true, value: face.sizeAdjust })
  }
  await font.load()
  document.fonts.add(font)
}

export async function loadBundledFallbackFonts(
  fontNames: string[],
  text: string,
  read: BundledFontAssetReader,
  register: BundledFontRegistrar = registerInDocument,
): Promise<void> {
  for (const family of requiredBundledFontFamilies(fontNames, text)) {
    for (const face of descriptorsByFamily.get(family) ?? []) {
      const key = [face.family, face.weight, face.style, face.unicodeRange ?? '', face.sizeAdjust ?? ''].join('|')
      if (loadedFaces.has(key)) continue
      const dataPromise = assetPromises.get(face.fileName) ?? read(face.fileName)
      assetPromises.set(face.fileName, dataPromise)
      await register(face, await dataPromise)
      loadedFaces.add(key)
    }
  }
}

export function resetBundledFontLoaderForTests(): void {
  loadedFaces.clear()
  assetPromises.clear()
}
