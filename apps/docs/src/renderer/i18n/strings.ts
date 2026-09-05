import { appStrings } from './strings-app'
import { editorStrings } from './strings-editor'
import { ribbonStrings } from './strings-ribbon'
import { shortcutStringsEn, shortcutStringsZh } from './strings-shortcuts'
import { docxNativeStringsEn, docxNativeStringsZh } from './strings-docx-native'

export const strings = {
  zh: { ...appStrings.zh, ...ribbonStrings.zh, ...editorStrings.zh, ...shortcutStringsZh, ...docxNativeStringsZh },
  en: { ...appStrings.en, ...ribbonStrings.en, ...editorStrings.en, ...shortcutStringsEn, ...docxNativeStringsEn },
  ja: { ...appStrings.ja, ...ribbonStrings.ja, ...editorStrings.ja, ...shortcutStringsEn, ...docxNativeStringsEn },
  ko: { ...appStrings.ko, ...ribbonStrings.ko, ...editorStrings.ko, ...shortcutStringsEn, ...docxNativeStringsEn },
  fr: { ...appStrings.fr, ...ribbonStrings.fr, ...editorStrings.fr, ...shortcutStringsEn, ...docxNativeStringsEn },
  de: { ...appStrings.de, ...ribbonStrings.de, ...editorStrings.de, ...shortcutStringsEn, ...docxNativeStringsEn },
  es: { ...appStrings.es, ...ribbonStrings.es, ...editorStrings.es, ...shortcutStringsEn, ...docxNativeStringsEn },
  th: { ...appStrings.th, ...ribbonStrings.th, ...editorStrings.th, ...shortcutStringsEn, ...docxNativeStringsEn },
  id: { ...appStrings.id, ...ribbonStrings.id, ...editorStrings.id, ...shortcutStringsEn, ...docxNativeStringsEn },
  ru: { ...appStrings.ru, ...ribbonStrings.ru, ...editorStrings.ru, ...shortcutStringsEn, ...docxNativeStringsEn },
  ar: { ...appStrings.ar, ...ribbonStrings.ar, ...editorStrings.ar, ...shortcutStringsEn, ...docxNativeStringsEn },
  pt: { ...appStrings.pt, ...ribbonStrings.pt, ...editorStrings.pt, ...shortcutStringsEn, ...docxNativeStringsEn },
  it: { ...appStrings.it, ...ribbonStrings.it, ...editorStrings.it, ...shortcutStringsEn, ...docxNativeStringsEn },
  pl: { ...appStrings.pl, ...ribbonStrings.pl, ...editorStrings.pl, ...shortcutStringsEn, ...docxNativeStringsEn },
  nl: { ...appStrings.nl, ...ribbonStrings.nl, ...editorStrings.nl, ...shortcutStringsEn, ...docxNativeStringsEn },
  ms: { ...appStrings.ms, ...ribbonStrings.ms, ...editorStrings.ms, ...shortcutStringsEn, ...docxNativeStringsEn },
  he: { ...appStrings.he, ...ribbonStrings.he, ...editorStrings.he, ...shortcutStringsEn, ...docxNativeStringsEn },
  hi: { ...appStrings.hi, ...ribbonStrings.hi, ...editorStrings.hi, ...shortcutStringsEn, ...docxNativeStringsEn },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...ribbonStrings['zh-TW'],
    ...editorStrings['zh-TW'],
    ...shortcutStringsZh,
    ...docxNativeStringsZh,
  },
}
