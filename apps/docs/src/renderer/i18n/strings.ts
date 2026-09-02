import { appStrings } from './strings-app'
import { editorStrings } from './strings-editor'
import { ribbonStrings } from './strings-ribbon'

export const strings = {
  zh: { ...appStrings.zh, ...ribbonStrings.zh, ...editorStrings.zh },
  en: { ...appStrings.en, ...ribbonStrings.en, ...editorStrings.en },
  ja: { ...appStrings.ja, ...ribbonStrings.ja, ...editorStrings.ja },
  ko: { ...appStrings.ko, ...ribbonStrings.ko, ...editorStrings.ko },
  fr: { ...appStrings.fr, ...ribbonStrings.fr, ...editorStrings.fr },
  de: { ...appStrings.de, ...ribbonStrings.de, ...editorStrings.de },
  es: { ...appStrings.es, ...ribbonStrings.es, ...editorStrings.es },
  th: { ...appStrings.th, ...ribbonStrings.th, ...editorStrings.th },
  id: { ...appStrings.id, ...ribbonStrings.id, ...editorStrings.id },
  ru: { ...appStrings.ru, ...ribbonStrings.ru, ...editorStrings.ru },
  ar: { ...appStrings.ar, ...ribbonStrings.ar, ...editorStrings.ar },
  pt: { ...appStrings.pt, ...ribbonStrings.pt, ...editorStrings.pt },
  it: { ...appStrings.it, ...ribbonStrings.it, ...editorStrings.it },
  pl: { ...appStrings.pl, ...ribbonStrings.pl, ...editorStrings.pl },
  nl: { ...appStrings.nl, ...ribbonStrings.nl, ...editorStrings.nl },
  ms: { ...appStrings.ms, ...ribbonStrings.ms, ...editorStrings.ms },
  he: { ...appStrings.he, ...ribbonStrings.he, ...editorStrings.he },
  hi: { ...appStrings.hi, ...ribbonStrings.hi, ...editorStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...ribbonStrings['zh-TW'],
    ...editorStrings['zh-TW'],
  },
}
