import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { nativeFindStrings } from './strings-native-find'
import { nativeFileStrings } from './strings-native-file'
import { nativeNameStrings } from './strings-native-names'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh, ...nativeFindStrings.zh, ...nativeFileStrings.zh, ...nativeNameStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en, ...nativeFindStrings.en, ...nativeFileStrings.en, ...nativeNameStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja, ...nativeFindStrings.ja, ...nativeFileStrings.ja, ...nativeNameStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko, ...nativeFindStrings.ko, ...nativeFileStrings.ko, ...nativeNameStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr, ...nativeFindStrings.fr, ...nativeFileStrings.fr, ...nativeNameStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de, ...nativeFindStrings.de, ...nativeFileStrings.de, ...nativeNameStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es, ...nativeFindStrings.es, ...nativeFileStrings.es, ...nativeNameStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th, ...nativeFindStrings.th, ...nativeFileStrings.th, ...nativeNameStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id, ...nativeFindStrings.id, ...nativeFileStrings.id, ...nativeNameStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru, ...nativeFindStrings.ru, ...nativeFileStrings.ru, ...nativeNameStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar, ...nativeFindStrings.ar, ...nativeFileStrings.ar, ...nativeNameStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt, ...nativeFindStrings.pt, ...nativeFileStrings.pt, ...nativeNameStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it, ...nativeFindStrings.it, ...nativeFileStrings.it, ...nativeNameStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl, ...nativeFindStrings.pl, ...nativeFileStrings.pl, ...nativeNameStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl, ...nativeFindStrings.nl, ...nativeFileStrings.nl, ...nativeNameStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms, ...nativeFindStrings.ms, ...nativeFileStrings.ms, ...nativeNameStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he, ...nativeFindStrings.he, ...nativeFileStrings.he, ...nativeNameStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi, ...nativeFindStrings.hi, ...nativeFileStrings.hi, ...nativeNameStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
    ...nativeFindStrings['zh-TW'],
    ...nativeFileStrings['zh-TW'],
    ...nativeNameStrings['zh-TW'],
  },
}
