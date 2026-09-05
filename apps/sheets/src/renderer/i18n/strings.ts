import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { nativeFindStrings } from './strings-native-find'
import { nativeFileStrings } from './strings-native-file'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh, ...nativeFindStrings.zh, ...nativeFileStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en, ...nativeFindStrings.en, ...nativeFileStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja, ...nativeFindStrings.ja, ...nativeFileStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko, ...nativeFindStrings.ko, ...nativeFileStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr, ...nativeFindStrings.fr, ...nativeFileStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de, ...nativeFindStrings.de, ...nativeFileStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es, ...nativeFindStrings.es, ...nativeFileStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th, ...nativeFindStrings.th, ...nativeFileStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id, ...nativeFindStrings.id, ...nativeFileStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru, ...nativeFindStrings.ru, ...nativeFileStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar, ...nativeFindStrings.ar, ...nativeFileStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt, ...nativeFindStrings.pt, ...nativeFileStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it, ...nativeFindStrings.it, ...nativeFileStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl, ...nativeFindStrings.pl, ...nativeFileStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl, ...nativeFindStrings.nl, ...nativeFileStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms, ...nativeFindStrings.ms, ...nativeFileStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he, ...nativeFindStrings.he, ...nativeFileStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi, ...nativeFindStrings.hi, ...nativeFileStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
    ...nativeFindStrings['zh-TW'],
    ...nativeFileStrings['zh-TW'],
  },
}
