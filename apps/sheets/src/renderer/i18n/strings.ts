import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { nativeFindStrings } from './strings-native-find'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh, ...nativeFindStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en, ...nativeFindStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja, ...nativeFindStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko, ...nativeFindStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr, ...nativeFindStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de, ...nativeFindStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es, ...nativeFindStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th, ...nativeFindStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id, ...nativeFindStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru, ...nativeFindStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar, ...nativeFindStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt, ...nativeFindStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it, ...nativeFindStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl, ...nativeFindStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl, ...nativeFindStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms, ...nativeFindStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he, ...nativeFindStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi, ...nativeFindStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
    ...nativeFindStrings['zh-TW'],
  },
}
