import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { nativeFindStrings } from './strings-native-find'
import { nativeFileStrings } from './strings-native-file'
import { nativeNameStrings } from './strings-native-names'
import { nativeWatchStrings } from './strings-native-watch'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh, ...nativeFindStrings.zh, ...nativeFileStrings.zh, ...nativeNameStrings.zh, ...nativeWatchStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en, ...nativeFindStrings.en, ...nativeFileStrings.en, ...nativeNameStrings.en, ...nativeWatchStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja, ...nativeFindStrings.ja, ...nativeFileStrings.ja, ...nativeNameStrings.ja, ...nativeWatchStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko, ...nativeFindStrings.ko, ...nativeFileStrings.ko, ...nativeNameStrings.ko, ...nativeWatchStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr, ...nativeFindStrings.fr, ...nativeFileStrings.fr, ...nativeNameStrings.fr, ...nativeWatchStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de, ...nativeFindStrings.de, ...nativeFileStrings.de, ...nativeNameStrings.de, ...nativeWatchStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es, ...nativeFindStrings.es, ...nativeFileStrings.es, ...nativeNameStrings.es, ...nativeWatchStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th, ...nativeFindStrings.th, ...nativeFileStrings.th, ...nativeNameStrings.th, ...nativeWatchStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id, ...nativeFindStrings.id, ...nativeFileStrings.id, ...nativeNameStrings.id, ...nativeWatchStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru, ...nativeFindStrings.ru, ...nativeFileStrings.ru, ...nativeNameStrings.ru, ...nativeWatchStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar, ...nativeFindStrings.ar, ...nativeFileStrings.ar, ...nativeNameStrings.ar, ...nativeWatchStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt, ...nativeFindStrings.pt, ...nativeFileStrings.pt, ...nativeNameStrings.pt, ...nativeWatchStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it, ...nativeFindStrings.it, ...nativeFileStrings.it, ...nativeNameStrings.it, ...nativeWatchStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl, ...nativeFindStrings.pl, ...nativeFileStrings.pl, ...nativeNameStrings.pl, ...nativeWatchStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl, ...nativeFindStrings.nl, ...nativeFileStrings.nl, ...nativeNameStrings.nl, ...nativeWatchStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms, ...nativeFindStrings.ms, ...nativeFileStrings.ms, ...nativeNameStrings.ms, ...nativeWatchStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he, ...nativeFindStrings.he, ...nativeFileStrings.he, ...nativeNameStrings.he, ...nativeWatchStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi, ...nativeFindStrings.hi, ...nativeFileStrings.hi, ...nativeNameStrings.hi, ...nativeWatchStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
    ...nativeFindStrings['zh-TW'],
    ...nativeFileStrings['zh-TW'],
    ...nativeNameStrings['zh-TW'],
    ...nativeWatchStrings['zh-TW'],
  },
}
