import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { nativeFindStrings } from './strings-native-find'
import { nativeFileStrings } from './strings-native-file'
import { nativeNameStrings } from './strings-native-names'
import { nativeWatchStrings } from './strings-native-watch'
import { nativeMergeStrings } from './strings-native-merge'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh, ...nativeFindStrings.zh, ...nativeFileStrings.zh, ...nativeNameStrings.zh, ...nativeWatchStrings.zh, ...nativeMergeStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en, ...nativeFindStrings.en, ...nativeFileStrings.en, ...nativeNameStrings.en, ...nativeWatchStrings.en, ...nativeMergeStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja, ...nativeFindStrings.ja, ...nativeFileStrings.ja, ...nativeNameStrings.ja, ...nativeWatchStrings.ja, ...nativeMergeStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko, ...nativeFindStrings.ko, ...nativeFileStrings.ko, ...nativeNameStrings.ko, ...nativeWatchStrings.ko, ...nativeMergeStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr, ...nativeFindStrings.fr, ...nativeFileStrings.fr, ...nativeNameStrings.fr, ...nativeWatchStrings.fr, ...nativeMergeStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de, ...nativeFindStrings.de, ...nativeFileStrings.de, ...nativeNameStrings.de, ...nativeWatchStrings.de, ...nativeMergeStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es, ...nativeFindStrings.es, ...nativeFileStrings.es, ...nativeNameStrings.es, ...nativeWatchStrings.es, ...nativeMergeStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th, ...nativeFindStrings.th, ...nativeFileStrings.th, ...nativeNameStrings.th, ...nativeWatchStrings.th, ...nativeMergeStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id, ...nativeFindStrings.id, ...nativeFileStrings.id, ...nativeNameStrings.id, ...nativeWatchStrings.id, ...nativeMergeStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru, ...nativeFindStrings.ru, ...nativeFileStrings.ru, ...nativeNameStrings.ru, ...nativeWatchStrings.ru, ...nativeMergeStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar, ...nativeFindStrings.ar, ...nativeFileStrings.ar, ...nativeNameStrings.ar, ...nativeWatchStrings.ar, ...nativeMergeStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt, ...nativeFindStrings.pt, ...nativeFileStrings.pt, ...nativeNameStrings.pt, ...nativeWatchStrings.pt, ...nativeMergeStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it, ...nativeFindStrings.it, ...nativeFileStrings.it, ...nativeNameStrings.it, ...nativeWatchStrings.it, ...nativeMergeStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl, ...nativeFindStrings.pl, ...nativeFileStrings.pl, ...nativeNameStrings.pl, ...nativeWatchStrings.pl, ...nativeMergeStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl, ...nativeFindStrings.nl, ...nativeFileStrings.nl, ...nativeNameStrings.nl, ...nativeWatchStrings.nl, ...nativeMergeStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms, ...nativeFindStrings.ms, ...nativeFileStrings.ms, ...nativeNameStrings.ms, ...nativeWatchStrings.ms, ...nativeMergeStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he, ...nativeFindStrings.he, ...nativeFileStrings.he, ...nativeNameStrings.he, ...nativeWatchStrings.he, ...nativeMergeStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi, ...nativeFindStrings.hi, ...nativeFileStrings.hi, ...nativeNameStrings.hi, ...nativeWatchStrings.hi, ...nativeMergeStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
    ...nativeFindStrings['zh-TW'],
    ...nativeFileStrings['zh-TW'],
    ...nativeNameStrings['zh-TW'],
    ...nativeWatchStrings['zh-TW'],
    ...nativeMergeStrings['zh-TW'],
  },
}
