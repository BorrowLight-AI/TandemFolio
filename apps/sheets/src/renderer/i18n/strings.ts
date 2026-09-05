import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { pageBreakStrings } from './strings-page-breaks'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...pageBreakStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...pageBreakStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...pageBreakStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...pageBreakStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...pageBreakStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...pageBreakStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...pageBreakStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...pageBreakStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...pageBreakStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...pageBreakStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...pageBreakStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...pageBreakStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...pageBreakStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...pageBreakStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...pageBreakStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...pageBreakStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...pageBreakStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...pageBreakStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...pageBreakStrings['zh-TW'],
  },
}
