import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi },
  'zh-TW': { ...appStrings['zh-TW'], ...dialogStrings['zh-TW'] },
}
