import type { Lang } from '@genoffice/i18n'

type LangDicts<D extends Record<string, string>> = { zh: D } & {
  [L in Exclude<Lang, 'zh'>]: Record<keyof D, string>
}

/** Type-only locale coverage check without coupling the deferred copy to the editor entry. */
export function defineStrings<D extends Record<string, string>>(dicts: LangDicts<D>): LangDicts<D> {
  return dicts
}
