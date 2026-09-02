export const MARKDOWN_AUTO_SAVE_STORAGE_KEY = 'mdapp.autoSave'

export interface MarkdownPreferenceStorage {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
}

export function readMarkdownAutoSavePreference(
  storage: MarkdownPreferenceStorage = localStorage,
): boolean {
  return storage.getItem(MARKDOWN_AUTO_SAVE_STORAGE_KEY) === '1'
}

export function setMarkdownAutoSavePreference(
  enabled: boolean,
  storage: MarkdownPreferenceStorage = localStorage,
): boolean {
  storage.setItem(MARKDOWN_AUTO_SAVE_STORAGE_KEY, enabled ? '1' : '0')
  return enabled
}
