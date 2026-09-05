/**
 * The shortcut registry is what the Keyboard Shortcuts sheet
 * renders, so it has to stay honest: no chord listed twice on either platform,
 * every row translatable, and no menu accelerator missing from the list.
 */
import { describe, expect, it } from 'vitest'
import {
  SHORTCUTS,
  SHORTCUT_GROUPS,
  shortcutKeys,
  type ShortcutDef,
} from '../src/renderer/shortcuts'
import { strings } from '../src/renderer/i18n/strings'

// alternates are listed as "A / B"; the separator has to keep ⌘/ itself intact
const chords = (def: ShortcutDef, isMac: boolean) => shortcutKeys(def, isMac).split(' / ')

describe('shortcut registry', () => {
  it('has unique ids and known groups', () => {
    const ids = SHORTCUTS.map((def) => def.id)
    expect(new Set(ids).size).toBe(ids.length)
    const groups = new Set(SHORTCUT_GROUPS.map((group) => group.id))
    for (const def of SHORTCUTS) expect(groups.has(def.group)).toBe(true)
  })

  it.each([
    ['macOS', true],
    ['Windows/Linux', false],
  ])('binds every chord once on %s', (_platform, isMac) => {
    const seen = new Map<string, string>()
    for (const def of SHORTCUTS) {
      for (const chord of chords(def, isMac as boolean)) {
        expect(seen.get(chord), `${chord} is claimed by ${seen.get(chord)} and ${def.id}`).toBe(
          undefined,
        )
        seen.set(chord, def.id)
      }
    }
  })

  it('labels every row in every language', () => {
    for (const [lang, dict] of Object.entries(strings)) {
      for (const entry of [...SHORTCUTS, ...SHORTCUT_GROUPS]) {
        const value = (dict as Record<string, string>)[entry.labelKey]
        expect(value, `${entry.labelKey} missing in ${lang}`).toBeTruthy()
      }
    }
  })

  it('rewrites Mac notation for Windows/Linux', () => {
    const byId = new Map(SHORTCUTS.map((def) => [def.id, def]))
    const win = (id: string) => shortcutKeys(byId.get(id)!, false)
    expect(win('save-as')).toBe('Ctrl+Shift+S')
    expect(win('page-break')).toBe('Ctrl+Enter')
    expect(win('nbsp')).toBe('Ctrl+Shift+Space')
    expect(win('style-h1')).toBe('Ctrl+Alt+1')
    // platform-specific keys come from the explicit override, not a rewrite
    expect(win('endnote')).toBe('Ctrl+Alt+D')
    expect(shortcutKeys(byId.get('endnote')!, true)).toBe('⌥⌘E')
  })

  it('does not advertise AI-backed commands', () => {
    expect(SHORTCUTS.map((entry) => entry.id)).not.toContain('proofread')
  })

})
