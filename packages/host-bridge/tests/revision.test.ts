import { describe, expect, it } from 'vitest'
import { nextRevision, type EditorCommand } from '../src/index'

const command: EditorCommand = {
  commandId: 'command-1',
  sessionId: 'session-1',
  baseRevision: 4,
  operation: { name: 'replace_text', arguments: { from: 1, to: 2, text: 'Lite' } },
}

describe('nextRevision', () => {
  it('advances a command based on the current document revision', () => {
    expect(nextRevision(4, command)).toBe(5)
  })

  it('rejects stale commands instead of applying them to drifting state', () => {
    expect(() => nextRevision(5, command)).toThrow('revision_conflict:4:5')
  })
})
