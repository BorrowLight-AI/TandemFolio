import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'

import { setDocxProtection } from '../src/renderer/editor/protection-actions'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

describe('explicit DOCX protection state', () => {
  it('sets passwordless protection, reports a no-op, and saves/reopens both final states', async () => {
    const enabled = await setDocxProtection(null, { enabled: true, password: null })
    expect(enabled).toEqual({
      ok: true,
      enabled: true,
      passwordProtected: false,
      changed: true,
      protection: { edit: 'readOnly', enforced: true },
    })
    if (!enabled.ok) throw new Error(enabled.message)
    expect(await setDocxProtection(enabled.protection, { enabled: true, password: null })).toEqual({
      ok: true,
      enabled: true,
      passwordProtected: false,
      changed: false,
      protection: enabled.protection,
    })

    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Protected</w:t></w:r></w:p>' }),
    )
    const blocks = parsed.blocks
      .filter((block) => !block.hidden && block.docxIndex !== null)
      .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! }))
    const protectedDoc = await parseDocx(
      await saveDocx(parsed, blocks, { protection: enabled.protection }),
    )
    expect(protectedDoc.protection).toEqual({ edit: 'readOnly', enforced: true })

    const disabled = await setDocxProtection(protectedDoc.protection, {
      enabled: false,
      password: null,
    })
    expect(disabled).toMatchObject({
      ok: true,
      enabled: false,
      passwordProtected: false,
      changed: true,
      protection: null,
    })
    if (!disabled.ok) throw new Error(disabled.message)
    const reopened = await parseDocx(
      await saveDocx(protectedDoc, blocks, { protection: disabled.protection }),
    )
    expect(reopened.protection).toBeNull()
  })

  it('requires the current password before disabling password-protected editing', async () => {
    const enabled = await setDocxProtection(null, { enabled: true, password: 's3cret' })
    expect(enabled).toMatchObject({
      ok: true,
      enabled: true,
      passwordProtected: true,
      changed: true,
    })
    if (!enabled.ok) throw new Error(enabled.message)
    await expect(
      setDocxProtection(enabled.protection, { enabled: false, password: 'wrong' }),
    ).resolves.toMatchObject({ ok: false, error: 'execution_failed' })
    await expect(
      setDocxProtection(enabled.protection, { enabled: false, password: 's3cret' }),
    ).resolves.toMatchObject({ ok: true, enabled: false, changed: true, protection: null })
  }, 40_000)
})
