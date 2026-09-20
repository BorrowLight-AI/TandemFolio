import { LocaleType } from '@univerjs/core'
import { describe, expect, it, vi } from 'vitest'

import { applyUniverLocale, univerLocaleFor } from '../src/renderer/univer-locales'
import type { UniverRuntime } from '../src/renderer/univer-state'

describe('univerLocaleFor', () => {
  it('maps every app language Univer has packs for', () => {
    expect(univerLocaleFor('zh')).toBe(LocaleType.ZH_CN)
    expect(univerLocaleFor('zh-TW')).toBe(LocaleType.ZH_TW)
    expect(univerLocaleFor('ja')).toBe(LocaleType.JA_JP)
    expect(univerLocaleFor('ko')).toBe(LocaleType.KO_KR)
    expect(univerLocaleFor('ar')).toBe(LocaleType.AR_SA)
  })

  it('falls back to English where Univer ships no pack', () => {
    for (const lang of ['en', 'th', 'nl', 'ms', 'he', 'hi']) {
      expect(univerLocaleFor(lang)).toBeNull()
    }
  })

  it('zh packs resolve and localize the DV rule name', async () => {
    const pack = (await import('@univerjs/preset-sheets-data-validation/locales/zh-CN')) as {
      default: Record<string, { list?: { name?: string } }>
    }
    expect(pack.default['sheets-data-validation']?.list?.name).toBe('值必须是列表中的值')
  })

  it('switches the mounted Univer UI back to its English boot locale', async () => {
    const setLocale = vi.fn()
    const runtime = {
      univer: {
        __getInjector: () => ({ get: () => ({ setLocale }) }),
      },
    } as unknown as UniverRuntime

    await applyUniverLocale(runtime, 'en')

    expect(setLocale).toHaveBeenCalledWith(LocaleType.EN_US)
  })

  it('keeps the latest Univer UI language when an older locale pack finishes later', async () => {
    const selected: LocaleType[] = []
    const runtime = {
      univer: {
        __getInjector: () => ({
          get: () => ({
            load: () => undefined,
            setLocale: (locale: LocaleType) => selected.push(locale),
          }),
        }),
      },
    } as unknown as UniverRuntime

    const older = applyUniverLocale(runtime, 'zh')
    await applyUniverLocale(runtime, 'en')
    await older

    expect(selected.at(-1)).toBe(LocaleType.EN_US)
  })
})
