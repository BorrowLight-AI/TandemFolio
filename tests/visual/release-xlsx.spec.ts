import { expect, test } from '@playwright/test'

test('XLSX reports four cold-start phases only after the mounted workbook commits', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => {
    const state = window.__codexVisualHost
    return state?.initialized && state.firstPollAt != null
  })

  const observed = await page.evaluate(() => {
    const state = window.__codexVisualHost
    return {
      hostColdStartMs: state.firstPollAt! - state.startedAt,
      trace: state.firstPollArguments?.startupTrace,
      snapshot: state.firstPollSnapshot,
    }
  })

  expect(observed.trace).toEqual({
    operation: 'xlsx.editor.cold_start',
    phases: {
      bootstrapMs: expect.any(Number),
      univerCreateMs: expect.any(Number),
      worksheetInstallMs: expect.any(Number),
      firstCommitMs: expect.any(Number),
    },
    bootstrapPhases: {
      resourceReceiveMs: expect.any(Number),
      moduleGraphReadyMs: expect.any(Number),
      reactMountMs: expect.any(Number),
    },
  })
  const phases = (
    observed.trace as {
      phases: Record<string, number>
    }
  ).phases
  for (const value of Object.values(phases)) expect(value).toBeGreaterThanOrEqual(0)
  expect(Object.values(phases).reduce((sum, value) => sum + value, 0)).toBeLessThanOrEqual(
    observed.hostColdStartMs,
  )
  const bootstrapPhases = (
    observed.trace as {
      bootstrapPhases: Record<string, number>
    }
  ).bootstrapPhases
  for (const value of Object.values(bootstrapPhases)) expect(value).toBeGreaterThanOrEqual(0)
  expect(Object.values(bootstrapPhases).reduce((sum, value) => sum + value, 0)).toBeCloseTo(
    phases.bootstrapMs,
    5,
  )
  expect(observed.snapshot).toEqual({
    appShellVisible: true,
    canvasCount: expect.any(Number),
    activeSheet: 'Sheet1',
  })
  expect(observed.snapshot!.canvasCount).toBeGreaterThan(0)
})

test('an offscreen XLSX releases canvas backing stores and resumes the same workbook', async ({
  page,
}) => {
  await page.goto('/?format=xlsx&width=1332&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.firstPollAt != null)

  const editor = page.frameLocator('#editor-frame')
  const canvasPixels = () =>
    editor.locator('#univer-container').evaluate((container) =>
      [...container.querySelectorAll('canvas')].reduce(
        (sum, canvas) => sum + canvas.width * canvas.height,
        0,
      ),
    )
  await expect.poll(canvasPixels).toBeGreaterThan(0)

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: true }),
    )
  })
  await expect.poll(canvasPixels).toBe(0)

  await editor.locator('#root').evaluate((root) => {
    root.dispatchEvent(
      Object.assign(new Event('contentvisibilityautostatechange'), { skipped: false }),
    )
  })
  await expect.poll(canvasPixels).toBeGreaterThan(0)

  const state = await page.evaluate(() => window.__codexVisualHost)
  expect(state.editorLoads).toBe(1)
  expect(state.errors).toEqual([])
})
