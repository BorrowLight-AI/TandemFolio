import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests/visual',
  timeout: 45_000,
  expect: {
    timeout: 15_000,
    toHaveScreenshot: {
      animations: 'disabled',
      caret: 'hide',
      maxDiffPixelRatio: 0.03,
      threshold: 0.2,
    },
  },
  fullyParallel: false,
  // These integration specs share one host server and deliberately restart
  // broker state. Cross-file workers can interrupt another format's open/save
  // lifecycle, so the release suite must exercise that shared authority serially.
  workers: 1,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  reporter: [['list']],
  snapshotPathTemplate: '{testDir}/__screenshots__/{testFilePath}/{arg}{ext}',
  use: {
    // Missing controls are deterministic test failures; do not let one locator
    // consume the full per-test budget before CI can report the real cause.
    actionTimeout: 15_000,
    baseURL: 'http://127.0.0.1:4178',
    browserName: 'chromium',
    colorScheme: 'light',
    deviceScaleFactor: 1,
    locale: 'zh-CN',
    reducedMotion: 'reduce',
    viewport: { width: 1440, height: 1360 },
  },
  webServer: {
    command: 'node tests/visual/serve-host.mjs',
    url: 'http://127.0.0.1:4178/health',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
})
