import { expect, test } from '@playwright/test'
import JSZip from 'jszip'

test('a new blank XLSX accepts a native chart and saves it', async ({ page }) => {
  await page.goto('/?format=xlsx&width=720&height=900')
  await page.waitForFunction(() => window.__codexVisualHost?.initialized)
  const editor = page.frameLocator('#editor-frame')

  const commands = [
    {
      commandId: 'blank-chart-sheet',
      operation: 'xlsx.sheet.add',
      arguments: { name: 'Dashboard' },
    },
    {
      commandId: 'blank-chart-values',
      operation: 'xlsx.range.set_values',
      arguments: {
        sheet: 'Dashboard',
        range: 'A1:B3',
        values: [
          ['Quarter', 'Revenue'],
          ['Q1', 10],
          ['Q2', 20],
        ],
      },
    },
    {
      commandId: 'blank-chart-add',
      operation: 'xlsx.chart.add',
      arguments: { sheet: 'Dashboard', dataRange: 'A1:B3', type: 'line', anchorCell: 'D2' },
    },
    {
      commandId: 'blank-chart-heading-style',
      operation: 'xlsx.range.apply_cell_style',
      arguments: { sheet: 'Dashboard', range: 'A1:B1', preset: 'heading-1' },
    },
    {
      commandId: 'blank-chart-number-format',
      operation: 'xlsx.range.set_number_format',
      arguments: { sheet: 'Dashboard', range: 'B2:B3', pattern: '¥#,##0.00' },
    },
    {
      commandId: 'blank-chart-theme',
      operation: 'xlsx.document.set_theme',
      arguments: { mode: 'theme', scheme: 'office' },
    },
  ] as const

  for (const [baseRevision, command] of commands.entries()) {
    await page.evaluate(
      ({ baseRevision, command }) => {
        window.__codexVisualHost.enqueueCommand({
          ...command,
          baseRevision,
        })
      },
      { baseRevision, command },
    )
    await expect
      .poll(() =>
        page.evaluate(
          (commandId) =>
            window.__codexVisualHost.acknowledgements.find(
              (entry) => entry.commandId === commandId,
            ),
          command.commandId,
        ),
      )
      .toMatchObject({ ok: true, revision: baseRevision + 1 })
  }

  await expect(editor.locator('.xlsx-chart')).toHaveCount(1)

  const downloadPromise = page.waitForEvent('download')
  await page.evaluate(() => {
    window.__codexVisualHost.enqueueCommand({
      commandId: 'blank-chart-save',
      baseRevision: 6,
      operation: 'xlsx.document.save',
      arguments: {},
    })
  })
  await expect
    .poll(() =>
      page.evaluate(() =>
        window.__codexVisualHost.acknowledgements.find(
          (entry) => entry.commandId === 'blank-chart-save',
        ),
      ),
    )
    .toMatchObject({ ok: true, revision: 7 })

  const savedPath = await (await downloadPromise).path()
  expect(savedPath).not.toBeNull()

  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(savedPath!)),
  )
  expect(zip.file('xl/charts/chart1.xml')).not.toBeNull()
  expect(zip.file('xl/styles.xml')).not.toBeNull()
  expect(zip.file('xl/theme/theme1.xml')).not.toBeNull()
})
