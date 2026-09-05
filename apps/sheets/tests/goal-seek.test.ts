import { describe, expect, it } from 'vitest'

import { solveGoalSeek } from '../src/renderer/goal-seek'
import { executeXlsxOperation } from '../src/renderer/operations/registry'
import { journalSuppression } from '../src/renderer/univer-state'

function runtimeFor(evaluate: (x: number) => number | string, initial = 0) {
  let changing = initial
  let calculationEnd: (() => void) | null = null
  const writes: { value: number; suppressed: boolean }[] = []
  const target = { getFormula: () => '=2*B1+1', getValue: () => evaluate(changing) }
  const by = {
    getFormula: () => '',
    getValue: () => changing,
    setValue: (value: number | { v: null }) => {
      changing = typeof value === 'number' ? value : 0
      writes.push({ value: changing, suppressed: journalSuppression.active })
      queueMicrotask(() => calculationEnd?.())
    },
  }
  const worksheet = {
    getRange: (row: number) => (row === 0 ? target : by),
  }
  return {
    runtime: {
      univer: {
        __getInjector: () => ({
          get: () => ({ beforeCommandExecuted: () => undefined }),
        }),
      },
      univerAPI: {
        getActiveWorkbook: () => ({ getActiveSheet: () => worksheet }),
        getFormula: () => ({
          calculationEnd: (callback: () => void) => {
            calculationEnd = callback
            return { dispose: () => { calculationEnd = null } }
          },
        }),
      },
    },
    value: () => changing,
    writes,
  }
}

describe('Goal Seek', () => {
  it('solves a numeric target through the mounted worksheet', async () => {
    const { runtime, value, writes } = runtimeFor((x) => 2 * x + 1)
    const result = await solveGoalSeek(runtime as never, { setCell: 'A1', toValue: 11, byCell: 'A2' })
    expect(result.found).toBe(true)
    expect(result.solution).toBeCloseTo(5)
    expect(value()).toBeCloseTo(5)
    expect(writes.filter((write) => !write.suppressed)).toEqual([{ value: 5, suppressed: false }])
  })

  it('restores the original changing value after a nonnumeric target result', async () => {
    const { runtime, value, writes } = runtimeFor(() => '#VALUE!', 7)
    await expect(solveGoalSeek(runtime as never, { setCell: 'A1', toValue: 11, byCell: 'A2' })).rejects.toThrow()
    expect(value()).toBe(7)
    expect(writes.every((write) => write.suppressed)).toBe(true)
  })

  it('is exposed as a typed XLSX calculation operation', async () => {
    const { runtime } = runtimeFor((x) => 3 * x)
    const result = await executeXlsxOperation(
      {
        operation: 'xlsx.calculation.goal_seek',
        arguments: { setCell: 'A1', toValue: 12, byCell: 'A2' },
      },
      { runtime: () => runtime as never },
    )
    expect(result).toMatchObject({
      handled: true,
      ok: true,
      output: { found: true, solution: 4 },
    })
  })
})
