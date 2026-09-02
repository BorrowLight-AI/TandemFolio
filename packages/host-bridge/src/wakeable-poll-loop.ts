export const LIVE_POLL_WAIT_MS = 10_000
export const LIVE_POLL_RETRY_MS = 500

interface PollClock {
  setTimeout(handler: () => void, timeout: number): number
  clearTimeout(handle: number): void
}

export interface WakeablePollLoop {
  start(): void
  wake(): void
  pause(): void
  resume(): void
  stop(): void
}

export function createWakeablePollLoop(
  poll: (waitMs: number) => Promise<void>,
  clock: PollClock = window,
): WakeablePollLoop {
  let started = false
  let stopped = false
  let paused = false
  let running = false
  let rerun = false
  let bootstrapped = false
  let timer = 0

  const schedule = (delayMs: number): void => {
    if (stopped || paused) return
    clock.clearTimeout(timer)
    timer = clock.setTimeout(() => void run(), delayMs)
  }

  const run = async (): Promise<void> => {
    if (stopped || paused) return
    if (running) {
      rerun = true
      return
    }
    running = true
    let failed = false
    try {
      await poll(bootstrapped ? LIVE_POLL_WAIT_MS : 0)
      bootstrapped = true
    } catch {
      failed = true
    } finally {
      running = false
      if (stopped || paused) return
      if (rerun) {
        rerun = false
        schedule(0)
      } else {
        schedule(failed ? LIVE_POLL_RETRY_MS : 0)
      }
    }
  }

  return {
    start() {
      if (stopped || started) return
      started = true
      schedule(0)
    },
    wake() {
      if (stopped || paused) return
      if (!started) {
        this.start()
      } else if (running) {
        rerun = true
      } else {
        schedule(0)
      }
    },
    pause() {
      if (stopped || paused) return
      paused = true
      clock.clearTimeout(timer)
    },
    resume() {
      if (stopped || !paused) return
      paused = false
      if (!started) return
      if (running) rerun = true
      else schedule(0)
    },
    stop() {
      stopped = true
      clock.clearTimeout(timer)
    },
  }
}
