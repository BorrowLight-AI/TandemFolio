import { execFileSync } from 'node:child_process'
import { platform } from 'node:os'

export type CaptureSchedulingExecutor = (command: string, arguments_: string[]) => void

const executeSchedulingCommand: CaptureSchedulingExecutor = (command, arguments_) => {
  execFileSync(command, arguments_, { stdio: 'ignore' })
}

/** Prevent a foreground macOS release capture from inheriting an app's background priority. */
export function ensureReleaseCaptureScheduling(
  hostPlatform = platform(),
  pid = process.pid,
  execute: CaptureSchedulingExecutor = executeSchedulingCommand,
): void {
  if (hostPlatform !== 'darwin') return
  execute('taskpolicy', ['-B', '-p', String(pid)])
}
