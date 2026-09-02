/// <reference types="vite/client" />

import type { DesktopApi } from '../shared/lite-api'

declare global {
  interface Window {
    desktop: DesktopApi
    __genofficeLite?: {
      editor: unknown
      openPath(path: string): Promise<void>
      save(): Promise<boolean | undefined>
      getStatus(): string
      exportPdfTo(path: string): Promise<void>
    }
  }
}

export {}
