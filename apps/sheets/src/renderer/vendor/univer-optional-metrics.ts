import { createIdentifier } from '@univerjs/core'

/**
 * Univer Sheets UI injects this token as an optional render-metrics sink.
 * Lite deliberately registers no implementation, so no event can leave the
 * mounted editor and the upstream telemetry package stays out of the build.
 */
export interface ITelemetryService {
  capture(eventName: string, properties?: Record<string, unknown>): void
}

export const ITelemetryService = createIdentifier<ITelemetryService>('lite.optional-render-metrics')
