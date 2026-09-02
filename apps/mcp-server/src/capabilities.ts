import type { LiveSession } from './session-store'
import type {
  OperationContextRequirement,
  OperationDescriptor,
  OperationEffect,
  OperationId,
} from '@tandemfolio/operation-contract'
import { getRegisteredOperationDescriptors } from './operation-manifest'
import releaseReadiness from './generated/release-readiness.json'

const SUMMARY_PAGE_LIMIT = 20

interface OperationSummary {
  id: OperationId
  family: string
  summary: string
  risk: 'low' | 'medium' | 'high'
  context: readonly OperationContextRequirement[]
  effects: readonly OperationEffect[]
  availability?: OperationAvailability
}

type OperationAvailability = {
  available: boolean
  reason: 'format_mismatch' | 'editor_offline' | 'selection_required' | null
}

interface FormatCapabilityMetadata {
  format: LiveSession['format']
  ready: boolean
  displayModes: ['inline', 'fullscreen']
  defaultDisplayMode: 'fullscreen'
  localFile: { picker: boolean; path: boolean; recovery: boolean }
  showTool: string
}

interface SummaryDiscovery {
  view: 'summary'
  operations: OperationSummary[]
  pagination: {
    limit: number
    total: number
    nextCursor: string | null
  }
}

interface DetailDiscovery {
  view: 'detail'
  operation: OperationDescriptor & {
    availability?: OperationAvailability
  }
}

export type FormatCapabilities = FormatCapabilityMetadata & {
  discovery: SummaryDiscovery | DetailDiscovery
}

export interface CapabilityDiscoveryQuery {
  view: 'summary' | 'detail'
  operation?: string
  session?: LiveSession
  family?: string
  limit?: number
  cursor?: string | null
}

const catalogs: Record<LiveSession['format'], FormatCapabilityMetadata> = {
  docx: {
    format: 'docx',
    ready: releaseReadiness.ready && releaseReadiness.formats.docx,
    displayModes: ['inline', 'fullscreen'],
    defaultDisplayMode: 'fullscreen',
    localFile: { picker: true, path: true, recovery: true },
    showTool: 'office_show_editor',
  },
  markdown: {
    format: 'markdown',
    ready: releaseReadiness.ready && releaseReadiness.formats.markdown,
    displayModes: ['inline', 'fullscreen'],
    defaultDisplayMode: 'fullscreen',
    localFile: { picker: true, path: true, recovery: true },
    showTool: 'office_show_markdown_editor',
  },
  xlsx: {
    format: 'xlsx',
    ready: releaseReadiness.ready && releaseReadiness.formats.xlsx,
    displayModes: ['inline', 'fullscreen'],
    defaultDisplayMode: 'fullscreen',
    localFile: { picker: true, path: true, recovery: true },
    showTool: 'office_show_xlsx_editor',
  },
  pptx: {
    format: 'pptx',
    ready: releaseReadiness.ready && releaseReadiness.formats.pptx,
    displayModes: ['inline', 'fullscreen'],
    defaultDisplayMode: 'fullscreen',
    localFile: { picker: true, path: true, recovery: true },
    showTool: 'office_show_pptx_editor',
  },
  pdf: {
    format: 'pdf',
    ready: releaseReadiness.ready && releaseReadiness.formats.pdf,
    displayModes: ['inline', 'fullscreen'],
    defaultDisplayMode: 'fullscreen',
    localFile: { picker: true, path: true, recovery: true },
    showTool: 'office_show_pdf_editor',
  },
}

function summaryDiscovery(
  format: LiveSession['format'],
  descriptors: readonly OperationDescriptor[],
  session: LiveSession | undefined,
  family: string | undefined,
  limit: number,
  cursor: string | null | undefined,
): SummaryDiscovery | null {
  const filtered = family
    ? descriptors.filter((descriptor) => descriptor.family === family)
    : descriptors
  const cursorIndex = cursor ? filtered.findIndex((descriptor) => descriptor.id === cursor) : -1
  if (cursor && cursorIndex < 0) return null
  const start = cursorIndex + 1
  const operations = filtered.slice(start, start + limit).map((descriptor) => {
    const availability = operationAvailability(format, descriptor, session)
    return {
      id: descriptor.id,
      family: descriptor.family,
      summary: descriptor.summary,
      risk: descriptor.risk,
      context: descriptor.context,
      effects: descriptor.effects,
      ...(availability ? { availability } : {}),
    }
  })
  return {
    view: 'summary',
    operations,
    pagination: {
      limit,
      total: filtered.length,
      nextCursor:
        start + operations.length < filtered.length ? (operations.at(-1)?.id ?? null) : null,
    },
  }
}

function operationAvailability(
  format: LiveSession['format'],
  descriptor: OperationDescriptor,
  session: LiveSession | undefined,
): OperationAvailability | undefined {
  if (!session) return undefined
  if (session.format !== format) return { available: false, reason: 'format_mismatch' }
  if (!session.connected) return { available: false, reason: 'editor_offline' }
  if (descriptor.context.includes('selection') && session.selection === null) {
    return { available: false, reason: 'selection_required' }
  }
  return { available: true, reason: null }
}

export function getFormatCapabilities(
  format: LiveSession['format'],
  query: CapabilityDiscoveryQuery = { view: 'summary' },
): FormatCapabilities | null {
  const descriptors = getRegisteredOperationDescriptors(format)
  if (query.view === 'summary' && query.operation) return null
  if (query.view === 'detail') {
    if (query.family || query.cursor !== undefined || query.limit !== undefined) return null
    const descriptor = descriptors.find((candidate) => candidate.id === query.operation)
    if (!descriptor) return null
    const availability = operationAvailability(format, descriptor, query.session)
    return {
      ...catalogs[format],
      discovery: {
        view: 'detail',
        operation: { ...descriptor, ...(availability ? { availability } : {}) },
      },
    }
  }
  const discovery = summaryDiscovery(
    format,
    descriptors,
    query.session,
    query.family,
    query.limit ?? SUMMARY_PAGE_LIMIT,
    query.cursor,
  )
  if (!discovery) return null
  return {
    ...catalogs[format],
    discovery,
  }
}
