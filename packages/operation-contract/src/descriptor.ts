export const OPERATION_FORMATS = ['docx', 'xlsx', 'pptx', 'pdf', 'markdown'] as const

export type OperationFormat = (typeof OPERATION_FORMATS)[number]
export type OperationId = `${OperationFormat}.${string}`

export type JsonSchema = Readonly<Record<string, unknown>>

export type OperationContextRequirement = 'document' | 'selection' | `${OperationFormat}.${string}`

export type OperationEffect = 'selection' | 'document' | 'persistence' | 'view'

export interface OperationDescriptor {
  readonly id: OperationId
  readonly format: OperationFormat
  readonly family: string
  readonly summary: string
  readonly visibility: 'agent' | 'internal'
  readonly inputSchema: JsonSchema
  readonly outputSchema: JsonSchema
  readonly risk: 'low' | 'medium' | 'high'
  readonly context: readonly OperationContextRequirement[]
  readonly effects: readonly OperationEffect[]
  readonly mutates: boolean
  readonly undoable: boolean
  readonly atomic: boolean
  readonly compatibilityAliases?: readonly string[]
}

export interface OperationCatalog {
  readonly format: OperationFormat
  readonly operations: readonly OperationDescriptor[]
}
