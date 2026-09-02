import { pdfOperationCatalog } from './catalog'

export type PdfRetainedProducerDisposition =
  | 'registry'
  | 'host-effect'
  | 'view-only'
  | 'missing'

export interface PdfRetainedProducer {
  readonly producer: string
  readonly disposition: PdfRetainedProducerDisposition
  readonly operationIds: readonly string[]
}

/**
 * Machine-checkable inventory of the retained community PDF command families.
 *
 * A producer may map to more than one exact final-state operation, but every
 * state-changing producer must end at `registry`; `host-effect` is reserved for
 * non-document output such as print/export and `view-only` for navigation state.
 */
export const pdfRetainedProducerBaseline = [
  {
    producer: 'document open',
    disposition: 'registry',
    operationIds: ['pdf.document.load_staged'],
  },
  {
    producer: 'document save/autosave',
    disposition: 'registry',
    operationIds: ['pdf.document.save'],
  },
  {
    producer: 'history undo',
    disposition: 'registry',
    operationIds: ['pdf.history.undo'],
  },
  {
    producer: 'history redo',
    disposition: 'registry',
    operationIds: ['pdf.history.redo'],
  },
  {
    producer: 'text markup add/toggle',
    disposition: 'registry',
    operationIds: ['pdf.markup.add', 'pdf.pending.delete', 'pdf.annotation.delete_saved'],
  },
  {
    producer: 'drawing and note add',
    disposition: 'registry',
    operationIds: ['pdf.drawing.add'],
  },
  {
    producer: 'drawing move/resize',
    disposition: 'registry',
    operationIds: ['pdf.drawing.update'],
  },
  {
    producer: 'signature placement',
    disposition: 'registry',
    operationIds: ['pdf.drawing.add'],
  },
  {
    producer: 'pending object delete',
    disposition: 'registry',
    operationIds: ['pdf.pending.delete'],
  },
  {
    producer: 'content text replace',
    disposition: 'registry',
    operationIds: ['pdf.text.replace'],
  },
  {
    producer: 'content text insert/edit',
    disposition: 'registry',
    operationIds: ['pdf.text.insert', 'pdf.text.update_inserted'],
  },
  {
    producer: 'content image insert',
    disposition: 'registry',
    operationIds: ['pdf.image.insert'],
  },
  {
    producer: 'content image move/resize/rotate/layer',
    disposition: 'registry',
    operationIds: ['pdf.image.transform'],
  },
  {
    producer: 'content image replace/crop/cutout/flip/opacity',
    disposition: 'registry',
    operationIds: ['pdf.image.replace'],
  },
  {
    producer: 'content image delete',
    disposition: 'registry',
    operationIds: ['pdf.image.delete'],
  },
  {
    producer: 'static form fill image state',
    disposition: 'registry',
    operationIds: ['pdf.static_form.set', 'pdf.pending.delete'],
  },
  {
    producer: 'AcroForm field edit',
    disposition: 'registry',
    operationIds: ['pdf.form.set_value'],
  },
  {
    producer: 'watermark/header/footer apply and clear',
    disposition: 'registry',
    operationIds: ['pdf.stamp.set'],
  },
  {
    producer: 'document metadata edit',
    disposition: 'registry',
    operationIds: ['pdf.document.set_metadata'],
  },
  {
    producer: 'page rotate',
    disposition: 'registry',
    operationIds: ['pdf.page.set_rotation'],
  },
  {
    producer: 'page delete',
    disposition: 'registry',
    operationIds: ['pdf.page.delete'],
  },
  {
    producer: 'page reorder',
    disposition: 'registry',
    operationIds: ['pdf.page.reorder'],
  },
  {
    producer: 'page insert from PDF',
    disposition: 'registry',
    operationIds: ['pdf.page.insert', 'pdf.page.insert_staged'],
  },
  {
    producer: 'save as/export images/extract/print',
    disposition: 'host-effect',
    operationIds: [],
  },
  {
    producer: 'search/navigation/zoom/sidebar/view modes/form focus',
    disposition: 'view-only',
    operationIds: [],
  },
] as const satisfies readonly PdfRetainedProducer[]

export function validatePdfRetainedProducerBaseline(): readonly string[] {
  const operationIds = new Set(pdfOperationCatalog.operations.map((operation) => operation.id))
  const baseline: readonly PdfRetainedProducer[] = pdfRetainedProducerBaseline
  return baseline.flatMap((entry) => {
    if (entry.disposition === 'missing') return [`${entry.producer}: missing`]
    if (entry.disposition !== 'registry') return []
    return entry.operationIds.flatMap((operationId) =>
      operationIds.has(operationId as never)
        ? []
        : [`${entry.producer}: unknown operation ${operationId}`],
    )
  })
}
