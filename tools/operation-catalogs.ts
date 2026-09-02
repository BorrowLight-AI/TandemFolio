import type { OperationCatalog } from '@tandemfolio/operation-contract'

import { docxOperationCatalog } from '../apps/docs/src/renderer/operations/catalog'
import { markdownOperationCatalog } from '../apps/markdown/src/renderer/operations/catalog'
import { pdfOperationCatalog } from '../apps/pdf/src/renderer/operations/catalog'
import { pptxOperationCatalog } from '../apps/slides/src/renderer/operations/catalog'
import { xlsxOperationCatalog } from '../apps/sheets/src/renderer/operations/catalog'

const operationCatalogs = [
  docxOperationCatalog,
  markdownOperationCatalog,
  pdfOperationCatalog,
  pptxOperationCatalog,
  xlsxOperationCatalog,
] satisfies OperationCatalog[]

export default operationCatalogs
