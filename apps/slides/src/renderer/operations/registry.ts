import { validateJsonSchemaValue } from '@tandemfolio/operation-contract'

import { pptxOperationCatalog } from './catalog'
import type { EditParagraph } from '../../shared/ipc'

type PptxOperationDescriptor = (typeof pptxOperationCatalog.operations)[number]
type PptxOperationId = PptxOperationDescriptor['id']

export interface PptxOperationCommand {
  readonly operation: string
  readonly arguments: Record<string, unknown>
}

export interface PptxOperationServices {
  readonly setAnimations?: (input: {
    readonly slideIndex: number
    readonly animations: readonly {
      readonly objectId: string
      readonly effect:
        | 'appear'
        | 'fade'
        | 'flyIn'
        | 'wipe'
        | 'wipeDown'
        | 'splitIn'
        | 'bounce'
        | 'flipIn'
        | 'zoom'
        | 'pulse'
        | 'spin'
        | 'grow'
        | 'teeter'
        | 'disappear'
        | 'fadeOut'
        | 'flyOut'
        | 'wipeOut'
        | 'shrink'
        | 'zoomOut'
        | 'motionPath'
      readonly trigger: 'onClick' | 'withPrev' | 'afterPrev'
      readonly durationMs: number
      readonly delayMs: number
      readonly motionPath?: string
      readonly paragraph?: number
    }[]
  }) => number | Promise<number>
  readonly addChart?: (input: {
    readonly slideIndex: number
    readonly kind:
      | 'bar'
      | 'barStacked'
      | 'barPercentStacked'
      | 'barH'
      | 'line'
      | 'area'
      | 'pie'
      | 'doughnut'
      | 'scatter'
      | 'radar'
      | 'comboBarLine'
    readonly title?: string
    readonly categories: readonly string[]
    readonly series: readonly { readonly name: string; readonly values: readonly number[] }[]
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
  }) => string | Promise<string>
  readonly updateChart?: (input: {
    readonly slideIndex: number
    readonly chartId: string
    readonly kind?:
      | 'bar'
      | 'barStacked'
      | 'barPercentStacked'
      | 'barH'
      | 'line'
      | 'area'
      | 'pie'
      | 'doughnut'
      | 'scatter'
      | 'radar'
      | 'comboBarLine'
    readonly title?: string
    readonly categories?: readonly string[]
    readonly series?: readonly { readonly name: string; readonly values: readonly number[] }[]
    readonly colors?: readonly string[]
    readonly legendPosition?: 'bottom' | 'top' | 'right' | 'left' | 'none'
    readonly dataLabels?: boolean
    readonly gridlines?: boolean
    readonly categoryAxisTitle?: string
    readonly valueAxisTitle?: string
    readonly gapWidthPercent?: number
    readonly switchRowsAndColumns?: boolean
    readonly pointColors?: readonly {
      readonly seriesIndex: number
      readonly pointIndex: number
      readonly color: string | null
    }[]
    readonly allowImportedSimplification: boolean
  }) => string | Promise<string>
  readonly addSmartArt?: (input: {
    readonly slideIndex: number
    readonly layout: 'list' | 'process' | 'cycle' | 'hierarchy' | 'pyramid' | 'matrix' | 'venn'
    readonly items: readonly string[]
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
  }) => string | Promise<string>
  readonly addImage?: (input: {
    readonly slideIndex: number
    readonly data: string
    readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly name?: string
  }) => string | Promise<string>
  readonly replaceImage?: (input: {
    readonly slideIndex: number
    readonly pictureId: string
    readonly data: string
    readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
    readonly preserveCrop: boolean
  }) => boolean | Promise<boolean>
  readonly addMedia?: (input: {
    readonly slideIndex: number
    readonly kind: 'video' | 'audio'
    readonly data: string
    readonly extension:
      'mp4' | 'm4v' | 'mov' | 'webm' | 'avi' | 'mp3' | 'wav' | 'm4a' | 'aac' | 'ogg'
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly name?: string
  }) => string | Promise<string>
  readonly addInk?: (input: {
    readonly slideIndex: number
    readonly data: string
    readonly payload: string
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
  }) => string | Promise<string>
  readonly addModel3d?: (input: {
    readonly slideIndex: number
    readonly data: string
    readonly extension: 'glb' | 'gltf'
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly name?: string
  }) => string | Promise<string>
  readonly applyTheme?: (input: {
    readonly preset:
      'office' | 'ember' | 'indigo' | 'forest' | 'cream' | 'rose' | 'graphite' | 'midnight'
  }) => number | Promise<number>
  readonly duplicateObjects?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
    readonly deltaXEmu: number
    readonly deltaYEmu: number
  }) => readonly string[] | Promise<readonly string[]>
  readonly copyObjectsTo?: (input: {
    readonly sourceSlideIndex: number
    readonly objectIds: readonly string[]
    readonly destinationSlideIndex: number
    readonly deltaXEmu: number
    readonly deltaYEmu: number
  }) => readonly string[] | Promise<readonly string[]>
  readonly copySlideTo?: (input: {
    readonly sourceSlideIndex: number
    readonly afterSlideIndex: number
    readonly mode: 'theme' | 'source' | 'picture'
    readonly pictureData?: string
  }) =>
    | { readonly slideIndex: number; readonly slideCount: number; readonly objectId?: string }
    | null
    | Promise<{
        readonly slideIndex: number
        readonly slideCount: number
        readonly objectId?: string
      } | null>
  readonly addComment?: (input: {
    readonly slideIndex: number
    readonly author: string
    readonly text: string
  }) =>
    | { readonly authorId: number; readonly index: number }
    | Promise<{ readonly authorId: number; readonly index: number }>
  readonly deleteComment?: (input: {
    readonly slideIndex: number
    readonly authorId: number
    readonly index: number
  }) => boolean | Promise<boolean>
  readonly createBlank?: () => Promise<{
    readonly fileName: string
    readonly slideCount: number
  }>
  readonly loadStaged: (input: {
    readonly name: string
    readonly data: ArrayBuffer
  }) => void | Promise<void>
  readonly save: () => Promise<
    | { readonly ok: true; readonly fileName: string }
    | { readonly ok: false; readonly message: string }
  >
  readonly deleteMasterObject?: (input: {
    readonly partPath: string
    readonly objectId: string
  }) => boolean | Promise<boolean>
  readonly setMasterText?: (input: {
    readonly partPath: string
    readonly objectId: string
    readonly paragraphs: readonly EditParagraph[]
  }) => boolean | Promise<boolean>
  readonly setMasterTransform?: (input: {
    readonly partPath: string
    readonly objectId: string
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly rotationDegrees: number
  }) => boolean | Promise<boolean>
  readonly setMasterFill?: (input: {
    readonly partPath: string
    readonly objectId: string
    readonly fill:
      | { readonly kind: 'none' }
      | { readonly kind: 'solid'; readonly color: string }
      | {
          readonly kind: 'gradient'
          readonly from: string
          readonly to: string
          readonly angleDegrees?: number
          readonly radial?: boolean
        }
  }) => boolean | Promise<boolean>
  readonly setMasterStroke?: (input: {
    readonly partPath: string
    readonly objectId: string
    readonly stroke: { readonly color: string; readonly widthEmu: number } | null
  }) => boolean | Promise<boolean>
  readonly saveAs?: (input: {
    readonly fileName: string
  }) => Promise<
    | { readonly ok: true; readonly fileName: string }
    | { readonly ok: false; readonly message: string }
  >
  readonly setSelection?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
  }) => number
  readonly replaceSelectedText?: (input: { readonly text: string }) => number | Promise<number>
  readonly moveSelectedObjects?: (input: {
    readonly deltaXEmu: number
    readonly deltaYEmu: number
  }) => number | Promise<number>
  readonly addBlankSlide?: (input: {
    readonly afterSlideIndex: number
  }) => Promise<{ readonly slideIndex: number; readonly slideCount: number }>
  readonly duplicateSlide?: (input: {
    readonly slideIndex: number
    readonly clearText: boolean
  }) => Promise<{ readonly slideIndex: number; readonly slideCount: number }>
  readonly deleteSlide?: (input: {
    readonly slideIndex: number
  }) => Promise<{ readonly slideIndex: number; readonly slideCount: number }>
  readonly moveSlide?: (input: {
    readonly fromIndex: number
    readonly toIndex: number
  }) => Promise<{ readonly slideIndex: number; readonly slideCount: number }>
  readonly undoHistory?: () => Promise<{
    readonly slideIndex: number
    readonly slideCount: number
  } | null>
  readonly redoHistory?: () => Promise<{
    readonly slideIndex: number
    readonly slideCount: number
  } | null>
  readonly deleteObject?: (input: {
    readonly slideIndex: number
    readonly objectId: string
  }) => boolean | Promise<boolean>
  readonly setObjectTransform?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly groupId?: string
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly rotationDegrees: number
  }) => boolean | Promise<boolean>
  readonly setElementFont?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
    readonly groupId?: string
    readonly fontFamily?: string
    readonly fontSizePt?: number
    readonly bold?: boolean
    readonly italic?: boolean
    readonly underline?: boolean
    readonly strike?: boolean
    readonly color?: string
  }) => number | Promise<number>
  readonly setParagraphFormat?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
    readonly groupId?: string
    readonly bullet?: 'char' | 'number' | 'none'
    readonly bulletChar?: string
    readonly bulletHangEmu?: number
    readonly bulletSizePct?: number
    readonly bulletColor?: string
    readonly lineSpacingPct?: number
    readonly spaceBeforePt?: number
    readonly spaceAfterPt?: number
    readonly align?: 'left' | 'center' | 'right' | 'justify'
    readonly indentDelta?: -1 | 1
  }) => number | Promise<number>
  readonly replaceAllText?: (input: {
    readonly find: string
    readonly replace: string
    readonly matchCase?: boolean
    readonly firstOnly?: boolean
    readonly slideIndex?: number
    readonly objectId?: string
  }) => number | Promise<number>
  readonly setTextParagraphs?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly groupId?: string
    readonly paragraphs: readonly EditParagraph[]
  }) => boolean | Promise<boolean>
  readonly addObject?: (input: {
    readonly slideIndex: number
    readonly kind: string
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
    readonly text?: string
    readonly paragraphs?: readonly EditParagraph[]
    readonly fillColor?: string
    readonly stroke?: { readonly color: string; readonly widthEmu: number }
  }) => string | Promise<string>
  readonly setObjectFlip?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
    readonly groupId?: string
    readonly horizontal?: boolean
    readonly vertical?: boolean
  }) => number | Promise<number>
  readonly setObjectFill?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly groupId?: string
    readonly fill: Record<string, unknown>
  }) => boolean | Promise<boolean>
  readonly setObjectImageFill?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly data: string
    readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
  }) => boolean | Promise<boolean>
  readonly setObjectStroke?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly groupId?: string
    readonly stroke: {
      readonly color: string
      readonly widthPt: number
      readonly dash?: string
    } | null
  }) => boolean | Promise<boolean>
  readonly setSlideBackground?: (input: {
    readonly scope: 'slide' | 'all'
    readonly slideIndex?: number
    readonly color: string
  }) => number | Promise<number>
  readonly groupObjects?: (input: {
    readonly slideIndex: number
    readonly objectIds: readonly string[]
  }) => string | Promise<string>
  readonly ungroupObject?: (input: {
    readonly slideIndex: number
    readonly groupId: string
  }) => number | Promise<number>
  readonly reorderObject?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly position: 'front' | 'back' | 'forward' | 'backward'
  }) => boolean | Promise<boolean>
  readonly setObjectTransforms?: (input: {
    readonly slideIndex: number
    readonly objects: readonly {
      readonly objectId: string
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
      readonly rotationDegrees: number
    }[]
  }) => number | Promise<number>
  readonly setConnectorEndpoints?: (input: {
    readonly slideIndex: number
    readonly connectorId: string
    readonly start: {
      readonly xEmu: number
      readonly yEmu: number
      readonly attachment: { readonly objectId: string; readonly connectionPoint: number } | null
    }
    readonly end: {
      readonly xEmu: number
      readonly yEmu: number
      readonly attachment: { readonly objectId: string; readonly connectionPoint: number } | null
    }
  }) => boolean | Promise<boolean>
  readonly setPictureCrop?: (input: {
    readonly slideIndex: number
    readonly pictureId: string
    readonly crop: {
      readonly left: number
      readonly top: number
      readonly right: number
      readonly bottom: number
    } | null
    readonly frame?: {
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
    }
  }) => boolean | Promise<boolean>
  readonly setPictureOpacity?: (input: {
    readonly slideIndex: number
    readonly pictureId: string
    readonly opacity: number
  }) => boolean | Promise<boolean>
  readonly setTextVerticalAnchor?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly anchor: 'top' | 'middle' | 'bottom'
  }) => boolean | Promise<boolean>
  readonly addTable?: (input: {
    readonly slideIndex: number
    readonly rows: number
    readonly columns: number
    readonly xEmu: number
    readonly yEmu: number
    readonly widthEmu: number
    readonly heightEmu: number
  }) => string | Promise<string>
  readonly setTableCellContent?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly row: number
    readonly column: number
    readonly paragraphs: readonly EditParagraph[]
  }) => boolean | Promise<boolean>
  readonly editTableStructure?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly action: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col'
    readonly index: number
    readonly before?: boolean
  }) => string | Promise<string>
  readonly mergeTableCells?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly action: 'merge-right' | 'merge-down' | 'split'
    readonly row: number
    readonly column: number
  }) => string | Promise<string>
  readonly setTableColumnWidth?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly column: number
    readonly widthEmu: number
  }) => boolean | Promise<boolean>
  readonly setTableRowHeight?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly row: number
    readonly heightEmu: number
  }) => boolean | Promise<boolean>
  readonly setTableCellAnchor?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly row: number
    readonly column: number
    readonly anchor: 'top' | 'middle' | 'bottom'
  }) => boolean | Promise<boolean>
  readonly setSlideLayout?: (input: {
    readonly slideIndex: number
    readonly layoutPath: string | null
  }) => boolean | Promise<boolean>
  readonly setSlideSize?: (input: {
    readonly widthEmu: number
    readonly heightEmu: number
  }) => number | Promise<number>
  readonly setSlideTransition?: (input: {
    readonly scope: 'slide' | 'all'
    readonly slideIndex?: number
    readonly transition:
      | 'none'
      | 'morph'
      | 'fade'
      | 'push'
      | 'wipe'
      | 'split'
      | 'circle'
      | 'cover'
      | 'pull'
      | 'dissolve'
      | 'zoom'
      | 'random'
  }) => number | Promise<number>
  readonly setSlideHidden?: (input: {
    readonly slideIndex: number
    readonly hidden: boolean
  }) => boolean | Promise<boolean>
  readonly setSlideAdvanceTimes?: (input: {
    readonly slides: readonly {
      readonly slideIndex: number
      readonly milliseconds: number | null
    }[]
  }) => number | Promise<number>
  readonly setNotes?: (input: {
    readonly slideIndex: number
    readonly text: string
  }) => boolean | Promise<boolean>
  readonly addSection?: (input: {
    readonly beforeSlideIndex: number
    readonly name: string
  }) => number | Promise<number>
  readonly renameSection?: (input: {
    readonly sectionId: string
    readonly name: string
  }) => number | Promise<number>
  readonly removeSection?: (input: { readonly sectionId: string }) => number | Promise<number>
  readonly moveSection?: (input: {
    readonly sectionId: string
    readonly direction: 'up' | 'down'
  }) => number | Promise<number>
  readonly setHyperlink?: (input: {
    readonly slideIndex: number
    readonly objectId: string
    readonly target:
      | { readonly kind: 'url'; readonly url: string }
      | { readonly kind: 'slide'; readonly slideIndex: number }
      | null
  }) => string | Promise<string>
  readonly applyHeaderFooter?: (input: {
    readonly footer: string | null
    readonly slideNumber: boolean
    readonly date: string | null
    readonly automaticDate: boolean
  }) => number | Promise<number>
  readonly addSlideWithLayout?: (input: {
    readonly afterSlideIndex: number
    readonly layoutPath: string
  }) => Promise<{ readonly slideIndex: number; readonly slideCount: number }>
  readonly setTableStyle?: (input: {
    readonly slideIndex: number
    readonly tableId: string
    readonly styleName?:
      | 'none'
      | 'lightGrid'
      | 'zebraBlue'
      | 'zebraGray'
      | 'headerDarkBlue'
      | 'headerOrange'
      | 'noBorder'
      | 'fullBorder'
    readonly firstRow?: boolean
    readonly bandedRows?: boolean
    readonly shadingColor?: string | null
    readonly borderColor?: string
    readonly borderWidthPt?: number
    readonly borderPreset?: 'all' | 'none'
    readonly cells?: readonly { readonly row: number; readonly column: number }[]
  }) => string | Promise<string>
}

type PptxOperationHandlerResult =
  | {
      readonly ok: true
      readonly output: Readonly<Record<string, unknown>>
      readonly checkpointRecovery?: false
      readonly refreshDocument?: true
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

type PptxOperationHandler = (
  arguments_: Record<string, unknown>,
  services: PptxOperationServices,
) => PptxOperationHandlerResult | Promise<PptxOperationHandlerResult>

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]'
}

function validatePptxParagraphs(paragraphs: readonly EditParagraph[]): string | null {
  let runCount = 0
  let textLength = 0
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      runCount += 1
      textLength += Array.from(run.text).length
      const link = run.link
      if (!link) continue
      if (link.kind === 'url' && ('slideIndex' in link || !link.url)) {
        return 'A URL run link must provide only url.'
      }
      if (link.kind === 'slide' && ('url' in link || !Number.isInteger(link.slideIndex))) {
        return 'A slide run link must provide only slideIndex.'
      }
    }
  }
  return runCount > 100_000 || textLength > 1_000_000
    ? 'PPTX rich text exceeds the bounded run or text limit.'
    : null
}

function validateImageData(data: string): string | null {
  if (data.length % 4 !== 0) return 'PPTX image data must be canonical padded base64.'
  const decodedSize =
    Math.floor((data.length * 3) / 4) - (data.endsWith('==') ? 2 : data.endsWith('=') ? 1 : 0)
  return decodedSize < 1 || decodedSize > 20 * 1_024 * 1_024
    ? 'PPTX image data must decode to between 1 byte and 20 MiB.'
    : null
}

function validateMediaData(data: string): string | null {
  if (data.length % 4 !== 0) return 'PPTX media data must be canonical padded base64.'
  const decodedSize =
    Math.floor((data.length * 3) / 4) - (data.endsWith('==') ? 2 : data.endsWith('=') ? 1 : 0)
  return decodedSize < 1 || decodedSize > 100 * 1_024 * 1_024
    ? 'PPTX media data must decode to between 1 byte and 100 MiB.'
    : null
}

function validateModelData(data: string): string | null {
  if (data.length % 4 !== 0) return 'PPTX 3D model data must be canonical padded base64.'
  const decodedSize =
    Math.floor((data.length * 3) / 4) - (data.endsWith('==') ? 2 : data.endsWith('=') ? 1 : 0)
  return decodedSize < 1 || decodedSize > 100 * 1_024 * 1_024
    ? 'PPTX 3D model data must decode to between 1 byte and 100 MiB.'
    : null
}

export type PptxOperationExecution =
  | { readonly handled: false }
  | ({ readonly handled: true; readonly operationId: PptxOperationId } & PptxOperationHandlerResult)

const handlers = {
  'pptx.animation.set': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setAnimations']>>[0]
    for (const item of input.animations) {
      if (item.effect === 'motionPath' && !item.motionPath)
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'A PPTX motionPath animation requires motionPath.',
        }
      if (item.effect !== 'motionPath' && item.motionPath !== undefined)
        return {
          ok: false,
          error: 'invalid_arguments',
          message: 'Only a PPTX motionPath animation may provide motionPath.',
        }
    }
    if (!services.setAnimations)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX animation service is unavailable.',
      }
    const animationCount = await services.setAnimations(input)
    if (animationCount < 0)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide or animation target does not exist.',
      }
    return { ok: true, output: { animationCount } }
  },
  'pptx.chart.add': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addChart']>>[0]
    if (input.series.some((series) => series.values.length !== input.categories.length))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Every PPTX chart series must contain one value per category.',
      }
    if (!services.addChart)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX chart-insertion service is unavailable.',
      }
    const chartId = await services.addChart(input)
    if (!chartId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { chartId }, refreshDocument: true }
  },
  'pptx.chart.update': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['updateChart']>>[0]
    const mutableKeys = Object.keys(input).filter(
      (key) => !['slideIndex', 'chartId', 'allowImportedSimplification'].includes(key),
    )
    if (!mutableKeys.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A PPTX chart update must provide at least one changed field.',
      }
    if (
      input.categories &&
      input.series?.some((series) => series.values.length !== input.categories?.length)
    )
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Every PPTX chart series must contain one value per category.',
      }
    const pointKeys = input.pointColors?.map((point) => `${point.seriesIndex}:${point.pointIndex}`)
    if (pointKeys && new Set(pointKeys).size !== pointKeys.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX chart point-color targets must be unique.',
      }
    if (!services.updateChart)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX chart-update service is unavailable.',
      }
    const chartId = await services.updateChart(input)
    if (!chartId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message:
          'The requested PPTX chart does not exist or imported-chart simplification was not allowed.',
      }
    return { ok: true, output: { chartId }, refreshDocument: true }
  },
  'pptx.comment.add': async (arguments_, services) => {
    if (!services.addComment)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX comment service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addComment']>>[0]
    const added = await services.addComment(input)
    if (!added)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: added }
  },
  'pptx.comment.delete': async (arguments_, services) => {
    if (!services.deleteComment)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX comment-deletion service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['deleteComment']>>[0]
    if (!(await services.deleteComment(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX comment does not exist.',
      }
    return { ok: true, output: { deleted: true } }
  },
  'pptx.connector.set_endpoints': async (arguments_, services) => {
    if (!services.setConnectorEndpoints)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX connector service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setConnectorEndpoints']>
    >[0]
    if (!(await services.setConnectorEndpoints(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX connector or attachment target does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.document.create_blank': async (_arguments, services) => {
    if (!services.createBlank) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX blank-presentation service is unavailable.',
      }
    }
    const created = await services.createBlank()
    return {
      ok: true,
      output: { opened: true, ...created },
      checkpointRecovery: false,
    }
  },
  'pptx.document.load_staged': async (arguments_, services) => {
    if (!isArrayBuffer(arguments_.data)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: '$.data must be a hydrated ArrayBuffer.',
      }
    }
    if (
      !(arguments_.name as string).toLowerCase().endsWith('.pptx') ||
      arguments_.data.byteLength !== arguments_.size
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'pptx.document.load_staged requires a valid staged PPTX descriptor.',
      }
    }
    await services.loadStaged({
      name: arguments_.name as string,
      data: arguments_.data,
    })
    return {
      ok: true,
      output: { opened: true, fileName: arguments_.name as string },
    }
  },
  'pptx.document.save': async (_arguments, services) => {
    const saved = await services.save()
    if (!saved.ok) {
      return {
        ok: false,
        error: 'execution_failed',
        message: saved.message,
      }
    }
    return {
      ok: true,
      output: { saved: true, fileName: saved.fileName },
    }
  },
  'pptx.document.save_as': async (arguments_, services) => {
    if (!services.saveAs) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX save-as service is unavailable.',
      }
    }
    const saved = await services.saveAs({ fileName: arguments_.fileName as string })
    if (!saved.ok) {
      return { ok: false, error: 'execution_failed', message: saved.message }
    }
    return { ok: true, output: { saved: true, fileName: saved.fileName } }
  },
  'pptx.history.redo': async (_arguments, services) => {
    if (!services.redoHistory) {
      return { ok: false, error: 'execution_failed', message: 'PPTX redo service is unavailable.' }
    }
    const restored = await services.redoHistory()
    if (!restored) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'There is no PPTX history entry to redo.',
      }
    }
    return {
      ok: true,
      output: { redone: true, ...restored },
      refreshDocument: true,
    }
  },
  'pptx.history.undo': async (_arguments, services) => {
    if (!services.undoHistory) {
      return { ok: false, error: 'execution_failed', message: 'PPTX undo service is unavailable.' }
    }
    const restored = await services.undoHistory()
    if (!restored) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'There is no PPTX history entry to undo.',
      }
    }
    return {
      ok: true,
      output: { undone: true, ...restored },
      refreshDocument: true,
    }
  },
  'pptx.hyperlink.set': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setHyperlink']>>[0]
    if (input.target?.kind === 'url' && (!('url' in input.target) || 'slideIndex' in input.target))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A URL hyperlink must provide only url.',
      }
    if (
      input.target?.kind === 'slide' &&
      (!('slideIndex' in input.target) || 'url' in input.target)
    )
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A slide hyperlink must provide only slideIndex.',
      }
    if (!services.setHyperlink)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX hyperlink service is unavailable.',
      }
    const objectId = await services.setHyperlink(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX object or hyperlink target does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.image.add_bytes': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addImage']>>[0]
    const invalidData = validateImageData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    if (!services.addImage)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX image-insertion service is unavailable.',
      }
    const objectId = await services.addImage(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide or image format does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.image.replace_bytes': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['replaceImage']>>[0]
    const invalidData = validateImageData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    if (!services.replaceImage)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX image-replacement service is unavailable.',
      }
    if (!(await services.replaceImage(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX picture or image format does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.ink.add': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addInk']>>[0]
    const invalidData = validateImageData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    try {
      JSON.parse(input.payload)
    } catch {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX ink payload must be valid bounded JSON.',
      }
    }
    if (!services.addInk)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX ink service is unavailable.',
      }
    const objectId = await services.addInk(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.master.object.delete': async (arguments_, services) => {
    if (!services.deleteMasterObject)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX master-object deletion service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['deleteMasterObject']>
    >[0]
    if (!(await services.deleteMasterObject(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX master part or object does not exist.',
      }
    return { ok: true, output: { deleted: true }, refreshDocument: true }
  },
  'pptx.master.object.set_fill': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setMasterFill']>>[0]
    const fill = input.fill
    if (fill.kind === 'solid' && !('color' in fill))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A solid master fill requires color.',
      }
    if (fill.kind === 'gradient' && (!('from' in fill) || !('to' in fill)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A gradient master fill requires from and to.',
      }
    if (
      fill.kind !== 'gradient' &&
      ('from' in fill || 'to' in fill || 'radial' in fill || 'angleDegrees' in fill)
    )
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Only a gradient master fill accepts gradient fields.',
      }
    if (!services.setMasterFill)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX master-fill service is unavailable.',
      }
    if (!(await services.setMasterFill(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX master fill target does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.master.object.set_stroke': async (arguments_, services) => {
    if (!services.setMasterStroke)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX master-stroke service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setMasterStroke']>>[0]
    if (!(await services.setMasterStroke(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX master stroke target does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.master.object.set_transform': async (arguments_, services) => {
    if (!services.setMasterTransform)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX master-transform service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setMasterTransform']>
    >[0]
    if (!(await services.setMasterTransform(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX master transform target does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.master.text.set_paragraphs': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setMasterText']>>[0]
    const paragraphError = validatePptxParagraphs(input.paragraphs)
    if (paragraphError) return { ok: false, error: 'invalid_arguments', message: paragraphError }
    if (!services.setMasterText)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX master-text service is unavailable.',
      }
    if (!(await services.setMasterText(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX master text target does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.media.add_bytes': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addMedia']>>[0]
    const invalidData = validateMediaData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    const videoExtensions = new Set(['mp4', 'm4v', 'mov', 'webm', 'avi'])
    if (videoExtensions.has(input.extension) !== (input.kind === 'video'))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX media extension must match its explicit audio or video kind.',
      }
    if (!services.addMedia)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX media-insertion service is unavailable.',
      }
    const objectId = await services.addMedia(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide or media format does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.model3d.add_bytes': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addModel3d']>>[0]
    const invalidData = validateModelData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    if (!services.addModel3d)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX 3D-model insertion service is unavailable.',
      }
    const objectId = await services.addModel3d(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide or 3D model format does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.notes.set': async (arguments_, services) => {
    if (!services.setNotes)
      return { ok: false, error: 'execution_failed', message: 'PPTX notes service is unavailable.' }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setNotes']>>[0]
    if (!(await services.setNotes(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.object.add': async (arguments_, services) => {
    if (arguments_.text !== undefined && arguments_.paragraphs !== undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Provide either text or paragraphs when adding a PPTX object, not both.',
      }
    }
    const kind = arguments_.kind as string
    const lineKinds = new Set(['line', 'lineArrow', 'lineArrowDouble', 'lineBent', 'lineCurved'])
    if (
      !lineKinds.has(kind) &&
      ((arguments_.widthEmu as number) === 0 || (arguments_.heightEmu as number) === 0)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Non-line PPTX objects require non-zero width and height.',
      }
    }
    const paragraphs = arguments_.paragraphs as unknown as EditParagraph[] | undefined
    if (paragraphs) {
      const paragraphError = validatePptxParagraphs(paragraphs)
      if (paragraphError) {
        return { ok: false, error: 'invalid_arguments', message: paragraphError }
      }
    }
    if (!services.addObject) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX element-insertion service is unavailable.',
      }
    }
    const optionalKeys = ['text', 'paragraphs', 'fillColor', 'stroke'] as const
    const input = {
      slideIndex: arguments_.slideIndex as number,
      kind,
      xEmu: arguments_.xEmu as number,
      yEmu: arguments_.yEmu as number,
      widthEmu: arguments_.widthEmu as number,
      heightEmu: arguments_.heightEmu as number,
      ...Object.fromEntries(
        optionalKeys
          .filter((key) => arguments_[key] !== undefined)
          .map((key) => [key, arguments_[key]]),
      ),
    }
    const objectId = await services.addObject(input)
    return {
      ok: true,
      output: { added: true, slideIndex: input.slideIndex, objectId },
    }
  },
  'pptx.object.copy_to': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['copyObjectsTo']>>[0]
    if (new Set(input.objectIds).size !== input.objectIds.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX copy objectIds must be unique.',
      }
    if (!services.copyObjectsTo)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-copy service is unavailable.',
      }
    const objectIds = [...(await services.copyObjectsTo(input))]
    if (objectIds.length !== input.objectIds.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A requested PPTX source object or destination slide does not exist.',
      }
    return { ok: true, output: { objectIds }, refreshDocument: true }
  },
  'pptx.object.delete': async (arguments_, services) => {
    if (!services.deleteObject) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-deletion service is unavailable.',
      }
    }
    const slideIndex = arguments_.slideIndex as number
    const objectId = arguments_.objectId as string
    const deleted = await services.deleteObject({ slideIndex, objectId })
    if (!deleted) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX object does not exist on the requested slide.',
      }
    }
    return { ok: true, output: { deleted: true, slideIndex, objectId } }
  },
  'pptx.object.duplicate': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['duplicateObjects']>
    >[0]
    if (new Set(input.objectIds).size !== input.objectIds.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX duplicate objectIds must be unique.',
      }
    if (!services.duplicateObjects)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-duplication service is unavailable.',
      }
    const objectIds = [...(await services.duplicateObjects(input))]
    if (objectIds.length !== input.objectIds.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'One or more requested PPTX duplicate targets do not exist.',
      }
    return { ok: true, output: { objectIds }, refreshDocument: true }
  },
  'pptx.object.move_selection': async (arguments_, services) => {
    const deltaXEmu = arguments_.deltaXEmu as number
    const deltaYEmu = arguments_.deltaYEmu as number
    if (deltaXEmu === 0 && deltaYEmu === 0) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'At least one movement delta must be non-zero.',
      }
    }
    if (!services.moveSelectedObjects) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX selected-object movement service is unavailable.',
      }
    }
    const moved = await services.moveSelectedObjects({ deltaXEmu, deltaYEmu })
    if (!moved) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Select at least one movable object first.',
      }
    }
    return { ok: true, output: { moved } }
  },
  'pptx.object.group': async (arguments_, services) => {
    const objectIds = arguments_.objectIds as string[]
    if (new Set(objectIds).size !== objectIds.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX group objectIds must be unique.',
      }
    if (!services.groupObjects)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX grouping service is unavailable.',
      }
    const input = { slideIndex: arguments_.slideIndex as number, objectIds }
    const groupId = await services.groupObjects(input)
    if (!groupId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX objects cannot be grouped.',
      }
    return { ok: true, output: { grouped: true, groupId }, refreshDocument: true }
  },
  'pptx.object.reorder': async (arguments_, services) => {
    if (!services.reorderObject)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-reorder service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['reorderObject']>>[0]
    if (!(await services.reorderObject(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX object cannot move farther in that direction.',
      }
    return { ok: true, output: { reordered: true } }
  },
  'pptx.object.set_fill': async (arguments_, services) => {
    const fill = arguments_.fill as Record<string, unknown>
    if (fill.kind === 'solid' && typeof fill.color !== 'string')
      return { ok: false, error: 'invalid_arguments', message: 'A solid PPTX fill requires color.' }
    if (fill.kind === 'gradient' && (typeof fill.from !== 'string' || typeof fill.to !== 'string'))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A gradient PPTX fill requires from and to colors.',
      }
    if (fill.kind === 'gradient' && fill.radial === true && fill.angleDegrees !== undefined)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A radial PPTX gradient cannot also specify angleDegrees.',
      }
    if (!services.setObjectFill)
      return { ok: false, error: 'execution_failed', message: 'PPTX fill service is unavailable.' }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectId: arguments_.objectId as string,
      ...(arguments_.groupId ? { groupId: arguments_.groupId as string } : {}),
      fill,
    }
    if (!(await services.setObjectFill(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX fill target does not exist.',
      }
    return {
      ok: true,
      output: { updated: true, slideIndex: input.slideIndex, objectId: input.objectId },
    }
  },
  'pptx.object.set_flip': async (arguments_, services) => {
    if (arguments_.horizontal === undefined && arguments_.vertical === undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Provide horizontal, vertical, or both for the final PPTX flip state.',
      }
    }
    if (!services.setObjectFlip) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-flip service is unavailable.',
      }
    }
    const optionalKeys = ['groupId', 'horizontal', 'vertical'] as const
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectIds: arguments_.objectIds as string[],
      ...(arguments_.groupId ? { groupId: arguments_.groupId as string } : {}),
      ...Object.fromEntries(
        optionalKeys
          .filter((key) => arguments_[key] !== undefined)
          .map((key) => [key, arguments_[key]]),
      ),
    }
    const changed = await services.setObjectFlip(input)
    return {
      ok: true,
      output: { changed },
      ...(changed === 0 ? { checkpointRecovery: false as const } : {}),
    }
  },
  'pptx.object.set_image_fill': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setObjectImageFill']>
    >[0]
    const invalidData = validateImageData(input.data)
    if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    if (!services.setObjectImageFill)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX image-fill service is unavailable.',
      }
    if (!(await services.setObjectImageFill(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX shape or image format does not exist.',
      }
    return {
      ok: true,
      output: {
        updated: true,
        slideIndex: input.slideIndex,
        objectId: input.objectId,
      },
      refreshDocument: true,
    }
  },
  'pptx.object.set_stroke': async (arguments_, services) => {
    if (!services.setObjectStroke)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX stroke service is unavailable.',
      }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectId: arguments_.objectId as string,
      ...(arguments_.groupId ? { groupId: arguments_.groupId as string } : {}),
      stroke: arguments_.stroke as { color: string; widthPt: number; dash?: string } | null,
    }
    if (!(await services.setObjectStroke(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX stroke target does not exist.',
      }
    return {
      ok: true,
      output: { updated: true, slideIndex: input.slideIndex, objectId: input.objectId },
    }
  },
  'pptx.object.set_transform': async (arguments_, services) => {
    if (!services.setObjectTransform) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX object-transform service is unavailable.',
      }
    }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectId: arguments_.objectId as string,
      ...(arguments_.groupId ? { groupId: arguments_.groupId as string } : {}),
      xEmu: arguments_.xEmu as number,
      yEmu: arguments_.yEmu as number,
      widthEmu: arguments_.widthEmu as number,
      heightEmu: arguments_.heightEmu as number,
      rotationDegrees: arguments_.rotationDegrees as number,
    }
    if (!(await services.setObjectTransform(input))) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX object does not exist on the requested slide.',
      }
    }
    return {
      ok: true,
      output: { transformed: true, slideIndex: input.slideIndex, objectId: input.objectId },
    }
  },
  'pptx.paragraph.set_format': async (arguments_, services) => {
    const patchKeys = [
      'bullet',
      'bulletChar',
      'bulletHangEmu',
      'bulletSizePct',
      'bulletColor',
      'lineSpacingPct',
      'spaceBeforePt',
      'spaceAfterPt',
      'align',
      'indentDelta',
    ] as const
    if (!patchKeys.some((key) => arguments_[key] !== undefined)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Provide at least one PPTX paragraph property to change.',
      }
    }
    if (!services.setParagraphFormat) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX paragraph-format service is unavailable.',
      }
    }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectIds: arguments_.objectIds as string[],
      ...(arguments_.groupId !== undefined ? { groupId: arguments_.groupId as string } : {}),
      ...Object.fromEntries(
        patchKeys
          .filter((key) => arguments_[key] !== undefined)
          .map((key) => [key, arguments_[key]]),
      ),
    }
    const changed = await services.setParagraphFormat(input)
    if (!changed) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'No editable PPTX paragraph matched the requested targets.',
      }
    }
    return { ok: true, output: { changed } }
  },
  'pptx.picture.set_crop': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setPictureCrop']>>[0]
    if (
      input.crop &&
      (input.crop.left + input.crop.right >= 1 || input.crop.top + input.crop.bottom >= 1)
    ) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX crop edges must leave a positive visible area.',
      }
    }
    if (!services.setPictureCrop)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX picture-crop service is unavailable.',
      }
    if (!(await services.setPictureCrop(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX picture does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.picture.set_opacity': async (arguments_, services) => {
    if (!services.setPictureOpacity)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX picture-opacity service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setPictureOpacity']>
    >[0]
    if (!(await services.setPictureOpacity(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX picture does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.selection.set': (arguments_, services) => {
    if (!services.setSelection) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX selection service is unavailable.',
      }
    }
    const slideIndex = arguments_.slideIndex as number
    const selected = services.setSelection({
      slideIndex,
      objectIds: arguments_.objectIds as string[],
    })
    return { ok: true, output: { selected, slideIndex }, checkpointRecovery: false }
  },
  'pptx.section.add': async (arguments_, services) => {
    if (!services.addSection)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX section service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addSection']>>[0]
    const sectionCount = await services.addSection(input)
    if (!sectionCount)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { sectionCount } }
  },
  'pptx.section.move': async (arguments_, services) => {
    if (!services.moveSection)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX section-move service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['moveSection']>>[0]
    const sectionCount = await services.moveSection(input)
    if (!sectionCount)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX section cannot move farther.',
      }
    return { ok: true, output: { sectionCount }, refreshDocument: true }
  },
  'pptx.section.remove': async (arguments_, services) => {
    if (!services.removeSection)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX section-removal service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['removeSection']>>[0]
    const sectionCount = await services.removeSection(input)
    if (sectionCount < 0)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX section does not exist.',
      }
    return { ok: true, output: { sectionCount } }
  },
  'pptx.section.rename': async (arguments_, services) => {
    if (!services.renameSection)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX section-rename service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['renameSection']>>[0]
    const sectionCount = await services.renameSection(input)
    if (!sectionCount)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX section does not exist.',
      }
    return { ok: true, output: { sectionCount } }
  },
  'pptx.slide.add_blank': async (arguments_, services) => {
    if (!services.addBlankSlide) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX blank-slide service is unavailable.',
      }
    }
    const added = await services.addBlankSlide({
      afterSlideIndex: arguments_.afterSlideIndex as number,
    })
    return { ok: true, output: { ...added }, refreshDocument: true }
  },
  'pptx.slide.add_with_layout': async (arguments_, services) => {
    if (!services.addSlideWithLayout)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX layout-slide service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['addSlideWithLayout']>
    >[0]
    const added = await services.addSlideWithLayout(input)
    return { ok: true, output: added, refreshDocument: true }
  },
  'pptx.slide.apply_header_footer': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['applyHeaderFooter']>
    >[0]
    if (input.automaticDate && !input.date)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'An automatic PPTX date requires a date display value.',
      }
    if (!services.applyHeaderFooter)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX header/footer service is unavailable.',
      }
    const changed = await services.applyHeaderFooter(input)
    if (!changed)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX header/footer state could not be applied.',
      }
    return { ok: true, output: { changed }, refreshDocument: true }
  },
  'pptx.slide.copy_to': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['copySlideTo']>>[0]
    if ((input.mode === 'picture') !== (input.pictureData !== undefined))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX picture-mode slide copy requires pictureData; other modes forbid it.',
      }
    if (input.pictureData) {
      const invalidData = validateImageData(input.pictureData)
      if (invalidData) return { ok: false, error: 'invalid_arguments', message: invalidData }
    }
    if (!services.copySlideTo)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-copy service is unavailable.',
      }
    const copied = await services.copySlideTo(input)
    if (!copied)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX source or destination slide does not exist.',
      }
    return { ok: true, output: copied, refreshDocument: true }
  },
  'pptx.slide.delete': async (arguments_, services) => {
    if (!services.deleteSlide) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-deletion service is unavailable.',
      }
    }
    const deleted = await services.deleteSlide({ slideIndex: arguments_.slideIndex as number })
    return { ok: true, output: { ...deleted } }
  },
  'pptx.slide.duplicate': async (arguments_, services) => {
    if (!services.duplicateSlide) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-duplication service is unavailable.',
      }
    }
    const duplicated = await services.duplicateSlide({
      slideIndex: arguments_.slideIndex as number,
      clearText: (arguments_.clearText as boolean | undefined) ?? false,
    })
    return { ok: true, output: { ...duplicated } }
  },
  'pptx.slide.move': async (arguments_, services) => {
    const fromIndex = arguments_.fromIndex as number
    const toIndex = arguments_.toIndex as number
    if (fromIndex === toIndex) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'fromIndex and toIndex must identify different slides.',
      }
    }
    if (!services.moveSlide) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-movement service is unavailable.',
      }
    }
    const moved = await services.moveSlide({ fromIndex, toIndex })
    return { ok: true, output: { ...moved } }
  },
  'pptx.slide.set_advance_times': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setSlideAdvanceTimes']>
    >[0]
    const indices = input.slides.map((item) => item.slideIndex)
    if (new Set(indices).size !== indices.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX advance-time slideIndex values must be unique.',
      }
    if (!services.setSlideAdvanceTimes)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX advance-time service is unavailable.',
      }
    const changed = await services.setSlideAdvanceTimes(input)
    if (!changed)
      return { ok: false, error: 'invalid_arguments', message: 'No requested PPTX slide exists.' }
    return { ok: true, output: { changed } }
  },
  'pptx.slide.set_background': async (arguments_, services) => {
    const scope = arguments_.scope as 'slide' | 'all'
    const slideIndex = arguments_.slideIndex as number | undefined
    if (scope === 'slide' && slideIndex === undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A slide-scoped PPTX background requires slideIndex.',
      }
    }
    if (scope === 'all' && slideIndex !== undefined) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'An all-slide PPTX background must not provide slideIndex.',
      }
    }
    if (!services.setSlideBackground) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-background service is unavailable.',
      }
    }
    const changed = await services.setSlideBackground({
      scope,
      ...(slideIndex !== undefined ? { slideIndex } : {}),
      color: arguments_.color as string,
    })
    if (!changed) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX background target does not exist.',
      }
    }
    return { ok: true, output: { changed }, refreshDocument: true }
  },
  'pptx.slide.set_hidden': async (arguments_, services) => {
    if (!services.setSlideHidden)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-visibility service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setSlideHidden']>>[0]
    if (!(await services.setSlideHidden(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.slide.set_layout': async (arguments_, services) => {
    if (!services.setSlideLayout)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-layout service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setSlideLayout']>>[0]
    if (!(await services.setSlideLayout(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide or layout does not exist.',
      }
    return { ok: true, output: { updated: true }, refreshDocument: true }
  },
  'pptx.slide.set_size': async (arguments_, services) => {
    if (!services.setSlideSize)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX slide-size service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setSlideSize']>>[0]
    const changed = await services.setSlideSize(input)
    if (!changed)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide size matches the current size or cannot be applied.',
      }
    return { ok: true, output: { changed }, refreshDocument: true }
  },
  'pptx.slide.set_transition': async (arguments_, services) => {
    const scope = arguments_.scope as 'slide' | 'all'
    const slideIndex = arguments_.slideIndex as number | undefined
    if ((scope === 'slide') !== (slideIndex !== undefined))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A slide-scoped PPTX transition requires slideIndex; all scope forbids it.',
      }
    if (!services.setSlideTransition)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX transition service is unavailable.',
      }
    const input = {
      scope,
      ...(slideIndex !== undefined ? { slideIndex } : {}),
      transition: arguments_.transition as NonNullable<
        Parameters<NonNullable<PptxOperationServices['setSlideTransition']>>[0]
      >['transition'],
    }
    const changed = await services.setSlideTransition(input)
    if (!changed)
      return { ok: false, error: 'invalid_arguments', message: 'No requested PPTX slide exists.' }
    return { ok: true, output: { changed } }
  },
  'pptx.smartart.add': async (arguments_, services) => {
    if (!services.addSmartArt)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX SmartArt-insertion service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addSmartArt']>>[0]
    const objectId = await services.addSmartArt(input)
    if (!objectId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX slide does not exist.',
      }
    return { ok: true, output: { objectId }, refreshDocument: true }
  },
  'pptx.table.add': async (arguments_, services) => {
    if (!services.addTable)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table insertion service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['addTable']>>[0]
    const tableId = await services.addTable(input)
    if (!tableId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX table target slide does not exist.',
      }
    return { ok: true, output: { tableId }, refreshDocument: true }
  },
  'pptx.table.edit_structure': async (arguments_, services) => {
    if (!services.editTableStructure)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-structure service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['editTableStructure']>
    >[0]
    const tableId = await services.editTableStructure(input)
    if (!tableId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The table structure edit is invalid or targets merged cells.',
      }
    return { ok: true, output: { tableId }, refreshDocument: true }
  },
  'pptx.table.merge_cells': async (arguments_, services) => {
    if (!services.mergeTableCells)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table merge service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['mergeTableCells']>>[0]
    const tableId = await services.mergeTableCells(input)
    if (!tableId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table merge or split is invalid.',
      }
    return { ok: true, output: { tableId }, refreshDocument: true }
  },
  'pptx.table.set_cell_anchor': async (arguments_, services) => {
    if (!services.setTableCellAnchor)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-cell anchor service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setTableCellAnchor']>
    >[0]
    if (!(await services.setTableCellAnchor(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table cell does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.table.set_cell_content': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setTableCellContent']>
    >[0]
    const paragraphError = validatePptxParagraphs(input.paragraphs)
    if (paragraphError) return { ok: false, error: 'invalid_arguments', message: paragraphError }
    if (!services.setTableCellContent)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-cell content service is unavailable.',
      }
    if (!(await services.setTableCellContent(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table cell does not exist or is merged.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.table.set_column_width': async (arguments_, services) => {
    if (!services.setTableColumnWidth)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-column service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setTableColumnWidth']>
    >[0]
    if (!(await services.setTableColumnWidth(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table column does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.table.set_row_height': async (arguments_, services) => {
    if (!services.setTableRowHeight)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-row service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setTableRowHeight']>
    >[0]
    if (!(await services.setTableRowHeight(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table row does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.table.set_style': async (arguments_, services) => {
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['setTableStyle']>>[0]
    const patchKeys = [
      'styleName',
      'firstRow',
      'bandedRows',
      'shadingColor',
      'borderColor',
      'borderWidthPt',
      'borderPreset',
    ] as const
    if (!patchKeys.some((key) => input[key] !== undefined))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Provide at least one PPTX table style field.',
      }
    if (input.styleName && patchKeys.slice(1).some((key) => input[key] !== undefined))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'A PPTX table preset cannot be combined with direct style fields.',
      }
    if (!services.setTableStyle)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX table-style service is unavailable.',
      }
    const tableId = await services.setTableStyle(input)
    if (!tableId)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX table or style target does not exist.',
      }
    return { ok: true, output: { tableId }, refreshDocument: true }
  },
  'pptx.text.replace_all': async (arguments_, services) => {
    if (!services.replaceAllText) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX find/replace service is unavailable.',
      }
    }
    const optionalKeys = ['matchCase', 'firstOnly', 'slideIndex', 'objectId'] as const
    const input = {
      find: arguments_.find as string,
      replace: arguments_.replace as string,
      ...Object.fromEntries(
        optionalKeys
          .filter((key) => arguments_[key] !== undefined)
          .map((key) => [key, arguments_[key]]),
      ),
    }
    const changed = await services.replaceAllText(input)
    return {
      ok: true,
      output: { changed },
      ...(changed === 0 ? { checkpointRecovery: false as const } : {}),
    }
  },
  'pptx.text.replace_selection': async (arguments_, services) => {
    if (!services.replaceSelectedText) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX selected-text service is unavailable.',
      }
    }
    const changed = await services.replaceSelectedText({ text: arguments_.text as string })
    if (!changed) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Select at least one editable text object first.',
      }
    }
    return { ok: true, output: { changed } }
  },
  'pptx.text.set_font': async (arguments_, services) => {
    const patchKeys = [
      'fontFamily',
      'fontSizePt',
      'bold',
      'italic',
      'underline',
      'strike',
      'color',
    ] as const
    if (!patchKeys.some((key) => arguments_[key] !== undefined)) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'Provide at least one PPTX font property to change.',
      }
    }
    if (!services.setElementFont) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX element-font service is unavailable.',
      }
    }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      objectIds: arguments_.objectIds as string[],
      ...(arguments_.groupId ? { groupId: arguments_.groupId as string } : {}),
      ...Object.fromEntries(
        patchKeys
          .filter((key) => arguments_[key] !== undefined)
          .map((key) => [key, arguments_[key]]),
      ),
    }
    const changed = await services.setElementFont(input)
    if (!changed) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'No editable PPTX text object matched the requested targets.',
      }
    }
    return { ok: true, output: { changed } }
  },
  'pptx.text.set_paragraphs': async (arguments_, services) => {
    const slideIndex = arguments_.slideIndex as number
    const objectId = arguments_.objectId as string
    const groupId = arguments_.groupId as string | undefined
    const paragraphs = arguments_.paragraphs as unknown as EditParagraph[]
    const paragraphError = validatePptxParagraphs(paragraphs)
    if (paragraphError) {
      return { ok: false, error: 'invalid_arguments', message: paragraphError }
    }
    if (!services.setTextParagraphs) {
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX rich-text service is unavailable.',
      }
    }
    const updated = await services.setTextParagraphs({
      slideIndex,
      objectId,
      ...(groupId ? { groupId } : {}),
      paragraphs,
    })
    if (!updated) {
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The PPTX text object does not exist on the requested slide.',
      }
    }
    return { ok: true, output: { updated: true, slideIndex, objectId } }
  },
  'pptx.text.set_vertical_anchor': async (arguments_, services) => {
    if (!services.setTextVerticalAnchor)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX text-anchor service is unavailable.',
      }
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setTextVerticalAnchor']>
    >[0]
    if (!(await services.setTextVerticalAnchor(input)))
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX text object does not exist.',
      }
    return { ok: true, output: { updated: true } }
  },
  'pptx.theme.apply': async (arguments_, services) => {
    if (!services.applyTheme)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX theme service is unavailable.',
      }
    const input = arguments_ as Parameters<NonNullable<PptxOperationServices['applyTheme']>>[0]
    const changedSlides = await services.applyTheme(input)
    if (!changedSlides)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX theme did not change this presentation.',
      }
    return { ok: true, output: { changedSlides }, refreshDocument: true }
  },
  'pptx.object.set_transforms': async (arguments_, services) => {
    const input = arguments_ as Parameters<
      NonNullable<PptxOperationServices['setObjectTransforms']>
    >[0]
    const ids = input.objects.map((item) => item.objectId)
    if (new Set(ids).size !== ids.length)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'PPTX batch transform objectIds must be unique.',
      }
    if (!services.setObjectTransforms)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX batch-transform service is unavailable.',
      }
    const changed = await services.setObjectTransforms(input)
    if (!changed)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'One or more PPTX transform targets do not exist.',
      }
    return { ok: true, output: { changed } }
  },
  'pptx.object.ungroup': async (arguments_, services) => {
    if (!services.ungroupObject)
      return {
        ok: false,
        error: 'execution_failed',
        message: 'PPTX ungrouping service is unavailable.',
      }
    const input = {
      slideIndex: arguments_.slideIndex as number,
      groupId: arguments_.groupId as string,
    }
    const ungrouped = await services.ungroupObject(input)
    if (!ungrouped)
      return {
        ok: false,
        error: 'invalid_arguments',
        message: 'The requested PPTX group does not exist.',
      }
    return { ok: true, output: { ungrouped }, refreshDocument: true }
  },
} satisfies Record<PptxOperationId, PptxOperationHandler>

export const pptxOperationHandlerIds = Object.freeze(
  Object.keys(handlers).sort() as PptxOperationId[],
)

const descriptorsByName = new Map<string, PptxOperationDescriptor>()
for (const descriptor of pptxOperationCatalog.operations) {
  descriptorsByName.set(descriptor.id, descriptor)
  for (const alias of descriptor.compatibilityAliases) {
    descriptorsByName.set(alias, descriptor)
  }
}

export async function executePptxOperation(
  command: PptxOperationCommand,
  services: PptxOperationServices,
): Promise<PptxOperationExecution> {
  const descriptor = descriptorsByName.get(command.operation)
  if (!descriptor) return { handled: false }

  const validation = validateJsonSchemaValue(descriptor.inputSchema, command.arguments)
  if (!validation.ok) {
    return {
      handled: true,
      operationId: descriptor.id,
      ok: false,
      error: 'invalid_arguments',
      message: validation.error.message,
    }
  }

  try {
    const execution = await handlers[descriptor.id](command.arguments, services)
    return {
      handled: true,
      operationId: descriptor.id,
      ...execution,
    }
  } catch (error) {
    return {
      handled: true,
      operationId: descriptor.id,
      ok: false,
      error: 'execution_failed',
      message: error instanceof Error ? error.message : 'The presentation was not saved.',
    }
  }
}
