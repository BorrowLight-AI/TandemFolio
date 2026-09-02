import { Buffer } from 'buffer'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import type {
  ElementClipboardItem,
  OpenedPptx,
  Paragraph,
  Slide,
  SlideBundle,
  SlideElement,
  TextElement,
} from '@genoffice/pptx-engine'
import type {
  DeleteElementOp,
  EditParagraph,
  EditTextOp,
  EditTransformOp,
  FindReplaceOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
} from '../../shared/ipc'
import {
  applyEditParagraphs,
  collectParagraphFormatPatches,
  levelsChanged,
} from '../../main/edit-text'

;(globalThis as typeof globalThis & { Buffer?: typeof Buffer }).Buffer ??= Buffer

type PptxEngine = typeof import('@genoffice/pptx-engine')
let enginePromise: Promise<PptxEngine> | null = null
const engine = (): Promise<PptxEngine> => (enginePromise ??= import('@genoffice/pptx-engine'))

function elementText(element: SlideElement): string | undefined {
  if (element.type !== 'text' && element.type !== 'shape') return undefined
  return (element as TextElement).text?.paragraphs
    .flatMap((paragraph) => paragraph.runs.map((run) => run.text))
    .join('')
}

function summary(element: SlideElement): Record<string, unknown> {
  const text = elementText(element)
  return {
    id: element.id,
    type: element.type,
    box: { ...element.transform.offset },
    ...(text !== undefined ? { text: text.slice(0, 500) } : {}),
    ...(element.type === 'group' ? { childCount: element.children.length } : {}),
    ...(element.type === 'table'
      ? { rows: element.rows.length, columns: element.rows[0]?.length ?? 0 }
      : {}),
    ...(element.type === 'chart' ? { chartKind: element.chart.kind } : {}),
  }
}

function replaceElementText(element: SlideElement, text: string): boolean {
  if (element.type !== 'text' && element.type !== 'shape') return false
  const target = element as TextElement
  const firstParagraph = target.text?.paragraphs[0]
  const firstRun = firstParagraph?.runs[0]
  if (!target.text || !firstParagraph || !firstRun) return false
  target.text.paragraphs = [
    {
      ...firstParagraph,
      runs: [{ ...firstRun, text }],
    },
  ]
  target.dirty = true
  return true
}

function canReplaceElementText(element: SlideElement): boolean {
  if (element.type !== 'text' && element.type !== 'shape') return false
  const target = element as TextElement
  return Boolean(target.text?.paragraphs[0]?.runs[0])
}

interface PresentationHistoryEntry {
  readonly bytes: Uint8Array
  readonly activeSlide: number
  readonly selectedIds: string[]
  readonly dirty: boolean
}

export class BrowserPresentation {
  #opened: OpenedPptx
  #dirty = false
  #activeSlide = 0
  #selectedIds: string[] = []
  #elementClipboard: { items: ElementClipboardItem[]; pasteCount: number } | null = null
  #slideClipboard: { bundle: SlideBundle; pictureData?: string } | null = null
  #lastSlidePaste: { afterIndex: number; historyLength: number } | null = null
  #masterParts = new Map<string, Slide>()
  #transformPreview = false
  readonly #history: PresentationHistoryEntry[] = []
  readonly #redo: PresentationHistoryEntry[] = []

  private constructor(
    readonly name: string,
    opened: OpenedPptx,
  ) {
    this.#opened = opened
  }

  static async open(name: string, data: ArrayBuffer): Promise<BrowserPresentation> {
    const opened = await (await engine()).openPptx(new Uint8Array(data))
    return new BrowserPresentation(name, opened)
  }

  static async blank(): Promise<BrowserPresentation> {
    const bytes = await (await engine()).createBlankPptx()
    const data = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
    return BrowserPresentation.open('', data as ArrayBuffer)
  }

  get dirty(): boolean {
    return this.#dirty
  }

  get activeSlide(): number {
    return this.#activeSlide
  }

  get slideCount(): number {
    return this.#opened.deck.slides.length
  }

  get size(): { cx: number; cy: number } {
    return { ...this.#opened.deck.size }
  }

  get canUndo(): boolean {
    return this.#history.length > 0
  }

  get canRedo(): boolean {
    return this.#redo.length > 0
  }

  async #snapshotHistory(): Promise<PresentationHistoryEntry> {
    return {
      bytes: await (await engine()).savePptx(this.#opened),
      activeSlide: this.#activeSlide,
      selectedIds: [...this.#selectedIds],
      dirty: this.#dirty,
    }
  }

  async #restoreHistory(entry: PresentationHistoryEntry): Promise<void> {
    this.#opened = await (await engine()).openPptx(entry.bytes)
    this.#activeSlide = Math.min(entry.activeSlide, this.slideCount - 1)
    this.#selectedIds = entry.selectedIds.filter((id) =>
      this.#opened.deck.slides[this.#activeSlide]?.elements.some((element) => element.id === id),
    )
    this.#dirty = entry.dirty
    this.#transformPreview = false
    this.#masterParts.clear()
  }

  async #recordHistory(): Promise<void> {
    this.#history.push(await this.#snapshotHistory())
    if (this.#history.length > 50) this.#history.shift()
    this.#redo.length = 0
  }

  firstTextObject(slideIndex: number): { id: string; text: string } | null {
    for (const element of this.#opened.deck.slides[slideIndex]?.elements ?? []) {
      const text = elementText(element)
      if (text !== undefined) return { id: element.id, text }
    }
    return null
  }

  select(slideIndex: number, ids: readonly string[]): number {
    if (!this.#opened.deck.slides[slideIndex]) throw new Error(`Unknown slide: ${slideIndex + 1}`)
    const existing = new Set(
      this.#opened.deck.slides[slideIndex]!.elements.map((element) => element.id),
    )
    if (ids.some((id) => !existing.has(id)))
      throw new Error('The selection contains an unknown object.')
    this.#activeSlide = slideIndex
    this.#selectedIds = [...new Set(ids)].slice(0, 100)
    return this.#selectedIds.length
  }

  async replaceSelectedText(text: string): Promise<number> {
    const slide = this.#opened.deck.slides[this.#activeSlide]
    if (!slide) return 0
    const targets = this.#selectedIds
      .map((id) => slide.elements.find((candidate) => candidate.id === id))
      .filter((element): element is SlideElement =>
        Boolean(element && canReplaceElementText(element)),
      )
    if (!targets.length) return 0
    await this.#recordHistory()
    let changed = 0
    for (const element of targets) if (replaceElementText(element, text)) changed += 1
    if (changed) this.#dirty = true
    return changed
  }

  async moveSelected(deltaXEmu: number, deltaYEmu: number): Promise<number> {
    const slide = this.#opened.deck.slides[this.#activeSlide]
    if (!slide) return 0
    const targets = this.#selectedIds
      .map((id) => slide.elements.find((candidate) => candidate.id === id))
      .filter((element): element is SlideElement => Boolean(element))
    if (!targets.length) return 0
    await this.#recordHistory()
    const moved: string[] = []
    for (const element of targets) {
      element.transform.offset.x += deltaXEmu
      element.transform.offset.y += deltaYEmu
      element.dirtyTransform = true
      moved.push(element.id)
    }
    if (moved.length) {
      ;(await engine()).updateConnectorsForMoved(slide, moved)
      this.#dirty = true
    }
    return moved.length
  }

  async addBlankSlide(afterIndex = this.#activeSlide): Promise<number> {
    if (!this.#opened.deck.slides[afterIndex]) throw new Error(`Unknown slide: ${afterIndex + 1}`)
    await this.#recordHistory()
    const added = (await engine()).insertBlankSlide(this.#opened, afterIndex)
    if (!added) throw new Error('Unable to add a slide after the active slide.')
    this.#activeSlide = afterIndex + 1
    this.#selectedIds = []
    this.#dirty = true
    return this.#activeSlide
  }

  async duplicateActiveSlide(sourceIndex = this.#activeSlide, clearText = false): Promise<number> {
    if (!this.#opened.deck.slides[sourceIndex]) throw new Error(`Unknown slide: ${sourceIndex + 1}`)
    await this.#recordHistory()
    const duplicated = (await engine()).duplicateSlide(this.#opened, sourceIndex, { clearText })
    if (!duplicated) throw new Error('Unable to duplicate the active slide.')
    this.#activeSlide = sourceIndex + 1
    this.#selectedIds = []
    this.#dirty = true
    return this.#activeSlide
  }

  async deleteActiveSlide(slideIndex = this.#activeSlide): Promise<number> {
    if (this.slideCount <= 1) throw new Error('A presentation must keep at least one slide.')
    if (!this.#opened.deck.slides[slideIndex]) throw new Error(`Unknown slide: ${slideIndex + 1}`)
    await this.#recordHistory()
    const removed = (await engine()).deleteSlide(this.#opened, slideIndex)
    if (!removed) throw new Error('A presentation must keep at least one slide.')
    this.#activeSlide = Math.min(slideIndex, this.slideCount - 1)
    this.#selectedIds = []
    this.#dirty = true
    return this.#activeSlide
  }

  async moveActiveSlideForward(): Promise<number> {
    const targetIndex = this.#activeSlide + 1
    if (targetIndex >= this.slideCount)
      throw new Error('The active slide cannot move farther back.')
    return this.moveSlide(this.#activeSlide, targetIndex)
  }

  async moveSlide(fromIndex: number, toIndex: number): Promise<number> {
    if (!this.#opened.deck.slides[fromIndex]) throw new Error(`Unknown slide: ${fromIndex + 1}`)
    if (!this.#opened.deck.slides[toIndex]) throw new Error(`Unknown slide: ${toIndex + 1}`)
    await this.#recordHistory()
    const moved = (await engine()).moveSlide(this.#opened, fromIndex, toIndex)
    if (!moved) throw new Error('The active slide cannot move farther back.')
    this.#activeSlide = toIndex
    this.#dirty = true
    return this.#activeSlide
  }

  async findReplace(
    op: FindReplaceOp,
    width = 960,
  ): Promise<{ count: number; slides: RenderSlide[] | null }> {
    await this.#recordHistory()
    const { count } = (await engine()).replaceAllInDeck(this.#opened.deck, op.find, op.replace, {
      matchCase: op.matchCase,
      firstOnly: op.firstOnly,
      slideIndex: op.slideIndex,
      elementId: op.elementId,
    })
    if (!count) {
      this.#history.pop()
      return { count: 0, slides: null }
    }
    this.#dirty = true
    return { count, slides: this.renderAll(width) }
  }

  async #applyAutofit(
    slideIndex: number,
    sourceIds: string[],
    width: number,
  ): Promise<RenderSlide> {
    const pptx = await engine()
    let rendered = this.render(slideIndex, width)
    for (const sourceId of sourceIds) {
      const slide = this.#opened.deck.slides[slideIndex]
      const element = slide?.elements.find(
        (candidate): candidate is TextElement =>
          candidate.id === sourceId && (candidate.type === 'text' || candidate.type === 'shape'),
      )
      if (!slide || !element?.text) continue
      if (element.text.autofit === 'resize') {
        const node = rendered.nodes.find((candidate) => candidate.sourceId === sourceId)
        if (node && (node.type === 'shape' || node.type === 'text') && node.text) {
          const requiredHeight = node.text.contentHeight + node.text.insets.t + node.text.insets.b
          if (Math.abs(requiredHeight - node.box.h) >= 1) {
            const baseWidthPx = this.#opened.deck.size.cx / EMU_PER_PX_96
            const scale = width / baseWidthPx
            element.transform = {
              ...element.transform,
              offset: {
                ...element.transform.offset,
                cy: Math.max(Math.round((requiredHeight / scale) * EMU_PER_PX_96), 1),
              },
            }
            element.dirtyTransform = true
            rendered = this.render(slideIndex, width)
          }
        }
      }
      if (element.text.autofit === 'shrink') {
        const node = rendered.nodes.find((candidate) => candidate.sourceId === sourceId)
        if (node && (node.type === 'shape' || node.type === 'text') && node.text) {
          const effective = node.text.fontScale
          const effectiveReduction = node.text.lnSpcReduction ?? 0
          if (
            Math.abs(effective - (element.text.fontScale ?? 1)) >= 0.005 ||
            Math.abs(effectiveReduction - (element.text.lnSpcReduction ?? 0)) >= 0.005
          ) {
            element.text.fontScale = effective
            if (effectiveReduction) element.text.lnSpcReduction = effectiveReduction
            else delete element.text.lnSpcReduction
            element.anchor.originalXml = pptx.patchBodyPrAutofit(
              element.anchor.originalXml,
              effective,
              effectiveReduction,
            )
            slide.structureDirty = true
          }
        }
      }
    }
    return rendered
  }

  async editText(op: EditTextOp, width = 960): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      const pptx = await engine()
      const found = pptx.findGroupChild(slide, op.groupId, op.sourceId)
      const child = found?.child
      if (!child || (child.type !== 'text' && child.type !== 'shape')) return null
      const textChild = child as TextElement
      if (!textChild.text) return null
      await this.#recordHistory()
      textChild.text.paragraphs = applyEditParagraphs(textChild.text.paragraphs, op.paragraphs)
      pptx.ensureRunLinkRels(this.#opened, op.slideIndex, textChild.text.paragraphs)
      if (!pptx.patchGroupChildText(slide, op.groupId, textChild)) {
        this.#history.pop()
        return null
      }
      for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
        pptx.setGroupChildParagraphFormat(slide, op.groupId, op.sourceId, patch, [index])
      }
      this.#dirty = true
      return this.render(op.slideIndex, width)
    }
    const element = slide.elements.find(
      (candidate): candidate is TextElement =>
        candidate.id === op.sourceId && (candidate.type === 'text' || candidate.type === 'shape'),
    )
    if (!element?.text) return null

    await this.#recordHistory()
    const pptx = await engine()
    const levelDirty = levelsChanged(element.text.paragraphs, op.paragraphs)
    element.text.paragraphs = applyEditParagraphs(element.text.paragraphs, op.paragraphs)
    pptx.ensureRunLinkRels(this.#opened, op.slideIndex, element.text.paragraphs)
    element.dirty = true
    for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
      pptx.setElementParagraphFormat(slide, op.sourceId, patch, [index])
    }
    if (levelDirty) {
      element.dirtyPPr = { ...element.dirtyPPr, level: true, indents: true }
      pptx.materializeSlide(this.#opened, op.slideIndex)
      this.#dirty = true
      return this.render(op.slideIndex, width)
    }

    const rendered = await this.#applyAutofit(op.slideIndex, [op.sourceId], width)
    this.#dirty = true
    return rendered
  }

  async setElementFont(op: SetElementFontOp, width = 960): Promise<RenderSlide | null> {
    return (await this.setElementFontWithCount(op, width))?.slide ?? null
  }

  async setElementFontWithCount(
    op: SetElementFontOp,
    width = 960,
  ): Promise<{ slide: RenderSlide; changed: number } | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    const pptx = await engine()
    let changed = 0
    for (const sourceId of op.sourceIds) {
      const changedTarget = op.groupId
        ? pptx.setGroupChildFont(slide, op.groupId, sourceId, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
        : pptx.setElementFont(slide, sourceId, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
      if (changedTarget) changed += 1
    }
    if (!changed) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return {
      slide: op.groupId
        ? this.render(op.slideIndex, width)
        : await this.#applyAutofit(op.slideIndex, op.sourceIds, width),
      changed,
    }
  }

  async setElementParagraphFormat(
    op: SetElementParagraphFormatOp,
    width = 960,
  ): Promise<RenderSlide | null> {
    return (await this.setElementParagraphFormatWithCount(op, width))?.slide ?? null
  }

  async setElementParagraphFormatWithCount(
    op: SetElementParagraphFormatOp,
    width = 960,
  ): Promise<{ slide: RenderSlide; changed: number } | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    const pptx = await engine()
    const patch = {
      bullet: op.bullet,
      bulletChar: op.bulletChar,
      bulletHangEmu: op.bulletHangEmu,
      bulletSizePct: op.bulletSizePct,
      bulletColor: op.bulletColor,
      lineSpacingPct: op.lineSpacingPct,
      spaceBeforePt: op.spaceBeforePt,
      spaceAfterPt: op.spaceAfterPt,
      align: op.align,
      indentDelta: op.indentDelta,
    }
    let changed = 0
    for (const sourceId of op.sourceIds) {
      const applied = op.groupId
        ? pptx.setGroupChildParagraphFormat(slide, op.groupId, sourceId, patch)
        : pptx.setElementParagraphFormat(slide, sourceId, patch)
      if (applied) changed += 1
    }
    if (!changed) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    if (op.indentDelta) {
      pptx.materializeSlide(this.#opened, op.slideIndex)
      return { slide: this.render(op.slideIndex, width), changed }
    }
    return { slide: await this.#applyAutofit(op.slideIndex, op.sourceIds, width), changed }
  }

  async deleteElement(op: DeleteElementOp, width = 960): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide?.elements.some((candidate) => candidate.id === op.sourceId)) return null
    await this.#recordHistory()
    if (!(await engine()).deleteElement(slide, op.sourceId)) {
      this.#history.pop()
      return null
    }
    this.#selectedIds = this.#selectedIds.filter((id) => id !== op.sourceId)
    this.#dirty = true
    return this.render(op.slideIndex, width)
  }

  async addElementEmu(
    input: {
      slideIndex: number
      kind: string
      xEmu: number
      yEmu: number
      widthEmu: number
      heightEmu: number
      text?: string
      paragraphs?: EditParagraph[]
      fillColor?: string
      stroke?: { color: string; widthEmu: number }
    },
    width = 960,
  ): Promise<{ slide: RenderSlide; objectId: string } | null> {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    const pptx = await engine()
    const paragraphs = input.paragraphs?.length
      ? input.paragraphs
      : input.text !== undefined
        ? input.text.split('\n').map((line) => ({ runs: [{ text: line }] }))
        : undefined
    const element = pptx.addElement(slide, {
      kind: input.kind,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      ...(paragraphs ? { paragraphs } : {}),
      ...(input.fillColor ? { fillColor: input.fillColor } : {}),
      ...(input.stroke ? { stroke: input.stroke } : {}),
    })
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: element.id }
  }

  async addElement(op: import('../../shared/ipc').AddElementOp): Promise<{
    slide: RenderSlide
    objectId: string
  } | null> {
    const baseWidthPx = this.#opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    return this.addElementEmu(
      {
        slideIndex: op.slideIndex,
        kind: op.kind,
        xEmu: toEmu(op.xPx),
        yEmu: toEmu(op.yPx),
        widthEmu: toEmu(op.wPx),
        heightEmu: toEmu(op.hPx),
        ...(op.text !== undefined ? { text: op.text } : {}),
        ...(op.paragraphs ? { paragraphs: op.paragraphs } : {}),
        ...(op.fillColor ? { fillColor: op.fillColor } : {}),
        ...(op.stroke
          ? {
              stroke: {
                color: op.stroke.color,
                widthEmu: Math.round(op.stroke.widthPt * 12_700),
              },
            }
          : {}),
      },
      op.fitWidthPx,
    )
  }

  #flipAxis(element: SlideElement, axis: 'h' | 'v'): void {
    const transform = element.transform
    const orbit = () => {
      const radians = (((transform.rot ?? 0) / 60_000) * Math.PI) / 180
      const baseX = transform.flipH ? transform.offset.cx : 0
      const baseY = transform.flipV ? transform.offset.cy : 0
      const vectorX = ((transform.flipH ? -1 : 1) * transform.offset.cx) / 2
      const vectorY = ((transform.flipV ? -1 : 1) * transform.offset.cy) / 2
      return {
        x: baseX + vectorX * Math.cos(radians) - vectorY * Math.sin(radians),
        y: baseY + vectorX * Math.sin(radians) + vectorY * Math.cos(radians),
      }
    }
    const before = orbit()
    if (axis === 'h') transform.flipH = !transform.flipH
    else transform.flipV = !transform.flipV
    const after = orbit()
    transform.offset.x += Math.round(before.x - after.x)
    transform.offset.y += Math.round(before.y - after.y)
    element.dirtyTransform = true
  }

  async #flipTargets(
    slideIndex: number,
    objectIds: readonly string[],
    groupId?: string,
  ): Promise<{
    slide: NonNullable<OpenedPptx['deck']['slides'][number]>
    targets: SlideElement[]
  } | null> {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    const pptx = await engine()
    const targets = objectIds
      .map((id) =>
        groupId
          ? pptx.findGroupChild(slide, groupId, id)?.child
          : slide.elements.find((element) => element.id === id),
      )
      .filter((element): element is SlideElement => Boolean(element?.transform))
    return targets.length ? { slide, targets } : null
  }

  async setFlipState(
    input: {
      slideIndex: number
      objectIds: readonly string[]
      groupId?: string
      horizontal?: boolean
      vertical?: boolean
    },
    width = 960,
  ): Promise<{ slide: RenderSlide; changed: number } | null> {
    const found = await this.#flipTargets(input.slideIndex, input.objectIds, input.groupId)
    if (!found) return null
    const changes = found.targets.filter(
      (element) =>
        (input.horizontal !== undefined && element.transform.flipH !== input.horizontal) ||
        (input.vertical !== undefined && element.transform.flipV !== input.vertical),
    )
    if (!changes.length) return { slide: this.render(input.slideIndex, width), changed: 0 }
    await this.#recordHistory()
    for (const element of changes) {
      if (input.horizontal !== undefined && element.transform.flipH !== input.horizontal) {
        this.#flipAxis(element, 'h')
      }
      if (input.vertical !== undefined && element.transform.flipV !== input.vertical) {
        this.#flipAxis(element, 'v')
      }
    }
    ;(await engine()).updateConnectorsForMoved(
      found.slide,
      changes.map((element) => element.id),
    )
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), changed: changes.length }
  }

  async toggleElements(
    op: import('../../shared/ipc').FlipElementOp,
    width = 960,
  ): Promise<RenderSlide | null> {
    const found = await this.#flipTargets(op.slideIndex, op.sourceIds, op.groupId)
    if (!found) return null
    await this.#recordHistory()
    for (const element of found.targets) this.#flipAxis(element, op.axis)
    ;(await engine()).updateConnectorsForMoved(
      found.slide,
      found.targets.map((element) => element.id),
    )
    this.#dirty = true
    return this.render(op.slideIndex, width)
  }

  async setFill(
    op: import('../../shared/ipc').EditFillOp,
    width = 960,
  ): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const pptx = await engine()
    await this.#recordHistory()
    const gradient =
      typeof op.fill === 'string'
        ? undefined
        : {
            stops: [
              { pos: 0, color: op.fill.gradient.from },
              { pos: 1, color: op.fill.gradient.to },
            ],
            ...(op.fill.gradient.radial
              ? { radial: true as const }
              : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60_000) }),
          }
    const fill = typeof op.fill === 'string' ? op.fill : gradient!
    if (op.groupId) {
      if (!pptx.editGroupChildFill(slide, op.groupId, op.sourceId, fill)) {
        this.#history.pop()
        return null
      }
    } else {
      const element = slide.elements.find(
        (item) => item.id === op.sourceId && (item.type === 'text' || item.type === 'shape'),
      ) as TextElement | undefined
      if (!element) {
        this.#history.pop()
        return null
      }
      element.fill =
        typeof op.fill === 'string'
          ? op.fill === 'none'
            ? { type: 'none' }
            : { type: 'solid', color: op.fill }
          : {
              type: 'gradient',
              stops: gradient!.stops,
              ...('radial' in gradient! ? { path: 'circle' as const } : { angle: gradient!.angle }),
            }
      element.dirtyFill = true
    }
    this.#dirty = true
    return this.render(op.slideIndex, width)
  }

  async setImageFill(
    input: {
      readonly slideIndex: number
      readonly objectId: string
      readonly data: string
      readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
    },
    width = 960,
  ): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (
      !slide?.elements.some(
        (element) =>
          element.id === input.objectId && (element.type === 'text' || element.type === 'shape'),
      )
    )
      return null
    await this.#recordHistory()
    if (
      !(await engine()).setElementImageFill(
        this.#opened,
        slide,
        input.objectId,
        new Uint8Array(Buffer.from(input.data, 'base64')),
        input.extension,
      )
    ) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async setStroke(
    op: import('../../shared/ipc').EditStrokeOp,
    width = 960,
  ): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const pptx = await engine()
    await this.#recordHistory()
    const stroke = op.stroke
      ? {
          color: op.stroke.color,
          widthEmu: Math.round(op.stroke.widthPt * 12_700),
          ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
        }
      : null
    if (op.groupId) {
      if (!pptx.editGroupChildStroke(slide, op.groupId, op.sourceId, stroke)) {
        this.#history.pop()
        return null
      }
    } else {
      const element = slide.elements.find(
        (item) =>
          item.id === op.sourceId &&
          (item.type === 'text' || item.type === 'shape' || item.type === 'picture'),
      ) as TextElement | undefined
      if (!element) {
        this.#history.pop()
        return null
      }
      element.stroke = op.stroke
        ? {
            fill: { type: 'solid', color: op.stroke.color },
            width: Math.round(op.stroke.widthPt * 12_700),
            ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          }
        : undefined
      element.dirtyStroke = true
    }
    this.#dirty = true
    return this.render(op.slideIndex, width)
  }

  async setBackground(
    input: {
      readonly scope: 'slide' | 'all'
      readonly slideIndex?: number
      readonly color: string
    },
    width = 960,
  ): Promise<{ readonly slides: RenderSlide[]; readonly changed: number } | null> {
    const targets =
      input.scope === 'all'
        ? this.#opened.deck.slides
        : [this.#opened.deck.slides[input.slideIndex!]].filter(
            (slide): slide is NonNullable<typeof slide> => Boolean(slide),
          )
    if (!targets.length) return null
    const pptx = await engine()
    await this.#recordHistory()
    for (const slide of targets) {
      pptx.setSlideBackground(slide, input.color)
      for (const element of slide.elements) {
        if (element.type !== 'shape' && element.type !== 'text') continue
        const shaped = element as TextElement
        if (shaped.fill?.type !== 'solid' && shaped.fill?.type !== 'gradient') continue
        if (
          shaped.text?.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.trim()))
        ) {
          continue
        }
        const { x, y, cx, cy } = element.transform.offset
        const coversX =
          x <= this.#opened.deck.size.cx * 0.05 && x + cx >= this.#opened.deck.size.cx * 0.95
        const coversY =
          y <= this.#opened.deck.size.cy * 0.05 && y + cy >= this.#opened.deck.size.cy * 0.95
        if (!coversX || !coversY) continue
        shaped.fill = { type: 'solid', color: input.color }
        shaped.dirtyFill = true
      }
    }
    this.#dirty = true
    return { slides: this.renderAll(width), changed: targets.length }
  }

  async setSlideLayout(slideIndex: number, layoutPath: string | null, width = 960) {
    if (!this.#opened.deck.slides[slideIndex]) return null
    await this.#recordHistory()
    const pptx = await engine()
    const result = layoutPath
      ? pptx.setSlideLayout(this.#opened, slideIndex, layoutPath)
      : pptx.resetSlideLayout(this.#opened, slideIndex)
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setSlideSize(widthEmu: number, heightEmu: number, width = 960) {
    await this.#recordHistory()
    if (!(await engine()).setSlideSize(this.#opened, Math.round(widthEmu), Math.round(heightEmu))) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.renderAll(width)
  }

  async setSlideTransition(
    input: {
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
    },
    width = 960,
  ) {
    const targets =
      input.scope === 'all'
        ? this.#opened.deck.slides
        : [this.#opened.deck.slides[input.slideIndex!]].filter(
            (slide): slide is NonNullable<typeof slide> => Boolean(slide),
          )
    if (!targets.length) return null
    await this.#recordHistory()
    for (const slide of targets) (await engine()).setSlideTransition(slide, input.transition)
    this.#dirty = true
    return { slides: this.renderAll(width), changed: targets.length }
  }

  async setSlideHidden(slideIndex: number, hidden: boolean, width = 960) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    ;(await engine()).setSlideHidden(slide, hidden)
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setSlideAdvanceTimes(
    slides: readonly { readonly slideIndex: number; readonly milliseconds: number | null }[],
    width = 960,
  ) {
    const targets = slides.map((item) => ({
      item,
      slide: this.#opened.deck.slides[item.slideIndex],
    }))
    if (targets.some(({ slide }) => !slide)) return null
    await this.#recordHistory()
    const pptx = await engine()
    for (const { item, slide } of targets) pptx.setSlideAdvanceTime(slide!, item.milliseconds)
    this.#dirty = true
    return { slides: this.renderAll(width), changed: targets.length }
  }

  async animations(slideIndex: number) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return []
    const pptx = await engine()
    const bySpid = new Map<number, SlideElement>()
    for (const element of slide.elements) {
      const spid = pptx.elementSpid(element)
      if (spid !== null && !bySpid.has(spid)) bySpid.set(spid, element)
    }
    return pptx.getSlideAnimations(slide).flatMap((animation) => {
      const element = bySpid.get(animation.spid)
      return element
        ? [
            {
              sourceId: element.id,
              targetName: element.name || element.type,
              effect: animation.effect,
              trigger: animation.trigger,
              durationMs: animation.durationMs,
              delayMs: animation.delayMs,
              ...(animation.motionPath !== undefined ? { motionPath: animation.motionPath } : {}),
              ...(animation.paragraph !== undefined ? { paragraph: animation.paragraph } : {}),
            },
          ]
        : []
    })
  }

  async shapeKeys(slideIndex: number) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return []
    const pptx = await engine()
    return slide.elements.map((element) => ({
      sourceId: element.id,
      spid: pptx.elementSpid(element),
      name: element.name ?? '',
    }))
  }

  async setAnimations(
    slideIndex: number,
    animations: readonly {
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
    }[],
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return -1
    const pptx = await engine()
    const mapped = animations.map((item) => {
      const element = slide.elements.find((candidate) => candidate.id === item.objectId)
      const spid = element ? pptx.elementSpid(element) : null
      return spid === null
        ? null
        : {
            spid,
            effect: item.effect,
            trigger: item.trigger,
            durationMs: Math.round(item.durationMs),
            delayMs: Math.round(item.delayMs),
            ...(item.motionPath !== undefined ? { motionPath: item.motionPath } : {}),
            ...(item.paragraph !== undefined ? { paragraph: item.paragraph } : {}),
          }
    })
    if (mapped.some((item) => item === null)) return -1
    await this.#recordHistory()
    pptx.setSlideAnimations(
      slide,
      mapped.filter((item): item is NonNullable<typeof item> => item !== null),
    )
    this.#dirty = true
    return mapped.length
  }

  async setNotes(slideIndex: number, text: string) {
    if (!this.#opened.deck.slides[slideIndex]) return false
    await this.#recordHistory()
    if (!(await engine()).setSlideNotes(this.#opened, slideIndex, text)) {
      this.#history.pop()
      return false
    }
    this.#dirty = true
    return true
  }

  async setHyperlink(
    slideIndex: number,
    objectId: string,
    target:
      | { readonly kind: 'url'; readonly url: string }
      | { readonly kind: 'slide'; readonly slideIndex: number }
      | null,
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    const index = slide?.elements.findIndex((element) => element.id === objectId) ?? -1
    if (!slide || index < 0) return null
    await this.#recordHistory()
    const fresh = (await engine()).setElementLink(this.#opened, slideIndex, objectId, target)
    if (!fresh) {
      this.#history.pop()
      return null
    }
    const newObjectId = fresh.elements[index]?.id
    if (!newObjectId) {
      this.#history.pop()
      return null
    }
    this.#selectedIds = this.#selectedIds.map((id) => (id === objectId ? newObjectId : id))
    this.#dirty = true
    return { slide: this.render(slideIndex, width), objectId: newObjectId }
  }

  async hyperlink(slideIndex: number, objectId: string) {
    return (await engine()).getElementLink(this.#opened, slideIndex, objectId)
  }

  async slideLinks(slideIndex: number) {
    return (await engine())
      .getSlideLinks(this.#opened, slideIndex)
      .map(({ elementId, target }) => ({ sourceId: elementId, target }))
  }

  async runLinks(slideIndex: number) {
    return (await engine())
      .getRunLinks(this.#opened, slideIndex)
      .map(({ elementId, ...rest }) => ({ sourceId: elementId, ...rest }))
  }

  async applyHeaderFooter(
    input: {
      readonly footer: string | null
      readonly slideNumber: boolean
      readonly date: string | null
      readonly automaticDate: boolean
    },
    width = 960,
  ) {
    await this.#recordHistory()
    if (
      !(await engine()).applyHeaderFooter(this.#opened, {
        footer: input.footer,
        slideNum: input.slideNumber,
        date: input.date,
        ...(input.automaticDate ? { dateAuto: true } : {}),
      })
    ) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slides: this.renderAll(width), changed: this.slideCount }
  }

  async headerFooter(slideIndex: number) {
    const slide = this.#opened.deck.slides[slideIndex]
    return slide
      ? (await engine()).readHeaderFooter(slide)
      : { footer: null, slideNum: false, date: null }
  }

  async addSlideWithLayout(afterSlideIndex: number, layoutPath: string, width = 960) {
    if (!this.#opened.deck.slides[afterSlideIndex]) return null
    const pptx = await engine()
    const resolved = layoutPath.startsWith(pptx.BUILTIN_LAYOUT_PREFIX)
      ? pptx.ensureBuiltinLayout(
          this.#opened.archive,
          this.#opened.deck.size,
          layoutPath.slice(pptx.BUILTIN_LAYOUT_PREFIX.length),
        )
      : layoutPath
    if (!resolved) return null
    await this.#recordHistory()
    if (!pptx.insertSlideWithLayout(this.#opened, afterSlideIndex, resolved)) {
      this.#history.pop()
      return null
    }
    this.#activeSlide = afterSlideIndex + 1
    this.#selectedIds = []
    this.#dirty = true
    return { slides: this.renderAll(width), slideIndex: this.#activeSlide }
  }

  async addComment(slideIndex: number, author: string, text: string) {
    if (!this.#opened.deck.slides[slideIndex]) return null
    await this.#recordHistory()
    const added = (await engine()).addSlideComment(this.#opened, slideIndex, { author, text })
    if (!added) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return added
  }

  async deleteComment(slideIndex: number, authorId: number, index: number) {
    if (!this.#opened.deck.slides[slideIndex]) return false
    await this.#recordHistory()
    if (!(await engine()).deleteSlideComment(this.#opened, slideIndex, { authorId, idx: index })) {
      this.#history.pop()
      return false
    }
    this.#dirty = true
    return true
  }

  async addSection(beforeSlideIndex: number, name: string) {
    await this.#recordHistory()
    const sections = (await engine()).addSection(this.#opened, beforeSlideIndex, name)
    if (!sections) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return sections
  }

  async renameSection(sectionId: string, name: string) {
    await this.#recordHistory()
    const sections = (await engine()).renameSection(this.#opened, sectionId, name)
    if (!sections) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return sections
  }

  async removeSection(sectionId: string) {
    await this.#recordHistory()
    const sections = (await engine()).removeSection(this.#opened, sectionId, { keepSlides: true })
    if (!sections) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return sections
  }

  async moveSection(sectionId: string, direction: 'up' | 'down', width = 960) {
    await this.#recordHistory()
    const sections = (await engine()).moveSection(this.#opened, sectionId, direction)
    if (!sections) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { sections, slides: this.renderAll(width) }
  }

  async groupObjects(slideIndex: number, objectIds: readonly string[], width = 960) {
    await this.#recordHistory()
    const result = (await engine()).groupElements(this.#opened, slideIndex, [...objectIds])
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#selectedIds = [result.groupId]
    this.#activeSlide = slideIndex
    this.#dirty = true
    return { slide: this.render(slideIndex, width), groupId: result.groupId }
  }

  async ungroupObject(slideIndex: number, groupId: string, width = 960) {
    const slide = this.#opened.deck.slides[slideIndex]
    const group = slide?.elements.find(
      (element) => element.id === groupId && element.type === 'group',
    )
    if (!group || group.type !== 'group') return null
    const childCount = group.children.length
    await this.#recordHistory()
    if (!(await engine()).ungroupElement(this.#opened, slideIndex, groupId)) {
      this.#history.pop()
      return null
    }
    this.#selectedIds = []
    this.#activeSlide = slideIndex
    this.#dirty = true
    return { slide: this.render(slideIndex, width), childCount }
  }

  async reorderObject(
    slideIndex: number,
    objectId: string,
    position: 'front' | 'back' | 'forward' | 'backward',
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    const index = slide.elements.findIndex((element) => element.id === objectId)
    if (index < 0) return null
    const destination =
      position === 'front'
        ? slide.elements.length - 1
        : position === 'back'
          ? 0
          : position === 'forward'
            ? index + 1
            : index - 1
    if (destination === index || destination < 0 || destination >= slide.elements.length)
      return null
    await this.#recordHistory()
    if (!(await engine()).reorderElement(slide, objectId, position)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setTransformsEmu(
    slideIndex: number,
    objects: readonly {
      readonly objectId: string
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
      readonly rotationDegrees: number
    }[],
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    const pairs = objects.map((item) => ({
      item,
      element: slide.elements.find((candidate) => candidate.id === item.objectId),
    }))
    if (pairs.some(({ element }) => !element)) return null
    await this.#recordHistory()
    for (const { item, element } of pairs) {
      element!.transform = {
        ...element!.transform,
        offset: {
          x: Math.round(item.xEmu),
          y: Math.round(item.yEmu),
          cx: Math.round(item.widthEmu),
          cy: Math.round(item.heightEmu),
        },
        rot: Math.round(item.rotationDegrees * 60_000),
      }
      element!.dirtyTransform = true
    }
    ;(await engine()).updateConnectorsForMoved(
      slide,
      objects.map((item) => item.objectId),
    )
    this.#dirty = true
    return { slide: this.render(slideIndex, width), changed: objects.length }
  }

  async setConnectorEndpoints(
    input: {
      readonly slideIndex: number
      readonly connectorId: string
      readonly start: {
        readonly xEmu: number
        readonly yEmu: number
        readonly attachment?: { readonly objectId: string; readonly connectionPoint: number } | null
      }
      readonly end: {
        readonly xEmu: number
        readonly yEmu: number
        readonly attachment?: { readonly objectId: string; readonly connectionPoint: number } | null
      }
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    const connector = slide?.elements.find((element) => element.id === input.connectorId)
    if (!slide || !connector || !connector.anchor.originalXml.includes('cNvCxnSpPr')) return null
    const pptx = await engine()
    const toAttachment = (attachment: typeof input.start.attachment) => {
      if (attachment === undefined || attachment === null) return attachment
      const target = slide.elements.find((element) => element.id === attachment.objectId)
      const id = target ? pptx.elementSpid(target) : null
      return id === null ? false : { id, idx: attachment.connectionPoint }
    }
    const start = toAttachment(input.start.attachment)
    const end = toAttachment(input.end.attachment)
    if (start === false || end === false) return null
    await this.#recordHistory()
    const p1 = input.start
    const p2 = input.end
    connector.transform = {
      ...connector.transform,
      offset: {
        x: Math.min(p1.xEmu, p2.xEmu),
        y: Math.min(p1.yEmu, p2.yEmu),
        cx: Math.abs(p2.xEmu - p1.xEmu),
        cy: Math.abs(p2.yEmu - p1.yEmu),
      },
      flipH: p1.xEmu > p2.xEmu,
      flipV: p1.yEmu > p2.yEmu,
    }
    connector.dirtyTransform = true
    if (!pptx.setElementConnection(slide, input.connectorId, { start, end })) {
      const previous = this.#history.pop()!
      await this.#restoreHistory(previous)
      return null
    }
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async setPictureCrop(
    input: {
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
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    const picture = slide?.elements.find(
      (element) => element.id === input.pictureId && element.type === 'picture',
    )
    if (!slide || !picture) return null
    await this.#recordHistory()
    if (
      !(await engine()).editPictureSrcRect(
        slide,
        input.pictureId,
        input.crop
          ? { l: input.crop.left, t: input.crop.top, r: input.crop.right, b: input.crop.bottom }
          : null,
      )
    ) {
      this.#history.pop()
      return null
    }
    if (input.frame) {
      picture.transform = {
        ...picture.transform,
        offset: {
          x: Math.round(input.frame.xEmu),
          y: Math.round(input.frame.yEmu),
          cx: Math.round(input.frame.widthEmu),
          cy: Math.round(input.frame.heightEmu),
        },
      }
      picture.dirtyTransform = true
      ;(await engine()).updateConnectorsForMoved(slide, [input.pictureId])
    }
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async setPictureOpacity(slideIndex: number, pictureId: string, opacity: number, width = 960) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide?.elements.some((element) => element.id === pictureId && element.type === 'picture'))
      return null
    await this.#recordHistory()
    if (!(await engine()).setPictureOpacity(slide, pictureId, opacity)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setTextVerticalAnchor(
    slideIndex: number,
    objectId: string,
    anchor: 'top' | 'middle' | 'bottom',
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    if (!(await engine()).setElementTextAnchor(slide, objectId, anchor)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async addTableEmu(
    input: {
      readonly slideIndex: number
      readonly rows: number
      readonly columns: number
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
    },
    width = 960,
  ) {
    if (!this.#opened.deck.slides[input.slideIndex]) return null
    await this.#recordHistory()
    const result = (await engine()).addTable(this.#opened, input.slideIndex, {
      rows: input.rows,
      cols: input.columns,
      offset: {
        x: Math.round(input.xEmu),
        y: Math.round(input.yEmu),
        cx: Math.round(input.widthEmu),
        cy: Math.round(input.heightEmu),
      },
    })
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), tableId: result.elementId }
  }

  async setTableCellContent(
    input: {
      readonly slideIndex: number
      readonly tableId: string
      readonly row: number
      readonly column: number
      readonly paragraphs: readonly EditParagraph[]
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    if (
      !(await engine()).editTableCellText(
        slide,
        input.tableId,
        input.row,
        input.column,
        input.paragraphs as Paragraph[],
      )
    ) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async editTableStructure(
    input: {
      readonly slideIndex: number
      readonly tableId: string
      readonly action: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col'
      readonly index: number
      readonly before?: boolean
    },
    width = 960,
  ) {
    await this.#recordHistory()
    const result = (await engine()).editTableStructure(
      this.#opened,
      input.slideIndex,
      input.tableId,
      { kind: input.action, index: input.index, ...(input.before ? { before: true } : {}) },
    )
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), tableId: result.elementId }
  }

  async mergeTableCells(
    input: {
      readonly slideIndex: number
      readonly tableId: string
      readonly action: 'merge-right' | 'merge-down' | 'split'
      readonly row: number
      readonly column: number
    },
    width = 960,
  ) {
    await this.#recordHistory()
    const result = (await engine()).mergeTableCells(this.#opened, input.slideIndex, input.tableId, {
      kind: input.action,
      row: input.row,
      col: input.column,
    })
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), tableId: result.elementId }
  }

  async setTableColumnWidth(
    slideIndex: number,
    tableId: string,
    column: number,
    widthEmu: number,
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    if (!(await engine()).setTableColWidth(slide, tableId, column, widthEmu)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setTableRowHeight(
    slideIndex: number,
    tableId: string,
    row: number,
    heightEmu: number,
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    if (!(await engine()).setTableRowHeight(slide, tableId, row, heightEmu)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setTableCellAnchor(
    slideIndex: number,
    tableId: string,
    row: number,
    column: number,
    anchor: 'top' | 'middle' | 'bottom',
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    if (!(await engine()).setTableCellAnchor(slide, tableId, row, column, anchor)) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(slideIndex, width)
  }

  async setTableStyle(
    input: {
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
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    const index = slide?.elements.findIndex((element) => element.id === input.tableId) ?? -1
    if (!slide || index < 0) return null
    const pptx = await engine()
    let edit: Parameters<typeof pptx.editTableStyle>[2]
    let stylePart: { styleId: string; styleDefXml: string } | null = null
    if (input.styleName) {
      const preset = pptx.TABLE_STYLE_PRESETS[input.styleName]
      if (!preset) return null
      if (preset.styleId && preset.styleDefXml) {
        stylePart = { styleId: preset.styleId, styleDefXml: preset.styleDefXml }
      }
      edit = {
        tblPrXml: preset.tblPrXml,
        clearDirectFormatting: true,
        ...(preset.border
          ? {
              borderPreset: 'all' as const,
              borderColor: preset.border.color,
              borderWidthEmu: preset.border.widthEmu,
            }
          : {}),
      }
    } else {
      edit = {
        ...(input.firstRow !== undefined ? { firstRow: input.firstRow } : {}),
        ...(input.bandedRows !== undefined ? { bandRow: input.bandedRows } : {}),
        ...(input.shadingColor !== undefined ? { shadingColor: input.shadingColor } : {}),
        ...(input.borderColor !== undefined ? { borderColor: input.borderColor } : {}),
        ...(input.borderWidthPt !== undefined
          ? { borderWidthEmu: Math.round(input.borderWidthPt * 12_700) }
          : {}),
        ...(input.borderPreset !== undefined ? { borderPreset: input.borderPreset } : {}),
        ...(input.cells
          ? { cells: input.cells.map((cell) => ({ row: cell.row, col: cell.column })) }
          : {}),
      }
    }
    await this.#recordHistory()
    if (stylePart) {
      pptx.ensureTableStylePart(this.#opened, stylePart.styleId, stylePart.styleDefXml)
    }
    if (
      !pptx.editTableStyle(slide, input.tableId, edit) ||
      !pptx.materializeSlide(this.#opened, input.slideIndex)
    ) {
      this.#history.pop()
      return null
    }
    const tableId = this.#opened.deck.slides[input.slideIndex]?.elements[index]?.id
    if (!tableId) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), tableId }
  }

  async addChartEmu(
    input: {
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
    },
    width = 960,
  ) {
    if (!this.#opened.deck.slides[input.slideIndex]) return null
    await this.#recordHistory()
    const pptx = await engine()
    const added = pptx.addChart(this.#opened, input.slideIndex, {
      kind: input.kind === 'barH' ? 'bar' : input.kind,
      ...(input.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(input.title !== undefined ? { title: input.title } : {}),
      categories: [...input.categories],
      series: input.series.map((series) => ({
        name: series.name,
        values: [...series.values],
      })),
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
    })
    if (!added) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), chartId: added.elementId }
  }

  async addSmartArtEmu(
    input: {
      readonly slideIndex: number
      readonly layout: 'list' | 'process' | 'cycle' | 'hierarchy' | 'pyramid' | 'matrix' | 'venn'
      readonly items: readonly string[]
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
    },
    width = 960,
  ) {
    if (!this.#opened.deck.slides[input.slideIndex]) return null
    await this.#recordHistory()
    const added = (await engine()).addSmartArt(this.#opened, input.slideIndex, {
      layout: input.layout,
      items: [...input.items],
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
    })
    if (!added) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: added.elementId }
  }

  async addImageEmu(
    input: {
      readonly slideIndex: number
      readonly data: string
      readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
      readonly name?: string
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    const picture = (await engine()).addPicture(this.#opened, slide, {
      bytes: new Uint8Array(Buffer.from(input.data, 'base64')),
      ext: input.extension,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      ...(input.name ? { name: input.name } : {}),
    })
    if (!picture) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: picture.id }
  }

  async replaceImageBytes(
    input: {
      readonly slideIndex: number
      readonly pictureId: string
      readonly data: string
      readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tif' | 'tiff'
      readonly preserveCrop: boolean
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (
      !slide ||
      !slide.elements.some(
        (element) => element.id === input.pictureId && element.type === 'picture',
      )
    )
      return null
    await this.#recordHistory()
    if (
      !(await engine()).replacePictureBytes(
        this.#opened,
        slide,
        input.pictureId,
        new Uint8Array(Buffer.from(input.data, 'base64')),
        input.extension,
        { keepSrcRect: input.preserveCrop },
      )
    ) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async addInkEmu(
    input: {
      readonly slideIndex: number
      readonly data: string
      readonly payload: string
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    await this.#recordHistory()
    const picture = (await engine()).addPicture(this.#opened, slide, {
      bytes: new Uint8Array(Buffer.from(input.data, 'base64')),
      ext: 'png',
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      name: `aislides-ink ${Date.now().toString(36)}`,
      descr: input.payload,
    })
    if (!picture) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: picture.id }
  }

  async duplicateObjects(
    input: {
      readonly slideIndex: number
      readonly objectIds: readonly string[]
      readonly deltaXEmu: number
      readonly deltaYEmu: number
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    const pptx = await engine()
    const items = input.objectIds.map((objectId) => {
      const element = slide.elements.find((candidate) => candidate.id === objectId)
      return element ? pptx.copyElementData(this.#opened, slide, element) : null
    })
    if (items.some((item) => item === null)) return null
    await this.#recordHistory()
    const pasted = pptx.pasteElements(
      this.#opened,
      input.slideIndex,
      items.filter((item): item is NonNullable<typeof item> => item !== null),
      { dx: input.deltaXEmu, dy: input.deltaYEmu },
    )
    if (!pasted) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectIds: pasted.elementIds }
  }

  async copyObjects(slideIndex: number, objectIds: readonly string[]): Promise<number> {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) return 0
    const pptx = await engine()
    const items = objectIds
      .map((objectId) => slide.elements.find((element) => element.id === objectId))
      .filter((element): element is SlideElement => Boolean(element))
      .map((element) => pptx.copyElementData(this.#opened, slide, element))
    this.#elementClipboard = items.length ? { items, pasteCount: 0 } : null
    return items.length
  }

  async pasteCopiedObjects(slideIndex: number, shiftStepEmu: number, width = 960) {
    const clipboard = this.#elementClipboard
    if (!clipboard?.items.length || !this.#opened.deck.slides[slideIndex]) return null
    const shift = shiftStepEmu * (clipboard.pasteCount + 1)
    await this.#recordHistory()
    const pasted = (await engine()).pasteElements(this.#opened, slideIndex, clipboard.items, {
      dx: shift,
      dy: shift,
    })
    if (!pasted) {
      this.#history.pop()
      return null
    }
    clipboard.pasteCount++
    this.#dirty = true
    return { slide: this.render(slideIndex, width), objectIds: pasted.elementIds }
  }

  async copyObjectsTo(
    input: {
      readonly sourceSlideIndex: number
      readonly objectIds: readonly string[]
      readonly destinationSlideIndex: number
      readonly deltaXEmu: number
      readonly deltaYEmu: number
    },
    width = 960,
  ) {
    const source = this.#opened.deck.slides[input.sourceSlideIndex]
    if (!source || !this.#opened.deck.slides[input.destinationSlideIndex]) return null
    const pptx = await engine()
    const items = input.objectIds.map((objectId) => {
      const element = source.elements.find((candidate) => candidate.id === objectId)
      return element ? pptx.copyElementData(this.#opened, source, element) : null
    })
    if (items.some((item) => item === null)) return null
    await this.#recordHistory()
    const pasted = pptx.pasteElements(
      this.#opened,
      input.destinationSlideIndex,
      items.filter((item): item is ElementClipboardItem => item !== null),
      { dx: input.deltaXEmu, dy: input.deltaYEmu },
    )
    if (!pasted) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return {
      slide: this.render(input.destinationSlideIndex, width),
      objectIds: pasted.elementIds,
    }
  }

  async #performSlidePaste(
    bundle: SlideBundle,
    afterSlideIndex: number,
    mode: 'theme' | 'source' | 'picture',
    pictureData: string | undefined,
    width: number,
  ) {
    const pptx = await engine()
    if (mode === 'picture') {
      const slideIndex = Math.min(Math.max(afterSlideIndex, 0), this.#opened.deck.slides.length - 1)
      const slide = this.#opened.deck.slides[slideIndex]
      if (!slide || !pictureData) return null
      const picture = pptx.addPicture(this.#opened, slide, {
        bytes: new Uint8Array(Buffer.from(pictureData, 'base64')),
        ext: 'png',
        offset: { x: 0, y: 0, cx: this.#opened.deck.size.cx, cy: this.#opened.deck.size.cy },
      })
      return picture
        ? {
            slides: this.renderAll(width),
            slideIndex,
            slideCount: this.slideCount,
            objectId: picture.id,
          }
        : null
    }
    const slide = pptx.pasteSlide(this.#opened, afterSlideIndex, bundle, {
      keepSourceFormatting: mode === 'source',
    })
    if (!slide) return null
    const slideIndex = this.#opened.deck.slides.indexOf(slide)
    return { slides: this.renderAll(width), slideIndex, slideCount: this.slideCount }
  }

  async copySlideTo(
    input: {
      readonly sourceSlideIndex: number
      readonly afterSlideIndex: number
      readonly mode: 'theme' | 'source' | 'picture'
      readonly pictureData?: string
    },
    width = 960,
  ) {
    const bundle = (await engine()).copySlide(this.#opened, input.sourceSlideIndex)
    if (!bundle) return null
    await this.#recordHistory()
    const result = await this.#performSlidePaste(
      bundle,
      input.afterSlideIndex,
      input.mode,
      input.pictureData,
      width,
    )
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return result
  }

  async copySlideForPaste(slideIndex: number, pictureData?: string): Promise<boolean> {
    const bundle = (await engine()).copySlide(this.#opened, slideIndex)
    if (!bundle) return false
    this.#slideClipboard = { bundle, ...(pictureData ? { pictureData } : {}) }
    return true
  }

  get hasSlideClipboard(): boolean {
    return this.#slideClipboard !== null
  }

  async pasteCopiedSlide(
    afterSlideIndex: number,
    mode: 'theme' | 'source' | 'picture',
    width = 960,
  ) {
    const clipboard = this.#slideClipboard
    if (!clipboard) return null
    await this.#recordHistory()
    const result = await this.#performSlidePaste(
      clipboard.bundle,
      afterSlideIndex,
      mode,
      clipboard.pictureData,
      width,
    )
    if (!result) {
      this.#history.pop()
      return null
    }
    this.#lastSlidePaste = { afterIndex: afterSlideIndex, historyLength: this.#history.length }
    this.#dirty = true
    return result
  }

  async repasteCopiedSlide(mode: 'theme' | 'source' | 'picture', width = 960) {
    const clipboard = this.#slideClipboard
    const record = this.#lastSlidePaste
    if (!clipboard || !record || this.#history.length !== record.historyLength) return null
    const snapshot = this.#history.pop()
    if (!snapshot) return null
    await this.#restoreHistory(snapshot)
    await this.#recordHistory()
    const result = await this.#performSlidePaste(
      clipboard.bundle,
      record.afterIndex,
      mode,
      clipboard.pictureData,
      width,
    )
    if (!result) {
      this.#history.pop()
      await this.#restoreHistory(snapshot)
      return null
    }
    record.historyLength = this.#history.length
    this.#dirty = true
    return result
  }

  async addMediaBytes(
    input: {
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
    },
    width = 960,
  ) {
    if (!this.#opened.deck.slides[input.slideIndex]) return null
    await this.#recordHistory()
    const added = (await engine()).addMedia(this.#opened, input.slideIndex, {
      kind: input.kind,
      bytes: new Uint8Array(Buffer.from(input.data, 'base64')),
      ext: input.extension,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      ...(input.name ? { name: input.name } : {}),
    })
    if (!added) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: added.elementId }
  }

  async addModel3dBytes(
    input: {
      readonly slideIndex: number
      readonly data: string
      readonly extension: 'glb' | 'gltf'
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
      readonly name?: string
    },
    width = 960,
  ) {
    if (!this.#opened.deck.slides[input.slideIndex]) return null
    await this.#recordHistory()
    const added = (await engine()).addModel3d(this.#opened, input.slideIndex, {
      bytes: new Uint8Array(Buffer.from(input.data, 'base64')),
      ext: input.extension,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      ...(input.name ? { name: input.name } : {}),
    })
    if (!added) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), objectId: added.elementId }
  }

  async applyThemeSpec(
    spec: {
      readonly name: string
      readonly colors: Readonly<Record<string, string>>
      readonly majorFont?: string
      readonly minorFont?: string
    },
    width = 960,
  ) {
    await this.#recordHistory()
    const previous = this.#history.at(-1)
    const pptx = await engine()
    try {
      pptx.commitSaved(this.#opened)
      const normalized = {
        name: spec.name,
        colors: { ...spec.colors },
        ...(spec.majorFont ? { majorFont: spec.majorFont } : {}),
        ...(spec.minorFont ? { minorFont: spec.minorFont } : {}),
      }
      const patched = pptx.applyThemeToArchive(this.#opened, normalized)
      const remapped = pptx.remapDeckColors(this.#opened, normalized)
      if (!patched && !remapped) {
        this.#history.pop()
        return null
      }
      this.#opened = pptx.reparseDeck(this.#opened)
      const background = spec.colors.lt1
      if (background) {
        for (const slide of this.#opened.deck.slides) {
          if (!slide.background) pptx.setSlideBackground(slide, `#${background.replace(/^#/, '')}`)
        }
      }
      this.#selectedIds = []
      this.#dirty = true
      return this.renderAll(width)
    } catch (error) {
      this.#history.pop()
      if (previous) await this.#restoreHistory(previous)
      throw error
    }
  }

  mediaData(slideIndex: number, objectId: string) {
    const slide = this.#opened.deck.slides[slideIndex]
    const element = slide?.elements.find((candidate) => candidate.id === objectId)
    if (!element || element.type !== 'picture' || !element.media?.target) return null
    if (element.media.external) return { kind: element.media.kind, dataUrl: element.media.target }
    const bytes = this.#opened.archive.readBytes(element.media.target)
    if (!bytes) return null
    const extension = element.media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime: Record<string, string> = {
      mp4: 'video/mp4',
      m4v: 'video/mp4',
      mov: 'video/mp4',
      webm: 'video/webm',
      avi: 'video/x-msvideo',
      mp3: 'audio/mpeg',
      wav: 'audio/wav',
      m4a: 'audio/mp4',
      aac: 'audio/aac',
      ogg: 'audio/ogg',
    }
    return {
      kind: element.media.kind,
      dataUrl: `data:${mime[extension] ?? (element.media.kind === 'video' ? 'video/mp4' : 'audio/mpeg')};base64,${Buffer.from(bytes).toString('base64')}`,
    }
  }

  async updateChart(
    input: {
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
    },
    width = 960,
  ) {
    const slide = this.#opened.deck.slides[input.slideIndex]
    const index = slide?.elements.findIndex((element) => element.id === input.chartId) ?? -1
    const chart = index < 0 ? undefined : slide?.elements[index]
    if (!slide || !chart || chart.type !== 'chart') return null
    const imported = chart.descr !== 'aislides-chart'
    if (imported && !input.allowImportedSimplification) return null

    const pointColors: Record<number, Record<number, string | null>> = {}
    for (const point of input.pointColors ?? []) {
      ;(pointColors[point.seriesIndex] ??= {})[point.pointIndex] = point.color
    }
    const legendPositions = {
      bottom: 'b',
      top: 't',
      right: 'r',
      left: 'l',
      none: 'none',
    } as const
    const pptx = await engine()
    await this.#recordHistory()
    if (imported && !pptx.markChartEditable(slide, input.chartId)) {
      this.#history.pop()
      return null
    }
    const edited = pptx.editChartElement(this.#opened, input.slideIndex, input.chartId, {
      ...(input.kind !== undefined
        ? {
            kind: input.kind === 'barH' ? ('bar' as const) : input.kind,
            ...(input.kind === 'barH' ? { barDir: 'bar' as const } : {}),
          }
        : {}),
      ...(input.title !== undefined ? { title: input.title } : {}),
      ...(input.categories ? { categories: [...input.categories] } : {}),
      ...(input.series
        ? {
            series: input.series.map((series) => ({
              name: series.name,
              values: [...series.values],
            })),
          }
        : {}),
      ...(input.colors ? { colorScheme: [...input.colors] } : {}),
      ...(input.legendPosition ? { legendPos: legendPositions[input.legendPosition] } : {}),
      ...(input.dataLabels !== undefined ? { dataLabels: input.dataLabels } : {}),
      ...(input.gridlines !== undefined ? { gridlines: input.gridlines } : {}),
      ...(input.categoryAxisTitle !== undefined ? { catAxisTitle: input.categoryAxisTitle } : {}),
      ...(input.valueAxisTitle !== undefined ? { valAxisTitle: input.valueAxisTitle } : {}),
      ...(input.gapWidthPercent !== undefined ? { gapWidthPct: input.gapWidthPercent } : {}),
      ...(input.switchRowsAndColumns !== undefined
        ? { switchRowCol: input.switchRowsAndColumns }
        : {}),
      ...(input.pointColors ? { pointColors } : {}),
    })
    if (!edited || !pptx.materializeSlide(this.#opened, input.slideIndex)) {
      this.#history.pop()
      return null
    }
    const chartId = this.#opened.deck.slides[input.slideIndex]?.elements[index]?.id
    if (!chartId) {
      this.#history.pop()
      return null
    }
    this.#dirty = true
    return { slide: this.render(input.slideIndex, width), chartId }
  }

  async chartData(slideIndex: number, chartId: string) {
    const slide = this.#opened.deck.slides[slideIndex]
    return slide ? (await engine()).getChartElementData(slide, chartId) : null
  }

  async editTransform(op: EditTransformOp): Promise<RenderSlide | null> {
    const baseWidthPx = this.#opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    return this.#setTransformEmu(
      {
        slideIndex: op.slideIndex,
        objectId: op.sourceId,
        xEmu: toEmu(op.xPx),
        yEmu: toEmu(op.yPx),
        widthEmu: toEmu(op.wPx),
        heightEmu: toEmu(op.hPx),
        rotationDegrees: op.rotationDeg,
        ...(op.groupId ? { groupId: op.groupId } : {}),
      },
      op.fitWidthPx,
      op.preview ?? false,
    )
  }

  async setTransformEmu(
    input: {
      slideIndex: number
      objectId: string
      groupId?: string
      xEmu: number
      yEmu: number
      widthEmu: number
      heightEmu: number
      rotationDegrees: number
    },
    width = 960,
  ): Promise<RenderSlide | null> {
    return this.#setTransformEmu(input, width, false)
  }

  async #setTransformEmu(
    input: {
      slideIndex: number
      objectId: string
      groupId?: string
      xEmu: number
      yEmu: number
      widthEmu: number
      heightEmu: number
      rotationDegrees: number
    },
    width: number,
    preview: boolean,
  ): Promise<RenderSlide | null> {
    const slide = this.#opened.deck.slides[input.slideIndex]
    if (!slide) return null
    const pptx = await engine()
    const element = input.groupId
      ? pptx.findGroupChild(slide, input.groupId, input.objectId)?.child
      : slide.elements.find((candidate) => candidate.id === input.objectId)
    if (!element) return null

    if (preview) {
      if (!this.#transformPreview) {
        await this.#recordHistory()
        this.#transformPreview = true
      }
    } else if (this.#transformPreview) {
      this.#transformPreview = false
    } else {
      await this.#recordHistory()
    }

    if (input.groupId) {
      if (
        !pptx.editGroupChildTransform(
          slide,
          input.groupId,
          input.objectId,
          {
            x: input.xEmu,
            y: input.yEmu,
            cx: input.widthEmu,
            cy: input.heightEmu,
          },
          input.rotationDegrees,
        )
      ) {
        this.#history.pop()
        this.#transformPreview = false
        return null
      }
      this.#dirty = true
      return this.render(input.slideIndex, width)
    }
    const isTable = element.type === 'table'
    if (isTable) pptx.resizeTable(slide, input.objectId, input.widthEmu, input.heightEmu)
    element.transform = {
      ...element.transform,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: isTable ? element.transform.offset.cx : input.widthEmu,
        cy: isTable ? element.transform.offset.cy : input.heightEmu,
      },
      rot: Math.round(input.rotationDegrees * 60000),
    }
    element.dirtyTransform = true
    pptx.updateConnectorsForMoved(slide, [input.objectId])
    this.#dirty = true
    return this.render(input.slideIndex, width)
  }

  async undo(): Promise<boolean> {
    const previous = this.#history.pop()
    if (!previous) return false
    this.#redo.push(await this.#snapshotHistory())
    if (this.#redo.length > 50) this.#redo.shift()
    await this.#restoreHistory(previous)
    return true
  }

  async redo(): Promise<boolean> {
    const next = this.#redo.pop()
    if (!next) return false
    this.#history.push(await this.#snapshotHistory())
    if (this.#history.length > 50) this.#history.shift()
    await this.#restoreHistory(next)
    return true
  }

  render(slideIndex: number, width = 960): RenderSlide {
    const slide = this.#opened.deck.slides[slideIndex]
    if (!slide) throw new Error(`Unknown slide: ${slideIndex + 1}`)
    return buildRenderSlide(slide, this.#opened.deck.size, {
      fitWidthPx: width,
      slideNo: slideIndex + 1,
    })
  }

  renderAll(width = 960): RenderSlide[] {
    return Array.from({ length: this.slideCount }, (_, slideIndex) =>
      this.render(slideIndex, width),
    )
  }

  async masterParts(width = 960) {
    const pptx = await engine()
    return pptx.listMasterParts(this.#opened.archive).flatMap((part) => {
      const slide = pptx.parseMasterPart(this.#opened.archive, part.partPath)
      if (slide) this.#masterParts.set(part.partPath, slide)
      return slide
        ? [
            {
              ...part,
              slide: buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width }),
            },
          ]
        : []
    })
  }

  async masterPart(partPath: string, width = 960): Promise<RenderSlide | null> {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(partPath) ?? pptx.parseMasterPart(this.#opened.archive, partPath)
    if (slide) this.#masterParts.set(partPath, slide)
    return slide ? buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width }) : null
  }

  async deleteMasterObject(
    input: { readonly partPath: string; readonly objectId: string },
    width = 960,
  ) {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(input.partPath) ??
      pptx.parseMasterPart(this.#opened.archive, input.partPath)
    if (!slide || !slide.elements.some((element) => element.id === input.objectId)) return null
    await this.#recordHistory()
    if (!pptx.deleteElement(slide, input.objectId)) {
      this.#history.pop()
      return null
    }
    this.#opened.archive.entries.set(input.partPath, Buffer.from(pptx.patchSlideXml(slide), 'utf8'))
    for (let index = 0; index < this.slideCount; index++) {
      pptx.materializeSlide(this.#opened, index)
    }
    this.#dirty = true
    return buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width })
  }

  async #commitMasterPart(partPath: string, slide: Slide, pptx: PptxEngine): Promise<void> {
    this.#opened.archive.entries.set(partPath, Buffer.from(pptx.patchSlideXml(slide), 'utf8'))
    for (let index = 0; index < this.slideCount; index++) pptx.materializeSlide(this.#opened, index)
    this.#dirty = true
  }

  async setMasterText(
    input: {
      readonly partPath: string
      readonly objectId: string
      readonly paragraphs: readonly EditParagraph[]
    },
    width = 960,
  ) {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(input.partPath) ??
      pptx.parseMasterPart(this.#opened.archive, input.partPath)
    const element = slide?.elements.find(
      (candidate) =>
        candidate.id === input.objectId &&
        (candidate.type === 'text' || candidate.type === 'shape'),
    ) as TextElement | undefined
    if (!slide || !element?.text) return null
    this.#masterParts.set(input.partPath, slide)
    await this.#recordHistory()
    element.text.paragraphs = applyEditParagraphs(element.text.paragraphs, [...input.paragraphs])
    element.dirty = true
    await this.#commitMasterPart(input.partPath, slide, pptx)
    return buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width })
  }

  async setMasterTransform(
    input: {
      readonly partPath: string
      readonly objectId: string
      readonly xEmu: number
      readonly yEmu: number
      readonly widthEmu: number
      readonly heightEmu: number
      readonly rotationDegrees: number
      readonly preview?: boolean
    },
    width = 960,
  ) {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(input.partPath) ??
      pptx.parseMasterPart(this.#opened.archive, input.partPath)
    const element = slide?.elements.find((candidate) => candidate.id === input.objectId)
    if (!slide || !element) return null
    this.#masterParts.set(input.partPath, slide)
    if (input.preview) {
      if (!this.#transformPreview) {
        await this.#recordHistory()
        this.#transformPreview = true
      }
    } else if (this.#transformPreview) {
      this.#transformPreview = false
    } else {
      await this.#recordHistory()
    }
    element.transform = {
      ...element.transform,
      offset: {
        x: input.xEmu,
        y: input.yEmu,
        cx: input.widthEmu,
        cy: input.heightEmu,
      },
      rot: Math.round(input.rotationDegrees * 60_000),
    }
    element.dirtyTransform = true
    if (!input.preview) await this.#commitMasterPart(input.partPath, slide, pptx)
    return buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width })
  }

  async setMasterFill(
    input: {
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
    },
    width = 960,
  ) {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(input.partPath) ??
      pptx.parseMasterPart(this.#opened.archive, input.partPath)
    const element = slide?.elements.find(
      (candidate) =>
        candidate.id === input.objectId &&
        (candidate.type === 'text' || candidate.type === 'shape'),
    ) as TextElement | undefined
    if (!slide || !element) return null
    this.#masterParts.set(input.partPath, slide)
    await this.#recordHistory()
    element.fill =
      input.fill.kind === 'none'
        ? { type: 'none' }
        : input.fill.kind === 'solid'
          ? { type: 'solid', color: input.fill.color }
          : {
              type: 'gradient',
              stops: [
                { pos: 0, color: input.fill.from },
                { pos: 1, color: input.fill.to },
              ],
              ...(input.fill.radial
                ? { path: 'circle' as const }
                : { angle: Math.round((input.fill.angleDegrees ?? 0) * 60_000) }),
            }
    element.dirtyFill = true
    await this.#commitMasterPart(input.partPath, slide, pptx)
    return buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width })
  }

  async setMasterStroke(
    input: {
      readonly partPath: string
      readonly objectId: string
      readonly stroke: { readonly color: string; readonly widthEmu: number } | null
    },
    width = 960,
  ) {
    const pptx = await engine()
    const slide =
      this.#masterParts.get(input.partPath) ??
      pptx.parseMasterPart(this.#opened.archive, input.partPath)
    const element = slide?.elements.find(
      (candidate) =>
        candidate.id === input.objectId &&
        (candidate.type === 'text' || candidate.type === 'shape'),
    ) as TextElement | undefined
    if (!slide || !element) return null
    this.#masterParts.set(input.partPath, slide)
    await this.#recordHistory()
    element.stroke = input.stroke
      ? { fill: { type: 'solid', color: input.stroke.color }, width: input.stroke.widthEmu }
      : undefined
    element.dirtyStroke = true
    await this.#commitMasterPart(input.partPath, slide, pptx)
    return buildRenderSlide(slide, this.#opened.deck.size, { fitWidthPx: width })
  }

  async layouts() {
    const pptx = await engine()
    return this.#opened.deck.slides.length > 0
      ? {
          layouts: pptx.listSlideLayouts(this.#opened.archive),
          size: this.size,
        }
      : null
  }

  async transition(slideIndex: number) {
    const slide = this.#opened.deck.slides[slideIndex]
    return slide ? (await engine()).getSlideTransition(slide) : 'none'
  }

  async notes(slideIndex: number): Promise<string> {
    const slide = this.#opened.deck.slides[slideIndex]
    return slide ? (await engine()).getSlideNotes(this.#opened.archive, slide.path) : ''
  }

  async comments(slideIndex: number) {
    const slide = this.#opened.deck.slides[slideIndex]
    return slide ? (await engine()).getSlideComments(this.#opened.archive, slide.path) : []
  }

  async sections() {
    return (await engine()).getSections(this.#opened)
  }

  async context(): Promise<Record<string, unknown>> {
    const slide = this.#opened.deck.slides[this.#activeSlide]
    const selected = (slide?.elements ?? [])
      .filter((element) => this.#selectedIds.includes(element.id))
      .slice(0, 20)
      .map(summary)
    const notes = slide
      ? (await engine()).getSlideNotes(this.#opened.archive, slide.path).slice(0, 2_000)
      : ''
    return {
      presentation: this.name,
      slideCount: this.slideCount,
      activeSlide: this.#activeSlide,
      selectedObjects: selected,
      notes,
    }
  }

  async checkpoint(): Promise<Uint8Array> {
    return (await engine()).savePptx(this.#opened)
  }

  async save(): Promise<Uint8Array> {
    const bytes = await this.checkpoint()
    ;(await engine()).commitSaved(this.#opened)
    this.#dirty = false
    return bytes
  }
}
