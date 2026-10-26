// @vitest-environment node
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { PictureRenderNode, RenderSlide } from '@genoffice/pptx-render'
import { BrowserPresentation } from '../src/renderer/host/browser-presentation'

// Synthetic 16px swatches, with no document or third-party image content.
const BLUE =
  'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFUlEQVR4nGNgyF1FGhrVMKph+GoAAKoGFxBobgn7AAAAAElFTkSuQmCC'
const GREEN =
  'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFUlEQVR4nGNg2DeNNDSqYVTD8NUAACuIVBD9pF++AAAAAElFTkSuQmCC'
const picture = (slide: RenderSlide) =>
  slide.nodes.find((node): node is PictureRenderNode => node.type === 'picture')!

async function imageDeck() {
  const presentation = await BrowserPresentation.blank()
  const added = await presentation.addImageEmu({
    slideIndex: 0,
    data: BLUE,
    extension: 'png',
    xEmu: 914400,
    yEmu: 914400,
    widthEmu: 1828800,
    heightEmu: 914400,
  })
  return { presentation, objectId: added!.objectId }
}

async function reopen(bytes: Uint8Array) {
  return BrowserPresentation.open('images.pptx', Uint8Array.from(bytes).buffer)
}

describe('PPTX browser media rendering', () => {
  it('renders package pictures immediately, after replacement/undo/redo, and after save/reopen', async () => {
    const { presentation, objectId } = await imageDeck()
    const originalBox = picture(presentation.render(0)).box
    expect(picture(presentation.render(0)).dataUrl).toBe(`data:image/png;base64,${BLUE}`)
    expect(picture(presentation.renderAll()[0]!).dataUrl).toBe(`data:image/png;base64,${BLUE}`)
    const replaced = await presentation.replaceImageBytes({
      slideIndex: 0,
      pictureId: objectId,
      data: GREEN,
      extension: 'png',
      preserveCrop: true,
    })
    expect(picture(replaced!).dataUrl).toBe(`data:image/png;base64,${GREEN}`)
    expect(picture(replaced!).box).toEqual(originalBox)
    await presentation.undo()
    expect(picture(presentation.render(0)).dataUrl).toBe(`data:image/png;base64,${BLUE}`)
    await presentation.redo()
    expect(picture(presentation.render(0)).dataUrl).toBe(`data:image/png;base64,${GREEN}`)
    const restored = await reopen(await presentation.save())
    expect(picture(restored.render(0)).dataUrl).toBe(`data:image/png;base64,${GREEN}`)
  })

  it('resolves background image fills through the same package', async () => {
    const { presentation } = await imageDeck()
    const result = await presentation.setBackground({
      scope: 'slide',
      slideIndex: 0,
      kind: 'image',
      mode: 'tile',
      data: GREEN,
      extension: 'png',
    })
    expect(result!.slides[0]!.background).toMatchObject({
      kind: 'image',
      dataUrl: `data:image/png;base64,${GREEN}`,
      mode: 'tile',
    })
    const restored = await reopen(await presentation.save())
    expect(restored.render(0).background).toMatchObject({
      dataUrl: `data:image/png;base64,${GREEN}`,
    })
  })

  it('resolves inherited master pictures and every master preview refresh', async () => {
    const { presentation } = await imageDeck()
    const zip = await JSZip.loadAsync(await presentation.save())
    const slidePath = 'ppt/slides/slide1.xml'
    const masterPath = 'ppt/slideMasters/slideMaster1.xml'
    const slideXml = await zip.file(slidePath)!.async('string')
    const pic = slideXml.match(/<p:pic\b[\s\S]*?<\/p:pic>/)![0]
    const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    const imageRel = slideRels.match(/<Relationship\b[^>]*Type="[^"]*\/image"[^>]*\/>/)![0]
    const rid = imageRel.match(/Id="([^"]+)"/)![1]!
    zip.file(slidePath, slideXml.replace(pic, ''))
    zip.file(
      masterPath,
      (await zip.file(masterPath)!.async('string')).replace(
        '</p:spTree>',
        `${pic.replace(`r:embed="${rid}"`, 'r:embed="masterImage"')}</p:spTree>`,
      ),
    )
    const relsPath = 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
    zip.file(
      relsPath,
      (await zip.file(relsPath)!.async('string')).replace(
        '</Relationships>',
        `${imageRel.replace(`Id="${rid}"`, 'Id="masterImage"')}</Relationships>`,
      ),
    )
    const restored = await reopen(await zip.generateAsync({ type: 'uint8array' }))
    expect(picture(restored.render(0))).toMatchObject({
      decoration: true,
      dataUrl: `data:image/png;base64,${BLUE}`,
    })
    const master = (await restored.masterParts()).find((part) => part.partPath === masterPath)!
    expect(picture(master.slide).dataUrl).toBe(`data:image/png;base64,${BLUE}`)
    const masterPicture = picture((await restored.masterPart(masterPath))!)
    const updated = await restored.setMasterTransform({
      partPath: masterPath,
      objectId: masterPicture.sourceId!,
      xEmu: 1828800,
      yEmu: 914400,
      widthEmu: 1828800,
      heightEmu: 914400,
      rotationDegrees: 0,
    })
    expect(picture(updated!).dataUrl).toBe(`data:image/png;base64,${BLUE}`)
  })
})
