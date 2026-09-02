import { createWriteStream } from 'node:fs'
import { pipeline } from 'node:stream/promises'
import type { OpenedPptx } from './index'
import { buildPptxZip } from './index'

/** Node-only streaming output. Browser consumers never import this subpath. */
export async function savePptxToFile(opened: OpenedPptx, filePath: string): Promise<void> {
  const source = buildPptxZip(opened).generateNodeStream({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    streamFiles: true,
  })
  await pipeline(source, createWriteStream(filePath))
}
