// Browser-safe adaptation of genspark-ai/genoffice Markdown asset scanning and Save As rewriting.
export interface MarkdownAssetBytes {
  readonly source: string
  readonly mime: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly data: ArrayBuffer
}

export interface MarkdownCompanionFile {
  readonly relativePath: string
  readonly data: ArrayBuffer
}

interface SourceRange {
  readonly start: number
  readonly end: number
  readonly source: string
}

function escapedAt(text: string, index: number): boolean {
  let slashes = 0
  for (let cursor = index - 1; cursor >= 0 && text[cursor] === '\\'; cursor -= 1) slashes += 1
  return slashes % 2 === 1
}

function codeRanges(markdown: string): Array<{ start: number; end: number }> {
  const ranges: Array<{ start: number; end: number }> = []
  let fence: { marker: string; length: number; start: number } | null = null
  let offset = 0
  for (const lineWithBreak of markdown.match(/.*(?:\n|$)/g) ?? []) {
    if (!lineWithBreak) continue
    const line = lineWithBreak.replace(/\r?\n$/, '')
    const end = offset + lineWithBreak.length
    if (fence) {
      const close = new RegExp(`^ {0,3}\\${fence.marker}{${fence.length},}[ \\t]*$`)
      if (close.test(line)) {
        ranges.push({ start: fence.start, end })
        fence = null
      }
    } else {
      const open = /^ {0,3}(`{3,}|~{3,})/.exec(line)
      if (open) fence = { marker: open[1]![0]!, length: open[1]!.length, start: offset }
      else if (/^(?: {4}|\t)/.test(line)) ranges.push({ start: offset, end })
    }
    offset = end
  }
  if (fence) ranges.push({ start: fence.start, end: markdown.length })
  for (const match of markdown.matchAll(/(`+)([^\n]*?)\1/g)) {
    const start = match.index
    if (!ranges.some((range) => start >= range.start && start < range.end)) {
      ranges.push({ start, end: start + match[0].length })
    }
  }
  return ranges.sort((left, right) => left.start - right.start)
}

function inside(ranges: ReadonlyArray<{ start: number; end: number }>, position: number): boolean {
  return ranges.some((range) => position >= range.start && position < range.end)
}

function markdownImageRanges(
  markdown: string,
  excluded: ReturnType<typeof codeRanges>,
): SourceRange[] {
  const ranges: SourceRange[] = []
  for (let index = 0; index < markdown.length - 2; index += 1) {
    if (
      markdown[index] !== '!' ||
      markdown[index + 1] !== '[' ||
      escapedAt(markdown, index) ||
      inside(excluded, index)
    ) {
      continue
    }
    const altEnd = markdown.indexOf('](', index + 2)
    if (altEnd < 0) continue
    let cursor = altEnd + 2
    let depth = 1
    let quote = ''
    let close = -1
    for (; cursor < markdown.length; cursor += 1) {
      const character = markdown[cursor]!
      if (escapedAt(markdown, cursor)) continue
      if (quote) {
        if (character === quote) quote = ''
      } else if (character === '"' || character === "'") quote = character
      else if (character === '(') depth += 1
      else if (character === ')' && --depth === 0) {
        close = cursor
        break
      }
    }
    if (close < 0) continue
    let start = altEnd + 2
    let end = close
    while (start < end && /\s/.test(markdown[start]!)) start += 1
    while (end > start && /\s/.test(markdown[end - 1]!)) end -= 1
    if (markdown[start] === '<' && markdown[end - 1] === '>') {
      start += 1
      end -= 1
    } else {
      const raw = markdown.slice(start, end)
      const title = /\s+(?:"[^"]*"|'[^']*')\s*$/.exec(raw)
      if (title?.index !== undefined) end = start + title.index
    }
    if (end > start) {
      ranges.push({
        start,
        end,
        source: markdown.slice(start, end).replace(/\\([\\()[\]<> ])/g, '$1'),
      })
    }
    index = close
  }
  return ranges
}

function decodeHtmlSource(value: string): string {
  const named: Readonly<Record<string, string>> = {
    amp: '&',
    apos: "'",
    gt: '>',
    lt: '<',
    nbsp: '\u00a0',
    quot: '"',
  }
  return value.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z]+);/gi, (_entity, body: string) => {
    if (body[0] === '#') {
      return String.fromCodePoint(
        Number.parseInt(
          body.slice(body[1]?.toLowerCase() === 'x' ? 2 : 1),
          body[1]?.toLowerCase() === 'x' ? 16 : 10,
        ),
      )
    }
    return named[body.toLowerCase()] ?? _entity
  })
}

function htmlImageRanges(markdown: string, excluded: ReturnType<typeof codeRanges>): SourceRange[] {
  const ranges: SourceRange[] = []
  for (const tag of markdown.matchAll(/<img\b[^>]*>/gi)) {
    const tagStart = tag.index
    if (inside(excluded, tagStart)) continue
    const sources = [...tag[0].matchAll(/\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi)]
    if (sources.length !== 1) continue
    const match = sources[0]!
    const raw = match[1] ?? match[2] ?? match[3] ?? ''
    const valueOffset = match[0].indexOf(raw)
    const start = tagStart + match.index + valueOffset
    ranges.push({ start, end: start + raw.length, source: decodeHtmlSource(raw) })
  }
  return ranges
}

function sourceRanges(markdown: string): SourceRange[] {
  const excluded = codeRanges(markdown)
  return [...markdownImageRanges(markdown, excluded), ...htmlImageRanges(markdown, excluded)].sort(
    (left, right) => left.start - right.start,
  )
}

export function extractMarkdownImageSources(markdown: string): string[] {
  return sourceRanges(markdown).map((range) => range.source)
}

export function rewriteMarkdownImageSources(
  markdown: string,
  rewrites: ReadonlyMap<string, string>,
): string {
  let output = ''
  let cursor = 0
  for (const range of sourceRanges(markdown)) {
    const replacement = rewrites.get(range.source)
    if (!replacement) continue
    output += markdown.slice(cursor, range.start) + replacement
    cursor = range.end
  }
  return cursor ? output + markdown.slice(cursor) : markdown
}

function safeRelativeSource(source: string): string | null {
  if (
    !source ||
    source.includes('\0') ||
    source.includes('?') ||
    source.includes('#') ||
    source.startsWith('/') ||
    source.startsWith('\\') ||
    /^[a-z][a-z0-9+.-]*:/i.test(source)
  ) {
    return null
  }
  let decoded: string
  try {
    decoded = decodeURIComponent(source).replace(/\\/g, '/')
  } catch {
    return null
  }
  const segments = decoded.split('/')
  return segments.some((segment) => !segment || segment === '.' || segment === '..')
    ? null
    : decoded
}

function decodeDataImage(source: string): MarkdownAssetBytes | null {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,([a-z0-9+/=\s]+)$/i.exec(source)
  if (!match) return null
  try {
    const binary = atob(match[2]!.replace(/\s/g, ''))
    if (binary.length < 1 || binary.length > 20_971_520) return null
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
    return {
      source,
      mime: match[1]!.toLowerCase() as MarkdownAssetBytes['mime'],
      data: bytes.buffer,
    }
  } catch {
    return null
  }
}

function safeStem(source: string): string {
  const name = source.replace(/\\/g, '/').split('/').pop() ?? 'image'
  const withoutExtension = name.replace(/\.[a-z0-9]{1,10}$/i, '')
  return (
    withoutExtension
      .replace(/[\x00-\x1f/\\:*?"<>|#%]+/g, '_')
      .replace(/\s+/g, '_')
      .replace(/^\.+/, '')
      .slice(0, 80)
      .replace(/[. ]+$/g, '') || 'image'
  )
}

async function digestPrefix(data: ArrayBuffer): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', data))
  return [...digest]
    .slice(0, 6)
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('')
}

export async function prepareMarkdownAssetSave(
  text: string,
  available: readonly MarkdownAssetBytes[],
  options: { readonly copyRelative: boolean },
): Promise<{
  readonly text: string
  readonly companionFiles: readonly MarkdownCompanionFile[]
  readonly rewrites: ReadonlyMap<string, string>
}> {
  const availableBySource = new Map(available.map((asset) => [asset.source, asset]))
  const rewrites = new Map<string, string>()
  const companionByPath = new Map<string, MarkdownCompanionFile>()
  for (const source of new Set(extractMarkdownImageSources(text))) {
    const embedded = decodeDataImage(source)
    const relative = safeRelativeSource(source)
    const asset = embedded ?? (relative ? availableBySource.get(source) : undefined)
    if (!asset) continue
    if (!embedded && /^assets\/[a-z0-9_.-]+-[a-f0-9]{12}\.(?:png|jpe?g|gif)$/i.test(relative!)) {
      companionByPath.set(relative!, { relativePath: relative!, data: asset.data })
      continue
    }
    if (!embedded && !options.copyRelative) {
      continue
    }
    const extension = asset.mime === 'image/jpeg' ? 'jpg' : asset.mime.slice('image/'.length)
    const relativePath = `assets/${safeStem(embedded ? 'image' : relative!)}-${await digestPrefix(asset.data)}.${extension}`
    rewrites.set(source, relativePath)
    companionByPath.set(relativePath, { relativePath, data: asset.data })
  }
  return {
    text: rewriteMarkdownImageSources(text, rewrites),
    companionFiles: [...companionByPath.values()],
    rewrites,
  }
}
