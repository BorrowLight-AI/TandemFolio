import type { OpenFileResult } from '../../shared/lite-api'

type LoadFile = (file: OpenFileResult) => Promise<unknown>

export async function createDocxRecoverySnapshot(
  fileName: string,
  buildBytes: () => Promise<Uint8Array | null>,
): Promise<{ fileName: string; data: ArrayBuffer } | undefined> {
  try {
    const bytes = await buildBytes()
    if (!bytes) return undefined
    return {
      fileName,
      data: bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer,
    }
  } catch {
    return undefined
  }
}

async function sha256(data: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', data)
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

export async function loadStagedDocx(
  input: { readonly blobId: string; readonly name: string; readonly data: ArrayBuffer },
  loadFile: LoadFile,
): Promise<void> {
  await loadFile({
    path: `mcp-local:${input.blobId}/${input.name}`,
    name: input.name,
    data: input.data,
    hash: await sha256(input.data),
  })
}
