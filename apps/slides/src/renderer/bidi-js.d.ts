declare module 'bidi-js' {
  interface BidiFactory {
    getEmbeddingLevels(text: string, explicitDirection?: 'ltr' | 'rtl'): { levels: Uint8Array }
    getReorderedIndices(text: string, embeddingLevels: { levels: Uint8Array }): number[]
  }
  export default function bidiFactory(): BidiFactory
}
