export type MarkdownSaveResult =
  { readonly ok: true; readonly fileName: string } | { readonly ok: false }

/** Serialize save requests so a second save waits for the in-flight write. */
export function createMarkdownSaveQueue(
  run: (saveAs: boolean) => Promise<MarkdownSaveResult>,
): (saveAs: boolean) => Promise<MarkdownSaveResult> {
  let tail: Promise<void> = Promise.resolve()
  return (saveAs) => {
    const result = tail.then(
      () => run(saveAs),
      () => run(saveAs),
    )
    tail = result.then(
      () => undefined,
      () => undefined,
    )
    return result
  }
}
