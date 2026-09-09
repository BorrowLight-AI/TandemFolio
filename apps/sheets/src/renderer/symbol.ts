/// Appends a picked symbol to whatever the cell already holds. Numbers and
/// booleans coerce to text: Excel's Symbol dialog also turns the cell into
/// text when characters are appended.
export function appendSymbol(
  existing: string | number | boolean | null | undefined,
  char: string,
): string {
  if (existing === null || existing === undefined || existing === '') return char
  return `${String(existing)}${char}`
}
