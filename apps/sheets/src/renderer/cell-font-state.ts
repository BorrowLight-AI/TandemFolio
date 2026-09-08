type CellFontRuntime = typeof globalThis & {
  __genofficeSubstitutedCellFamilies?: Set<string>
}

const runtime = globalThis as CellFontRuntime
const substitutedFamilies = (runtime.__genofficeSubstitutedCellFamilies ??= new Set<string>())

export function markSubstitutedCellFamily(family: string): void {
  substitutedFamilies.add(family.toLowerCase())
}

export function isSubstitutedCellFamily(family: string): boolean {
  return substitutedFamilies.has(family.toLowerCase())
}
