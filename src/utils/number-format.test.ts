import { createRequire } from 'node:module'
import { expect, test } from 'vitest'
import { BUILTIN_FORMATS, builtinFormatCode, builtinFormatId } from './number-format'

test('known builtin ids map to their canonical codes', () => {
  expect(builtinFormatCode(0)).toBe('General')
  expect(builtinFormatCode(1)).toBe('0')
  expect(builtinFormatCode(2)).toBe('0.00')
  expect(builtinFormatCode(9)).toBe('0%')
  expect(builtinFormatCode(49)).toBe('@')
  // id 22 canonical (exceljs has quoted-h rendering bug; we use ECMA-376 §18.8.30 form)
  expect(builtinFormatCode(22)).toBe('m/d/yy h:mm')
})

test('unknown ids return undefined', () => {
  expect(builtinFormatCode(5)).toBeUndefined() // 5 is a gap in the builtin table
  expect(builtinFormatCode(999)).toBeUndefined()
})

test('builtinFormatId reverses builtinFormatCode for unambiguous codes', () => {
  expect(builtinFormatId('General')).toBe(0)
  expect(builtinFormatId('0.00')).toBe(2)
  expect(builtinFormatId('@')).toBe(49)
  expect(builtinFormatId('not-a-builtin')).toBeUndefined()
})

test('the table agrees with the exceljs oracle defaults (real cross-check)', () => {
  // Load exceljs's authoritative built-in format table via createRequire (CJS module).
  // Export shape: { [id: number]: { f: string } | { [locale: string]: string } }
  // Only entries with an `f` property are universal (non-locale-specific).
  // Locale-specific entries (ids 27-36, 50-58, 59-70, 81) are intentionally omitted
  // from our table — they require locale context to resolve.
  type ExcelEntry = { f?: string } & Record<string, string | undefined>
  type ExcelDefaults = Record<number, ExcelEntry>

  const require = createRequire(import.meta.url)
  // eslint-disable-next-line typescript/no-unsafe-type-assertion
  const excelDefaults = require('exceljs/lib/xlsx/defaultnumformats.js') as unknown as ExcelDefaults

  // For every id in the INTERSECTION of our table and exceljs's universal entries,
  // assert our code string matches exceljs's `f` property exactly.
  const excelUniversalIds: number[] = Object.keys(excelDefaults)
    .map(Number)
    .filter((id) => {
      const entry = excelDefaults[id]
      return entry !== undefined && typeof entry.f === 'string'
    })

  const ourIds = new Set(Object.keys(BUILTIN_FORMATS).map(Number))

  let checkedCount = 0
  for (const id of excelUniversalIds) {
    if (ourIds.has(id)) {
      if (id === 22) {
        // exceljs stores 'm/d/yy "h":mm' (quoted-h = literal). ECMA-376 §18.8.30 canonical
        // is 'm/d/yy h:mm' (hour token). We intentionally deviate; builtin codes are never
        // written to files, so there is no interop cost. Excluded from strict equality.
        checkedCount++
        continue
      }
      const entry = excelDefaults[id]
      const excelCode = entry?.f
      if (excelCode === undefined) continue
      expect(
        BUILTIN_FORMATS[id],
        `id ${id}: our code '${String(BUILTIN_FORMATS[id])}' should match exceljs '${excelCode}'`,
      ).toBe(excelCode)
      checkedCount++
    }
  }

  // Sanity: every id in our table must have been verified against exceljs's universal set.
  // If checkedCount < ourIds.size, we have ids in our table that exceljs treats as locale-only
  // or does not include at all — that would be a gap in the cross-check.
  expect(checkedCount).toBe(ourIds.size)

  // Coverage delta: ids present in exceljs universal set but absent from our table.
  // None expected — all universal exceljs entries are encoded in our table.
  // Ids 27-36, 50-58, 59-70, 81 are locale-specific (no `f` property) and intentionally omitted.
  const excelOnlyIds = excelUniversalIds.filter((id) => !ourIds.has(id))
  if (excelOnlyIds.length > 0) {
    console.info(`[oracle] exceljs universal ids not in our table: ${excelOnlyIds.join(', ')}`)
  }
})
