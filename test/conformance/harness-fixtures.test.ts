import { expect, test } from 'vitest'
import { canonicalizeXml, diffParts, listFixtures, parseFixture, unzipToParts } from './harness'

test('canonicalizeXml sorts attributes so order does not matter', () => {
  expect(canonicalizeXml('<c r="A1" s="2" t="n"/>')).toBe(
    canonicalizeXml('<c t="n" s="2" r="A1"/>'),
  )
  expect(canonicalizeXml('<a x="1"/>')).not.toBe(canonicalizeXml('<a x="2"/>'))
})

test('unzipToParts canonicalizes every xlsx fixture and finds [Content_Types].xml', async () => {
  // missing-bits.xlsx is intentionally an incomplete ZIP with only 2 entries
  // (xl/_rels/workbook.xml.rels + xl/worksheets/sheet1.xml, no [Content_Types].xml).
  // It is designed to test missing-part handling — the reader correctly reads it.
  const KNOWN_INCOMPLETE = new Set(['missing-bits.xlsx'])

  for (const path of listFixtures()) {
    const name = path.split('/').pop()!
    const parts = await unzipToParts(await parseFixture(path))
    if (!KNOWN_INCOMPLETE.has(name)) {
      expect(Object.keys(parts)).toContain('[Content_Types].xml')
    }
  }
}, 60_000)

test('diffParts reports no differences for identical part maps and flags real ones', async () => {
  const parts = await unzipToParts(await parseFixture(listFixtures()[0]!))
  expect(diffParts(parts, parts)).toEqual([])
  const mutated = { ...parts, 'extra.xml': '<x/>' }
  expect(diffParts(parts, mutated).length).toBeGreaterThan(0)
}, 60_000)
