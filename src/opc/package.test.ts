import { expect, test } from 'vitest'
import { OpcPackage } from './package'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

test('reads content types and root relationships from a real fixture', async () => {
  const { readFile } = await import('node:fs/promises')
  const bytes = new Uint8Array(
    await readFile(new URL('../../test/fixtures/1904.xlsx', import.meta.url)),
  )
  const pkg = await OpcPackage.read(bytes)
  expect(pkg.partNames()).toContain('xl/workbook.xml')
  expect(pkg.contentTypeOf('xl/workbook.xml')).toContain('spreadsheetml')
  // the root rels point at the workbook (type ends with .../officeDocument, not extended-properties)
  const officeDoc = pkg
    .rootRelationships()
    .find((r) => r.type.endsWith('relationships/officeDocument'))
  expect(officeDoc?.target).toContain('workbook.xml')
})

test('builds a package and round-trips through read', async () => {
  const pkg = await OpcPackage.read(await buildMinimal())
  expect(pkg.contentTypeOf('xl/workbook.xml')).toBe('application/test+xml')
  expect(dec(pkg.getPart('xl/workbook.xml')!)).toBe('<workbook/>')
})

async function buildMinimal(): Promise<Uint8Array> {
  // Construct a minimal package via setPart + toBytes, then return its bytes.
  const { OpcPackage: P } = await import('./package')
  const pkg = P.empty()
  pkg.setPart('xl/workbook.xml', 'application/test+xml', enc('<workbook/>'))
  return pkg.toBytes()
}
