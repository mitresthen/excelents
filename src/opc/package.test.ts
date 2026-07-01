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

test('getPart returns a defensive copy: mutating the result does not affect subsequent calls', () => {
  const pkg = OpcPackage.empty()
  pkg.setPart('xl/workbook.xml', 'application/test+xml', enc('<workbook/>'))
  const first = pkg.getPart('xl/workbook.xml')!
  first[0] = 0xff // mutate the returned buffer
  const second = pkg.getPart('xl/workbook.xml')!
  expect(second[0]).not.toBe(0xff)
})

test('setDefault + addPart: media resolves via the Default content type (no Override)', async () => {
  const pkg = OpcPackage.empty()
  pkg.setDefault('png', 'image/png')
  pkg.addPart('xl/media/image1.png', new Uint8Array([1, 2, 3]))
  const rt = await OpcPackage.read(await pkg.toBytes())
  expect(Array.from(rt.getPart('xl/media/image1.png')!)).toEqual([1, 2, 3])
  expect(rt.contentTypeOf('xl/media/image1.png')).toBe('image/png')
})

async function buildMinimal(): Promise<Uint8Array> {
  // Construct a minimal package via setPart + toBytes, then return its bytes.
  const { OpcPackage: P } = await import('./package')
  const pkg = P.empty()
  pkg.setPart('xl/workbook.xml', 'application/test+xml', enc('<workbook/>'))
  return pkg.toBytes()
}
