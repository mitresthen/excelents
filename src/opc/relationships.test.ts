import { expect, test } from 'vitest'
import { OpcPackage } from './package'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

const REL = (id: string, type: string, target: string, mode?: string): string =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`

async function pkgWith(parts: Record<string, string>): Promise<OpcPackage> {
  const pkg = OpcPackage.empty()
  for (const [name, body] of Object.entries(parts)) pkg.setPart(name, 'application/xml', enc(body))
  // round-trip through bytes so we exercise the real read path
  return OpcPackage.read(await pkg.toBytes())
}

test('relationshipsFor resolves targets relative to the part directory', async () => {
  const pkg = await pkgWith({
    'xl/workbook.xml': '<workbook/>',
    'xl/_rels/workbook.xml.rels':
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      REL('rId1', 'http://x/worksheet', 'worksheets/sheet1.xml') +
      REL('rId2', 'http://x/styles', '/xl/styles.xml') +
      REL('rId3', 'http://x/hyperlink', 'https://example.com/', 'External') +
      `</Relationships>`,
  })
  const rels = pkg.relationshipsFor('xl/workbook.xml')
  expect(rels.find((r) => r.id === 'rId1')?.target).toBe('xl/worksheets/sheet1.xml')
  expect(rels.find((r) => r.id === 'rId2')?.target).toBe('xl/styles.xml')
  // External targets are returned verbatim.
  expect(rels.find((r) => r.id === 'rId3')?.target).toBe('https://example.com/')
})

test('relationshipsFor returns [] when the part has no rels', async () => {
  const pkg = await pkgWith({ 'xl/workbook.xml': '<workbook/>' })
  expect(pkg.relationshipsFor('xl/workbook.xml')).toEqual([])
})

test('relationshipsFor normalizes parent-directory (..) segments in targets', async () => {
  const pkg = await pkgWith({
    'xl/worksheets/sheet1.xml': '<worksheet/>',
    'xl/worksheets/_rels/sheet1.xml.rels':
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      REL('rId1', 'http://x/image', '../media/image1.png') +
      `</Relationships>`,
  })
  const rels = pkg.relationshipsFor('xl/worksheets/sheet1.xml')
  expect(rels.find((r) => r.id === 'rId1')?.target).toBe('xl/media/image1.png')
})
