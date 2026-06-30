import { expect, test } from 'vitest'
import { readZip } from './zip-reader'
import { writeZip } from './zip-writer'

test('throws on non-zip input', async () => {
  await expect(readZip(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow()
})

test('throws on unsupported compression method', async () => {
  const archive = await writeZip([{ name: 'x.txt', data: new Uint8Array([97, 98, 99]) }])
  // Work on a mutable copy (slice() guarantees byteOffset=0 and a fresh buffer)
  const copy = archive.slice()
  const view = new DataView(copy.buffer)
  // Locate EOCD signature by scanning backward
  const EOCD_SIG = 0x06054b50
  let eocd = -1
  for (let i = copy.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  // Patch the compression method field in the central directory to 6 (unsupported)
  const cdOffset = view.getUint32(eocd + 16, true)
  view.setUint16(cdOffset + 10, 6, true)
  await expect(readZip(copy)).rejects.toThrow('Unsupported ZIP compression method 6')
})

test('reads a real xlsx fixture and finds the core OPC parts', async () => {
  const { readFile } = await import('node:fs/promises')
  const url = new URL('../../test/fixtures/1904.xlsx', import.meta.url)
  const bytes = new Uint8Array(await readFile(url))
  const parts = await readZip(bytes)
  expect(parts.has('[Content_Types].xml')).toBe(true)
  expect(parts.has('xl/workbook.xml')).toBe(true)
  // the workbook part is real XML
  expect(new TextDecoder().decode(parts.get('xl/workbook.xml'))).toContain('<workbook')
})
