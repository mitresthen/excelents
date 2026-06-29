import { expect, test } from 'vitest'
import { readZip } from './zip-reader'
import { writeZip } from './zip-writer'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

test('writeZip then readZip round-trips multiple entries', async () => {
  const archive = await writeZip([
    { name: '[Content_Types].xml', data: enc('<Types/>') },
    { name: 'xl/workbook.xml', data: enc('<workbook>' + 'x'.repeat(5000) + '</workbook>') },
    { name: 'empty.bin', data: new Uint8Array(0) },
  ])
  const parts = await readZip(archive)
  expect(new TextDecoder().decode(parts.get('[Content_Types].xml'))).toBe('<Types/>')
  expect(new TextDecoder().decode(parts.get('xl/workbook.xml'))).toBe(
    '<workbook>' + 'x'.repeat(5000) + '</workbook>',
  )
  const empty = parts.get('empty.bin')
  if (empty === undefined) throw new Error('missing empty.bin')
  expect(empty.length).toBe(0)
})

test('produces a valid EOCD signature', async () => {
  const archive = await writeZip([{ name: 'a.txt', data: enc('hello') }])
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)
  // EOCD is the last 22 bytes (no comment)
  expect(view.getUint32(archive.length - 22, true)).toBe(0x06054b50)
})

test('round-trips binary data with a correct CRC (no corruption)', async () => {
  const data = new Uint8Array(512)
  for (let i = 0; i < data.length; i++) data[i] = (i * 7) & 0xff
  const parts = await readZip(await writeZip([{ name: 'b.bin', data }]))
  const bin = parts.get('b.bin')
  if (bin === undefined) throw new Error('missing b.bin')
  expect([...bin]).toEqual([...data])
})
