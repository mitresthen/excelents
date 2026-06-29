import { expect, test } from 'vitest'
import { readZip } from './zip-reader'

test('throws on non-zip input', async () => {
  await expect(readZip(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow()
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
