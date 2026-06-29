import { expect, test } from 'vitest'
import { readableToBytes } from './streams'
import { createMemoryFileSystem } from './fs'

test('writeFile then readFile round-trips', async () => {
  const fs = createMemoryFileSystem()
  await fs.writeFile('a.bin', new Uint8Array([1, 2, 3]))
  expect([...(await fs.readFile('a.bin'))]).toEqual([1, 2, 3])
})

test('readFile of a missing path rejects', async () => {
  const fs = createMemoryFileSystem()
  await expect(fs.readFile('nope')).rejects.toThrow()
})

test('seeds from initial contents', async () => {
  const fs = createMemoryFileSystem({ 'seed.txt': new Uint8Array([9]) })
  expect([...(await fs.readFile('seed.txt'))]).toEqual([9])
})

test('createReadable streams a written file', async () => {
  const fs = createMemoryFileSystem()
  await fs.writeFile('r.bin', new Uint8Array([5, 6, 7]))
  expect([...(await readableToBytes(fs.createReadable('r.bin')))]).toEqual([5, 6, 7])
})

test('createWritable collects chunks into a readable file', async () => {
  const fs = createMemoryFileSystem()
  const w = fs.createWritable('w.bin')
  const writer = w.getWriter()
  await writer.write(new Uint8Array([1, 2]))
  await writer.write(new Uint8Array([3]))
  await writer.close()
  expect([...(await fs.readFile('w.bin'))]).toEqual([1, 2, 3])
})

test('mutation isolation on write: mutating original does not affect stored bytes', async () => {
  const fs = createMemoryFileSystem()
  const buf = new Uint8Array([1, 2, 3])
  await fs.writeFile('a.bin', buf)
  buf[0] = 99
  expect([...(await fs.readFile('a.bin'))]).toEqual([1, 2, 3])
})

test('mutation isolation on read: mutating returned buffer does not affect store', async () => {
  const fs = createMemoryFileSystem()
  await fs.writeFile('a.bin', new Uint8Array([1, 2, 3]))
  const result = await fs.readFile('a.bin')
  result[0] = 99
  expect([...(await fs.readFile('a.bin'))]).toEqual([1, 2, 3])
})

test('seed isolation: mutating seed buffer does not affect stored bytes', async () => {
  const seed = new Uint8Array([7, 8, 9])
  const fs = createMemoryFileSystem({ 'seed.bin': seed })
  seed[0] = 99
  expect([...(await fs.readFile('seed.bin'))]).toEqual([7, 8, 9])
})

test('createReadable throws synchronously on missing path', () => {
  const fs = createMemoryFileSystem()
  expect(() => fs.createReadable('missing')).toThrow()
})
