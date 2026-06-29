import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, expect, test } from 'vitest'
import { readableToBytes } from './io/streams'
import { nodeFileSystem } from './node'

let dir = ''
beforeAll(() => {
  dir = mkdtempSync(join(tmpdir(), 'excelents-node-'))
})
afterAll(() => {
  rmSync(dir, { recursive: true, force: true })
})

test('writeFile then readFile round-trips on disk', async () => {
  const p = join(dir, 'a.bin')
  await nodeFileSystem.writeFile(p, new Uint8Array([1, 2, 3]))
  expect([...(await nodeFileSystem.readFile(p))]).toEqual([1, 2, 3])
})

test('createReadable streams a file from disk', async () => {
  const p = join(dir, 'r.bin')
  await nodeFileSystem.writeFile(p, new Uint8Array([7, 8, 9]))
  expect([...(await readableToBytes(nodeFileSystem.createReadable(p)))]).toEqual([7, 8, 9])
})

test('createWritable writes a file to disk', async () => {
  const p = join(dir, 'w.bin')
  const writer = nodeFileSystem.createWritable(p).getWriter()
  await writer.write(new Uint8Array([4, 5]))
  await writer.write(new Uint8Array([6]))
  await writer.close()
  expect([...(await nodeFileSystem.readFile(p))]).toEqual([4, 5, 6])
})
