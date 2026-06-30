import { expect, test } from 'vitest'
import { readableToBytes } from '../io/streams'
import { readZip } from './zip-reader'
import { writeZipStream } from './zip-stream'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

async function* chunks(...parts: string[]): AsyncGenerator<Uint8Array> {
  for (const p of parts) yield enc(p)
}

test('writeZipStream output round-trips through readZip (multi-chunk + multi-entry)', async () => {
  async function* entries() {
    yield { name: 'a.txt', data: chunks('hello ', 'world') }
    yield { name: 'dir/b.bin', data: chunks('x'.repeat(2000)) }
  }
  const bytes = await readableToBytes(writeZipStream(entries()))
  const parts = await readZip(bytes)
  expect(dec(parts.get('a.txt')!)).toBe('hello world')
  expect(dec(parts.get('dir/b.bin')!)).toBe('x'.repeat(2000))
  expect(parts.size).toBe(2)
})

test('an empty entry round-trips', async () => {
  async function* entries() {
    yield { name: 'empty', data: chunks() }
  }
  const bytes = await readableToBytes(writeZipStream(entries()))
  expect(dec((await readZip(bytes)).get('empty')!)).toBe('')
})

test('the output is emitted as multiple chunks (truly streamed, not one buffer)', async () => {
  async function* entries() {
    yield { name: 'big', data: chunks('y'.repeat(50000)) }
  }
  const reader = writeZipStream(entries()).getReader()
  let count = 0
  for (;;) {
    const { done } = await reader.read()
    if (done) break
    count++
  }
  expect(count).toBeGreaterThan(1)
})
