import { expect, test } from 'vitest'
import { bytesToReadable, concatUint8, readableToBytes } from './streams'

test('concatUint8 joins chunks in order', () => {
  const out = concatUint8([new Uint8Array([1, 2]), new Uint8Array([]), new Uint8Array([3])])
  expect([...out]).toEqual([1, 2, 3])
})

test('concatUint8 of nothing is empty', () => {
  expect(concatUint8([]).length).toBe(0)
})

test('bytesToReadable then readableToBytes round-trips', async () => {
  const input = new Uint8Array([10, 20, 30, 40])
  const out = await readableToBytes(bytesToReadable(input))
  expect([...out]).toEqual([10, 20, 30, 40])
})

test('readableToBytes collects a multi-chunk stream', async () => {
  const stream = new ReadableStream<Uint8Array>({
    start(c) {
      c.enqueue(new Uint8Array([1]))
      c.enqueue(new Uint8Array([2, 3]))
      c.close()
    },
  })
  expect([...(await readableToBytes(stream))]).toEqual([1, 2, 3])
})
