import { expect, test } from 'vitest'
import { nativeCodec } from './codec'

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

test('deflateRaw then inflateRaw recovers the original bytes', async () => {
  const input = enc('hello world '.repeat(50))
  const compressed = await nativeCodec.deflateRaw(input)
  const restored = await nativeCodec.inflateRaw(compressed)
  expect(dec(restored)).toBe(dec(input))
})

test('deflateRaw actually compresses redundant input', async () => {
  const input = enc('A'.repeat(10_000))
  const compressed = await nativeCodec.deflateRaw(input)
  expect(compressed.length).toBeLessThan(input.length)
})

test('round-trips empty input', async () => {
  const out = await nativeCodec.inflateRaw(await nativeCodec.deflateRaw(new Uint8Array(0)))
  expect(out.length).toBe(0)
})

test('round-trips binary (non-text) bytes', async () => {
  const input = new Uint8Array(256)
  for (let i = 0; i < 256; i++) input[i] = i
  const out = await nativeCodec.inflateRaw(await nativeCodec.deflateRaw(input))
  expect([...out]).toEqual([...input])
})
