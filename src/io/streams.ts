/** Concatenate byte chunks into a single `Uint8Array`. */
export function concatUint8(chunks: Uint8Array[]): Uint8Array {
  let total = 0
  for (const c of chunks) total += c.length
  const out = new Uint8Array(total)
  let offset = 0
  for (const c of chunks) {
    out.set(c, offset)
    offset += c.length
  }
  return out
}

/** Wrap bytes in a single-chunk `ReadableStream`. */
export function bytesToReadable(bytes: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes)
      controller.close()
    },
  })
}

/** Collect a byte stream into one `Uint8Array`. */
export async function readableToBytes(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    if (value !== undefined) chunks.push(value)
  }
  return concatUint8(chunks)
}
