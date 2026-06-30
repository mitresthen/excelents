import { readableToBytes } from './streams'

/** A raw-DEFLATE compression codec (no zlib/gzip header — the format ZIP uses). */
export interface Codec {
  deflateRaw(data: Uint8Array): Promise<Uint8Array>
  inflateRaw(data: Uint8Array): Promise<Uint8Array>
}

async function transform(
  data: Uint8Array,
  writable: WritableStream<BufferSource>,
  readable: ReadableStream<Uint8Array>,
): Promise<Uint8Array> {
  // new Uint8Array(typedArray) always produces Uint8Array<ArrayBuffer> — satisfies BufferSource
  const copy: Uint8Array<ArrayBuffer> = new Uint8Array(data)
  const writer = writable.getWriter()
  // Write and read must run concurrently: writer.close() can deadlock if the
  // readable side isn't being consumed (backpressure blocks the flush).
  const [, result] = await Promise.all([
    (async () => {
      await writer.write(copy)
      await writer.close()
    })(),
    readableToBytes(readable),
  ])
  return result
}

/** Default codec using the platform Compression Streams API (`deflate-raw`). */
export const nativeCodec: Codec = {
  deflateRaw(data: Uint8Array): Promise<Uint8Array> {
    const cs = new CompressionStream('deflate-raw')
    return transform(data, cs.writable, cs.readable)
  },
  inflateRaw(data: Uint8Array): Promise<Uint8Array> {
    const ds = new DecompressionStream('deflate-raw')
    return transform(data, ds.writable, ds.readable)
  },
}
