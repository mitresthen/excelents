import { crc32Final, crc32Init, crc32Update } from './crc32'

/** One streamed ZIP entry: a name plus its (chunked) uncompressed data. */
export interface ZipStreamEntry {
  name: string
  data: AsyncIterable<Uint8Array> | Iterable<Uint8Array>
}

const LOCAL_SIG = 0x04034b50
const CENTRAL_SIG = 0x02014b50
const EOCD_SIG = 0x06054b50
const DD_SIG = 0x08074b50 // data descriptor
const UTF8_FLAG = 0x0800 // bit 11: UTF-8 name
const DD_FLAG = 0x0008 // bit 3: crc + sizes follow in a data descriptor
const DOS_DATE = 0x0021 // 1980-01-01
const DEFLATE = 8

/**
 * Stream a ZIP archive as bytes become available. Each entry is DEFLATE'd through the
 * platform `CompressionStream` while its CRC-32 and sizes are computed incrementally, so no
 * entry is ever held whole in memory. Sizes are unknown when the local header is written, so
 * bit 3 is set and a data descriptor follows each entry; the central directory (authoritative
 * for `readZip`) carries the final sizes.
 */
export function writeZipStream(
  entries: AsyncIterable<ZipStreamEntry> | Iterable<ZipStreamEntry>,
): ReadableStream<Uint8Array> {
  const gen = zipChunks(entries)
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      const { done, value } = await gen.next()
      if (done) controller.close()
      else controller.enqueue(value)
    },
    async cancel(reason) {
      await gen.return(reason)
    },
  })
}

async function* zipChunks(
  entries: AsyncIterable<ZipStreamEntry> | Iterable<ZipStreamEntry>,
): AsyncGenerator<Uint8Array> {
  const encoder = new TextEncoder()
  const centrals: Uint8Array[] = []
  let offset = 0
  let count = 0

  for await (const entry of entries) {
    const nameBytes = encoder.encode(entry.name)
    const localOffset = offset

    const local = new Uint8Array(30 + nameBytes.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, LOCAL_SIG, true)
    lv.setUint16(4, 20, true)
    lv.setUint16(6, UTF8_FLAG | DD_FLAG, true)
    lv.setUint16(8, DEFLATE, true)
    lv.setUint16(12, DOS_DATE, true)
    // crc (14), compSize (18), uncompSize (22) stay 0 — supplied by the data descriptor.
    lv.setUint16(26, nameBytes.length, true)
    local.set(nameBytes, 30)
    yield local
    offset += local.length

    const cs = new CompressionStream('deflate-raw')
    const writer = cs.writable.getWriter()
    const reader = cs.readable.getReader()
    let crc = crc32Init()
    let uncompSize = 0
    let compSize = 0

    // Feed source chunks into the compressor in the background while we drain its output.
    const fed = (async () => {
      for await (const chunk of entry.data) {
        crc = crc32Update(crc, chunk)
        uncompSize += chunk.length
        await writer.write(new Uint8Array(chunk))
      }
      await writer.close()
    })()

    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      if (value !== undefined && value.length > 0) {
        compSize += value.length
        offset += value.length
        yield value
      }
    }
    await fed // surface any error from the source

    const finalCrc = crc32Final(crc)

    const dd = new Uint8Array(16)
    const ddv = new DataView(dd.buffer)
    ddv.setUint32(0, DD_SIG, true)
    ddv.setUint32(4, finalCrc, true)
    ddv.setUint32(8, compSize, true)
    ddv.setUint32(12, uncompSize, true)
    yield dd
    offset += dd.length

    const central = new Uint8Array(46 + nameBytes.length)
    const cv = new DataView(central.buffer)
    cv.setUint32(0, CENTRAL_SIG, true)
    cv.setUint16(4, 20, true)
    cv.setUint16(6, 20, true)
    cv.setUint16(8, UTF8_FLAG | DD_FLAG, true)
    cv.setUint16(10, DEFLATE, true)
    cv.setUint16(14, DOS_DATE, true)
    cv.setUint32(16, finalCrc, true)
    cv.setUint32(20, compSize, true)
    cv.setUint32(24, uncompSize, true)
    cv.setUint16(28, nameBytes.length, true)
    cv.setUint32(42, localOffset, true)
    central.set(nameBytes, 46)
    centrals.push(central)
    count++
  }

  const cdOffset = offset
  let cdSize = 0
  for (const c of centrals) {
    cdSize += c.length
    yield c
  }

  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, EOCD_SIG, true)
  ev.setUint16(8, count, true)
  ev.setUint16(10, count, true)
  ev.setUint32(12, cdSize, true)
  ev.setUint32(16, cdOffset, true)
  yield eocd
}
