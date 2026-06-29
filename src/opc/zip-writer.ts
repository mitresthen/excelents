import { type Codec, nativeCodec } from '../io/codec'
import { concatUint8 } from '../io/streams'
import { crc32 } from './crc32'

export interface ZipEntry {
  name: string
  data: Uint8Array
}

const DOS_DATE = 0x0021 // 1980-01-01 (minimum DOS date)
const UTF8_FLAG = 0x0800 // bit 11: filename is UTF-8

/** Produce a standard ZIP archive; each entry DEFLATE'd, or STORE'd if smaller. */
export async function writeZip(
  entries: ZipEntry[],
  codec: Codec = nativeCodec,
): Promise<Uint8Array> {
  const encoder = new TextEncoder()
  const locals: Uint8Array[] = []
  const centrals: Uint8Array[] = []
  let offset = 0

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name)
    const uncompSize = entry.data.length
    const crc = crc32(entry.data)
    const deflated = await codec.deflateRaw(entry.data)
    const store = deflated.length >= uncompSize
    const method = store ? 0 : 8
    const body = store ? entry.data : deflated
    const compSize = body.length

    const local = new Uint8Array(30 + nameBytes.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true)
    lv.setUint16(4, 20, true)
    lv.setUint16(6, UTF8_FLAG, true)
    lv.setUint16(8, method, true)
    lv.setUint16(10, 0, true)
    lv.setUint16(12, DOS_DATE, true)
    lv.setUint32(14, crc, true)
    lv.setUint32(18, compSize, true)
    lv.setUint32(22, uncompSize, true)
    lv.setUint16(26, nameBytes.length, true)
    lv.setUint16(28, 0, true)
    local.set(nameBytes, 30)
    locals.push(local, body)

    const central = new Uint8Array(46 + nameBytes.length)
    const cv = new DataView(central.buffer)
    cv.setUint32(0, 0x02014b50, true)
    cv.setUint16(4, 20, true)
    cv.setUint16(6, 20, true)
    cv.setUint16(8, UTF8_FLAG, true)
    cv.setUint16(10, method, true)
    cv.setUint16(12, 0, true)
    cv.setUint16(14, DOS_DATE, true)
    cv.setUint32(16, crc, true)
    cv.setUint32(20, compSize, true)
    cv.setUint32(24, uncompSize, true)
    cv.setUint16(28, nameBytes.length, true)
    cv.setUint16(30, 0, true)
    cv.setUint16(32, 0, true)
    cv.setUint16(34, 0, true)
    cv.setUint16(36, 0, true)
    cv.setUint32(38, 0, true)
    cv.setUint32(42, offset, true)
    central.set(nameBytes, 46)
    centrals.push(central)

    offset += local.length + body.length
  }

  const cdOffset = offset
  let cdSize = 0
  for (const c of centrals) cdSize += c.length

  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(4, 0, true)
  ev.setUint16(6, 0, true)
  ev.setUint16(8, entries.length, true)
  ev.setUint16(10, entries.length, true)
  ev.setUint32(12, cdSize, true)
  ev.setUint32(16, cdOffset, true)
  ev.setUint16(20, 0, true)

  return concatUint8([...locals, ...centrals, eocd])
}
