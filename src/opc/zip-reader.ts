import { type Codec, nativeCodec } from '../io/codec'

const EOCD_SIG = 0x06054b50
const CD_SIG = 0x02014b50

/** Read a standard ZIP archive into a map of entry name -> decompressed bytes. */
export async function readZip(
  bytes: Uint8Array,
  codec: Codec = nativeCodec,
): Promise<Map<string, Uint8Array>> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd === -1) throw new Error('Not a ZIP archive: no end-of-central-directory record')

  const count = view.getUint16(eocd + 10, true)
  let p = view.getUint32(eocd + 16, true)
  const decoder = new TextDecoder()
  const out = new Map<string, Uint8Array>()

  for (let n = 0; n < count; n++) {
    if (view.getUint32(p, true) !== CD_SIG) {
      throw new Error('Corrupt ZIP: bad central-directory signature')
    }
    const method = view.getUint16(p + 10, true)
    const compSize = view.getUint32(p + 20, true)
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    const localOffset = view.getUint32(p + 42, true)
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen))

    const localNameLen = view.getUint16(localOffset + 26, true)
    const localExtraLen = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + localNameLen + localExtraLen
    const body = bytes.subarray(dataStart, dataStart + compSize)
    let data: Uint8Array
    if (method === 8) {
      data = await codec.inflateRaw(body)
    } else if (method === 0) {
      data = body.slice()
    } else {
      throw new Error(`Unsupported ZIP compression method ${method} for ${name}`)
    }
    out.set(name, data)

    p += 46 + nameLen + extraLen + commentLen
  }
  return out
}
