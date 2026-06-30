import { type Codec, nativeCodec } from '../io/codec'

const EOCD_SIG = 0x06054b50
const CD_SIG = 0x02014b50

/** One ZIP entry located via the central directory, with its raw (possibly compressed) body. */
export interface ZipEntryLocation {
  readonly name: string
  /** Compression method: 0 = stored, 8 = DEFLATE. */
  readonly method: number
  /** The entry's raw stored/compressed bytes, sliced from the archive (not decompressed). */
  readonly body: Uint8Array
}

/**
 * Walk the central directory and return every entry's name, method and raw body WITHOUT
 * decompressing — the random-access primitive a streaming reader uses to inflate one part
 * (e.g. a worksheet's `sheetData`) at a time. The central directory carries the final sizes,
 * so this works for both Excel files and data-descriptor archives (our streaming writer).
 */
export function readZipEntries(bytes: Uint8Array): ZipEntryLocation[] {
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
  const out: ZipEntryLocation[] = []

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
    out.push({ name, method, body })

    p += 46 + nameLen + extraLen + commentLen
  }
  return out
}

/** Decompress one located entry's raw body to its uncompressed bytes. */
export async function inflateEntry(
  entry: ZipEntryLocation,
  codec: Codec = nativeCodec,
): Promise<Uint8Array> {
  if (entry.method === 8) return codec.inflateRaw(entry.body)
  if (entry.method === 0) return entry.body.slice()
  throw new Error(`Unsupported ZIP compression method ${entry.method} for ${entry.name}`)
}

/** Read a standard ZIP archive into a map of entry name -> decompressed bytes. */
export async function readZip(
  bytes: Uint8Array,
  codec: Codec = nativeCodec,
): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>()
  for (const entry of readZipEntries(bytes)) {
    out.set(entry.name, await inflateEntry(entry, codec))
  }
  return out
}
