function buildTable(): Uint32Array {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) {
      c = (c & 1) === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    }
    table[n] = c >>> 0
  }
  return table
}

const CRC_TABLE = buildTable()

/** Begin an incremental CRC-32; returns the running state (not yet finalized). */
export function crc32Init(): number {
  return 0xffffffff
}

/** Fold another chunk into a running CRC-32 state (from {@link crc32Init}). */
export function crc32Update(state: number, data: Uint8Array): number {
  let c = state
  for (let i = 0; i < data.length; i++) {
    c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8)
  }
  return c
}

/** Finalize a running CRC-32 state into the unsigned 32-bit checksum. */
export function crc32Final(state: number): number {
  return (state ^ 0xffffffff) >>> 0
}

/** Standard ISO-HDLC CRC-32 (polynomial 0xEDB88320), unsigned 32-bit. */
export function crc32(data: Uint8Array): number {
  return crc32Final(crc32Update(crc32Init(), data))
}
