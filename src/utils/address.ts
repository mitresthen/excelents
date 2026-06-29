const ADDRESS_RE = /^\$?([A-Za-z]+)\$?(\d+)$/

/** Convert a column letter (`'A'`, `'AA'`) to a 1-based column number. */
export function colToNumber(col: string): number {
  let n = 0
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64)
  }
  return n
}

/** Convert a 1-based column number to its column letter. */
export function numberToCol(num: number): string {
  let n = num
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** Parse a cell address (`'A1'`, `'$A$1'`) into 1-based row/col. */
export function decodeAddress(address: string): { row: number; col: number } {
  const m = ADDRESS_RE.exec(address)
  if (m === null) {
    throw new Error(`Invalid cell address: ${JSON.stringify(address)}`)
  }
  return { row: Number(m[2]!), col: colToNumber(m[1]!.toUpperCase()) }
}

/** Build a cell address from 1-based row/col. */
export function encodeAddress(row: number, col: number): string {
  return `${numberToCol(col)}${row}`
}
