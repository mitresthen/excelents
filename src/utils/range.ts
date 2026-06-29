import { decodeAddress, encodeAddress } from './address'

export type RangeBox = { top: number; left: number; bottom: number; right: number }

/** Parse an `A1:B2` (or single-cell `A1`) range into a normalized box. */
export function decodeRange(range: string): RangeBox {
  const parts = range.split(':')
  const a = decodeAddress(parts[0]!)
  const b = parts.length > 1 ? decodeAddress(parts[1]!) : a
  return {
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  }
}

/** Build an `A1:B2` range (or single-cell `A1` when the box is 1x1). */
export function encodeRange(box: RangeBox): string {
  const tl = encodeAddress(box.top, box.left)
  if (box.top === box.bottom && box.left === box.right) {
    return tl
  }
  return `${tl}:${encodeAddress(box.bottom, box.right)}`
}
