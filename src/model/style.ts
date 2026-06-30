/** Font facet. `color` is ARGB hex, e.g. `'FF0000FF'`. */
export interface Font {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
}

/** Fill facet (pattern fills only for now). */
export interface Fill {
  type: 'pattern'
  pattern: 'solid' | 'none'
  fgColor?: string
}

export interface BorderEdge {
  style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double'
  color?: string
}

export interface Borders {
  top?: BorderEdge
  bottom?: BorderEdge
  left?: BorderEdge
  right?: BorderEdge
}

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
}

/** A cell's visual style. All facets optional; `numberFormat` is a format code. */
export interface CellStyle {
  font?: Font
  fill?: Fill
  border?: Borders
  alignment?: Alignment
  numberFormat?: string
}

/** Overlay `patch`'s facets over `base` (facet-level, not deep). Inputs are not mutated. */
export function mergeStyles(base: CellStyle, patch: CellStyle): CellStyle {
  return { ...base, ...patch }
}
