import type { Alignment, BorderEdge, Borders, CellStyle, Fill, Font } from '../model/style'
import { builtinFormatCode } from '../utils/number-format'
import { tokenize } from '../xml/tokenizer'

export interface ParsedStyles {
  /** CellStyle for each cellXfs index. Index 0 is the default empty style. */
  readonly cellStyles: CellStyle[]
  /** numFmtId -> format code, for every custom numFmt declared in the part. */
  readonly numFmtById: Map<number, string>
}

const VERTICAL_FROM_OOXML: Record<string, NonNullable<Alignment['vertical']>> = {
  top: 'top',
  center: 'middle',
  bottom: 'bottom',
}

/** Narrow a raw border-style attribute to the model union (control-flow narrowing, no cast). */
function asBorderStyle(s: string | undefined): NonNullable<BorderEdge['style']> | undefined {
  switch (s) {
    case 'thin':
    case 'medium':
    case 'thick':
    case 'dashed':
    case 'dotted':
    case 'double':
      return s
    default:
      return undefined
  }
}

type Section = 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs' | null

/** Parse `xl/styles.xml` into CellStyles indexed by cellXfs position (inverse of writeStylesXml). */
export function readStyles(xml: string): ParsedStyles {
  const numFmtById = new Map<number, string>()
  const fonts: Font[] = []
  const fills: Array<Fill | undefined> = []
  const borders: Array<Borders | undefined> = []
  const cellStyles: CellStyle[] = []

  let section: Section = null
  let font: Font | undefined
  let fill: Fill | undefined
  let border: Borders | undefined
  let edge: keyof Borders | undefined
  let xf: CellStyle | undefined

  for (const tok of tokenize(xml)) {
    if (tok.type !== 'open' && tok.type !== 'close') continue
    const open = tok.type === 'open'

    switch (tok.name) {
      case 'numFmts':
      case 'fonts':
      case 'fills':
      case 'borders':
      case 'cellXfs':
        section = open ? tok.name : null
        continue
      case 'numFmt':
        if (open) {
          const id = Number(tok.attributes['numFmtId'])
          const code = tok.attributes['formatCode']
          if (!Number.isNaN(id) && code !== undefined) numFmtById.set(id, code)
        }
        continue
    }

    if (section === 'fonts') {
      if (tok.name === 'font') {
        if (open) font = {}
        else if (font !== undefined) {
          fonts.push(font)
          font = undefined
        }
      } else if (open && font !== undefined) {
        if (tok.name === 'b') font.bold = true
        else if (tok.name === 'i') font.italic = true
        else if (tok.name === 'u') font.underline = true
        else if (tok.name === 'sz') font.size = Number(tok.attributes['val'])
        else if (tok.name === 'color') font.color = tok.attributes['rgb']
        else if (tok.name === 'name') font.name = tok.attributes['val']
      }
    } else if (section === 'fills') {
      if (tok.name === 'fill') {
        if (open) fill = undefined
        else fills.push(fill)
      } else if (open && tok.name === 'patternFill') {
        fill =
          tok.attributes['patternType'] === 'solid'
            ? { type: 'pattern', pattern: 'solid' }
            : undefined
      } else if (open && tok.name === 'fgColor' && fill !== undefined) {
        fill = { ...fill, fgColor: tok.attributes['rgb'] }
      }
    } else if (section === 'borders') {
      if (tok.name === 'border') {
        if (open) border = {}
        else {
          borders.push(border)
          border = undefined
        }
      } else if (
        border !== undefined &&
        (tok.name === 'left' || tok.name === 'right' || tok.name === 'top' || tok.name === 'bottom')
      ) {
        if (open) {
          const style = asBorderStyle(tok.attributes['style'])
          if (style !== undefined) {
            edge = tok.name
            border[edge] = { style }
          }
          if (tok.selfClosing) edge = undefined
        } else {
          edge = undefined
        }
      } else if (open && tok.name === 'color' && edge !== undefined && border !== undefined) {
        const e = border[edge]
        if (e !== undefined) border[edge] = { ...e, color: tok.attributes['rgb'] }
      }
    } else if (section === 'cellXfs') {
      if (tok.name === 'xf') {
        if (open) {
          xf = {}
          const numFmtId = Number(tok.attributes['numFmtId'] ?? '0')
          const fontId = Number(tok.attributes['fontId'] ?? '0')
          const fillId = Number(tok.attributes['fillId'] ?? '0')
          const borderId = Number(tok.attributes['borderId'] ?? '0')
          const code = numFmtById.get(numFmtId) ?? builtinFormatCode(numFmtId)
          if (numFmtId !== 0 && code !== undefined && code !== 'General') xf.numberFormat = code
          if (fontId !== 0 && fonts[fontId] !== undefined) xf.font = fonts[fontId]
          if (fillId > 1 && fills[fillId] !== undefined) xf.fill = fills[fillId]
          const b = borders[borderId]
          if (borderId !== 0 && b !== undefined && Object.keys(b).length > 0) xf.border = b
          if (tok.selfClosing) {
            cellStyles.push(xf)
            xf = undefined
          }
        } else if (xf !== undefined) {
          cellStyles.push(xf)
          xf = undefined
        }
      } else if (open && tok.name === 'alignment' && xf !== undefined) {
        const a: Alignment = {}
        const h = tok.attributes['horizontal']
        const vv = tok.attributes['vertical']
        if (h === 'left' || h === 'center' || h === 'right') a.horizontal = h
        if (vv !== undefined && VERTICAL_FROM_OOXML[vv] !== undefined) {
          a.vertical = VERTICAL_FROM_OOXML[vv]
        }
        if (tok.attributes['wrapText'] === '1') a.wrapText = true
        if (Object.keys(a).length > 0) xf.alignment = a
      }
    }
  }
  return { cellStyles, numFmtById }
}
