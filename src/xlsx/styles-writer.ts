import type { Alignment, Borders, CellStyle, Fill, Font } from '../model/style'
import { builtinFormatId } from '../utils/number-format'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Internal fill entry — superset of the public Fill type to allow the mandatory 'gray125' pattern. */
interface FillEntry {
  type: 'pattern'
  pattern: 'none' | 'gray125' | 'solid'
  fgColor?: string
}

/** Internal XF (cell format) entry referencing interned table indices. */
interface XfEntry {
  numFmtId: number
  fontId: number
  fillId: number
  borderId: number
  /** Inline alignment, when the style sets one. Alignment is not interned into a separate table. */
  alignment?: Alignment
}

/** OOXML `vertical` uses `center`; our model exposes the more conventional `middle`. */
const VERTICAL_TO_OOXML: Record<NonNullable<Alignment['vertical']>, string> = {
  top: 'top',
  middle: 'center',
  bottom: 'bottom',
}

/** True when the alignment carries at least one meaningful facet (empty `{}` is treated as none). */
function hasAlignment(a: Alignment | undefined): a is Alignment {
  return (
    a !== undefined &&
    (a.horizontal !== undefined || a.vertical !== undefined || a.wrapText === true)
  )
}

/** Stable, key-order-independent dedup key for an alignment facet. */
function alignmentKey(a: Alignment | undefined): string {
  if (a === undefined) return ''
  return `h=${a.horizontal ?? ''}|v=${a.vertical ?? ''}|w=${a.wrapText === true ? '1' : ''}`
}

const DEFAULT_FONT: Font = { name: 'Calibri', size: 11 }
/** Mandatory Excel fill at index 0. */
const FILL_NONE: FillEntry = { type: 'pattern', pattern: 'none' }
/** Mandatory Excel fill at index 1 (Excel rejects files that lack it). */
const FILL_GRAY125: FillEntry = { type: 'pattern', pattern: 'gray125' }
const DEFAULT_BORDER: Borders = {}

/**
 * Interns CellStyle facets into OOXML tables (numFmts, fonts, fills, borders, cellXfs)
 * and maps each unique CellStyle to a cellXfs index.
 *
 * Index 0 of every table is the default / empty entry as required by OOXML.
 * The empty style `{}` always maps to cellXfs index 0.
 */
export class StyleRegistry {
  private readonly numFmtMap = new Map<string, number>()
  private readonly customNumFmts: Array<{ id: number; code: string }> = []
  private nextCustomId = 164

  private readonly fontKeyMap = new Map<string, number>()
  private readonly fontList: Font[] = [DEFAULT_FONT]

  private readonly fillKeyMap = new Map<string, number>()
  private readonly fillList: FillEntry[] = [FILL_NONE, FILL_GRAY125]

  private readonly borderKeyMap = new Map<string, number>()
  private readonly borderList: Borders[] = [DEFAULT_BORDER]

  private readonly xfKeyMap = new Map<string, number>()
  private readonly xfList: XfEntry[] = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]

  constructor() {
    this.fontKeyMap.set(JSON.stringify(DEFAULT_FONT), 0)
    this.fillKeyMap.set(JSON.stringify(FILL_NONE), 0)
    this.fillKeyMap.set(JSON.stringify(FILL_GRAY125), 1)
    this.borderKeyMap.set(JSON.stringify(DEFAULT_BORDER), 0)
    // Must match the key shape produced by xfIndexFor (trailing empty alignment segment).
    this.xfKeyMap.set(`0,0,0,0,${alignmentKey(undefined)}`, 0)
  }

  private internNumFmt(code: string | undefined): number {
    if (code === undefined) return 0
    const builtinId = builtinFormatId(code)
    if (builtinId !== undefined) return builtinId
    const existing = this.numFmtMap.get(code)
    if (existing !== undefined) return existing
    const id = this.nextCustomId++
    this.numFmtMap.set(code, id)
    this.customNumFmts.push({ id, code })
    return id
  }

  private internFont(font: Font | undefined): number {
    if (font === undefined) return 0
    const key = JSON.stringify(font)
    const existing = this.fontKeyMap.get(key)
    if (existing !== undefined) return existing
    const idx = this.fontList.length
    this.fontList.push(font)
    this.fontKeyMap.set(key, idx)
    return idx
  }

  private internFill(fill: Fill | undefined): number {
    if (fill === undefined) return 0
    const entry: FillEntry =
      fill.pattern === 'solid'
        ? { type: 'pattern', pattern: 'solid', fgColor: fill.fgColor }
        : { type: 'pattern', pattern: 'none' }
    const key = JSON.stringify(entry)
    const existing = this.fillKeyMap.get(key)
    if (existing !== undefined) return existing
    const idx = this.fillList.length
    this.fillList.push(entry)
    this.fillKeyMap.set(key, idx)
    return idx
  }

  private internBorder(border: Borders | undefined): number {
    if (border === undefined) return 0
    const key = JSON.stringify(border)
    const existing = this.borderKeyMap.get(key)
    if (existing !== undefined) return existing
    const idx = this.borderList.length
    this.borderList.push(border)
    this.borderKeyMap.set(key, idx)
    return idx
  }

  /**
   * Returns the cellXfs index for the given style, interning each facet
   * into its table on first encounter. The empty style `{}` returns 0.
   */
  xfIndexFor(style: CellStyle): number {
    const numFmtId = this.internNumFmt(style.numberFormat)
    const fontId = this.internFont(style.font)
    const fillId = this.internFill(style.fill)
    const borderId = this.internBorder(style.border)
    const alignment = hasAlignment(style.alignment) ? style.alignment : undefined
    const key = `${numFmtId},${fontId},${fillId},${borderId},${alignmentKey(alignment)}`
    const existing = this.xfKeyMap.get(key)
    if (existing !== undefined) return existing
    const idx = this.xfList.length
    this.xfList.push({ numFmtId, fontId, fillId, borderId, alignment })
    this.xfKeyMap.set(key, idx)
    return idx
  }

  get customFormats(): ReadonlyArray<{ id: number; code: string }> {
    return this.customNumFmts
  }
  get fonts(): ReadonlyArray<Font> {
    return this.fontList
  }
  get fills(): ReadonlyArray<FillEntry> {
    return this.fillList
  }
  get borders(): ReadonlyArray<Borders> {
    return this.borderList
  }
  get xfs(): ReadonlyArray<XfEntry> {
    return this.xfList
  }
}

function writeFontXml(w: XmlWriter, font: Font): void {
  w.open('font')
  if (font.bold === true) w.leaf('b')
  if (font.italic === true) w.leaf('i')
  if (font.underline === true) w.leaf('u')
  if (font.size !== undefined) w.leaf('sz', { val: font.size })
  if (font.color !== undefined) w.leaf('color', { rgb: font.color })
  if (font.name !== undefined) w.leaf('name', { val: font.name })
  w.close('font')
}

function writeFillXml(w: XmlWriter, fill: FillEntry): void {
  w.open('fill')
  if (fill.pattern === 'solid') {
    w.open('patternFill', { patternType: 'solid' })
    if (fill.fgColor !== undefined) w.leaf('fgColor', { rgb: fill.fgColor })
    w.close('patternFill')
  } else {
    // 'none' or 'gray125' — self-closing
    w.leaf('patternFill', { patternType: fill.pattern })
  }
  w.close('fill')
}

function writeBorderEdge(w: XmlWriter, tag: string, edge: Borders[keyof Borders]): void {
  if (edge === undefined) {
    w.leaf(tag)
    return
  }
  w.open(tag, { style: edge.style })
  if (edge.color !== undefined) w.leaf('color', { rgb: edge.color })
  w.close(tag)
}

function writeBorderXml(w: XmlWriter, border: Borders): void {
  w.open('border')
  writeBorderEdge(w, 'left', border.left)
  writeBorderEdge(w, 'right', border.right)
  writeBorderEdge(w, 'top', border.top)
  writeBorderEdge(w, 'bottom', border.bottom)
  w.leaf('diagonal')
  w.close('border')
}

function writeAlignmentXml(w: XmlWriter, a: Alignment): void {
  w.leaf('alignment', {
    horizontal: a.horizontal,
    vertical: a.vertical !== undefined ? VERTICAL_TO_OOXML[a.vertical] : undefined,
    wrapText: a.wrapText === true ? 1 : undefined,
  })
}

/**
 * Serialize the style registry to `xl/styles.xml`.
 *
 * Children of `<styleSheet>` are emitted in OOXML schema order:
 * numFmts → fonts → fills → borders → cellXfs.
 */
export function writeStylesXml(registry: StyleRegistry): string {
  const w = new XmlWriter().declaration().open('styleSheet', { xmlns: MAIN_NS })

  // numFmts — only custom formats (id >= 164); builtins need no entry
  const customFmts = registry.customFormats
  if (customFmts.length > 0) {
    w.open('numFmts', { count: customFmts.length })
    for (const fmt of customFmts) {
      w.leaf('numFmt', { numFmtId: fmt.id, formatCode: fmt.code })
    }
    w.close('numFmts')
  }

  // fonts (index 0 = default Calibri 11)
  const fonts = registry.fonts
  w.open('fonts', { count: fonts.length })
  for (const font of fonts) writeFontXml(w, font)
  w.close('fonts')

  // fills (index 0 = none, index 1 = gray125 — both mandatory)
  const fills = registry.fills
  w.open('fills', { count: fills.length })
  for (const fill of fills) writeFillXml(w, fill)
  w.close('fills')

  // borders (index 0 = default empty border)
  const borders = registry.borders
  w.open('borders', { count: borders.length })
  for (const border of borders) writeBorderXml(w, border)
  w.close('borders')

  // cellXfs (index 0 = default xf)
  const xfs = registry.xfs
  w.open('cellXfs', { count: xfs.length })
  for (const xf of xfs) {
    const attrs = {
      numFmtId: xf.numFmtId,
      fontId: xf.fontId,
      fillId: xf.fillId,
      borderId: xf.borderId,
      xfId: 0,
      applyNumberFormat: xf.numFmtId !== 0 ? 1 : undefined,
      applyFont: xf.fontId !== 0 ? 1 : undefined,
      applyFill: xf.fillId !== 0 ? 1 : undefined,
      applyBorder: xf.borderId !== 0 ? 1 : undefined,
      applyAlignment: xf.alignment !== undefined ? 1 : undefined,
    }
    if (xf.alignment === undefined) {
      w.leaf('xf', attrs)
    } else {
      w.open('xf', attrs)
      writeAlignmentXml(w, xf.alignment)
      w.close('xf')
    }
  }
  w.close('cellXfs')

  return w.close('styleSheet').toString()
}
