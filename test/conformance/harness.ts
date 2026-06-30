import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { readZip } from '../../src/opc/zip-reader'
import { tokenize } from '../../src/xml/tokenizer'
import { XmlWriter } from '../../src/xml/writer'

const FIXTURES_DIR = new URL('../fixtures/', import.meta.url).pathname

/** Absolute paths of all xlsx fixtures. */
export function listFixtures(ext = '.xlsx'): string[] {
  return readdirSync(FIXTURES_DIR)
    .filter((f) => f.endsWith(ext))
    .map((f) => join(FIXTURES_DIR, f))
}

/** Builds the same workbook with the old exceljs oracle and the new library. */
export async function withOracle<T>(
  build: (lib: 'oracle' | 'excelents') => Promise<T>,
): Promise<{ oracle: T; excelents: T }> {
  return { oracle: await build('oracle'), excelents: await build('excelents') }
}

/** Loads a fixture as bytes. */
export async function parseFixture(path: string): Promise<Uint8Array> {
  const { readFile } = await import('node:fs/promises')
  return new Uint8Array(await readFile(path))
}

function isXmlPart(name: string): boolean {
  return name.endsWith('.xml') || name.endsWith('.rels')
}

/**
 * Tokenize, sort each element's attributes, and re-emit compact — a stable canonical form.
 *
 * Whitespace-only text runs (insignificant inter-element formatting, e.g. the indentation
 * exceljs pretty-prints) are dropped so two producers with different formatting compare equal.
 * Known limitation: a string whose content is itself only whitespace (`<t xml:space="preserve">
 *   </t>`) is also dropped — acceptable because both producers drop it identically.
 */
export function canonicalizeXml(xml: string): string {
  const w = new XmlWriter()
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open') {
      const sorted: Record<string, string> = {}
      for (const key of Object.keys(tok.attributes).sort()) sorted[key] = tok.attributes[key]!
      if (tok.selfClosing) w.leaf(tok.name, sorted)
      else w.open(tok.name, sorted)
    } else if (tok.type === 'close') {
      w.close(tok.name)
    } else if (tok.value.trim() !== '') {
      w.text(tok.value)
    }
  }
  return w.toString()
}

/** Unzip xlsx bytes; canonicalize XML parts, mark binary parts by length. */
export async function unzipToParts(bytes: Uint8Array): Promise<Record<string, string>> {
  const entries = await readZip(bytes)
  const out: Record<string, string> = {}
  const decoder = new TextDecoder()
  for (const [name, data] of entries) {
    out[name] = isXmlPart(name) ? canonicalizeXml(decoder.decode(data)) : `bytes:${data.length}`
  }
  return out
}

/** Structural diff over canonical part maps. */
export function diffParts(
  a: Record<string, string>,
  b: Record<string, string>,
): { part: string; detail: string }[] {
  const diffs: { part: string; detail: string }[] = []
  const names = new Set([...Object.keys(a), ...Object.keys(b)])
  for (const name of [...names].sort()) {
    if (!(name in a)) diffs.push({ part: name, detail: 'missing on left' })
    else if (!(name in b)) diffs.push({ part: name, detail: 'missing on right' })
    else if (a[name] !== b[name]) diffs.push({ part: name, detail: 'content differs' })
  }
  return diffs
}

/**
 * Build the same workbook with both producers, canonicalize their parts, and diff.
 * `filter` narrows to the parts we own (e.g. `xl/worksheets/*`); incidental parts
 * exceljs adds (theme, docProps, …) are excluded by default-only when filtered.
 */
export async function compareWriteParts(
  build: (lib: 'oracle' | 'excelents') => Promise<Uint8Array>,
  filter: (part: string) => boolean = () => true,
): Promise<{ part: string; detail: string }[]> {
  const { oracle, excelents } = await withOracle(build)
  const a = await unzipToParts(oracle)
  const b = await unzipToParts(excelents)
  const pick = (m: Record<string, string>): Record<string, string> =>
    Object.fromEntries(Object.entries(m).filter(([k]) => filter(k)))
  return diffParts(pick(a), pick(b))
}

/** Resolve a `sharedStrings.xml` part to its index-ordered plain-text values. */
function parseSharedStrings(bytes: Uint8Array | undefined): string[] {
  if (bytes === undefined) return []
  const xml = new TextDecoder().decode(bytes)
  const values: string[] = []
  let current = ''
  let inSi = false // <si> does not nest in OOXML, so a boolean suffices
  let inText = false
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' && tok.name === 'si') {
      current = ''
      inSi = true
    } else if (tok.type === 'close' && tok.name === 'si') {
      values.push(current)
      inSi = false
    } else if (tok.type === 'open' && tok.name === 't') {
      inText = true
    } else if (tok.type === 'close' && tok.name === 't') {
      inText = false
    } else if (tok.type === 'text' && inText && inSi) {
      current += tok.value
    }
  }
  return values
}

/** A semantically-resolved cell: its type tag and value (shared strings already dereferenced). */
export interface ResolvedCell {
  t: string
  v: string
  f?: string
}

/** The semantic content of a worksheet: resolved cells by ref, plus merge ranges. */
export interface SheetContent {
  cells: Record<string, ResolvedCell>
  merges: string[]
}

/**
 * Extract a worksheet's semantic content (resolved cell values + merges) from xlsx bytes,
 * independent of each producer's incidental XML structure or shared-string index ordering.
 *
 * Limitation: only `<v>`-bearing cells are read. Inline strings (`t="inlineStr"` with
 * `<is><t>…`) carry no `<v>` and are skipped — exceljs defaults to shared strings, so this
 * does not arise in the conformance suite, but a producer emitting inline strings would
 * appear to have missing cells here.
 */
export async function extractSheetContent(
  bytes: Uint8Array,
  sheetPath = 'xl/worksheets/sheet1.xml',
): Promise<SheetContent> {
  const entries = new Map(await readZip(bytes))
  const sst = parseSharedStrings(entries.get('xl/sharedStrings.xml'))
  const wsBytes = entries.get(sheetPath)
  if (wsBytes === undefined) throw new Error(`worksheet part not found: ${sheetPath}`)
  const xml = new TextDecoder().decode(wsBytes)

  const cells: Record<string, ResolvedCell> = {}
  const merges: string[] = []
  let ref: string | undefined
  let type = 'n'
  let v: string | undefined
  let f: string | undefined
  let inV = false
  let inF = false

  const finalize = (): void => {
    if (ref === undefined) return
    if (v === undefined && f === undefined) return // empty cell
    const resolved = type === 's' && v !== undefined ? (sst[Number(v)] ?? '') : (v ?? '')
    cells[ref] = f !== undefined ? { t: type, v: resolved, f } : { t: type, v: resolved }
  }

  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' || tok.type === 'close') {
      if (tok.name === 'c' && tok.type === 'open') {
        finalize()
        ref = tok.attributes.r
        type = tok.attributes.t ?? 'n'
        v = undefined
        f = undefined
        // Reset accumulator flags defensively in case a prior cell's </v>/</f> was malformed.
        inV = false
        inF = false
        if (tok.selfClosing) {
          ref = undefined // empty self-closed cell
        }
      } else if (tok.name === 'v') {
        inV = tok.type === 'open'
      } else if (tok.name === 'f') {
        inF = tok.type === 'open'
      } else if (tok.name === 'mergeCell' && tok.type === 'open') {
        if (tok.attributes.ref !== undefined) merges.push(tok.attributes.ref)
      }
    } else if (tok.type === 'text') {
      if (inV) v = (v ?? '') + tok.value
      else if (inF) f = (f ?? '') + tok.value
    }
  }
  finalize()
  return { cells, merges: merges.sort() }
}
