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

/** Tokenize, sort each element's attributes, and re-emit compact — a stable canonical form. */
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
    } else {
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
