import { readdirSync } from 'node:fs'
import { join } from 'node:path'

const FIXTURES_DIR = new URL('../fixtures/', import.meta.url).pathname

/** Marker for harness functions whose real bodies land in SP-1 (OPC + XML). */
export class NotImplemented extends Error {
  constructor(what: string) {
    super(`${what} is implemented in SP-1 (requires OPC + XML substrate)`)
    this.name = 'NotImplemented'
  }
}

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

/** Unzips xlsx bytes into a map of part name → canonical XML. Real body: SP-1. */
export function unzipToParts(_bytes: Uint8Array): Promise<Record<string, string>> {
  throw new NotImplemented('unzipToParts')
}

/** Structural XML-part diff with normalization. Real body: SP-1. */
export function diffParts(
  _a: Record<string, string>,
  _b: Record<string, string>,
): { part: string; detail: string }[] {
  throw new NotImplemented('diffParts')
}

/** Loads a fixture as bytes. */
export async function parseFixture(path: string): Promise<Uint8Array> {
  const { readFile } = await import('node:fs/promises')
  return new Uint8Array(await readFile(path))
}
