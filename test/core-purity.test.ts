/**
 * Enforces the "web-standard-core" invariant:
 * `src/index.ts` and `src/csv.ts` must not import any `node:` builtins.
 *
 * `src/node.ts` is explicitly allowed — it is the Node-only adapter.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, test } from 'vitest'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')

function readSrc(file: string): string {
  return readFileSync(join(root, 'src', file), 'utf-8')
}

const NODE_IMPORT_RE = /(['"])node:/

describe('core-purity: no node: imports in universal core', () => {
  test('src/index.ts has no node: imports', () => {
    const src = readSrc('index.ts')
    expect(NODE_IMPORT_RE.test(src)).toBe(false)
  })

  test('src/csv.ts has no node: imports', () => {
    const src = readSrc('csv.ts')
    expect(NODE_IMPORT_RE.test(src)).toBe(false)
  })

  test('src/node.ts IS allowed to use node: imports (sanity check)', () => {
    const src = readSrc('node.ts')
    expect(NODE_IMPORT_RE.test(src)).toBe(true)
  })
})
