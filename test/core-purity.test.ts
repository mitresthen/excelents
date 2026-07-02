import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { expect, test } from 'vitest'

const SRC = new URL('../src/', import.meta.url).pathname

function tsFiles(dir: string): string[] {
  const out: string[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...tsFiles(full))
    else if (entry.name.endsWith('.ts') && !entry.name.endsWith('.test.ts')) out.push(full)
  }
  return out
}

test('core src (except the node.ts adapter and migrate.ts CLI) imports no node: builtin', () => {
  const offenders: string[] = []
  for (const file of tsFiles(SRC)) {
    if (file.endsWith('/node.ts') || file.endsWith('/migrate.ts')) continue
    const src = readFileSync(file, 'utf8')
    if (
      /from\s+['"]node:/.test(src) ||
      /import\s+['"]node:/.test(src) ||
      /require\(['"]node:/.test(src)
    ) {
      offenders.push(file)
    }
  }
  expect(offenders).toEqual([])
})
