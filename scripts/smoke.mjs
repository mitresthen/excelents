/**
 * Dual-import smoke test: packs the package, installs it into a temporary
 * consumer directory, then verifies both ESM and CJS entry points resolve and
 * expose `version` as a string. Cleans up the tarball and temp dir on exit.
 *
 * Usage: node scripts/smoke.mjs   (or: pnpm smoke)
 */
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(import.meta.url), '..', '..')

// Step 1: pack
console.log('[smoke] packing...')
const packOutput = execFileSync('pnpm', ['pack'], { cwd: root, encoding: 'utf-8' }).trim()
const lastLine = packOutput.split('\n').at(-1).trim()
const tgzName = lastLine.endsWith('.tgz')
  ? lastLine
  : readdirSync(root).find((f) => f.startsWith('excelents-') && f.endsWith('.tgz'))
if (!tgzName) throw new Error('Could not find packed tarball')
const tgzPath = join(root, tgzName)
console.log(`[smoke] tarball: ${tgzName}`)

// Step 2: create temp consumer dir and install
const tmpDir = mkdtempSync(join(tmpdir(), 'excelents-smoke-'))

try {
  execFileSync('npm', ['init', '-y'], { cwd: tmpDir, stdio: 'pipe' })
  execFileSync('npm', ['install', tgzPath], { cwd: tmpDir, stdio: 'pipe' })

  // Step 3a: ESM check (top-level await via --input-type=module)
  console.log('[smoke] checking ESM import...')
  execFileSync(
    'node',
    [
      '--input-type=module',
      '-e',
      "const m = await import('excelents'); if (typeof m.version !== 'string') throw new Error('esm export missing'); console.log('esm ok')",
    ],
    { cwd: tmpDir, stdio: 'inherit' },
  )

  // Step 3b: CJS check
  console.log('[smoke] checking CJS require...')
  execFileSync(
    'node',
    [
      '-e',
      "const m = require('excelents'); if (typeof m.version !== 'string') throw new Error('cjs export missing'); console.log('cjs ok')",
    ],
    { cwd: tmpDir, stdio: 'inherit' },
  )
} finally {
  // Step 4: clean up
  rmSync(tmpDir, { recursive: true, force: true })
  rmSync(tgzPath, { force: true })
  console.log('[smoke] cleaned up.')
}
