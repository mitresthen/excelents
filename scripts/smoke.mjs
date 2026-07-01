/**
 * Import smoke test: packs the package, installs it into a temporary consumer
 * directory, then verifies the ESM entry point resolves and exposes `version`
 * as a string. Cleans up the tarball and temp dir on exit. The package is
 * ESM-only, so there is no CJS entry point to check.
 *
 * Usage: node scripts/smoke.mjs   (or: pnpm smoke)
 */
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(import.meta.url), '..', '..')

// Step 1: clean up any pre-existing tarballs
readdirSync(root)
  .filter((f) => f.startsWith('mitresthen-excelents-') && f.endsWith('.tgz'))
  .forEach((f) => rmSync(join(root, f), { force: true }))

// Step 2: pack
console.log('[smoke] packing...')
execFileSync('pnpm', ['pack'], { cwd: root, encoding: 'utf-8' })

// Step 3: find the tarball by glob pattern (robust, not dependent on pnpm stdout)
const tgzName = readdirSync(root).find(
  (f) => f.startsWith('mitresthen-excelents-') && f.endsWith('.tgz'),
)
if (!tgzName) throw new Error('pnpm pack did not produce a tarball')
const tgzPath = join(root, tgzName)
console.log(`[smoke] tarball: ${tgzName}`)

// Step 2: create temp consumer dir and install
const tmpDir = mkdtempSync(join(tmpdir(), 'excelents-smoke-'))

try {
  execFileSync('npm', ['init', '-y'], { cwd: tmpDir, stdio: 'pipe' })
  execFileSync('npm', ['install', tgzPath], { cwd: tmpDir, stdio: 'pipe' })

  // Step 3: ESM check (top-level await via --input-type=module)
  console.log('[smoke] checking ESM import...')
  execFileSync(
    'node',
    [
      '--input-type=module',
      '-e',
      "const m = await import('@mitresthen/excelents'); if (typeof m.version !== 'string') throw new Error('esm export missing'); console.log('esm ok')",
    ],
    { cwd: tmpDir, stdio: 'inherit' },
  )
} finally {
  // Step 4: clean up
  rmSync(tmpDir, { recursive: true, force: true })
  rmSync(tgzPath, { force: true })
  console.log('[smoke] cleaned up.')
}
