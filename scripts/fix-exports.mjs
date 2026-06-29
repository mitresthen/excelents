/**
 * Post-build fixup: set the top-level `types` field to the ESM declaration.
 *
 * tsdown (0.22.3) writes `cjsTypes || esmTypes` for `types` when `legacy` is
 * true, which results in `./dist/index.d.cts`. For a `"type": "module"`
 * package the top-level `types` entry must point to the ESM declaration
 * (`.d.ts`), not the CJS one — the CJS declaration is correctly placed under
 * `exports["."].require.types` by the `customExports` hook in tsdown.config.ts.
 *
 * This script is invoked as part of the `build` npm script, AFTER `tsdown`,
 * so tsdown's overwrite is fixed before the package is used.
 */
import { readFileSync, writeFileSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'url'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const pkgPath = join(root, 'package.json')

const raw = readFileSync(pkgPath, 'utf-8')
const pkg = JSON.parse(raw)

pkg.types = './dist/index.d.ts'

const indent = raw.match(/^(\s+)/m)?.[1]?.length ?? 2
const eol = raw.includes('\r\n') ? '\r\n' : '\n'
const trailing = raw.endsWith('\n') ? eol : ''

writeFileSync(pkgPath, JSON.stringify(pkg, null, indent) + trailing, 'utf-8')
console.log('[fix-exports] set package.json#types → ./dist/index.d.ts')
