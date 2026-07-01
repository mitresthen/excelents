import { readFileSync, writeFileSync } from 'node:fs'

const VERSION_LINE = /^(export const version: string = ')[^']*(')/m

/**
 * Rewrite the `export const version: string = '...'` literal in the given
 * src/index.ts source to `version`, returning the updated source.
 * @param {string} indexSource
 * @param {string} version
 * @returns {string}
 */
export function syncVersion(indexSource, version) {
  if (!VERSION_LINE.test(indexSource)) {
    throw new Error('sync-version: could not find `export const version` line in src/index.ts')
  }
  return indexSource.replace(VERSION_LINE, (_m, prefix, suffix) => `${prefix}${version}${suffix}`)
}

// CLI entry: node scripts/sync-version.mjs  (invoked by the npm `version` lifecycle)
if (import.meta.url === `file://${process.argv[1]}`) {
  const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf-8'))
  const indexUrl = new URL('../src/index.ts', import.meta.url)
  const src = readFileSync(indexUrl, 'utf-8')
  const updated = syncVersion(src, pkg.version)
  if (updated !== src) {
    writeFileSync(indexUrl, updated)
    console.log(`sync-version: src/index.ts -> ${pkg.version}`)
  } else {
    console.log(`sync-version: already at ${pkg.version}`)
  }
}
