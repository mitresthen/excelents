import { readdirSync, readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { gzipSync } from 'node:zlib'

/**
 * @param {Buffer} buf
 * @returns {number}
 */
export function gzipSize(buf) {
  return gzipSync(buf).length
}

/**
 * @param {Record<string, number>} sizes
 * @param {Record<string, number>} budgets
 * @returns {{ ok: boolean, violations: Array<{ entry: string, size: number | null, budget: number, reason?: string }> }}
 */
export function checkBudgets(sizes, budgets) {
  const violations = []
  for (const [entry, budget] of Object.entries(budgets)) {
    const size = sizes[entry]
    if (size === undefined) {
      violations.push({ entry, size: null, budget, reason: 'missing' })
    } else if (size > budget) {
      violations.push({ entry, size, budget })
    }
  }
  return { ok: violations.length === 0, violations }
}

/**
 * @param {string} distDir
 * @returns {Record<string, number>}
 */
function measureDist(distDir) {
  const sizes = {}
  if (!existsSync(distDir)) return sizes
  for (const file of readdirSync(distDir)) {
    if (file.endsWith('.js')) sizes[file] = gzipSize(readFileSync(join(distDir, file)))
  }
  return sizes
}

// CLI entry: node scripts/size.mjs
if (import.meta.url === `file://${process.argv[1]}`) {
  const budgets = JSON.parse(readFileSync(new URL('../size-budget.json', import.meta.url)))
  const sizes = measureDist(new URL('../dist', import.meta.url).pathname)
  const { ok, violations } = checkBudgets(sizes, budgets)
  for (const [entry, size] of Object.entries(sizes)) {
    console.log(`${entry}: ${size} B gzip`)
  }
  if (!ok) {
    console.error('Bundle size budget exceeded:', JSON.stringify(violations, null, 2))
    process.exit(1)
  }
  console.log('All entries within budget.')
}
