/**
 * `excelents-migrate` — scan a project for exceljs usage and report how it maps
 * to excelents. Advisory only: it never modifies files.
 *
 *   npx @mitresthen/excelents [path] [--json]
 *
 * The scanner itself (scanSource) is pure and exported for programmatic use.
 */
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'
import { type MigrationCategory, type MigrationFinding, scanSource } from './migrate/scan'

export { scanSource, referencesExcelJs } from './migrate/scan'
export type { MigrationCategory, MigrationFinding } from './migrate/scan'

const GUIDE_URL =
  'https://github.com/mitresthen/excelents/blob/master/docs/MIGRATING-FROM-EXCELJS.md'
const SKIP_DIRS = new Set(['node_modules', 'dist', 'build', 'out', 'coverage', 'vendor'])
const SOURCE_EXT = /\.(?:js|jsx|ts|tsx|mjs|cjs|mts|cts)$/
const MAX_FILE_BYTES = 2 * 1024 * 1024 // skip bundles/minified artifacts

export interface MigrationReport {
  scannedFiles: number
  findings: MigrationFinding[]
  summary: Record<MigrationCategory, number>
}

function collectFiles(dir: string, out: string[]): void {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith('.') || SKIP_DIRS.has(entry.name)) continue
    const full = join(dir, entry.name)
    if (entry.isDirectory()) collectFiles(full, out)
    else if (entry.isFile() && SOURCE_EXT.test(entry.name)) out.push(full)
  }
}

/** Scan every source file under `root`. */
export function scanProject(root: string): MigrationReport {
  const files: string[] = []
  collectFiles(root, files)
  const findings: MigrationFinding[] = []
  for (const file of files) {
    if (statSync(file).size > MAX_FILE_BYTES) continue
    findings.push(...scanSource(readFileSync(file, 'utf8'), relative(root, file) || file))
  }
  const summary: Record<MigrationCategory, number> = { map: 0, restructure: 0, blocked: 0 }
  for (const f of findings) summary[f.category]++
  return { scannedFiles: files.length, findings, summary }
}

const SECTIONS: ReadonlyArray<{ category: MigrationCategory; title: string; blurb: string }> = [
  {
    category: 'map',
    title: 'AUTO-MAPPABLE',
    blurb: 'mechanical renames — safe to apply one-for-one',
  },
  {
    category: 'restructure',
    title: 'NEEDS RESTRUCTURING',
    blurb: 'same capability, different shape — read the linked guide section first',
  },
  {
    category: 'blocked',
    title: 'NO EQUIVALENT',
    blurb: 'dropped or not-yet-supported features — decide before migrating',
  },
]

function renderReport(report: MigrationReport): string {
  const lines: string[] = []
  const { scannedFiles, findings, summary } = report
  lines.push(`excelents migration scan — ${String(scannedFiles)} source files scanned`)
  if (findings.length === 0) {
    lines.push('No exceljs usage found.')
    return lines.join('\n')
  }
  const fileWidth = Math.max(...findings.map((f) => `${f.file}:${String(f.line)}`.length))
  for (const { category, title, blurb } of SECTIONS) {
    const group = findings.filter((f) => f.category === category)
    if (group.length === 0) continue
    lines.push('', `${title} (${String(group.length)}) — ${blurb}`)
    for (const f of group) {
      const loc = `${f.file}:${String(f.line)}`.padEnd(fileWidth)
      lines.push(`  ${loc}  ${f.found} → ${f.advice}  [§${f.guideSection}]`)
    }
  }
  lines.push(
    '',
    `Summary: ${String(summary.map)} auto-mappable, ${String(summary.restructure)} need restructuring, ${String(summary.blocked)} without an equivalent.`,
    `Guide (§anchors above): ${GUIDE_URL}`,
  )
  return lines.join('\n')
}

/** CLI body — exported so the bin shim stays a one-liner and this stays testable. */
export function runCli(argv: string[]): number {
  const json = argv.includes('--json')
  const target = argv.find((a) => !a.startsWith('-')) ?? '.'
  let report: MigrationReport
  try {
    report = scanProject(target)
  } catch (err) {
    console.error(
      `excelents-migrate: cannot scan '${target}': ${err instanceof Error ? err.message : String(err)}`,
    )
    return 2
  }
  console.log(json ? JSON.stringify(report, null, 2) : renderReport(report))
  return 0
}
