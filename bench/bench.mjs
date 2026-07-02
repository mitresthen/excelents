/**
 * excelents vs exceljs benchmark. Each case runs in a fresh Node child process
 * (see run-case.mjs) so heap state and module caches are never shared; peak RSS
 * is sampled inside the child. Run `pnpm build` first — excelents is benchmarked
 * from dist/, exactly as consumers get it.
 *
 * Usage: node bench/bench.mjs [rows]   (default 100000)
 */
import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('..', import.meta.url))
const ROWS = Number(process.argv[2] ?? 100_000)

if (!existsSync(join(root, 'dist/index.js'))) {
  console.error('dist/ not found — run `pnpm build` first')
  process.exit(1)
}

const runCase = (name, inputFile) => {
  const args = [join(root, 'bench/run-case.mjs'), name, String(ROWS)]
  if (inputFile) args.push(inputFile)
  const out = execFileSync(process.execPath, args, { cwd: root, encoding: 'utf-8' })
  return JSON.parse(out)
}

// Shared input for the read cases, written by excelents in a child process.
console.error(`[bench] preparing ${ROWS}-row input file...`)
const workDir = mkdtempSync(join(tmpdir(), 'excelents-bench-'))
const inputFile = join(workDir, 'input.xlsx')
const { createWorkbook, writeXlsx } = await import(join(root, 'dist/index.js'))
{
  const wb = createWorkbook()
  const ws = wb.addSheet('Data')
  for (let r = 0; r < ROWS; r++) {
    const row = new Array(10)
    for (let c = 0; c < 10; c++) row[c] = c % 2 === 0 ? `cell ${r}:${c}` : r * 10 + c
    ws.addRow(row)
  }
  writeFileSync(inputFile, await writeXlsx(wb))
}

const MATCHUPS = [
  ['buffered write', 'excelents-write', 'exceljs-write', false],
  ['buffered read', 'excelents-read', 'exceljs-read', true],
  ['streaming write', 'excelents-stream-write', 'exceljs-stream-write', false],
  ['streaming read', 'excelents-stream-read', 'exceljs-stream-read', true],
]

const rows = []
for (const [label, ours, theirs, needsInput] of MATCHUPS) {
  console.error(`[bench] ${label}...`)
  const a = runCase(ours, needsInput ? inputFile : undefined)
  const b = runCase(theirs, needsInput ? inputFile : undefined)
  rows.push({ label, excelents: a, exceljs: b })
}
rmSync(workDir, { recursive: true, force: true })

console.log(`\n${ROWS.toLocaleString('en-US')} rows x 10 cols (5 string + 5 number per row)\n`)
console.log('| Scenario | excelents | exceljs | excelents peak RSS | exceljs peak RSS |')
console.log('| --- | --- | --- | --- | --- |')
for (const { label, excelents, exceljs } of rows) {
  console.log(
    `| ${label} | ${(excelents.ms / 1000).toFixed(1)} s | ${(exceljs.ms / 1000).toFixed(1)} s ` +
      `| ${excelents.peakRssMb} MB | ${exceljs.peakRssMb} MB |`,
  )
}
