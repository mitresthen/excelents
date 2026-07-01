import { expect, test } from 'vitest'
import { syncVersion } from './sync-version.mjs'

const SAMPLE = [
  '/** The package version, kept in sync with package.json (asserted by src/index.test.ts). */',
  "export const version: string = '0.0.0'",
  "export { createWorkbook, Workbook } from './model/workbook'",
].join('\n')

test('syncVersion rewrites the version literal', () => {
  const out = syncVersion(SAMPLE, '1.2.3')
  expect(out).toContain("export const version: string = '1.2.3'")
  expect(out).not.toContain("'0.0.0'")
})

test('syncVersion leaves the rest of the file untouched', () => {
  const out = syncVersion(SAMPLE, '1.2.3')
  expect(out).toContain("export { createWorkbook, Workbook } from './model/workbook'")
  expect(out).toContain('/** The package version, kept in sync with package.json')
  expect(out.split('\n').length).toBe(SAMPLE.split('\n').length)
})

test('syncVersion handles prerelease/build versions', () => {
  const out = syncVersion(SAMPLE, '1.2.3-rc.1')
  expect(out).toContain("export const version: string = '1.2.3-rc.1'")
})

test('syncVersion throws when the version line is absent', () => {
  expect(() => syncVersion('export const foo = 1\n', '1.2.3')).toThrow(/version/)
})
