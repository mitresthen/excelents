import { gzipSync } from 'node:zlib'
import { expect, test } from 'vitest'
import { checkBudgets, gzipSize } from './size.mjs'

test('gzipSize matches node zlib gzip length', () => {
  const buf = Buffer.from('hello world '.repeat(50))
  expect(gzipSize(buf)).toBe(gzipSync(buf).length)
})

test('checkBudgets flags entries over budget', () => {
  const result = checkBudgets({ 'index.js': 2048 }, { 'index.js': 1024 })
  expect(result.ok).toBe(false)
  expect(result.violations).toEqual([{ entry: 'index.js', size: 2048, budget: 1024 }])
})

test('checkBudgets passes when under budget', () => {
  const result = checkBudgets({ 'index.js': 512 }, { 'index.js': 1024 })
  expect(result.ok).toBe(true)
  expect(result.violations).toEqual([])
})
