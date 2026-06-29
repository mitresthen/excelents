import { expect, test } from 'vitest'
import { version } from './index'

test('exposes a version string', () => {
  expect(typeof version).toBe('string')
})
