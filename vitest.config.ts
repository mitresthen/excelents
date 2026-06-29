import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    projects: [
      {
        test: {
          name: 'node',
          environment: 'node',
          include: ['src/**/*.test.ts', 'test/**/*.test.ts'],
          exclude: ['test/browser/**'],
        },
      },
      // Browser project (enabled in SP-1 when substrate browser tests exist):
      // {
      //   test: {
      //     name: 'browser',
      //     include: ['test/browser/**/*.test.ts'],
      //     browser: { enabled: true, provider: 'playwright', headless: true,
      //       instances: [{ browser: 'chromium' }] },
      //   },
      // },
    ],
  },
})
