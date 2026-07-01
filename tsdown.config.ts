import { defineConfig } from 'tsdown'

export default defineConfig({
  entry: ['./src/index.ts', './src/node.ts', './src/csv.ts', './src/stream.ts'],
  format: ['esm'],
  platform: 'neutral',
  target: 'node24',
  dts: true,
  deps: { neverBundle: [/^node:/] },
  tsconfig: './tsconfig.build.json',
  exports: {
    // For a single-format (ESM) build, tsdown emits a flat `"<subpath>":
    // "./dist/x.js"` per entry. Expand each into an explicit `{ types, default }`
    // so declarations resolve under node16/nodenext as well as bundler resolution.
    customExports(exports) {
      for (const [key, value] of Object.entries(exports)) {
        if (key === './package.json') continue
        if (typeof value === 'string' && value.endsWith('.js')) {
          exports[key] = {
            types: value.replace(/\.js$/, '.d.ts'),
            default: value,
          }
        }
      }
      return exports
    },
  },
  clean: true,
  publint: 'ci-only',
  attw: { profile: 'esm-only', enabled: 'ci-only' },
})
