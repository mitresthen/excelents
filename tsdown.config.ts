import { defineConfig } from 'tsdown'

export default defineConfig({
  entry: ['./src/index.ts', './src/node.ts', './src/csv.ts'],
  format: ['esm', 'cjs'],
  platform: 'neutral',
  target: 'node20',
  dts: true,
  tsconfig: './tsconfig.build.json',
  exports: {
    // Add per-condition `types` entries so every subpath resolves declarations
    // correctly under moduleResolution: node16 AND bundler, for both ESM and CJS.
    customExports(exports) {
      for (const [key, value] of Object.entries(exports)) {
        if (key === './package.json') continue
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          const newValue: Record<string, unknown> = {}
          for (const [cond, distPath] of Object.entries(value as Record<string, string>)) {
            if (cond === 'import' && typeof distPath === 'string') {
              newValue[cond] = {
                types: distPath.replace(/\.js$/, '.d.ts'),
                default: distPath,
              }
            } else if (cond === 'require' && typeof distPath === 'string') {
              newValue[cond] = {
                types: distPath.replace(/\.cjs$/, '.d.cts'),
                default: distPath,
              }
            } else {
              newValue[cond] = distPath
            }
          }
          exports[key] = newValue
        }
      }
      return exports
    },
  },
  clean: true,
  publint: 'ci-only',
  attw: { profile: 'node16', enabled: 'ci-only' },
})
