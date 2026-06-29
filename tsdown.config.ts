import { defineConfig } from 'tsdown'

export default defineConfig({
  entry: ['./src/index.ts', './src/node.ts', './src/csv.ts'],
  format: ['esm', 'cjs'],
  platform: 'neutral',
  target: 'node20',
  dts: true,
  tsconfig: './tsconfig.build.json',
  exports: true,
  clean: true,
  publint: 'ci-only',
  attw: 'ci-only',
})
