import { platform } from 'node:os'

/** Node adapter placeholder; implemented in SP-1 (`io/` FS adapter). */
export function nodeAdapterPlaceholder(): string {
  return platform()
}
