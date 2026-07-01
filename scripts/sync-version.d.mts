/**
 * Rewrite the `export const version: string = '...'` literal in the given
 * src/index.ts source to `version`, returning the updated source. Throws if the
 * version line is absent.
 */
export declare function syncVersion(indexSource: string, version: string): string
