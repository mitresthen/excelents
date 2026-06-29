import { bytesToReadable, concatUint8 } from './streams'

/** Abstracts file access so the universal core never imports `node:fs` directly. */
export interface FileSystemAdapter {
  readFile(path: string): Promise<Uint8Array>
  writeFile(path: string, data: Uint8Array): Promise<void>
  createReadable(path: string): ReadableStream<Uint8Array>
  createWritable(path: string): WritableStream<Uint8Array>
}

/** A pure in-memory `FileSystemAdapter` (no `node:`), for tests and edge/browser use. */
export function createMemoryFileSystem(
  initial: Record<string, Uint8Array> = {},
): FileSystemAdapter {
  const files = new Map<string, Uint8Array>(
    Object.entries(initial).map(([k, v]) => [k, new Uint8Array(v)]),
  )
  return {
    readFile(path: string): Promise<Uint8Array> {
      const data = files.get(path)
      if (data === undefined) return Promise.reject(new Error(`No such file: ${path}`))
      return Promise.resolve(new Uint8Array(data))
    },
    writeFile(path: string, data: Uint8Array): Promise<void> {
      files.set(path, new Uint8Array(data))
      return Promise.resolve()
    },
    createReadable(path: string): ReadableStream<Uint8Array> {
      const data = files.get(path)
      if (data === undefined) throw new Error(`No such file: ${path}`)
      return bytesToReadable(new Uint8Array(data))
    },
    createWritable(path: string): WritableStream<Uint8Array> {
      const chunks: Uint8Array[] = []
      return new WritableStream<Uint8Array>({
        write(chunk: Uint8Array): void {
          chunks.push(chunk)
        },
        close(): void {
          files.set(path, concatUint8(chunks))
        },
      })
    },
  }
}
