import { createReadStream, createWriteStream } from 'node:fs'
import { readFile as nodeReadFile, writeFile as nodeWriteFile } from 'node:fs/promises'
import { Readable, Writable } from 'node:stream'
import type { FileSystemAdapter } from './io/fs'

/** A `FileSystemAdapter` backed by Node's filesystem, bridging Node streams to Web Streams. */
export const nodeFileSystem: FileSystemAdapter = {
  async readFile(path: string): Promise<Uint8Array> {
    return new Uint8Array(await nodeReadFile(path))
  },
  async writeFile(path: string, data: Uint8Array): Promise<void> {
    await nodeWriteFile(path, data)
  },
  createReadable(path: string): ReadableStream<Uint8Array> {
    // Node's ReadableStream from toWeb is nominally distinct from the DOM lib type
    // oxlint-disable-next-line no-unsafe-type-assertion
    return Readable.toWeb(createReadStream(path)) as ReadableStream<Uint8Array>
  },
  createWritable(path: string): WritableStream<Uint8Array> {
    // Node's WritableStream from toWeb is nominally distinct from the DOM lib type
    // oxlint-disable-next-line no-unsafe-type-assertion
    return Writable.toWeb(createWriteStream(path)) as WritableStream<Uint8Array>
  },
}
