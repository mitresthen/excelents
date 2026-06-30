// Type-only import (erased at build): the SST stores rich runs in the model's shape.
import type { RichTextRun } from '../model/cell'

/** One entry of the shared-string table: either plain text or a sequence of formatted runs. */
export type SharedStringEntry =
  | { readonly type: 'plain'; readonly text: string }
  | { readonly type: 'rich'; readonly runs: readonly RichTextRun[] }

/** Interning table for the xlsx shared-strings part (`sharedStrings.xml`). */
export class SharedStrings {
  private readonly index = new Map<string, number>()
  private readonly list: SharedStringEntry[] = []
  private totalRefs = 0

  /** Intern a plain string; returns its 0-based index (existing index for duplicates). */
  add(value: string): number {
    return this.intern(`p\0${value}`, { type: 'plain', text: value })
  }

  /** Intern a rich (multi-run) string; returns its 0-based index (existing index for duplicates). */
  addRich(runs: readonly RichTextRun[]): number {
    return this.intern(`r\0${JSON.stringify(runs)}`, { type: 'rich', runs })
  }

  private intern(key: string, entry: SharedStringEntry): number {
    this.totalRefs += 1
    const existing = this.index.get(key)
    if (existing !== undefined) return existing
    const next = this.list.length
    this.index.set(key, next)
    this.list.push(entry)
    return next
  }

  getIndex(value: string): number | undefined {
    return this.index.get(`p\0${value}`)
  }

  getString(index: number): string | undefined {
    const entry = this.list[index]
    if (entry === undefined) return undefined
    return entry.type === 'plain' ? entry.text : entry.runs.map((r) => r.text).join('')
  }

  get count(): number {
    return this.totalRefs
  }

  get uniqueCount(): number {
    return this.list.length
  }

  /** All interned entries in index order (for serialization). */
  get entries(): readonly SharedStringEntry[] {
    return this.list
  }

  /** Plain-text projection of every entry (rich entries flattened to concatenated run text). */
  get values(): readonly string[] {
    return this.list.map((e) => (e.type === 'plain' ? e.text : e.runs.map((r) => r.text).join('')))
  }
}
