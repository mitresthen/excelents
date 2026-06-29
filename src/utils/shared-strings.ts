/** Interning table for the xlsx shared-strings part (`sharedStrings.xml`). */
export class SharedStrings {
  private readonly index = new Map<string, number>()
  private readonly list: string[] = []
  private totalRefs = 0

  /** Intern a string; returns its 0-based index (existing index for duplicates). */
  add(value: string): number {
    this.totalRefs += 1
    const existing = this.index.get(value)
    if (existing !== undefined) return existing
    const next = this.list.length
    this.index.set(value, next)
    this.list.push(value)
    return next
  }

  getIndex(value: string): number | undefined {
    return this.index.get(value)
  }

  getString(index: number): string | undefined {
    return this.list[index]
  }

  get count(): number {
    return this.totalRefs
  }

  get uniqueCount(): number {
    return this.list.length
  }

  get values(): readonly string[] {
    return this.list
  }
}
