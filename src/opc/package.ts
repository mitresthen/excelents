import { XmlWriter } from '../xml/writer'
import { tokenize } from '../xml/tokenizer'
import { readZip } from './zip-reader'
import { type ZipEntry, writeZip } from './zip-writer'

export interface OpcRelationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

const CONTENT_TYPES = '[Content_Types].xml'
const ROOT_RELS = '_rels/.rels'
const TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.')
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase()
}

/** An Open Packaging Conventions container over a standard ZIP. */
export class OpcPackage {
  private readonly parts = new Map<string, Uint8Array>()
  private readonly defaults = new Map<string, string>() // extension -> content type
  private readonly overrides = new Map<string, string>() // part name (no leading /) -> content type

  static empty(): OpcPackage {
    const pkg = new OpcPackage()
    pkg.defaults.set('rels', 'application/vnd.openxmlformats-package.relationships+xml')
    pkg.defaults.set('xml', 'application/xml')
    return pkg
  }

  static async read(bytes: Uint8Array): Promise<OpcPackage> {
    return OpcPackage.fromEntries(await readZip(bytes))
  }

  /**
   * Build a package from already-decompressed parts (name -> bytes). A streaming reader uses
   * this to assemble just the small structural parts (content-types, rels, workbook) it needs
   * to resolve the package, without decompressing the large worksheet bodies up front.
   */
  static fromEntries(entries: Map<string, Uint8Array>): OpcPackage {
    const pkg = new OpcPackage()
    for (const [name, data] of entries) {
      if (name === CONTENT_TYPES) continue
      pkg.parts.set(name, data)
    }
    const ct = entries.get(CONTENT_TYPES)
    if (ct !== undefined) pkg.parseContentTypes(new TextDecoder().decode(ct))
    return pkg
  }

  private parseContentTypes(xml: string): void {
    for (const tok of tokenize(xml)) {
      if (tok.type !== 'open') continue
      if (tok.name === 'Default') {
        const ext = tok.attributes['Extension']
        const type = tok.attributes['ContentType']
        if (ext !== undefined && type !== undefined) this.defaults.set(ext.toLowerCase(), type)
      } else if (tok.name === 'Override') {
        const part = tok.attributes['PartName']
        const type = tok.attributes['ContentType']
        if (part !== undefined && type !== undefined) {
          this.overrides.set(part.replace(/^\//, ''), type)
        }
      }
    }
  }

  getPart(name: string): Uint8Array | undefined {
    const data = this.parts.get(name)
    return data === undefined ? undefined : new Uint8Array(data)
  }

  partNames(): string[] {
    return [...this.parts.keys()]
  }

  contentTypeOf(partName: string): string | undefined {
    return this.overrides.get(partName) ?? this.defaults.get(extensionOf(partName))
  }

  rootRelationships(): OpcRelationship[] {
    const rels = this.parts.get(ROOT_RELS)
    if (rels === undefined) return []
    return this.parseRelationships(new TextDecoder().decode(rels))
  }

  private parseRelationships(xml: string): OpcRelationship[] {
    const out: OpcRelationship[] = []
    for (const tok of tokenize(xml)) {
      if (tok.type === 'open' && tok.name === 'Relationship') {
        const id = tok.attributes['Id']
        const type = tok.attributes['Type']
        const target = tok.attributes['Target']
        if (id !== undefined && type !== undefined && target !== undefined) {
          const targetMode = tok.attributes['TargetMode']
          out.push(
            targetMode === undefined ? { id, type, target } : { id, type, target, targetMode },
          )
        }
      }
    }
    return out
  }

  /** Directory of a part path (`xl/worksheets/sheet1.xml` -> `xl/worksheets`), '' for root. */
  private static dirOf(partName: string): string {
    const slash = partName.lastIndexOf('/')
    return slash === -1 ? '' : partName.slice(0, slash)
  }

  /** Collapse `seg/../` and `./` segments in a package path (e.g. `xl/worksheets/../media/x`). */
  private static normalizePath(path: string): string {
    const out: string[] = []
    for (const seg of path.split('/')) {
      if (seg === '' || seg === '.') continue
      if (seg === '..') out.pop()
      else out.push(seg)
    }
    return out.join('/')
  }

  /** Resolve a relationship Target against a base directory into an absolute part path. */
  private static resolveTarget(baseDir: string, target: string): string {
    if (target.startsWith('/')) return OpcPackage.normalizePath(target) // already package-absolute
    if (baseDir === '') return OpcPackage.normalizePath(target)
    return OpcPackage.normalizePath(`${baseDir}/${target}`)
  }

  /**
   * Relationships declared for a specific part (`<dir>/_rels/<file>.rels`), with internal
   * targets resolved to absolute part paths. External targets are returned verbatim.
   */
  relationshipsFor(partName: string): OpcRelationship[] {
    const dir = OpcPackage.dirOf(partName)
    const base = dir === '' ? partName : partName.slice(dir.length + 1)
    const relsPath = dir === '' ? `_rels/${base}.rels` : `${dir}/_rels/${base}.rels`
    const bytes = this.parts.get(relsPath)
    if (bytes === undefined) return []
    const rels = this.parseRelationships(new TextDecoder().decode(bytes))
    for (const r of rels) {
      if (r.targetMode !== 'External') r.target = OpcPackage.resolveTarget(dir, r.target)
    }
    return rels
  }

  setPart(name: string, contentType: string, data: Uint8Array): void {
    this.parts.set(name, data)
    this.overrides.set(name, contentType)
  }

  /** Register a `<Default>` content type for a file extension (e.g. `png` → `image/png`). */
  setDefault(extension: string, contentType: string): void {
    this.defaults.set(extension.toLowerCase(), contentType)
  }

  /**
   * Add a part whose content type is resolved by a registered `Default` extension (no per-part
   * `Override`) — used for binary media whose extension is declared via `setDefault`.
   */
  addPart(name: string, data: Uint8Array): void {
    this.parts.set(name, data)
  }

  private buildContentTypes(): string {
    const w = new XmlWriter().declaration().open('Types', { xmlns: TYPES_NS })
    for (const [ext, type] of this.defaults)
      w.leaf('Default', { Extension: ext, ContentType: type })
    for (const [part, type] of this.overrides) {
      w.leaf('Override', { PartName: `/${part}`, ContentType: type })
    }
    return w.close('Types').toString()
  }

  async toBytes(): Promise<Uint8Array> {
    const encoder = new TextEncoder()
    const entries: ZipEntry[] = [
      { name: CONTENT_TYPES, data: encoder.encode(this.buildContentTypes()) },
    ]
    for (const [name, data] of this.parts) entries.push({ name, data })
    return writeZip(entries)
  }
}
