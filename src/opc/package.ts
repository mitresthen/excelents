import { XmlWriter } from '../xml/writer'
import { tokenize } from '../xml/tokenizer'
import { readZip } from './zip-reader'
import { type ZipEntry, writeZip } from './zip-writer'

export interface OpcRelationship {
  id: string
  type: string
  target: string
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
    const entries = await readZip(bytes)
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
    return this.parts.get(name)
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
          out.push({ id, type, target })
        }
      }
    }
    return out
  }

  setPart(name: string, contentType: string, data: Uint8Array): void {
    this.parts.set(name, data)
    this.overrides.set(name, contentType)
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
