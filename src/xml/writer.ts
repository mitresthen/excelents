import { escapeAttr, escapeText } from './entities'

export type XmlAttrs = Record<string, string | number | undefined>

function renderAttrs(attrs: XmlAttrs | undefined): string {
  if (attrs === undefined) return ''
  let out = ''
  for (const key of Object.keys(attrs)) {
    const value = attrs[key]
    if (value === undefined) continue
    out += ` ${key}="${escapeAttr(String(value))}"`
  }
  return out
}

/** Builds well-formed XML as a string, with correct escaping. */
export class XmlWriter {
  private buf = ''

  declaration(): this {
    this.buf += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return this
  }

  open(name: string, attrs?: XmlAttrs): this {
    this.buf += `<${name}${renderAttrs(attrs)}>`
    return this
  }

  leaf(name: string, attrs?: XmlAttrs): this {
    this.buf += `<${name}${renderAttrs(attrs)}/>`
    return this
  }

  text(content: string): this {
    this.buf += escapeText(content)
    return this
  }

  close(name: string): this {
    this.buf += `</${name}>`
    return this
  }

  toString(): string {
    return this.buf
  }
}
