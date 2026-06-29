import { unescapeXml } from './entities'

export type XmlToken =
  | { type: 'open'; name: string; attributes: Record<string, string>; selfClosing: boolean }
  | { type: 'close'; name: string }
  | { type: 'text'; value: string }

function isSpace(ch: string): boolean {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r'
}

/** Lazily tokenize an OOXML-subset XML string. */
export function* tokenize(xml: string): Generator<XmlToken, void, unknown> {
  let i = 0
  const n = xml.length
  while (i < n) {
    if (xml[i] !== '<') {
      const next = xml.indexOf('<', i)
      const end = next === -1 ? n : next
      yield { type: 'text', value: unescapeXml(xml.slice(i, end)) }
      i = end
      continue
    }
    if (xml.startsWith('<?', i)) {
      const end = xml.indexOf('?>', i + 2)
      i = end === -1 ? n : end + 2
    } else if (xml.startsWith('<!--', i)) {
      const end = xml.indexOf('-->', i + 4)
      i = end === -1 ? n : end + 3
    } else if (xml.startsWith('<![CDATA[', i)) {
      const end = xml.indexOf(']]>', i + 9)
      yield { type: 'text', value: xml.slice(i + 9, end === -1 ? n : end) }
      i = end === -1 ? n : end + 3
    } else if (xml.startsWith('</', i)) {
      const end = xml.indexOf('>', i + 2)
      yield { type: 'close', name: xml.slice(i + 2, end).trim() }
      i = end + 1
    } else {
      i += 1
      let j = i
      while (j < n && !isSpace(xml[j]!) && xml[j] !== '>' && xml[j] !== '/') j++
      const name = xml.slice(i, j)
      i = j
      const attributes: Record<string, string> = {}
      for (;;) {
        while (i < n && isSpace(xml[i]!)) i++
        if (i >= n) break
        if (xml[i] === '/' && xml[i + 1] === '>') {
          i += 2
          yield { type: 'open', name, attributes, selfClosing: true }
          break
        }
        if (xml[i] === '>') {
          i += 1
          yield { type: 'open', name, attributes, selfClosing: false }
          break
        }
        let k = i
        while (k < n && xml[k] !== '=' && !isSpace(xml[k]!) && xml[k] !== '>' && xml[k] !== '/') k++
        const attrName = xml.slice(i, k)
        i = k
        while (i < n && isSpace(xml[i]!)) i++
        if (xml[i] === '=') {
          i += 1
          while (i < n && isSpace(xml[i]!)) i++
          const quote = xml[i]
          if (quote === '"' || quote === "'") {
            i += 1
            const close = xml.indexOf(quote, i)
            attributes[attrName] = unescapeXml(xml.slice(i, close === -1 ? n : close))
            i = close === -1 ? n : close + 1
          }
        } else {
          attributes[attrName] = ''
        }
      }
    }
  }
}
