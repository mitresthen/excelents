import type { SharedStrings } from '../utils/shared-strings'
import { XmlWriter } from '../xml/writer'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Serialize the shared-string table to `xl/sharedStrings.xml`. */
export function writeSharedStringsXml(sst: SharedStrings): string {
  const w = new XmlWriter()
    .declaration()
    .open('sst', { xmlns: MAIN_NS, count: sst.count, uniqueCount: sst.uniqueCount })
  for (const value of sst.values) w.open('si').open('t').text(value).close('t').close('si')
  return w.close('sst').toString()
}
