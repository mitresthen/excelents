import type { OpcPackage } from '../opc/package'
import { tokenize } from '../xml/tokenizer'

export interface WorkbookParts {
  /** Worksheet parts in workbook order: display name + absolute part path. */
  readonly sheets: ReadonlyArray<{ name: string; path: string }>
  /** Absolute path to sharedStrings.xml, or undefined if absent. */
  readonly sharedStringsPath: string | undefined
  /** Absolute path to styles.xml, or undefined if absent. */
  readonly stylesPath: string | undefined
  /** Defined names in document order; `localSheetId` is set for sheet-scoped names. */
  readonly definedNames: ReadonlyArray<{ name: string; formula: string; localSheetId?: number }>
  /** True when `<workbookPr date1904="1"/>` — serial dates count from the 1904 epoch. */
  readonly date1904: boolean
}

const OFFICE_DOC = '/officeDocument'
const SHARED_STRINGS = '/sharedStrings'
const STYLES = '/styles'

function workbookPartPath(pkg: OpcPackage): string {
  const rel = pkg.rootRelationships().find((r) => r.type.endsWith(OFFICE_DOC))
  // root rels targets are raw (`xl/workbook.xml`); strip a leading slash if present.
  return (rel?.target ?? 'xl/workbook.xml').replace(/^\//, '')
}

/** Locate the workbook's worksheet/sharedStrings/styles parts via the OPC relationships. */
export function readWorkbookParts(pkg: OpcPackage): WorkbookParts {
  const workbookPath = workbookPartPath(pkg)
  const rels = pkg.relationshipsFor(workbookPath)
  const byId = new Map(rels.map((r) => [r.id, r]))

  const sheets: Array<{ name: string; path: string }> = []
  const definedNames: Array<{ name: string; formula: string; localSheetId?: number }> = []
  let dnName: string | undefined
  let dnLocalSheetId: number | undefined
  let dnText = ''
  let inDefinedName = false

  let date1904 = false

  const xml = new TextDecoder().decode(pkg.getPart(workbookPath) ?? new Uint8Array())
  for (const tok of tokenize(xml)) {
    if (tok.type === 'open' && tok.name === 'workbookPr') {
      const flag = tok.attributes['date1904']
      date1904 = flag === '1' || flag === 'true'
    } else if (tok.type === 'open' && tok.name === 'sheet') {
      const name = tok.attributes['name']
      const rid = tok.attributes['r:id']
      const target = rid !== undefined ? byId.get(rid)?.target : undefined
      if (name !== undefined && target !== undefined) sheets.push({ name, path: target })
    } else if (tok.type === 'open' && tok.name === 'definedName') {
      inDefinedName = true
      dnName = tok.attributes['name']
      const lsid = tok.attributes['localSheetId']
      dnLocalSheetId = lsid !== undefined ? Number(lsid) : undefined
      dnText = ''
    } else if (tok.type === 'close' && tok.name === 'definedName') {
      if (dnName !== undefined) {
        definedNames.push(
          dnLocalSheetId === undefined
            ? { name: dnName, formula: dnText }
            : { name: dnName, formula: dnText, localSheetId: dnLocalSheetId },
        )
      }
      inDefinedName = false
    } else if (tok.type === 'text' && inDefinedName) {
      dnText += tok.value
    }
  }
  const sharedStringsPath = rels.find((r) => r.type.endsWith(SHARED_STRINGS))?.target
  const stylesPath = rels.find((r) => r.type.endsWith(STYLES))?.target
  return { sheets, sharedStringsPath, stylesPath, definedNames, date1904 }
}
