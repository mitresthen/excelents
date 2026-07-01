import type { ImagePlacement, WorkbookImage } from '../model/image'
import { decodeAddress } from '../utils/address'
import { XmlWriter } from '../xml/writer'

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** EMU (English Metric Units) per pixel at 96 DPI — the DrawingML pixel↔EMU conversion. */
const EMU_PER_PIXEL = 9525

/**
 * Serialize a sheet's image placements to `xl/drawings/drawingN.xml`. Each placement becomes a
 * `oneCellAnchor` (fixed pixel size anchored to a cell). The blip's `r:embed` is a drawing-local
 * relationship id (`rId{i+1}`) resolved by the drawing's `.rels` (see `writeDrawingRelsXml`).
 */
export function writeDrawingXml(placements: readonly ImagePlacement[]): string {
  const w = new XmlWriter().declaration().open('xdr:wsDr', { 'xmlns:xdr': XDR_NS, 'xmlns:a': A_NS })

  placements.forEach((p, i) => {
    const { row, col } = decodeAddress(p.tl) // 1-indexed
    const n = i + 1

    w.open('xdr:oneCellAnchor', { editAs: p.editAs ?? 'oneCell' })

    w.open('xdr:from')
      .open('xdr:col')
      .text(String(col - 1))
      .close('xdr:col')
      .open('xdr:colOff')
      .text('0')
      .close('xdr:colOff')
      .open('xdr:row')
      .text(String(row - 1))
      .close('xdr:row')
      .open('xdr:rowOff')
      .text('0')
      .close('xdr:rowOff')
      .close('xdr:from')

    w.leaf('xdr:ext', {
      cx: Math.floor(p.size.width * EMU_PER_PIXEL),
      cy: Math.floor(p.size.height * EMU_PER_PIXEL),
    })

    w.open('xdr:pic')
      .open('xdr:nvPicPr')
      .leaf('xdr:cNvPr', { id: n, name: `Picture ${n}` })
      .open('xdr:cNvPicPr')
      .leaf('a:picLocks', { noChangeAspect: 1 })
      .close('xdr:cNvPicPr')
      .close('xdr:nvPicPr')
      .open('xdr:blipFill')
      .leaf('a:blip', { 'xmlns:r': REL_NS, 'r:embed': `rId${n}`, cstate: 'print' })
      .open('a:stretch')
      .leaf('a:fillRect')
      .close('a:stretch')
      .close('xdr:blipFill')
      .open('xdr:spPr')
      .open('a:xfrm')
      .leaf('a:off', { x: 0, y: 0 })
      .leaf('a:ext', { cx: 0, cy: 0 })
      .close('a:xfrm')
      .open('a:prstGeom', { prst: 'rect' })
      .leaf('a:avLst')
      .close('a:prstGeom')
      .close('xdr:spPr')
      .close('xdr:pic')

    w.leaf('xdr:clientData')
    w.close('xdr:oneCellAnchor')
  })

  return w.close('xdr:wsDr').toString()
}

/**
 * Serialize `xl/drawings/_rels/drawingN.xml.rels` — one relationship per placement mapping its
 * drawing-local `rId{i+1}` to the workbook media file `../media/image{imageId+1}.{ext}`.
 */
export function writeDrawingRelsXml(
  placements: readonly ImagePlacement[],
  media: readonly WorkbookImage[],
): string {
  const w = new XmlWriter().declaration().open('Relationships', { xmlns: PKG_REL_NS })
  placements.forEach((p, i) => {
    const ext = media[p.imageId]?.extension ?? 'png'
    w.leaf('Relationship', {
      Id: `rId${i + 1}`,
      Type: IMAGE_REL,
      Target: `../media/image${p.imageId + 1}.${ext}`,
    })
  })
  return w.close('Relationships').toString()
}
