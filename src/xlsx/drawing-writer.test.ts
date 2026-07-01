import { expect, test } from 'vitest'
import type { ImagePlacement, WorkbookImage } from '../model/image'
import { writeDrawingRelsXml, writeDrawingXml } from './drawing-writer'

test('writeDrawingXml renders a oneCellAnchor with 0-indexed from, EMU ext, and blip rId', () => {
  const placements: ImagePlacement[] = [{ imageId: 0, tl: 'F1', size: { width: 180, height: 101 } }]
  const xml = writeDrawingXml(placements)
  expect(xml).toContain(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
  )
  expect(xml).toContain('<xdr:oneCellAnchor editAs="oneCell">')
  // F1 -> col 5, row 0 (0-indexed drawing coords)
  expect(xml).toContain('<xdr:col>5</xdr:col>')
  expect(xml).toContain('<xdr:row>0</xdr:row>')
  // 180*9525=1714500, 101*9525=962025
  expect(xml).toContain('<xdr:ext cx="1714500" cy="962025"/>')
  expect(xml).toContain(
    '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print"/>',
  )
  expect(xml).toContain('<a:picLocks noChangeAspect="1"/>')
  expect(xml).toContain('<xdr:clientData/>')
})

test('writeDrawingXml honors editAs and numbers anchors from 1', () => {
  const placements: ImagePlacement[] = [
    { imageId: 0, tl: 'A1', size: { width: 10, height: 10 }, editAs: 'absolute' },
    { imageId: 0, tl: 'B2', size: { width: 10, height: 10 } },
  ]
  const xml = writeDrawingXml(placements)
  expect(xml).toContain('<xdr:oneCellAnchor editAs="absolute">')
  expect(xml).toContain('name="Picture 1"')
  expect(xml).toContain('name="Picture 2"')
  expect(xml).toContain('r:embed="rId2"')
})

test('writeDrawingRelsXml maps each placement rId to its media target by extension', () => {
  const media: WorkbookImage[] = [
    { data: new Uint8Array(), extension: 'png' },
    { data: new Uint8Array(), extension: 'jpeg' },
  ]
  const placements: ImagePlacement[] = [{ imageId: 1, tl: 'A1', size: { width: 10, height: 10 } }]
  const xml = writeDrawingRelsXml(placements, media)
  expect(xml).toContain('Id="rId1"')
  expect(xml).toContain('Target="../media/image2.jpeg"')
  expect(xml).toContain('/relationships/image"')
})
