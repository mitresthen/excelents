# excelents — xlsx write features: images, autoFilter, frozen panes, alignment indent

**Status:** approved design · **Date:** 2026-07-01 · **Scope:** write-only

## 1. Goal & scope

Add four xlsx serialization capabilities that a consumer's report-export code needs and
excelents currently can't emit:

1. **Embedded images** (workbook media + per-sheet drawing) — e.g. a logo anchored top-right.
2. **AutoFilter** — `<autoFilter ref="A9:F123"/>` on a sheet.
3. **Frozen header rows** (and columns) — `<sheetViews>` with a frozen `<pane>`.
4. **Alignment indent** — `indent` attribute on cell alignment.

**Write-only.** No read/round-trip support for any of the four. **Buffered `writeXlsx` only**;
the streaming writer (`stream-writer.ts`) is out of scope. **Idiomatic excelents API** (not an
ExcelJS-compatible shim); the consumer updates its export call sites.

The exact XML shapes below are taken from exceljs 4.4.0's own serializers, because exceljs is our
conformance oracle: whatever we emit must round-trip through `exceljs`, which is also battle-tested
against real Excel.

## 2. Public API (model)

```ts
// Alignment gains one field (src/model/style.ts)
interface Alignment {
  horizontal?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
  indent?: number // NEW — non-negative; emitted only when > 0
}

// Worksheet (src/model/worksheet.ts)
ws.setAutoFilter(ref: string): void         // e.g. 'A9:F123'
ws.freeze(opts: { rows?: number; cols?: number }): void  // frozen split counts
ws.placeImage(imageId: number, placement: ImagePlacement): void

// Workbook (src/model/workbook.ts)
wb.addImage(image: WorkbookImageInput): number  // returns a workbook-global image id

type ImageExtension = 'png' | 'jpeg' | 'gif'
interface WorkbookImageInput {
  data: Uint8Array | string // string = base64 (no data: prefix)
  extension: ImageExtension
}
interface ImagePlacement {
  tl: string                                // top-left anchor cell, e.g. 'F1'
  size: { width: number; height: number }   // pixels (96 DPI)
  editAs?: 'oneCell' | 'absolute'           // default 'oneCell'
}
```

Rationale for the anchor being an address string (`tl: 'F1'`) rather than ExcelJS's 0-indexed
`{col,row}`: excelents already addresses cells as strings everywhere (`ws.cell('F1')`), so this is
consistent. Internally `'F1'` → `decodeAddress` (1-indexed) → drawing coords (0-indexed).

`wb.addImage` returns a 0-based id (index into workbook media). The same image reused across sheets
is stored once; each `placeImage` references it by id.

## 3. Per-feature serialization

### 3.1 Alignment indent (`styles-writer.ts`, `style.ts`)

- `Alignment.indent?: number` added.
- `hasAlignment(a)` also true when `a.indent !== undefined && a.indent > 0`.
- `alignmentKey(a)` gains an `|i=<n>` segment so indent participates in xf dedup.
- `writeAlignmentXml` emits `indent` when `> 0` (0 is the OOXML default → omit).

### 3.2 AutoFilter (`worksheet.ts`, `worksheet-writer.ts`)

- Worksheet stores an optional `autoFilterRef` string; `setAutoFilter(ref)` sets it,
  `get autoFilter(): string | undefined` reads it.
- Emit `<autoFilter ref="..."/>` **after `</sheetData>` and before `<mergeCells>`** (CT_Worksheet
  sequence). This is a direct child of `<worksheet>` — distinct from the `<autoFilter>` inside a
  `<table>` part.

### 3.3 Frozen panes (`worksheet.ts`, `worksheet-writer.ts`)

- Worksheet stores an optional `{ rows, cols }` (each defaulting to 0); `freeze({rows,cols})` sets it.
- Emit `<sheetViews>` **after `<dimension>` and before `<cols>`** (CT_Worksheet sequence):

```xml
<!-- example: freeze({ rows: 9 }) -->
<sheetViews>
  <sheetView workbookViewId="0">
    <pane ySplit="9" topLeftCell="A10" activePane="bottomLeft" state="frozen"/>
    <selection pane="bottomLeft" activeCell="A10" sqref="A10"/>
  </sheetView>
</sheetViews>
```

- `xSplit`/`ySplit` emitted only when > 0. `topLeftCell = encodeAddress(rows+1, cols+1)`.
- `activePane`: `cols>0 && rows>0 → 'bottomRight'`; `cols>0 → 'topRight'`; else `'bottomLeft'`
  (matches exceljs `sheet-view-xform`).

### 3.4 Images (biggest lift)

Follows the existing **tables** precedent (sheet → part → rels), with one extra hop
(sheet → drawing → media).

**Parts produced (per workbook / per sheet-with-images):**

| Part | Content type | Notes |
| --- | --- | --- |
| `xl/media/image{k}.{ext}` | `Default Extension` → `image/png` etc. | workbook-global, 1-based `k` |
| `xl/drawings/drawing{n}.xml` | `Override` → `…drawing+xml` | one per sheet that has images |
| `xl/drawings/_rels/drawing{n}.xml.rels` | rels | drawing → media (`…/relationships/image`) |
| sheet gains `<drawing r:id>` | — | sheet → drawing (`…/relationships/drawing`) |

**`xl/drawings/drawing{n}.xml`** (new `drawing-writer.ts`), one `oneCellAnchor` per placement:

```xml
<xdr:wsDr xmlns:xdr="…/drawingml/2006/spreadsheetDrawing" xmlns:a="…/drawingml/2006/main">
  <xdr:oneCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="1714500" cy="962025"/>            <!-- width*9525, height*9525, floored -->
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="…/officeDocument/2006/relationships" r:embed="rId1" cstate="print"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
```

- `EMU_PER_PIXEL = 9525`. `cNvPr id`/`name` use the 1-based anchor index within the drawing.
- Each anchor's `r:embed` is a drawing-local rId (`rId1`, `rId2`, …), mapped in the drawing rels to
  `../media/image{k}.{ext}` for that placement's image id. (One rel per placement — if the same
  image is placed twice, two rels target the same media file; valid and simplest.)

**OPC support (`opc/package.ts`)** — two small additions:
- `setDefault(ext, contentType)` → emits `<Default Extension="png" ContentType="image/png"/>`
  (dedupes across images sharing an extension; this is what Excel itself writes).
- `addPart(name, data)` → adds a part **without** a content-type Override (media relies on the
  Default). Existing `setPart` keeps writing an Override (used for the drawing xml).

**Pipeline wiring (`write.ts`, `worksheet-writer.ts`, `workbook-writer.ts`):**
- `worksheet-writer` emits `<drawing r:id>` **before `<tableParts>`** and returns
  `drawing?: { rid, placements }` in `WorksheetWriteResult`. The drawing rId is allocated after
  hyperlink and table rIds (`rId{pending.length + tables.length + 1}`) so existing table rId
  numbering is untouched.
- `write.ts` emits media parts once (`wb.media`), then per sheet-with-images: `drawingCounter++`,
  write `drawing{n}.xml` + its rels, and add the sheet→drawing rel.
- `writeWorksheetRelsXml(hyperlinks, tables, drawings)` gains a `drawings` param
  (`Type=…/relationships/drawing`, internal target `../drawings/drawing{n}.xml`).
- `content-types.ts` gains `drawing` content type; `image/png|jpeg|gif` via `imageContentType(ext)`.

**base64 → bytes** (`utils/base64.ts`): `atob`-based, `node:`-free and universal
(`Uint8Array.fromBase64` is not yet in Node 24; `atob` is the web-standard path that also works in
browsers/edge). `Uint8Array` input passes through unchanged.

## 4. New / changed files

**New:** `src/xlsx/drawing-writer.ts`, `src/utils/base64.ts` (+ their `.test.ts`).
**Changed:** `src/model/style.ts`, `src/model/worksheet.ts`, `src/model/workbook.ts`,
`src/xlsx/styles-writer.ts`, `src/xlsx/worksheet-writer.ts`, `src/xlsx/workbook-writer.ts`,
`src/xlsx/write.ts`, `src/xlsx/content-types.ts`, `src/opc/package.ts`, `src/index.ts` (export new
option types). Size budget (`size-budget.json`) may need a small bump for `index.js`.

## 5. Testing (TDD)

Each unit gets a focused test written first:
- `base64.test.ts` — round-trips known vectors; passes through `Uint8Array`.
- `styles-writer` — indent emitted only when > 0; participates in xf dedup.
- `worksheet-writer` — `<autoFilter>` slot/attr; `<sheetViews>`/`<pane>` for frozen rows & cols;
  `<drawing r:id>` placed before `<tableParts>`.
- `drawing-writer` — anchor XML for a placement (col/row/EMU/`r:embed`).
- `package` — `setDefault`/`addPart` produce the expected `[Content_Types].xml` and media part.

**Oracle round-trip** (extends the existing exceljs-based conformance tests): write with excelents,
read back with exceljs, assert — `worksheet.autoFilter` ref; `views[0]` state `frozen` + `ySplit`;
a cell's `alignment.indent`; `worksheet.getImages()` count + anchor + `workbook` media presence and
extent. Also assert the package opens cleanly (no repair).

## 6. Non-goals

- Reading any of these back into the excelents model.
- Streaming-writer support for images/drawings.
- `twoCellAnchor` / `absoluteAnchor`, image hyperlinks, background images, cropping, DPI scaling.
- ExcelJS API compatibility.

## 7. Consumer migration

The consumer's export code moves from ExcelJS shapes to excelents idiomatic calls:
`sheet.autoFilter = {from,to}` → `ws.setAutoFilter('A9:F123')`; `sheet.views = [{state,ySplit}]` →
`ws.freeze({rows})`; `workbook.addImage` + `sheet.addImage(id,{tl,ext,editAs})` →
`wb.addImage(...)` + `ws.placeImage(id, {tl, size})`; alignment `indent` is a drop-in field. This
lands in the consumer repo, separately from this library work.
