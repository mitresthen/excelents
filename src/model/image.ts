/** Supported embedded-image formats. */
export type ImageExtension = 'png' | 'jpeg' | 'gif'

/** Input to `Workbook.addImage` — raw bytes or a base64 string (no `data:` prefix). */
export interface WorkbookImageInput {
  data: Uint8Array | string
  extension: ImageExtension
}

/** An image registered on a workbook, decoded to bytes. */
export interface WorkbookImage {
  data: Uint8Array
  extension: ImageExtension
}

/** Where and how big to draw an image on a sheet (one-cell anchor, fixed pixel size). */
export interface ImageAnchor {
  /** Top-left anchor cell, e.g. `'F1'`. */
  tl: string
  /** Displayed size in pixels (96 DPI). */
  size: { width: number; height: number }
  /** Anchor behavior; defaults to `'oneCell'` (moves with the cell, fixed size). */
  editAs?: 'oneCell' | 'absolute'
}

/** An image placement on a worksheet: a workbook image id plus its anchor. */
export interface ImagePlacement extends ImageAnchor {
  imageId: number
}
