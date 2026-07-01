import type { ImageExtension } from '../model/image'

/** OOXML content types for the parts we emit. */
export const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
} as const

const IMAGE_CONTENT_TYPES: Record<ImageExtension, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

/** Content type for an image extension, registered as a `<Default>` on the package. */
export function imageContentType(ext: ImageExtension): string {
  return IMAGE_CONTENT_TYPES[ext]
}
