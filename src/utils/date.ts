export type DateMode = { date1904?: boolean }

const MS_PER_DAY = 86_400_000
const EPOCH_1900 = Date.UTC(1899, 11, 30) // absorbs the 1900 leap-year bug for dates >= 1900-03-01
const EPOCH_1904 = Date.UTC(1904, 0, 1)

function epoch(mode: DateMode): number {
  return mode.date1904 === true ? EPOCH_1904 : EPOCH_1900
}

/** Convert a UTC `Date` to an Excel serial number. */
export function dateToSerial(date: Date, mode: DateMode = {}): number {
  return (date.getTime() - epoch(mode)) / MS_PER_DAY
}

/** Convert an Excel serial number to a UTC `Date`. */
export function serialToDate(serial: number, mode: DateMode = {}): Date {
  return new Date(epoch(mode) + serial * MS_PER_DAY)
}
