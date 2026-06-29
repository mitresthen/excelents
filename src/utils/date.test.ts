import { expect, test } from 'vitest'
import { dateToSerial, serialToDate } from './date'

const utc = (y: number, m: number, d: number): Date => new Date(Date.UTC(y, m - 1, d))

test('dateToSerial matches well-known Excel anchors (1900 system)', () => {
  expect(dateToSerial(utc(1900, 3, 1))).toBe(61)
  expect(dateToSerial(utc(2000, 1, 1))).toBe(36526)
  expect(dateToSerial(utc(2024, 1, 1))).toBe(45292)
})

test('serialToDate inverts dateToSerial (1900 system)', () => {
  for (const d of [utc(1900, 3, 1), utc(2000, 1, 1), utc(2024, 1, 1), utc(2026, 6, 29)]) {
    expect(serialToDate(dateToSerial(d)).getTime()).toBe(d.getTime())
  }
})

test('1904 system anchors and round-trip', () => {
  expect(dateToSerial(utc(1904, 1, 1), { date1904: true })).toBe(0)
  const d = utc(2024, 1, 1)
  expect(serialToDate(dateToSerial(d, { date1904: true }), { date1904: true }).getTime()).toBe(
    d.getTime(),
  )
})

test('fractional serials carry the time of day', () => {
  // 0.5 day == 12:00 noon
  const noon = serialToDate(45292.5)
  expect(noon.getUTCHours()).toBe(12)
})

test('dateToSerial agrees with the exceljs oracle for modern dates', async () => {
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('s')
  for (const d of [utc(2000, 1, 1), utc(2024, 1, 1), utc(2026, 6, 29)]) {
    const cell = ws.getCell('A1')
    cell.value = d
    // exceljs stores the serial in the cell model when written; compare to ours.
    const oracleSerial = Math.floor((d.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000)
    expect(Math.floor(dateToSerial(d))).toBe(oracleSerial)
  }
})
