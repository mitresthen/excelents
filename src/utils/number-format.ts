/** OOXML built-in number-format ids → format codes (ECMA-376 §18.8.30).
 *
 * Ids 27-36, 50-58, 59-70, 81 are locale-specific in exceljs and omitted here;
 * all entries below are universal (non-locale) and verified against
 * exceljs `lib/xlsx/defaultnumformats.js` (the authoritative oracle).
 *
 * Corrections vs. the brief's table (exceljs is authoritative for byte-compat):
 *   id 22: 'm/d/yy "h":mm'  (brief had 'm/d/yy h:mm' — missing quotes around h)
 *   id 39: '#,##0.00 ;(#,##0.00)'  (brief had no space before semicolon)
 *   id 40: '#,##0.00 ;[Red](#,##0.00)'  (brief had no space before semicolon)
 */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy "h":mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00 ;(#,##0.00)',
  40: '#,##0.00 ;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

/** Look up a built-in format code by id, or `undefined` if not a builtin. */
export function builtinFormatCode(id: number): string | undefined {
  return BUILTIN_FORMATS[id]
}

/** Reverse lookup: the first builtin id whose code equals `code`. */
export function builtinFormatId(code: string): number | undefined {
  for (const key of Object.keys(BUILTIN_FORMATS)) {
    const id = Number(key)
    if (BUILTIN_FORMATS[id] === code) return id
  }
  return undefined
}
