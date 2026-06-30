/** A cell-range input restriction (dropdown list, numeric bounds, etc.). */
export interface DataValidation {
  /** Target range(s), e.g. `'A1:A10'` or space-separated `'A1 C1'`. */
  readonly sqref: string
  readonly type: 'list' | 'whole' | 'decimal' | 'textLength' | 'date'
  readonly operator?:
    | 'between'
    | 'notBetween'
    | 'equal'
    | 'notEqual'
    | 'greaterThan'
    | 'lessThan'
    | 'greaterThanOrEqual'
    | 'lessThanOrEqual'
  readonly formula1: string
  readonly formula2?: string
  readonly allowBlank?: boolean
  /** When `false`, the in-cell dropdown is hidden (OOXML stores this inverted). */
  readonly showDropDown?: boolean
  readonly showErrorMessage?: boolean
}
