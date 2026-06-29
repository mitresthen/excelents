export declare function gzipSize(buf: Buffer): number

export interface BudgetViolation {
  entry: string
  size: number | null
  budget: number
  reason?: string
}

export interface BudgetResult {
  ok: boolean
  violations: BudgetViolation[]
}

export declare function checkBudgets(
  sizes: Record<string, number>,
  budgets: Record<string, number>,
): BudgetResult
