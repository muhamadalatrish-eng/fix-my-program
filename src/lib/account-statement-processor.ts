export type StatementCell = string | number | boolean | Date | null | undefined

export type StatementSheet = { rows: StatementCell[][] }

export type StatementFormat = "ledger" | "bank"

export type ExtractedStatement = {
  format: StatementFormat
  headers: string[]
  rows: StatementCell[][]
  headerBlocks: number
  sourceRows: number
}

type HeaderIndexes =
  | {
      format: "bank"
      indexes: number[]
      transactionIndexes: number[]
    }
  | {
      format: "ledger"
      indexes: number[]
      entryNumber: number
      transactionIndexes: number[]
    }

export const LEDGER_HEADERS = ["رقم القيد", "التاريخ", "المدين", "الدائن", "الرصيد", "البيان", "بيان القيد"]
export const BANK_HEADERS = ["التاريخ البنكي", "التاريخ الفعلي", "الإيضاحات", "تاريخ الحق", "مبالغ مدفوعة", "مبالغ مستلمة", "الرصيد (شيكل)"]

function normalizeHeader(value: StatementCell): string {
  return String(value ?? "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/\s+/g, "")
}

function findHeaderIndexes(row: StatementCell[]): HeaderIndexes | null {
  const headers = row.map(normalizeHeader)
  const find = (predicate: (header: string) => boolean) => headers.findIndex(predicate)
  const bankIndexes = [
    find((header) => /تاريخ.*بنكي|بنكي.*تاريخ/.test(header)),
    find((header) => /تاريخ.*فعلي|فعلي.*تاريخ/.test(header)),
    find((header) => /ايضاحات|ملاحظات|clarification|description/.test(header)),
    find((header) => /تاريخ.*حق|حق.*تاريخ/.test(header)),
    find((header) => /مبالغ?مدفوعة|مدفوع|paid/.test(header)),
    find((header) => /مبالغ?مستلمة|مستلم|received/.test(header)),
    find((header) => /رصيد|balance/.test(header)),
  ]
  if (bankIndexes.slice(0, 6).every((index) => index >= 0)) {
    return { format: "bank", indexes: bankIndexes, transactionIndexes: bankIndexes }
  }

  const entryNumber = find((header) => /رقم/.test(header) && /قيد|حركة|معاملة|عملية|سند|مرجع/.test(header))
  const transactionIndexes = [
    entryNumber,
    find((header) => /تاريخ/.test(header) || header === "date"),
    find((header) => /مدين|debit/.test(header)),
    find((header) => /دائن|credit/.test(header)),
    find((header) => /رصيد|balance/.test(header)),
    find((header) => /^(ال)?بيان$|وصف|تفاصيل|description|details/.test(header)),
    find((header) => /بيان.*قيد|قيد.*بيان|نوعالحركة|نوعالعملية/.test(header)),
  ]
  if (entryNumber >= 0 && transactionIndexes[1] >= 0 && transactionIndexes[2] >= 0 && transactionIndexes[3] >= 0 && transactionIndexes.slice(5).some((index) => index >= 0)) {
    return {
      format: "ledger",
      indexes: transactionIndexes,
      entryNumber,
      transactionIndexes: transactionIndexes.filter((index) => index >= 0),
    }
  }
  return null
}

function hasValue(value: StatementCell): boolean {
  return value !== null && value !== undefined && String(value).trim() !== ""
}

export function extractStatementRows(sheets: StatementSheet[]): ExtractedStatement | null {
  const rows: StatementCell[][] = []
  let format: StatementFormat | null = null
  let headers: string[] = []
  let headerBlocks = 0
  let sourceRows = 0

  for (const sheet of sheets) {
    let active: HeaderIndexes | null = null
    let readingTransactions = false
    for (const sourceRow of sheet.rows) {
      sourceRows += 1
      const detected = findHeaderIndexes(sourceRow)
      if (detected) {
        active = detected
        readingTransactions = true
        headerBlocks += 1
        if (!format) {
          format = detected.format
          headers = detected.format === "bank" ? [...BANK_HEADERS] : [...LEDGER_HEADERS]
        } else if (format !== detected.format) {
          active = null
          readingTransactions = false
        }
        continue
      }
      if (!sourceRow.some(hasValue)) {
        readingTransactions = false
        continue
      }
      if (!active || !readingTransactions || active.format !== format) continue
      if (active.format === "bank") {
        const transaction = active.transactionIndexes.map((index) => sourceRow[index] ?? "")
        if (!hasValue(transaction[0]) || !hasValue(transaction[1])) continue
        if (!transaction.slice(2, 6).some(hasValue)) continue
        rows.push(transaction)
      } else {
        const transaction = active.indexes.map((index) => index < 0 ? "" : sourceRow[index] ?? "")
        if (!hasValue(transaction[0])) continue
        if (!transaction.some(hasValue)) continue
        rows.push(transaction)
      }
    }
  }
  return format ? { format, headers, rows, headerBlocks, sourceRows } : null
}

export function mergeSelectedStatements(
  documents: Array<{ headers: string[]; rows: StatementCell[][]; selectedColumns: boolean[]; selectedRows: boolean[] }>,
): { headers: string[]; rows: StatementCell[][] } {
  const headers: string[] = []
  const indexes = new Map<string, number>()
  const normalize = (value: string) => normalizeHeader(value)
  documents.forEach((document) => document.headers.forEach((header, index) => {
    if (!document.selectedColumns[index]) return
    const key = normalize(header)
    if (!indexes.has(key)) {
      indexes.set(key, headers.length)
      headers.push(header)
    }
  }))
  const rows = documents.flatMap((document) => document.rows.flatMap((row, rowIndex) => {
    if (!document.selectedRows[rowIndex]) return []
    const merged = Array.from({ length: headers.length }, () => "" as StatementCell)
    document.headers.forEach((header, columnIndex) => {
      if (!document.selectedColumns[columnIndex]) return
      const target = indexes.get(normalize(header))
      if (target !== undefined) merged[target] = row[columnIndex] ?? ""
    })
    return [merged]
  }))
  return { headers, rows }
}
