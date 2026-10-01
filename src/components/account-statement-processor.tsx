import { useMemo, useRef, useState, type ChangeEvent, type DragEvent } from "react"
import * as XLSX from "xlsx"
import { AlertTriangle, ArrowDownToLine, ChevronDown, FileSpreadsheet, Plus, RefreshCw, ShieldCheck, Sparkles, Upload } from "lucide-react"
import { extractStatementRows, mergeSelectedStatements, type StatementCell, type StatementFormat } from "../lib/account-statement-processor"

type StatementDocument = {
  id: string
  fileName: string
  format: StatementFormat
  headers: string[]
  rows: StatementCell[][]
  selectedColumns: boolean[]
  selectedRows: boolean[]
  pages: number
}

let nextId = 0

function readFile(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => reader.result instanceof ArrayBuffer ? resolve(reader.result) : reject(new Error("تعذر قراءة الملف."))
    reader.onerror = () => reject(new Error("حدث خطأ أثناء قراءة الملف."))
    reader.readAsArrayBuffer(file)
  })
}

function cellText(value: StatementCell): string {
  if (value instanceof Date) return value.toLocaleDateString("en-GB")
  return value == null ? "" : String(value)
}

async function parseFile(file: File): Promise<StatementDocument> {
  if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) throw new Error("صيغة الملف غير مدعومة. ارفع XLSX أو XLS أو XLSM أو CSV.")
  const workbook = XLSX.read(await readFile(file), { type: "array", cellDates: true })
  const statement = extractStatementRows(workbook.SheetNames.map((name) => ({
    rows: XLSX.utils.sheet_to_json<StatementCell[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false }),
  })))
  if (!statement) throw new Error("لم أتعرف على ترويسة كشف الحساب في هذا الملف.")
  if (!statement.rows.length) throw new Error("تم التعرف على الأعمدة، لكن لم يتم العثور على حركات.")
  return {
    id: `${Date.now()}-${nextId++}`,
    fileName: file.name,
    format: statement.format,
    headers: statement.headers,
    rows: statement.rows,
    selectedColumns: statement.headers.map((header) => !/رصيد|balance/i.test(header)),
    selectedRows: statement.rows.map(() => true),
    pages: statement.headerBlocks,
  }
}

export function AccountStatementProcessor() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [documents, setDocuments] = useState<StatementDocument[]>([])
  const [expanded, setExpanded] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [dragging, setDragging] = useState(false)
  const [errors, setErrors] = useState<string[]>([])
  const merged = useMemo(() => mergeSelectedStatements(documents), [documents])
  const selectedCount = documents.reduce((count, item) => count + item.selectedRows.filter(Boolean).length, 0)

  const addFiles = async (fileList: FileList | File[]) => {
    const files = Array.from(fileList)
    if (!files.length) return
    setBusy(true)
    setErrors([])
    const loaded: StatementDocument[] = []
    const failed: string[] = []
    for (const file of files) {
      try {
        loaded.push(await parseFile(file))
      } catch (error) {
        failed.push(`${file.name}: ${error instanceof Error ? error.message : "فشل تحليل الملف."}`)
      }
    }
    if (loaded.length) {
      setDocuments((current) => [...current, ...loaded])
      setExpanded(loaded.at(-1)?.id ?? null)
    }
    setErrors(failed)
    setBusy(false)
    if (inputRef.current) inputRef.current.value = ""
  }

  const updateDocument = (id: string, updater: (document: StatementDocument) => StatementDocument) =>
    setDocuments((current) => current.map((document) => document.id === id ? updater(document) : document))

  const download = () => {
    if (!merged.headers.length || !merged.rows.length) return
    const sheet = XLSX.utils.aoa_to_sheet([merged.headers, ...merged.rows])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, "الكشف المنظف")
    XLSX.writeFile(workbook, "كشف-الحسابات-المنظف.xlsx")
  }

  return (
    <main dir="rtl" className="min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-6xl space-y-6">
        <header className="rounded-3xl bg-[#10243c] px-6 py-8 text-white shadow-xl sm:px-9">
          <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><Sparkles className="h-3.5 w-3.5" />معالجة محلية وآمنة</div>
          <h1 className="text-2xl font-extrabold sm:text-3xl">معالج كشوفات الحسابات والبيانات الذكي</h1>
          <p className="mt-3 max-w-3xl text-sm leading-7 text-slate-300">ارفع كشف المحفظة أو البنك، اختر الصفوف والأعمدة التي تريدها، واجمع أي عدد من الكشوف في ملف Excel واحد.</p>
          <div className="mt-4 flex items-center gap-2 text-xs text-emerald-200"><ShieldCheck className="h-4 w-4" />تتم المعالجة على جهازك فقط</div>
        </header>

        <section className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
          <div className="mb-4">
            <h2 className="text-sm font-bold">رفع الكشوفات</h2>
            <p className="mt-1 text-xs leading-6 text-slate-500">يمكن رفع عدة ملفات معاً أو إضافة ملفات لاحقاً بزر +.</p>
          </div>
          <input ref={inputRef} type="file" multiple accept=".xlsx,.xls,.xlsm,.csv" className="hidden" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files && void addFiles(event.target.files)} />
          <div
            onDragEnter={(event) => { event.preventDefault(); setDragging(true) }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setDragging(false) }}
            onDrop={(event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); void addFiles(event.dataTransfer.files) }}
            className={`flex min-h-48 flex-col items-center justify-center rounded-2xl border-2 border-dashed px-5 py-8 text-center ${dragging ? "border-cyan-500 bg-cyan-50" : "border-slate-200 bg-slate-50"}`}
          >
            <div className="mb-3 text-cyan-700">{busy ? <RefreshCw className="h-7 w-7 animate-spin" /> : documents.length ? <Plus className="h-7 w-7" /> : <Upload className="h-7 w-7" />}</div>
            <p className="text-sm font-bold">{busy ? "جارٍ قراءة الملفات..." : "اسحب ملفاً أو أكثر وأفلته هنا"}</p>
            <button type="button" disabled={busy} onClick={() => inputRef.current?.click()} className="mt-4 inline-flex items-center gap-2 rounded-xl bg-[#10243c] px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">
              {documents.length > 0 && <Plus className="h-4 w-4" />}{documents.length ? "إضافة ملف" : "اختيار ملف"}
            </button>
            <span className="mt-2 text-[11px] text-slate-400">XLSX · XLS · XLSM · CSV</span>
          </div>
          {!!errors.length && <div role="alert" className="mt-4 space-y-1 rounded-xl border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800"><AlertTriangle className="ml-2 inline h-4 w-4" />{errors.map((error) => <p key={error}>{error}</p>)}</div>}
        </section>

        {!!documents.length && <>
          <section className="space-y-3">
            <div><h2 className="text-lg font-bold">اختر ما تريد الاحتفاظ به من كل ملف</h2><p className="mt-1 text-xs text-slate-500">عمود الرصيد غير محدد تلقائياً. يمكنك تعديله، وتحديد الحركات المستبعدة يدوياً.</p></div>
            {documents.map((document, documentIndex) => {
              const isExpanded = expanded === document.id
              return <article key={document.id} className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
                <button type="button" onClick={() => setExpanded(isExpanded ? null : document.id)} aria-expanded={isExpanded} className="flex w-full items-center justify-between gap-4 px-4 py-4 text-right sm:px-6">
                  <span className="flex min-w-0 items-center gap-3"><FileSpreadsheet className="h-5 w-5 shrink-0 text-cyan-700" /><span className="min-w-0"><span className="block truncate text-sm font-bold">{documentIndex + 1}. {document.fileName}</span><span className="mt-1 block text-xs text-slate-500">{document.format === "bank" ? "كشف بنكي" : "كشف محاسبي"} · {document.selectedRows.filter(Boolean).length} من {document.rows.length} حركة محددة</span></span></span>
                  <ChevronDown className={`h-5 w-5 shrink-0 transition-transform ${isExpanded ? "rotate-180" : ""}`} />
                </button>
                {isExpanded && <div className="border-t border-slate-100 p-4 sm:p-6">
                  <div className="mb-2 flex flex-wrap items-center justify-between gap-2"><h3 className="text-sm font-bold">الأعمدة</h3><div className="flex gap-2"><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.headers.map(() => true) }))} className="rounded-lg border px-3 py-1.5 text-xs">تحديد الكل</button><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.headers.map(() => false) }))} className="rounded-lg border px-3 py-1.5 text-xs">إلغاء الكل</button></div></div>
                  <div className="mb-5 flex flex-wrap gap-2">{document.headers.map((header, index) => <label key={`${header}-${index}`} className={`inline-flex cursor-pointer items-center gap-2 rounded-xl border px-3 py-2 text-sm ${document.selectedColumns[index] ? "border-cyan-200 bg-cyan-50" : "border-slate-200 text-slate-500"}`}><input type="checkbox" checked={document.selectedColumns[index]} onChange={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.selectedColumns.map((selected, column) => column === index ? !selected : selected) }))} className="h-4 w-4 accent-cyan-700" />{header}</label>)}</div>
                  <div className="mb-2 flex flex-wrap items-center justify-between gap-2"><h3 className="text-sm font-bold">صفوف الحركات</h3><div className="flex gap-2"><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedRows: item.selectedRows.map(() => true) }))} className="rounded-lg border px-3 py-1.5 text-xs">تحديد الكل</button><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedRows: item.selectedRows.map(() => false) }))} className="rounded-lg border px-3 py-1.5 text-xs">إلغاء الكل</button></div></div>
                  <div className="max-h-[360px] overflow-auto rounded-xl border border-slate-200"><table className="w-full min-w-[760px] border-collapse text-right text-xs"><thead className="sticky top-0 bg-slate-100"><tr><th className="px-3 py-2">إبقاء</th><th className="px-3 py-2">#</th>{document.headers.map((header, index) => <th key={`${header}-${index}`} className="px-3 py-2">{header}</th>)}</tr></thead><tbody>{document.rows.map((row, rowIndex) => <tr key={rowIndex} className="border-t border-slate-100"><td className="px-3 py-2 text-center"><input type="checkbox" checked={document.selectedRows[rowIndex]} onChange={() => updateDocument(document.id, (item) => ({ ...item, selectedRows: item.selectedRows.map((selected, index) => index === rowIndex ? !selected : selected) }))} aria-label={`إبقاء الصف ${rowIndex + 1}`} /></td><td className="px-3 py-2">{rowIndex + 1}</td>{document.headers.map((header, columnIndex) => <td key={`${header}-${columnIndex}`} className="max-w-80 px-3 py-2">{cellText(row[columnIndex])}</td>)}</tr>)}</tbody></table></div>
                </div>}
              </article>
            })}
          </section>

          <section className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-sm">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b p-5"><div><h2 className="font-bold">الملف الموحّد</h2><p className="mt-1 text-xs text-slate-500">{documents.length} ملف · {selectedCount} حركة محفوظة · الرصيد مستبعد افتراضياً</p></div><button onClick={download} disabled={!merged.rows.length || !merged.headers.length} className="inline-flex items-center gap-2 rounded-xl bg-emerald-700 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50"><ArrowDownToLine className="h-4 w-4" />تنزيل Excel</button></div>
            {!!merged.headers.length && <div className="max-h-[420px] overflow-auto"><table className="w-full min-w-[760px] border-collapse text-right text-xs"><thead className="sticky top-0 bg-slate-100"><tr>{merged.headers.map((header, index) => <th key={`${header}-${index}`} className="px-4 py-3">{header}</th>)}</tr></thead><tbody>{merged.rows.slice(0, 30).map((row, rowIndex) => <tr key={rowIndex} className="border-t"><>{merged.headers.map((header, index) => <td key={`${header}-${index}`} className="px-4 py-3">{cellText(row[index])}</td>)}</></tr>)}</tbody></table></div>}
          </section>
        </>}
      </div>
    </main>
  )
}
