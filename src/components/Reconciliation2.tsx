"use client"
/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars, no-constant-binary-expression */
import React, { useState, useCallback, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, Download, ChevronDown, X,
  Check, ChevronLeft, CreditCard,
  Shield, Trash2, Sparkles, Clock, Info, RotateCcw
} from "lucide-react";

// ─── Excel helpers ────────────────────────────────────────────────────────────
function readFileBuf(f: File): Promise<ArrayBuffer> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target!.result as ArrayBuffer);
    r.onerror = () => rej(new Error("فشل قراءة الملف"));
    r.readAsArrayBuffer(f);
  });
}
function parseSheet(buf: ArrayBuffer) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  return { headers: rows.length ? Object.keys(rows[0]) : [], rows };
}
function fmtDate(v: unknown): string {
  if (!v) return "";
  if (v instanceof Date) return v.toLocaleDateString("ar-SA");
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? String(v) : d.toLocaleDateString("ar-SA");
}
function toNum(v: unknown): number {
  if (v === "" || v == null) return 0;
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/,/g, "")) || 0;
}
function fmtNum(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ─── تخزين دائم (IndexedDB) ────────────────────────────────────────────────────
const IDB_NAME = "recon_store";
const IDB_STORE = "kv";
let idbReady: Promise<IDBDatabase> | null = null;

function openIDB(): Promise<IDBDatabase> {
  if (idbReady) return idbReady;
  idbReady = new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return idbReady;
}

async function idbGet<T>(key: string): Promise<T | undefined> {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = () => resolve(req.result as T | undefined);
    req.onerror = () => reject(req.error);
  });
}

async function idbSet(key: string, value: unknown): Promise<void> {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function storageGet<T>(key: string, fallback: T): Promise<T> {
  try {
    const w = window as any;
    if (w?.storage?.get) {
      const res = await w.storage.get(key);
      return res ? (JSON.parse(res.value) as T) : fallback;
    }
  } catch { /* */ }
  try {
    const val = await idbGet<T>(key);
    return val !== undefined ? val : fallback;
  } catch {
    try {
      const raw = localStorage.getItem(key);
      return raw ? (JSON.parse(raw) as T) : fallback;
    } catch { return fallback; }
  }
}

async function storageSet(key: string, value: unknown): Promise<void> {
  try {
    const w = window as any;
    if (w?.storage?.set) { await w.storage.set(key, JSON.stringify(value)); return; }
  } catch (e) { console.error("storage.set فشل:", e); }
  try {
    await idbSet(key, value);
    return;
  } catch (e) {
    console.error("IndexedDB فشل، محاولة localStorage:", e);
  }
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { console.error("فشل الحفظ:", e); }
}

// ─── Name matching (نفس خوارزمية المنصة الأولى) ──────────────────────────────
const AR_EN: Record<string, string> = {
  "ا":"a","أ":"a","إ":"a","آ":"a","ى":"a","ة":"h","ب":"b","ت":"t","ث":"th",
  "ج":"j","ح":"h","خ":"kh","د":"d","ذ":"th","ر":"r","ز":"z","س":"s","ش":"sh",
  "ص":"s","ض":"d","ط":"t","ظ":"z","ع":"a","غ":"g","ف":"f","ق":"q","ك":"k",
  "ل":"l","م":"m","ن":"n","ه":"h","و":"w","ي":"i","ئ":"i","ء":"","ّ":"",
  "َ":"","ُ":"","ِ":"","ْ":"","ً":"","ٌ":"","ٍ":"",
};
function normName(s: string) {
  const c = s.toLowerCase()
    .replace(/اجمد/g,"احمد")
    .replace(/تحويل\s+الكتروني\s+موبايل:/g,"")
    .replace(/الدفع\s+لصديق\s+من/g,"")
    .replace(/تحويل\s+اي-براق/g,"")
    .replace(/wallet/g,"").replace(/خبز/g,"").trim();
  let t = c.split("").map(ch => AR_EN[ch] ?? ch).join("");
  t = t.replace(/gh/g,"g").replace(/ou/g,"o").replace(/oo/g,"o").replace(/w/g,"o")
       .replace(/ee/g,"i").replace(/y/g,"i").replace(/h$/g,"a")
       .replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
  return t;
}
function levenshtein(a: string, b: string): number {
  const dp = Array.from({ length: a.length+1 }, (_, i) =>
    Array.from({ length: b.length+1 }, (_, j) => i===0?j:j===0?i:0));
  for (let i=1;i<=a.length;i++)
    for (let j=1;j<=b.length;j++)
      dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j-1],dp[i][j-1],dp[i-1][j]);
  return dp[a.length][b.length];
}
function nameSim(a: string, b: string): number {
  const na=normName(a), nb=normName(b);
  if (!na||!nb) return 0;
  if (na===nb) return 1;
  return 1-levenshtein(na,nb)/Math.max(na.length,nb.length);
}
function wordTypoMatch(a: string, b: string): boolean {
  if (!a || !b) return false;
  if (a === b) return true;
  const maxLen = Math.max(a.length, b.length);
  const d = levenshtein(a, b);
  const threshold = maxLen <= 4 ? 1 : maxLen <= 7 ? 2 : Math.ceil(maxLen * 0.3);
  return d <= threshold;
}
function nameTokens(s: string): string[] {
  return normName(s).split(" ").filter(t => t.length > 1 && t !== "al");
}
function advancedMatchCheck(cashierName: string, bankDesc: string): {
  isMatch: boolean; isApprox: boolean; matchType: string;
} {
  const cT = nameTokens(cashierName);
  const bT = nameTokens(bankDesc);
  if (!cT.length || !bT.length) return { isMatch: false, isApprox: false, matchType: "none" };
  const squished = (arr: string[]) => arr.slice().sort().join("");
  if (squished(cT) === squished(bT)) return { isMatch: true, isApprox: false, matchType: "exact" };
  if (cT.length >= 2 && bT.length >= 2) {
    if (wordTypoMatch(cT[0], bT[0]) && wordTypoMatch(cT[1], bT[1])) {
      return { isMatch: true, isApprox: !(cT[0]===bT[0] && cT[1]===bT[1]), matchType: "firstSecond" };
    }
  }
  if (cT.length >= 4 && bT.length >= 4) {
    if (wordTypoMatch(cT[3], bT[3])) return { isMatch: true, isApprox: cT[3]!==bT[3], matchType: "fourthName" };
  }
  if (cT.length === 1 || bT.length === 1) {
    const shortSide = cT.length === 1 ? cT : bT;
    const longSide = cT.length === 1 ? bT : cT;
    if (wordTypoMatch(shortSide[0], longSide[0])) return { isMatch: true, isApprox: shortSide[0]!==longSide[0], matchType: "typo" };
  }
  for (const ct of cT) {
    for (const bt of bT) {
      if (wordTypoMatch(ct, bt)) return { isMatch: true, isApprox: true, matchType: "typo" };
    }
  }
  return { isMatch: false, isApprox: false, matchType: "none" };
}

// ─── Visa number detection ────────────────────────────────────────────────────
const VISA_NUM_RE = new RegExp("[/-]\\s*(\\d{4})(?!\\d)");
function extractVisaNumber(rawName: string): string | null {
  const m = rawName.match(VISA_NUM_RE);
  return m ? m[1] : null;
}

// ─── Auto-detect columns ──────────────────────────────────────────────────────
function autoDetect(headers: string[], hints: string[]): string {
  const scores=headers.map(h=>({h,s:hints.reduce((a,hint)=>a+(h.toLowerCase().includes(hint.toLowerCase())?1:0),0)})).sort((a,b)=>b.s-a.s);
  return scores[0]?.s>0?scores[0].h:"";
}
const HINTS = {
  date:   ["date","تاريخ","value","posting"],
  desc:   ["description","narrative","detail","إيضاح","بيان","وصف","narr","payee","particular"],
  debit:  ["debit","مدين","سحب","paid","withdrawal","dr","مدفوع","خروج","مبالغ مدفوعة"],
  credit: ["credit","دائن","إيداع","received","deposit","cr","دخول","مستلم","مبالغ مستلمة"],
  name:   ["name","اسم","customer","client","employee","موظف","عميل","الزبون","بيان","وصف","narrative","detail","إيضاح"],
  accountType: ["account type","account_type","نوع الحساب","نوع حساب","نوع","حساب","account"],
  authNum: ["تفويض","authorization","auth","رقم","number","ref","مرجع"],
};
const CARD_PAID_HINTS = ["مدفوع بالكارد", "مدفوع بالكرت", "مدفوع كارد", "دفع بالكارد", "دفع بالكرت", "card paid", "card payment", "card amount"];

function detectCardPaidColumn(headers: string[]): string {
  const exact = headers.find(header => CARD_PAID_HINTS.some(hint => header.trim().toLowerCase() === hint.toLowerCase()));
  return exact || autoDetect(headers, CARD_PAID_HINTS);
}

// ─── Types ────────────────────────────────────────────────────────────────────
interface BankRow2 {
  id: number;
  date: string;
  description: string;
  rawAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  isVisaBank: boolean; // مشتريات عمولة تجار
  orig: Record<string, unknown>;
}

interface AzaRow {
  id: number;
  rawName: string;
  name: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  accountType: string; // محفظة / تطبيق / فيزا
  date: string;
  visaNumber: string | null;
  orig: Record<string, unknown>;
}

interface SonyVisaRow {
  id: number;
  rawName: string;
  name: string;
  amount: number;
  visaNumber: string | null;
  authNumber: string | null; // رقم التفويض
  date: string;
  orig: Record<string, unknown>;
}

interface MatchedVisa {
  id: string;
  sonyId: number;
  azaId: number;
  sonyName: string;
  azaName: string;
  visaNumber: string;
  authNumber: string | null;
  amount: number;
  sonyRawName: string;
  azaRawName: string;
  matchedAt: string;
  note?: string;
}

interface PendingVisaBank {
  id: string;
  bankId: number;
  bankDesc: string;
  authNumber: string | null;
  amount: number;
  date: string;
  movedAt: string;
}

type R2Tab = "overview" | "bank" | "aza" | "sony" | "auth" | "matched" | "pendingBank" | "unmatched";

// ─── DropZone ─────────────────────────────────────────────────────────────────
function DropZone({ file, onFile, onClear }: { file:File|null; onFile:(f:File)=>void; onClear:()=>void }) {
  const handleDrop=useCallback((e:React.DragEvent)=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)onFile(f);},[onFile]);
  if (file) return (
    <div className="flex items-center justify-between gap-3 p-3.5 bg-blue-500/5 border border-blue-500/25 rounded-lg">
      <div className="flex items-center gap-2.5">
        <FileSpreadsheet className="w-4 h-4 text-blue-600 shrink-0"/>
        <div><div className="text-sm font-medium">{file.name}</div><div className="text-xs text-muted-foreground">{(file.size/1024).toFixed(1)} KB</div></div>
      </div>
      <button onClick={onClear} className="p-1 hover:bg-muted rounded transition-colors text-muted-foreground"><X className="w-4 h-4"/></button>
    </div>
  );
  return (
    <label onDrop={handleDrop} onDragOver={e=>e.preventDefault()}
      className="flex flex-col items-center gap-2 p-6 border-2 border-dashed border-border rounded-lg cursor-pointer hover:border-blue-500/50 hover:bg-muted/10 transition-all">
      <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)onFile(f);}}/>
      <Upload className="w-6 h-6 text-muted-foreground"/>
      <span className="text-sm">اسحب الملف أو انقر هنا</span>
      <span className="text-xs text-muted-foreground">xlsx · xls · csv</span>
    </label>
  );
}

function Sel({ label, headers, value, onChange }: { label:string; headers:string[]; value:string; onChange:(v:string)=>void; }) {
  return (
    <div className="space-y-1">
      <label className="text-xs font-medium text-muted-foreground">{label}</label>
      <div className="relative">
        <select value={value} onChange={e=>onChange(e.target.value)}
          className="w-full appearance-none bg-background border border-border rounded-md px-3 py-2 text-sm pr-7 focus:outline-none">
          <option value="">— اختر العمود —</option>
          {headers.map(h=><option key={h} value={h}>{h}</option>)}
        </select>
        <ChevronDown className="absolute left-2 top-1/2 -translate-y-1/2 w-3 h-3 text-muted-foreground pointer-events-none"/>
      </div>
    </div>
  );
}

export interface StageAResult {
  id: number;
  source: string;
  name: string;
  originalName: string;
  amount: number;
  databaseName: string;
  databaseAmount: number;
  userId: string;
  registrationTime: string;
  customerNumber: string;
  invoiceNumber: string;
  accountType: string;
  invoice: Record<string, unknown>;
  matched: boolean;
  issue: string;
  category: "بنك فلسطين" | "محفظة تجارية" | "محفظة محمود" | "محفظة" | "جوال بي" | "فيزا";
  categoryIssue: string;
}

function cleanStageAName(value: string): string {
  return value.replace(/^\s*بطاقة\s+ائتمان\s*\/\s*/i, "").trim();
}

function stageANameKey(value: string): string {
  return normName(cleanStageAName(value)).split(" ").filter(token => /^[a-z]+$/.test(token)).sort().join(" ");
}

function classifyStageAInvoice(originalName: string, source: string): { category: "بنك فلسطين" | "محفظة تجارية" | "محفظة محمود" | "محفظة" | "جوال بي" | "فيزا"; categoryIssue: string } {
  const hasVisaNumber = /(?:^|[^\d])\d{4}(?:$|[^\d])/.test(originalName);
  const normalized = originalName.toLowerCase().replace(/[ـ_-]/g, " ").replace(/\s+/g, " ").trim();
  const jawwalPay = /جوال\s*بي+|jawwal\s*pay|jawwalpay/i.test(normalized);
  const walletWords = ["محفظة", "wallet"];
  const isWallet = walletWords.some(word => normalized.includes(word));
  if (hasVisaNumber) {
    return { category: "فيزا", categoryIssue: isWallet || source === "محفظة تجارية" ? "مسجلة في ملف المحفظة، لكن رقم 4 أرقام يدل على فيزا" : "" };
  }
  if (jawwalPay) return { category: "جوال بي", categoryIssue: "" };
  if (normalized.includes("محمود") && isWallet) return { category: "محفظة محمود", categoryIssue: "" };
  if (isWallet) return { category: "محفظة", categoryIssue: "" };
  if (source === "بنك فلسطين") return { category: "بنك فلسطين", categoryIssue: "" };
  if (source === "محفظة تجارية") return { category: "محفظة تجارية", categoryIssue: "" };
  return { category: "محفظة", categoryIssue: "لم يظهر رقم فيزا أو اسم محفظة واضح؛ تم تصنيفها كمحفظة" };
}

function stageAVisaNumber(value: string): string {
  return value.match(/(?:^|\D)(\d{4})(?:\D|$)/)?.[1] || "";
}

function StageAPlatform({ onBack, onTransferToB, onTransferVisaToC }: { onBack: () => void; onTransferToB: (results: StageAResult[]) => void; onTransferVisaToC: (results: StageAResult[]) => void }) {
  const [databaseFile, setDatabaseFile] = useState<File | null>(null);
  const [databaseHeaders, setDatabaseHeaders] = useState<string[]>([]);
  const [databaseRows, setDatabaseRows] = useState<Record<string, unknown>[]>([]);
  const [databaseMap, setDatabaseMap] = useState({ name: "", amount: "", userId: "", registrationTime: "", customerNumber: "", invoiceNumber: "" });
  const [invoiceFiles, setInvoiceFiles] = useState<Record<string, File | null>>({ bank: null, wallet: null, visa: null });
  const [invoiceHeaders, setInvoiceHeaders] = useState<Record<string, string[]>>({ bank: [], wallet: [], visa: [] });
  const [invoiceRows, setInvoiceRows] = useState<Record<string, Record<string, unknown>[]>>({ bank: [], wallet: [], visa: [] });
  const [invoiceMaps, setInvoiceMaps] = useState<Record<string, { name: string; paidAmount: string; receivedAmount: string; accountType: string; date: string; reference: string }>>({
    bank: { name: "", paidAmount: "", receivedAmount: "", accountType: "", date: "", reference: "" }, wallet: { name: "", paidAmount: "", receivedAmount: "", accountType: "", date: "", reference: "" }, visa: { name: "", paidAmount: "", receivedAmount: "", accountType: "", date: "", reference: "" }
  });
  const [invoiceSwaps, setInvoiceSwaps] = useState<Record<string, boolean>>({ bank: false, wallet: false, visa: false });
  const [results, setResults] = useState<StageAResult[]>([]);
  const [resultFilter, setResultFilter] = useState<"all" | "matched" | "unmatched" | "wallet" | "جوال بي" | "visa">("all");
  const [selectedResultIds, setSelectedResultIds] = useState<Set<string>>(new Set());
  const [resultNameSearch, setResultNameSearch] = useState("");
  const [resultAmountSearch, setResultAmountSearch] = useState("");
  const [stageASourceFilter, setStageASourceFilter] = useState("all");
  const [matchedSearch, setMatchedSearch] = useState("");
  const [selectedMatchedIds, setSelectedMatchedIds] = useState<Set<string>>(new Set());
  const [replacementSearch, setReplacementSearch] = useState("");
  const [stageAPage, setStageAPage] = useState(1);
  const [error, setError] = useState<string | null>(null);

  const loadDatabase = async (file: File) => {
    setDatabaseFile(file); setError(null);
    try {
      const parsed = parseSheet(await readFileBuf(file));
      setDatabaseHeaders(parsed.headers); setDatabaseRows(parsed.rows);
      setDatabaseMap({
        name: autoDetect(parsed.headers, ["اسم الزبون", "اسم العميل", "customer", "client", "name", "اسم"]),
        amount: autoDetect(parsed.headers, ["مدفوع بالكارد", "مدفوع بالكرت", "card paid", "card payment"]) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        userId: autoDetect(parsed.headers, ["user", "userid", "user id", "مستخدم", "رقم المستخدم", "كاشير"]),
        registrationTime: autoDetect(parsed.headers, ["الساعة", "وقت التسجيل", "وقت الفاتورة", "time", "timestamp"]),
        customerNumber: autoDetect(parsed.headers, ["رقم الزبون", "رقم العميل", "customer number", "customer id", "client id"]),
        invoiceNumber: autoDetect(parsed.headers, ["رقم الفاتورة", "رقم الفاتوره", "invoice number", "invoice no", "invoice id"])
      });
    } catch (e) { setError((e as Error).message); }
  };

  const loadInvoices = async (source: string, file: File) => {
    setInvoiceFiles(prev => ({ ...prev, [source]: file })); setError(null);
    try {
      const parsed = parseSheet(await readFileBuf(file));
      setInvoiceHeaders(prev => ({ ...prev, [source]: parsed.headers }));
      setInvoiceRows(prev => ({ ...prev, [source]: parsed.rows }));
      setInvoiceMaps(prev => ({ ...prev, [source]: {
        name: autoDetect(parsed.headers, HINTS.name),
        paidAmount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        receivedAmount: autoDetect(parsed.headers, HINTS.credit),
        accountType: autoDetect(parsed.headers, HINTS.accountType),
        date: autoDetect(parsed.headers, HINTS.date),
        reference: autoDetect(parsed.headers, HINTS.authNum)
      }}));
    } catch (e) { setError((e as Error).message); }
  };

  const runStageA = () => {
    if (!databaseRows.length || !databaseMap.name || !databaseMap.amount) { setError("ارفع ملف قاعدة البيانات وحدد عمود الاسم والمبلغ أولاً"); return; }
    const available = databaseRows.map((row, index) => ({ row, index, used: false }));
    const parseStageAAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
    const output: StageAResult[] = [];
    (["bank", "wallet", "visa"] as const).forEach(source => {
      const map = invoiceMaps[source];
      invoiceRows[source].forEach((invoice, index) => {
        const originalName = String(map.name ? invoice[map.name] : "").trim();
        const name = cleanStageAName(originalName);
        const paidAmount = parseStageAAmount(map.paidAmount ? invoice[map.paidAmount] : 0);
        const receivedAmount = parseStageAAmount(map.receivedAmount ? invoice[map.receivedAmount] : 0);
        const amount = invoiceSwaps[source] ? receivedAmount || paidAmount : paidAmount || receivedAmount;
        if (!name && !amount) return;
        const sameName = available.filter(item => !item.used && stageANameKey(String(item.row[databaseMap.name])) === stageANameKey(name));
        const match = sameName.find(item => Math.abs(parseStageAAmount(item.row[databaseMap.amount]) - amount) <= 0.01);
        if (match) match.used = true;
        // لا نعرض سجل قاعدة البيانات كأنه تطابق إلا إذا تطابق الاسم والمبلغ فعلاً.
        // المبالغ الأخرى لنفس الاسم تظهر في رسالة المشكلة فقط.
        const databaseCandidate = match;
        const issue = match ? "" : sameName.length ? `الاسم موجود، لكن المبلغ مختلف (المطلوب: ${fmtNum(amount)}؛ المتاح: ${sameName.map(item => fmtNum(parseStageAAmount(item.row[databaseMap.amount]))).join("، ")})` : "الاسم غير موجود في قاعدة البيانات";
        const sourceLabel = source === "bank" ? "بنك فلسطين" : source === "wallet" ? "محفظة تجارية" : "فيزا";
        const classification = classifyStageAInvoice(originalName, sourceLabel);
        output.push({
          id: index, source: sourceLabel, name, originalName, amount,
          databaseName: databaseCandidate ? String(databaseCandidate.row[databaseMap.name] ?? "") : "",
          databaseAmount: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.amount]) : 0,
          userId: databaseCandidate ? String(databaseCandidate.row[databaseMap.userId] ?? "") : "",
          registrationTime: databaseCandidate ? String(databaseCandidate.row[databaseMap.registrationTime] ?? "") : "",
          customerNumber: databaseCandidate ? String(databaseCandidate.row[databaseMap.customerNumber] ?? "") : "",
          invoiceNumber: databaseCandidate ? String(databaseCandidate.row[databaseMap.invoiceNumber] ?? "") : "",
          accountType: String(map.accountType ? invoice[map.accountType] : "").trim(), invoice, matched: !!match, issue, ...classification
        });
      });
    });
    setResults(output);
    setError(null);
  };

  const exportStageA = () => {
    if (!results.length) return;
    const workbook = XLSX.utils.book_new();
    const rows = results.map(result => [result.source, result.category, result.originalName, result.databaseName, result.amount, result.databaseAmount || "", result.userId, result.registrationTime, result.customerNumber, result.invoiceNumber, result.accountType, result.matched ? "مطابق" : "غير مطابق", result.categoryIssue || result.issue]);
    const worksheet = XLSX.utils.aoa_to_sheet([["المصدر الأصلي", "التصنيف", "البيان الأصلي", "البيان من قاعدة البيانات", "المبلغ", "مبلغ قاعدة البيانات", "رقم المستخدم", "الساعة", "رقم الزبون", "رقم الفاتورة", "نوع الحساب من ملف الفواتير", "حالة المطابقة", "المشكلة / التصحيح المطلوب"], ...rows]);
    worksheet["!cols"] = [{ wch: 18 }, { wch: 14 }, { wch: 32 }, { wch: 28 }, { wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 16 }, { wch: 48 }];
    XLSX.utils.book_append_sheet(workbook, worksheet, "نتائج المرحلة A");
    XLSX.writeFile(workbook, "نتائج_المرحلة_A_ومشاكل_اخرى.xlsx");
  };

  const resultKey = (result: StageAResult) => `${result.source}-${result.id}`;
  const toggleResultSelection = (result: StageAResult) => setSelectedResultIds(previous => {
    const next = new Set(previous); const key = resultKey(result);
    if (next.has(key)) next.delete(key); else next.add(key);
    return next;
  });
  const updateResultStatus = (result: StageAResult, matched: boolean) => {
    setResults(previous => previous.map(item => item === result ? { ...item, matched, issue: matched ? "تم اعتمادها يدوياً" : "تم رفضها يدوياً" } : item));
  };
  const rejectSelectedResults = () => {
    setResults(previous => previous.map(item => selectedResultIds.has(resultKey(item)) ? { ...item, matched: false, issue: "تم رفضها يدوياً" } : item));
    setSelectedResultIds(new Set());
  };

  const allActionResults = results.filter(result => {
    const nameQuery = resultNameSearch.trim().toLowerCase();
    const amountQuery = resultAmountSearch.trim();
    return (!nameQuery || `${result.originalName} ${result.databaseName} ${JSON.stringify(result.invoice)}`.toLowerCase().includes(nameQuery)) &&
      (stageASourceFilter === "all" || result.source === stageASourceFilter || result.category === stageASourceFilter) &&
      (!amountQuery || String(result.amount).includes(amountQuery) || String(result.databaseAmount).includes(amountQuery));
  });
  const pagedActionResults = allActionResults.slice((Math.min(stageAPage, Math.max(1, Math.ceil(allActionResults.length / 50))) - 1) * 50, Math.min(stageAPage, Math.max(1, Math.ceil(allActionResults.length / 50))) * 50);
    const stageAFilterBar = results.length > 0 && <div className="flex flex-row-reverse flex-wrap items-center gap-3 rounded-xl border border-slate-200 bg-white p-3"><input value={resultNameSearch} onChange={event => { setResultNameSearch(event.target.value); setStageAPage(1); }} placeholder="بحث برقم الفاتورة أو العميل أو المرجع..." className="order-1 min-w-[240px] flex-1 rounded-lg border border-slate-200 px-4 py-3 text-sm outline-none focus:border-blue-400 focus:ring-2 focus:ring-blue-100" /><select value={stageASourceFilter} onChange={event => { setStageASourceFilter(event.target.value); setStageAPage(1); }} className="order-2 rounded-lg border border-slate-200 bg-white px-4 py-3 text-sm outline-none"><option value="all">كل الوسائل</option><option value="بنك فلسطين">بنك فلسطين</option><option value="محفظة تجارية">محفظة تجارية</option><option value="محفظة محمود">محفظة محمود</option><option value="جوال بي">جوال بي</option><option value="فيزا">فيزا</option></select><button onClick={() => setSelectedResultIds(new Set(visibleActionResults.map(result => resultKey(result))))} className="order-3 rounded-lg border border-slate-200 bg-white px-5 py-3 text-sm font-medium hover:bg-slate-50">تحديد الكل</button></div>;
    const inlineResultActions = results.length > 0 && <div className="space-y-2 rounded-xl border border-slate-200 bg-white p-3"><div className="flex items-center justify-between"><strong className="text-xs">خيارات الفواتير</strong><span className="text-xs text-slate-500">حدد الفاتورة لإظهار خياراتها</span></div>{pagedActionResults.map(result => <div key={`inline-${resultKey(result)}`} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-100 bg-slate-50 p-2"><input type="checkbox" checked={selectedResultIds.has(resultKey(result))} onChange={() => toggleResultSelection(result)} /><span className="min-w-48 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span>{selectedResultIds.has(resultKey(result)) && <><button onClick={() => updateResultStatus(result, true)} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">اعتماد</button><button onClick={() => updateResultStatus(result, false)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button><select onChange={event => replaceMatchedDatabase(result, event.target.value)} defaultValue="-1" className="max-w-40 rounded border bg-white px-2 py-1 text-[10px]"><option value="-1">تغيير إلى شخص آخر</option>{databaseRows.map((row, index) => <option key={index} value={index}>{String(row[databaseMap.name] ?? "")} · {String(row[databaseMap.amount] ?? "")}</option>)}</select><button onClick={() => { if (window.confirm("حذف الفاتورة وسجلها من النتائج؟")) deleteMatchedPair(result); }} className="rounded border border-slate-300 bg-slate-100 px-2 py-1 text-[10px] text-slate-700">حذف الاثنين</button></>}</div>)}</div>;
  const stageAPageSize = 20;
  const totalStageAPages = Math.max(1, Math.ceil(allActionResults.length / stageAPageSize));
  const currentStageAPage = Math.min(stageAPage, totalStageAPages);
  const visibleActionResults = allActionResults.slice((currentStageAPage - 1) * stageAPageSize, currentStageAPage * stageAPageSize);
  const paginationControls = results.length > stageAPageSize && <div className="flex items-center justify-center gap-2 rounded-xl border border-slate-200 bg-white p-3 text-xs"><button onClick={() => setStageAPage(page => Math.max(1, page - 1))} disabled={currentStageAPage === 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button><span className="text-slate-500">صفحة {currentStageAPage} من {totalStageAPages} · عرض {visibleActionResults.length} من {allActionResults.length}</span><button onClick={() => setStageAPage(page => Math.min(totalStageAPages, page + 1))} disabled={currentStageAPage === totalStageAPages} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button></div>;
  const visibleMatchedResults = results.filter(result => {
    if (!result.matched) return false;
    const query = matchedSearch.trim().toLowerCase();
    if (!query) return true;
    return `${result.originalName} ${result.databaseName}`.toLowerCase().includes(query) || String(result.amount).includes(query) || String(result.databaseAmount).includes(query);
  });
  const rejectMatched = (result: StageAResult) => updateResultStatus(result, false);
  const deleteMatchedPair = (result: StageAResult) => setResults(previous => previous.filter(item => item !== result));
  const replaceMatchedDatabase = (result: StageAResult, value: string) => {
    const row = databaseRows[Number(value)];
    if (!row) return;
    setResults(previous => previous.map(item => item === result ? {
      ...item,
      databaseName: String(row[databaseMap.name] ?? ""),
      databaseAmount: Number.parseFloat(String(row[databaseMap.amount] ?? "").replace(/,/g, "")) || 0,
      userId: String(row[databaseMap.userId] ?? ""),
      registrationTime: String(row[databaseMap.registrationTime] ?? ""),
      customerNumber: String(row[databaseMap.customerNumber] ?? ""),
      invoiceNumber: String(row[databaseMap.invoiceNumber] ?? ""),
      matched: true, issue: "تم تعديل سجل قاعدة البيانات يدوياً"
    } : item));
  };

  const replacementRows = databaseRows.filter(row => {
    const query = replacementSearch.trim().toLowerCase();
    return !query || `${String(row[databaseMap.name] ?? "")} ${String(row[databaseMap.amount] ?? "")}`.toLowerCase().includes(query);
  });

  const matchedActionsPanel = results.length > 0 && <div className="rounded-xl border border-green-200 bg-green-50/40 p-3"><div className="mb-2 flex flex-wrap items-center gap-2"><strong className="text-xs text-green-800">المطابقات وإجراءاتها</strong><input value={matchedSearch} onChange={event => setMatchedSearch(event.target.value)} placeholder="بحث المطابقات بالاسم أو المبلغ" className="min-w-56 flex-1 rounded border bg-white px-2 py-1.5 text-[11px]" /><span className="text-xs text-green-700">{visibleMatchedResults.length} مطابقة</span></div><div className="space-y-2">{visibleMatchedResults.map(result => <div key={`matched-${resultKey(result)}`} className="flex flex-wrap items-center gap-2 rounded-lg border border-green-100 bg-white p-2"><input type="checkbox" checked={selectedMatchedIds.has(resultKey(result))} onChange={() => setSelectedMatchedIds(previous => { const next = new Set(previous); const key = resultKey(result); if (next.has(key)) next.delete(key); else next.add(key); return next; })} /><span className="min-w-44 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span><span className="text-[11px] text-green-700">↔ {result.databaseName} ({fmtNum(result.databaseAmount)})</span><select value={databaseRows.findIndex(row => String(row[databaseMap.name] ?? "") === result.databaseName && Number.parseFloat(String(row[databaseMap.amount] ?? "").replace(/,/g, "")) === result.databaseAmount)} onChange={event => replaceMatchedDatabase(result, event.target.value)} className="max-w-36 rounded border px-2 py-1 text-[10px]"><option value="-1">اختيار شخص آخر</option>{databaseRows.map((row, index) => <option key={index} value={index}>{String(row[databaseMap.name] ?? "")} · {String(row[databaseMap.amount] ?? "")}</option>)}</select><button onClick={() => rejectMatched(result)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">✕ رفض</button><button onClick={() => deleteMatchedPair(result)} className="rounded border border-slate-300 bg-slate-100 px-2 py-1 text-[10px] text-slate-700">حذف الاثنين</button></div>)}</div></div>;

  const selectionPanel = results.length > 0 && <div className="rounded-xl border border-slate-200 bg-white p-3"><div className="mb-2 flex flex-wrap items-center gap-2"><strong className="text-xs">إجراءات الفواتير</strong><span className="text-xs text-slate-500">محدد: {selectedResultIds.size}</span><input value={resultNameSearch} onChange={event => setResultNameSearch(event.target.value)} placeholder="بحث بالاسم أو البيان" className="min-w-48 flex-1 rounded border px-2 py-1.5 text-[11px]" /><input value={resultAmountSearch} onChange={event => setResultAmountSearch(event.target.value)} placeholder="بحث بالمبلغ" inputMode="decimal" className="w-28 rounded border px-2 py-1.5 text-[11px]" /></div><div className="space-y-2">{visibleActionResults.map(result => <div key={resultKey(result)} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-100 bg-slate-50 p-2"><input type="checkbox" checked={selectedResultIds.has(resultKey(result))} onChange={() => toggleResultSelection(result)} /><span className="min-w-48 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span><button onClick={() => updateResultStatus(result, true)} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">اعتماد</button><button onClick={() => updateResultStatus(result, false)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button><button onClick={() => { onTransferToB([result]); setSelectedResultIds(previous => { const next = new Set(previous); next.delete(resultKey(result)); return next; }); }} disabled={result.category === "فيزا"} className="rounded bg-emerald-600 px-2 py-1 text-[10px] text-white disabled:opacity-40">نقل B</button><button onClick={() => { onTransferVisaToC([result]); setSelectedResultIds(previous => { const next = new Set(previous); next.delete(resultKey(result)); return next; }); }} disabled={result.category !== "فيزا"} className="rounded bg-amber-600 px-2 py-1 text-[10px] text-white disabled:opacity-40">نقل C</button></div>)}</div></div>;

  const renderInvoice = (source: "bank" | "wallet" | "visa", label: string) => (
    <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4 shadow-[0_4px_12px_rgba(30,64,175,0.08)]">
      <h3 className="text-sm font-bold text-slate-800">{label}</h3>
      {!!invoiceHeaders[source].length && (
        <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
          <input type="checkbox" checked={invoiceSwaps[source]} onChange={event => setInvoiceSwaps(prev => ({ ...prev, [source]: event.target.checked }))} className="rounded" />
          عكس المدين والدائن
        </label>
      )}
      <DropZone file={invoiceFiles[source]} onFile={file => loadInvoices(source, file)} onClear={() => { setInvoiceFiles(prev => ({ ...prev, [source]: null })); setInvoiceHeaders(prev => ({ ...prev, [source]: [] })); setInvoiceRows(prev => ({ ...prev, [source]: [] })); }} />
      {!!invoiceHeaders[source].length && <>
        <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
          <Sel label="التاريخ" headers={invoiceHeaders[source]} value={invoiceMaps[source].date} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], date: value } }))} />
          <Sel label="البيان / اسم الزبون" headers={invoiceHeaders[source]} value={invoiceMaps[source].name} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], name: value } }))} />
          <Sel label="المبالغ المدفوعة (Debit)" headers={invoiceHeaders[source]} value={invoiceMaps[source].paidAmount} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], paidAmount: value } }))} />
          <Sel label="المبالغ المستلمة (Credit)" headers={invoiceHeaders[source]} value={invoiceMaps[source].receivedAmount} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], receivedAmount: value } }))} />
          <Sel label="نوع الحساب" headers={invoiceHeaders[source]} value={invoiceMaps[source].accountType} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], accountType: value } }))} />
          <Sel label="المرجع / رقم الفاتورة" headers={invoiceHeaders[source]} value={invoiceMaps[source].reference} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], reference: value } }))} />
        </div>
      </>}
    </div>
  );

  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-6xl space-y-5">
    <div className="flex flex-wrap items-center justify-between gap-3"><button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة للمنصة الرئيسية</button><div className="flex items-center gap-3"><span className="flex h-12 w-12 items-center justify-center rounded-xl bg-blue-600 text-2xl font-black text-white shadow-sm">A</span><div><h1 className="text-2xl font-black text-blue-700">فرز الفواتير حسب قاعدة البيانات</h1><p className="mt-1 text-sm text-slate-500">مطابقة الاسم والمبلغ مع قاعدة البيانات، ثم استخراج المطابقات ومشاكل أخرى.</p></div></div><div className="flex gap-2"><button onClick={runStageA} className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"><Sparkles className="h-4 w-4" />تشغيل الفرز</button>{results.length > 0 && <button onClick={exportStageA} className="flex items-center gap-2 rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white hover:bg-green-700"><Download className="h-4 w-4" />تصدير Excel</button>}</div></div>
    <div className="rounded-xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900">تتم المطابقة بالاسم والمبلغ معاً. كل صف في قاعدة البيانات يُستهلك مرة واحدة، لذلك يمكن معالجة فاتورتين للاسم نفسه بمبلغين مختلفين بشكل صحيح. تتم إزالة عبارة <strong>بطاقة ائتمان /</strong> فقط، وتبقى النجمة والشرطة والرقم كما هي.</div>
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}{stageAFilterBar}
    <div className="space-y-3 rounded-2xl border border-blue-100 bg-white p-5 shadow-sm"><div className="flex items-center justify-between"><h2 className="text-base font-bold text-slate-800">ملف قاعدة البيانات</h2><span className="rounded-full bg-blue-50 px-3 py-1 text-xs font-semibold text-blue-700">المصدر الأساسي للفرز</span></div><DropZone file={databaseFile} onFile={loadDatabase} onClear={() => { setDatabaseFile(null); setDatabaseHeaders([]); setDatabaseRows([]); }} />{!!databaseHeaders.length && <div className="grid grid-cols-1 gap-2 sm:grid-cols-3"><Sel label="اسم الزبون (البيان من قاعدة البيانات)" headers={databaseHeaders} value={databaseMap.name} onChange={value => setDatabaseMap(prev => ({ ...prev, name: value }))} /><Sel label="مدفوع بالكارد (مبلغ قاعدة البيانات)" headers={databaseHeaders} value={databaseMap.amount} onChange={value => setDatabaseMap(prev => ({ ...prev, amount: value }))} /><Sel label="رقم المستخدم" headers={databaseHeaders} value={databaseMap.userId} onChange={value => setDatabaseMap(prev => ({ ...prev, userId: value }))} /><Sel label="الساعة" headers={databaseHeaders} value={databaseMap.registrationTime} onChange={value => setDatabaseMap(prev => ({ ...prev, registrationTime: value }))} /><Sel label="رقم الزبون" headers={databaseHeaders} value={databaseMap.customerNumber} onChange={value => setDatabaseMap(prev => ({ ...prev, customerNumber: value }))} /><Sel label="رقم الفاتورة" headers={databaseHeaders} value={databaseMap.invoiceNumber} onChange={value => setDatabaseMap(prev => ({ ...prev, invoiceNumber: value }))} /></div>}</div>
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">{renderInvoice("bank", "ملف فواتير بنك فلسطين")}{renderInvoice("wallet", "ملف فواتير المحفظة التجارية")}{renderInvoice("visa", "ملف فواتير الفيزا")}</div>
    {inlineResultActions}{paginationControls}
    {!!results.length && <div className="overflow-hidden rounded-xl border border-slate-200 bg-white"><div className="flex flex-wrap items-center justify-between gap-3 border-b bg-slate-50 px-4 py-3 text-sm"><div className="flex flex-wrap items-center gap-1"><strong className="ml-3">نتائج الفرز</strong>{([ ["all", "الكل"], ["matched", "المطابق"], ["unmatched", "غير المطابق"], ["wallet", "محفظة"], ["visa", "فيزا"] ] as const).map(([value, label]) => <button key={value} onClick={() => setResultFilter(value)} className={`rounded-lg px-2.5 py-1.5 text-xs ${resultFilter === value ? "bg-blue-600 text-white" : "bg-white text-slate-600 hover:bg-blue-50"}`}>{label}</button>)}</div><div className="flex items-center gap-2"><span className="text-xs text-slate-500">مطابق: {results.filter(result => result.matched).length} · غير مطابق: {results.filter(result => !result.matched).length} · محفظة: {results.filter(result => result.category === "محفظة").length} · فيزا: {results.filter(result => result.category === "فيزا").length}</span>{results.some(result => result.matched) && <button onClick={() => onTransferToB(results.filter(result => result.matched))} className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-emerald-700">ترحيل المطابق إلى B ←</button>}{results.some(result => result.category === "فيزا") && <button onClick={() => onTransferVisaToC(results.filter(result => result.category === "فيزا"))} className="rounded-lg bg-amber-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-amber-700">ترحيل كل الفيزا إلى C ←</button>}</div></div><div className="overflow-x-auto"><table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right">المصدر الأصلي</th><th className="p-3 text-right">التصنيف</th><th className="p-3 text-right">البيان الأصلي</th><th className="p-3 text-right">اسم الزبون من قاعدة البيانات</th><th className="p-3 text-right">مبلغ الفاتورة</th><th className="p-3 text-right">مدفوع بالكارد من قاعدة البيانات</th><th className="p-3 text-right">نوع الحساب</th><th className="p-3 text-right">الحالة</th><th className="p-3 text-right">ملاحظة</th></tr></thead><tbody>{results.filter(result => resultFilter === "all" || (resultFilter === "matched" ? result.matched : resultFilter === "unmatched" ? !result.matched : result.category === (resultFilter === "wallet" ? "محفظة" : "فيزا"))).map(result => <tr key={`${result.source}-${result.id}`} className={`border-t ${result.matched ? "bg-green-50/30" : "bg-red-50/30"}`}><td className="p-3">{result.source}</td><td className="p-3 font-semibold"><span className={result.category === "فيزا" ? "text-amber-700" : "text-emerald-700"}>{result.category}</span>{result.categoryIssue && <div className="text-[10px] font-normal text-orange-700">{result.categoryIssue}</div>}</td><td className="p-3 font-medium">{result.originalName}</td><td className="p-3">{result.databaseName || "—"}</td><td className="p-3 font-mono">{fmtNum(result.amount)}</td><td className="p-3 font-mono">{result.databaseName ? fmtNum(result.databaseAmount) : "—"}</td><td className="p-3">{result.accountType || "—"}</td><td className={`p-3 font-semibold ${result.matched ? "text-green-700" : "text-red-700"}`}>{result.matched ? "مطابق" : "غير مطابق"}</td><td className="p-3 text-xs text-orange-700">{result.categoryIssue || result.issue || "—"}</td></tr>)}</tbody></table></div></div>}
  </div></div>;
}

interface StageBInvoice {
  id: string;
  name: string;
  originalName: string;
  amount: number;
  paidAmount: number;
  accountType: string;
  userId: string;
  registrationTime: string;
  customerNumber: string;
  invoiceNumber: string;
  date: string;
  source: string;
  raw: Record<string, unknown>;
}

interface StageBTransfer {
  id: string;
  name: string;
  description: string;
  amount: number;
  receivedAmount: number;
  accountType: string;
  type: "بنك فلسطين" | "محفظة تجارية";
  date: string;
  raw: Record<string, unknown>;
}

interface StageBMatch {
  id: string;
  invoice: StageBInvoice;
  transfer?: StageBTransfer;
  score: number;
  status: "pending" | "confirmed" | "rejected" | "held" | "heldInvoice" | "heldTransfer" | "visa" | "difference";
  reason: string;
}

function stageBNameParts(value: string): string[][] {
  const cleaned = value
    .replace(/pacs008|wallet|ct|ils|usd|eur/gi, " ")
    .replace(/\d+/g, " ")
    .replace(/[|_*:;,()[\]{}]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.split("/").map(part => normName(part).split(" ").filter(Boolean)).filter(Boolean);
}

function stageBNameKey(value: string): string {
  const parts = stageBNameParts(value);
  const names = parts[0] || [];
  return names.length >= 2 ? `${names[0]} ${names[names.length - 1]}` : names.join(" ");
}

function stageBNameMatch(invoiceName: string, transferName: string, aliases: Record<string, string>): number {
  const invoiceKey = stageBNameKey(invoiceName);
  const aliasKey = aliases[invoiceKey];
  const segments = stageBNameParts(transferName);
  if (aliasKey && segments.some(segment => stageBNameKey(segment.join(" ")) === aliasKey)) return 1;
  const invoiceParts = stageBNameParts(invoiceName);
  return segments.reduce((best, segment) => {
    const candidate = segment.length >= 2 ? `${segment[0]} ${segment[segment.length - 1]}` : segment.join(" ");
    return Math.max(best, nameSim(invoiceParts[0]?.join(" ") || "", candidate));
  }, 0);
}

function stageBAccountType(invoice: StageBInvoice): string {
  return invoice.accountType || (invoice.source.includes("بنك") ? "بنك فلسطين" : invoice.source.includes("محفظة") ? "محفظة تجارية" : "");
}

function stageBTransferAccountType(transfer: StageBTransfer): string {
  return transfer.accountType || transfer.type;
}

function StageBPlatform({ invoices, onBack, onTransferVisaToC }: { invoices: StageAResult[]; onBack: () => void; onTransferVisaToC: (results: StageAResult[]) => void }) {
  const emptyMap = useMemo(() => ({ name: "", amount: "", paidAmount: "", receivedAmount: "", accountType: "", userId: "", date: "" }), []);
  const [files, setFiles] = useState<Record<string, File | null>>({ invoiceBank: null, transferBank: null, invoiceWallet: null, transferWallet: null });
  const [headers, setHeaders] = useState<Record<string, string[]>>({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
  const [rows, setRows] = useState<Record<string, Record<string, unknown>[]>>({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
  const [maps, setMaps] = useState<Record<string, typeof emptyMap>>({ invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap });
  const [matches, setMatches] = useState<StageBMatch[]>([]);
  const [aliases, setAliases] = useState<Record<string, string>>({});
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [invoiceSearch, setInvoiceSearch] = useState("");
  const [transferSearch, setTransferSearch] = useState("");
  const [drawerMatchId, setDrawerMatchId] = useState<string | null>(null);
  const [drawerInvoiceSearch, setDrawerInvoiceSearch] = useState("");
  const [drawerTransferSearch, setDrawerTransferSearch] = useState("");
  const [drawerMinAmount, setDrawerMinAmount] = useState("");
  const [drawerMaxAmount, setDrawerMaxAmount] = useState("");
  const [drawerMode, setDrawerMode] = useState<"invoice" | "transfer">("transfer");
  const [stageBView, setStageBView] = useState<"pending" | "confirmed" | "manual">("pending");
  const rowRefs = useRef<Record<string, HTMLTableRowElement | null>>({});
  const [detailMatchId, setDetailMatchId] = useState<string | null>(null);
  const undoStack = useRef<StageBMatch[][]>([]);
  const [error, setError] = useState<string | null>(null);
  const [stageBPage, setStageBPage] = useState(1);
  const [showSettings, setShowSettings] = useState(false);
  const [amountTolerancePercent, setAmountTolerancePercent] = useState(0.5);
  const [stageBAccountFilter, setStageBAccountFilter] = useState("all");
  type StageBSort = "none" | "name" | "amountAsc" | "amountDesc";
  const [invoiceSort, setInvoiceSort] = useState<StageBSort>("none");
  const [transferSort, setTransferSort] = useState<StageBSort>("none");
  const [stageBSwaps, setStageBSwaps] = useState<Record<string, boolean>>({ invoiceBank: false, transferBank: false, invoiceWallet: false, transferWallet: false });
  const sessionFileRef = useRef<HTMLInputElement>(null);
  const hasTransferredInvoices = invoices.some(invoice => invoice.category !== "فيزا");

  useEffect(() => {
    storageGet<Record<string, string>>("stage_b_aliases", {}).then(setAliases);
    storageGet<any>("stage_b_session", null).then(session => {
      if (!session) return;
      setHeaders(session.headers || { invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
      setRows(session.rows || { invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
      setMaps(session.maps || { invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap });
      setMatches(session.matches || []);
      setAmountTolerancePercent(session.amountTolerancePercent ?? 0.5);
    });
  }, [emptyMap]);

  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadFile = async (source: "invoiceBank" | "transferBank" | "invoiceWallet" | "transferWallet", file: File) => {
    try {
      const parsed = parseSheet(await readFileBuf(file));
      setFiles(prev => ({ ...prev, [source]: file }));
      setHeaders(prev => ({ ...prev, [source]: parsed.headers }));
      setRows(prev => ({ ...prev, [source]: parsed.rows }));
      setMaps(prev => ({ ...prev, [source]: {
        name: autoDetect(parsed.headers, HINTS.name),
        amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        paidAmount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, ["paid", "مدفوع", "المبلغ المدفوع"]) || autoDetect(parsed.headers, HINTS.debit),
        receivedAmount: autoDetect(parsed.headers, ["received", "مستلم", "المبلغ المستلم"]) || autoDetect(parsed.headers, HINTS.credit),
        accountType: autoDetect(parsed.headers, HINTS.accountType),
        userId: "",
        date: autoDetect(parsed.headers, HINTS.date)
      } }));
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };

  const invoiceData = useMemo<StageBInvoice[]>(() => {
    const sources = (["invoiceBank", "invoiceWallet"] as const).flatMap(source => {
      const swap = stageBSwaps[source];
      return rows[source].map((row, index) => ({
        id: `${source}-${index}`, name: String(row[maps[source].name] ?? "").trim(),
        originalName: String(row[maps[source].name] ?? "").trim(), amount: parseAmount(swap ? (row[maps[source].receivedAmount] || row[maps[source].amount]) : (row[maps[source].paidAmount] || row[maps[source].amount])),
        paidAmount: parseAmount(swap ? (row[maps[source].receivedAmount] || row[maps[source].amount]) : (row[maps[source].paidAmount] || row[maps[source].amount])),
        accountType: source === "invoiceBank" ? "بنك فلسطين" : "محفظة تجارية", userId: "",
        registrationTime: "", customerNumber: "", invoiceNumber: "",
        date: String(row[maps[source].date] ?? "").trim(), source: source === "invoiceBank" ? "فاتورة بنك فلسطين" : "فاتورة محفظة تجارية", raw: row
      })).filter(row => row.name || row.amount);
    });
    if (sources.length) return sources;
    return invoices.filter(invoice => invoice.category !== "فيزا").map((invoice, index) => ({
      id: `stage-a-${invoice.id}-${index}`, name: invoice.name, originalName: invoice.originalName,
      amount: invoice.amount, accountType: invoice.accountType, userId: invoice.userId,
      paidAmount: invoice.amount,
      registrationTime: invoice.registrationTime, customerNumber: invoice.customerNumber, invoiceNumber: invoice.invoiceNumber,
      date: "", source: invoice.category, raw: invoice.invoice
    }));
  }, [rows, maps, invoices, stageBSwaps]);

  const transferData = useMemo<StageBTransfer[]>(() => (["transferBank", "transferWallet"] as const).flatMap(source => {
    const swap = stageBSwaps[source];
    return rows[source].map((row, index) => ({
      id: `${source}-${index}`, name: String(row[maps[source].name] ?? "").trim(),
      description: String(row[maps[source].name] ?? "").trim(), amount: parseAmount(swap ? (row[maps[source].paidAmount] || row[maps[source].amount]) : (row[maps[source].receivedAmount] || row[maps[source].amount])),
      receivedAmount: parseAmount(swap ? (row[maps[source].paidAmount] || row[maps[source].amount]) : (row[maps[source].receivedAmount] || row[maps[source].amount])),
      accountType: source === "transferBank" ? "بنك فلسطين" : "محفظة تجارية",
      type: (source === "transferBank" ? "بنك فلسطين" : "محفظة تجارية") as StageBTransfer["type"],
      date: String(row[maps[source].date] ?? "").trim(), raw: row
    })).filter(row => row.name || row.amount);
  }), [rows, maps, stageBSwaps]);

  const runStageB = () => {
    if (!invoiceData.length) { setError(invoices.length ? "لا توجد فواتير مرحلة من قاعدة البيانات." : "ارفع ملفي فواتير بنك فلسطين والمحفظة التجارية."); return; }
    if (!transferData.length) { setError("ارفع حوالات بنك فلسطين وحوالات المحفظة التجارية."); return; }
    const used = new Set<string>();
    const output = invoiceData.map(invoice => {
      const candidates = transferData
        .filter(transfer => !used.has(transfer.id))
        .map(transfer => {
          const amountDiff = Math.abs(transfer.receivedAmount - invoice.paidAmount);
          const nameScore = stageBNameMatch(invoice.name, transfer.description, aliases);
          const accountDiff = stageBAccountType(invoice) && stageBTransferAccountType(transfer) && normName(stageBAccountType(invoice)) !== normName(stageBTransferAccountType(transfer));
          const score = nameScore * 70 + (amountDiff <= 0.01 ? 30 : amountDiff <= Math.max(0.01, invoice.amount * 0.005) ? 20 : 0);
          return { transfer, score, amountDiff, accountDiff };
        })
        .filter(candidate => candidate.amountDiff <= Math.max(0.01, invoice.amount * (amountTolerancePercent / 100)))
        .sort((a, b) => b.score - a.score)[0];
      if (!candidates) return { id: `m-${invoice.id}`, invoice, score: 0, status: "held", reason: "لم توجد حوالة مناسبة" } satisfies StageBMatch;
      used.add(candidates.transfer.id);
      const status = candidates.amountDiff > 0.01 || candidates.accountDiff ? "difference" : candidates.score >= 90 ? "pending" : "held";
      const reasons = [
        stageBNameMatch(invoice.name, candidates.transfer.description, aliases) < 0.5 ? "اختلاف اسم يحتاج مراجعة" : "",
        candidates.accountDiff ? `فرق نوع الحساب: ${stageBAccountType(invoice)} مقابل ${stageBTransferAccountType(candidates.transfer)}` : "",
        candidates.amountDiff > 0.01 ? `فرق مبلغ: الحوالة ${candidates.transfer.receivedAmount > invoice.paidAmount ? "أكثر" : "أقل"} بـ ${fmtNum(candidates.amountDiff)} شيكل` : ""
      ].filter(Boolean);
      return { id: `m-${invoice.id}-${candidates.transfer.id}`, invoice, transfer: candidates.transfer, score: Math.round(candidates.score), status, reason: reasons.join("؛ ") || "مطابقة اسم ومبلغ" } satisfies StageBMatch;
    });
    setMatches(output);
    setSelected(new Set());
    setStageBPage(1);
    setStageBView("pending");
    setError(null);
  };

  const updateStatus = (ids: Set<string>, status: StageBMatch["status"]) => setMatches(prev => {
    undoStack.current.push(prev);
    return prev.map(match => ids.has(match.id) ? { ...match, status } : match);
  });
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setDrawerMatchId(null);
        setDetailMatchId(null);
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
        event.preventDefault();
        const previous = undoStack.current.pop();
        if (previous) setMatches(previous);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);
  const replaceTransfer = (match: StageBMatch, transfer: StageBTransfer) => {
    setMatches(prev => {
      undoStack.current.push(prev);
      return prev.map(item => item.id === match.id ? {
      ...item, transfer, status: Math.abs(transfer.receivedAmount - item.invoice.paidAmount) > 0.01 || normName(stageBAccountType(item.invoice)) !== normName(stageBTransferAccountType(transfer)) ? "difference" : "pending",
      score: Math.round(stageBNameMatch(item.invoice.name, transfer.description, aliases) * 70 + (Math.abs(transfer.receivedAmount - item.invoice.paidAmount) <= 0.01 ? 30 : 0)),
      reason: "تم اختيار حوالة بديلة يدوياً"
      } : item);
    });
    setDrawerMatchId(null);
  };
  const replaceInvoice = (match: StageBMatch, invoice: StageBInvoice) => {
    setMatches(prev => {
      undoStack.current.push(prev);
      return prev.map(item => item.id === match.id ? {
      ...item, invoice, status: item.transfer && (Math.abs(item.transfer.receivedAmount - invoice.paidAmount) > 0.01 || normName(stageBAccountType(invoice)) !== normName(stageBTransferAccountType(item.transfer))) ? "difference" : "pending",
      score: item.transfer ? Math.round(stageBNameMatch(invoice.name, item.transfer.description, aliases) * 70 + (Math.abs(item.transfer.receivedAmount - invoice.paidAmount) <= 0.01 ? 30 : 0)) : 0,
      reason: "تم اختيار فاتورة بديلة يدوياً"
      } : item);
    });
    setDrawerMatchId(null);
  };
  const saveAliases = (items: StageBMatch[]) => {
    const next = { ...aliases };
    items.forEach(item => { if (item.transfer) next[stageBNameKey(item.invoice.name)] = stageBNameKey(item.transfer.name); });
    setAliases(next);
    void storageSet("stage_b_aliases", next);
  };
  const confirmSelected = () => {
    const items = matches.filter(match => selected.has(match.id) && match.transfer);
    if (!items.length) return;
    updateStatus(new Set(items.map(item => item.id)), "confirmed");
    saveAliases(items);
    setSelected(new Set());
  };
  const visibleMatches = matches.filter(match => {
    const invoiceHit = !invoiceSearch.trim() || `${match.invoice.name} ${match.invoice.userId}`.toLowerCase().includes(invoiceSearch.trim().toLowerCase());
    const transferHit = !transferSearch.trim() || match.transfer?.description.toLowerCase().includes(transferSearch.trim().toLowerCase());
    const accountHit = stageBAccountFilter === "all" || stageBAccountFilter === stageBAccountType(match.invoice) || stageBAccountFilter === (match.transfer ? stageBTransferAccountType(match.transfer) : "");
    const viewHit = stageBView === "confirmed"
      ? match.status === "confirmed"
      : stageBView === "manual"
        ? ["heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)
        : !["confirmed", "heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status);
    return invoiceHit && transferHit && accountHit && viewHit;
  }).sort((a, b) => {
    const compare = (sort: StageBSort, left: { name: string; amount: number }, right: { name: string; amount: number }) => {
      if (sort === "name") return left.name.localeCompare(right.name, "ar");
      if (sort === "amountAsc") return left.amount - right.amount;
      if (sort === "amountDesc") return right.amount - left.amount;
      return 0;
    };
    const invoiceResult = compare(invoiceSort, a.invoice, b.invoice);
    if (invoiceResult) return invoiceResult;
    return compare(
      transferSort,
      { name: a.transfer?.description || "", amount: a.transfer?.receivedAmount || 0 },
      { name: b.transfer?.description || "", amount: b.transfer?.receivedAmount || 0 },
    );
  });
  const stageBPageSize = 20;
  const totalStageBPages = Math.max(1, Math.ceil(visibleMatches.length / stageBPageSize));
  const currentStageBPage = Math.min(stageBPage, totalStageBPages);
  const pagedMatches = visibleMatches.slice((currentStageBPage - 1) * stageBPageSize, currentStageBPage * stageBPageSize);
  const pageMatchIds = new Set(pagedMatches.map(match => match.id));
  const pageIsSelected = pagedMatches.length > 0 && pagedMatches.every(match => selected.has(match.id));
  const togglePage = () => setSelected(previous => {
    const next = new Set(previous);
    if (pageIsSelected) {
      pageMatchIds.forEach(id => next.delete(id));
    } else {
      pageMatchIds.forEach(id => next.add(id));
    }
    return next;
  });
  const toggleAll = togglePage;
  const openReplacement = (matchId: string, mode: "invoice" | "transfer") => {
    setDrawerMode(mode);
    setDrawerMatchId(matchId);
    window.setTimeout(() => rowRefs.current[matchId]?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" }), 0);
  };
  const suggestionClass = (amountDiff: number, nameMatch: number) =>
    amountDiff <= 0.01 && nameMatch >= 0.5
      ? "border-green-300 bg-green-50 text-green-800"
      : amountDiff <= 0.01
        ? "border-amber-300 bg-amber-50 text-amber-800"
        : "border-red-300 bg-red-50 text-red-800";

  const exportStageB = () => {
    const workbook = XLSX.utils.book_new();
    const rowsFor = (items: StageBMatch[]) => items.map(match => [
      match.invoice.userId, match.invoice.registrationTime, match.invoice.customerNumber, match.invoice.invoiceNumber,
      match.invoice.name, match.invoice.amount, match.invoice.accountType || "—",
      match.transfer?.description || "—", match.transfer?.amount ?? "", match.transfer?.accountType || "—",
      match.transfer?.type || "—", match.transfer ? fmtNum(match.transfer.receivedAmount - match.invoice.paidAmount) : "", stageBAccountType(match.invoice) || "—",
      match.transfer ? stageBTransferAccountType(match.transfer) : "—", match.score, match.reason
    ]);
    const headersOut = ["رقم المستخدم", "الساعة", "رقم الزبون", "رقم الفاتورة", "الفاتورة", "مبلغ الفاتورة", "نوع حساب الفاتورة", "الحوالة", "مبلغ الحوالة", "نوع حوالة الملف", "فرق المبلغ", "نوع الحساب الفعلي للفاتورة", "نوع الحساب الفعلي للحوالة", "الثقة", "الملاحظة"];
    const add = (name: string, items: StageBMatch[]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headersOut, ...rowsFor(items)]), name);
    add("المطابقة العامة", matches);
    add("المؤكد والمطابق", matches.filter(match => match.status === "confirmed"));
    add("المنتظر / قيد المطابقة", matches.filter(match => ["held", "heldInvoice", "heldTransfer", "rejected"].includes(match.status)));
    add("الفيزا", matches.filter(match => match.status === "visa"));
    add("فروقات تحتاج تدقيق", matches.filter(match => match.status === "difference"));
    XLSX.writeFile(workbook, "مطابقة_B_الفواتير_والحوالات.xlsx");
  };

  const saveStageBSession = () => {
    void storageSet("stage_b_session", { headers, rows, maps, matches, amountTolerancePercent });
    setError(null);
  };

  const exportStageBSession = () => {
    const blob = new Blob([JSON.stringify({ headers, rows, maps, matches, amountTolerancePercent })], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "جلسة_المرحلة_B.json";
    link.click();
    URL.revokeObjectURL(url);
  };

  const clearStageBSession = () => {
    if (!window.confirm("مسح ملفات ونتائج المرحلة B الحالية؟")) return;
    setFiles({ invoiceBank: null, transferBank: null, invoiceWallet: null, transferWallet: null });
    setHeaders({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
    setRows({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
    setMaps({ invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap });
    setMatches([]);
    setSelected(new Set());
    void storageSet("stage_b_session", null);
  };

  const importStageBSession = async (file: File) => {
    try {
      const session = JSON.parse(new TextDecoder().decode(await readFileBuf(file)));
      if (!session || typeof session !== "object" || !session.rows || !session.maps) throw new Error("ملف المرحلة B غير صالح.");
      setHeaders(session.headers || { invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [] });
      setRows(session.rows);
      setMaps(session.maps);
      setMatches(session.matches || []);
      setAmountTolerancePercent(session.amountTolerancePercent ?? 0.5);
      setError(null);
      void storageSet("stage_b_session", session);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const stageBStats = [
    { label: "مؤكد", value: matches.filter(match => match.status === "confirmed").length, color: "border-blue-300 bg-blue-50 text-green-700" },
    { label: "منتظر", value: matches.filter(match => ["pending", "held"].includes(match.status)).length, color: "border-slate-200 bg-white text-blue-700" },
    { label: "مرفوض", value: matches.filter(match => match.status === "rejected").length, color: "border-slate-200 bg-white text-red-600" },
    { label: "مطابقة يدوية", value: matches.filter(match => ["heldInvoice", "heldTransfer", "difference"].includes(match.status)).length, color: "border-slate-200 bg-white text-purple-700" },
    { label: "فيزا", value: matches.filter(match => match.status === "visa").length, color: "border-slate-200 bg-white text-teal-700" },
    { label: "فواتير", value: invoiceData.length, color: "border-slate-200 bg-white text-orange-600" },
    { label: "حوالات", value: transferData.length, color: "border-slate-200 bg-white text-orange-600" },
    { label: "إجمالي المطابقات", value: matches.length, color: "border-slate-200 bg-white text-slate-700" },
  ];

  const renderFile = (source: "invoiceBank" | "transferBank" | "invoiceWallet" | "transferWallet", label: string) => (
    <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4">
      <h2 className="text-sm font-bold">{label}</h2>
      <DropZone file={files[source]} onFile={file => loadFile(source, file)} onClear={() => { setFiles(prev => ({ ...prev, [source]: null })); setHeaders(prev => ({ ...prev, [source]: [] })); setRows(prev => ({ ...prev, [source]: [] })); }} />
      {!!headers[source].length && <>
        <div className="grid grid-cols-2 gap-2">
          <Sel label="الاسم / البيان" headers={headers[source]} value={maps[source].name} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], name: value } }))} />
          <Sel label="المبلغ المدفوع" headers={headers[source]} value={maps[source].paidAmount} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], paidAmount: value, amount: value } }))} />
          <Sel label="المبلغ المستلم" headers={headers[source]} value={maps[source].receivedAmount} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], receivedAmount: value } }))} />
          <Sel label="نوع الحساب" headers={headers[source]} value={maps[source].accountType} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], accountType: value } }))} />
        </div>
        <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
          <input type="checkbox" checked={stageBSwaps[source]} onChange={event => setStageBSwaps(prev => ({ ...prev, [source]: event.target.checked }))} className="rounded" />
          عكس المدين والدائن
        </label>
      </>}
    </div>
  );

  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-7xl space-y-5">
    <div className="sticky top-2 z-20 flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-slate-200 bg-white/95 p-5">
      <div className="flex items-center gap-3">
        <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة للمنصة الرئيسية</button>
        <span className="hidden h-8 w-px bg-slate-200 sm:block" />
        <div><h1 className="text-xl font-black text-slate-900">B — المطابقة الذكية للفواتير والحوالات</h1><p className="mt-1 text-xs text-slate-500">مطابقة المبالغ والأسماء بين الفواتير وحوالات بنك فلسطين والمحفظة التجارية</p></div>
      </div>
      <div className="flex flex-wrap gap-2">
        <button onClick={() => setShowSettings(value => !value)} className="rounded-lg border px-3 py-2 text-xs">الإعدادات</button>
        <button onClick={saveStageBSession} className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs text-blue-700">حفظ المشروع</button>
        <button onClick={exportStageBSession} className="rounded-lg border px-3 py-2 text-xs">تصدير المشروع</button>
        <button onClick={() => sessionFileRef.current?.click()} className="rounded-lg border px-3 py-2 text-xs">استيراد ملف سابق</button>
        <input ref={sessionFileRef} type="file" accept=".json,application/json" className="hidden" onChange={event => { const file = event.target.files?.[0]; if (file) void importStageBSession(file); event.target.value = ""; }} />
        <button onClick={clearStageBSession} className="rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-700">أدوات الحذف</button>
        <button onClick={runStageB} className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white"><Sparkles className="mr-1 inline h-4 w-4" />تشغيل المطابقة</button>
        {matches.length > 0 && <button onClick={exportStageB} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white"><Download className="mr-1 inline h-4 w-4" />تصدير الشيتات</button>}
      </div>
    </div>
    {showSettings && <div className="rounded-xl border border-blue-100 bg-blue-50/60 p-4">
      <label className="flex items-center gap-2 text-xs">نسبة التسامح في فرق المبلغ
        <input type="number" min="0" max="10" step="0.1" value={amountTolerancePercent} onChange={event => setAmountTolerancePercent(Math.min(10, Math.max(0, Number(event.target.value) || 0)))} className="w-20 rounded-lg border bg-white px-2 py-1.5 text-center" />%
      </label>
    </div>}
    <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4 text-sm leading-6 text-emerald-900">تتم المطابقة على أساس المبلغ والاسم معًا. عند الترحيل من A ارفع حوالتي بنك فلسطين والمحفظة التجارية فقط، وعند العمل المستقل ارفع ملفات الفواتير والحوالات الأربعة. نوع الحساب يُحدد تلقائياً من خانة الملف.</div>
    {!!invoices.filter(invoice => invoice.category === "فيزا").length && <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
      <div><strong>فواتير الفيزا المرحّلة من قاعدة البيانات</strong><div className="mt-1 text-xs">تم الاحتفاظ بنتيجة الفرز، وهذه الفواتير تُرحّل إلى منصة C ولا تدخل في مطابقة B.</div></div>
      <button onClick={() => onTransferVisaToC(invoices.filter(invoice => invoice.category === "فيزا"))} className="rounded-lg bg-amber-600 px-4 py-2 text-xs font-medium text-white hover:bg-amber-700">ترحيل الفيزا إلى منصة C ({invoices.filter(invoice => invoice.category === "فيزا").length})</button>
    </div>}
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
      {hasTransferredInvoices ? <>{renderFile("transferBank", "حوالات بنك فلسطين")}{renderFile("transferWallet", "حوالات المحفظة التجارية")}</> : <>{renderFile("invoiceBank", "فواتير بنك فلسطين")}{renderFile("invoiceWallet", "فواتير المحفظة التجارية")}{renderFile("transferBank", "حوالات بنك فلسطين")}{renderFile("transferWallet", "حوالات المحفظة التجارية")}</>}
    </div>
    <div className="grid grid-cols-2 gap-2 sm:grid-cols-4 lg:grid-cols-8">
      {stageBStats.map(stat => <div key={stat.label} className={`rounded-xl border p-3 shadow-[0_3px_10px_rgba(30,64,175,0.06)] ${stat.color}`}>
        <div className="text-[11px] font-medium text-slate-600">{stat.label}</div>
        <div className="mt-1 text-2xl font-bold">{stat.value}</div>
      </div>)}
    </div>
    {!!matches.length && <div className={`stage-b-results overflow-hidden rounded-xl border bg-white ${stageBView === "confirmed" ? "stage-b-confirmed" : ""}`}>
      <div className="flex flex-wrap items-center gap-2 border-b bg-slate-50 p-3"><div className="flex gap-1 rounded-lg bg-white p-1"><button onClick={() => { setStageBView("pending"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "pending" ? "bg-amber-100 text-amber-800" : "text-slate-600"}`}>المنتظر {matches.filter(match => !["confirmed", "heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)).length}</button><button onClick={() => { setStageBView("confirmed"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "confirmed" ? "bg-green-100 text-green-800" : "text-slate-600"}`}>المؤكد {matches.filter(match => match.status === "confirmed").length}</button><button onClick={() => { setStageBView("manual"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "manual" ? "bg-blue-100 text-blue-800" : "text-slate-600"}`}>مطابقة يدوية {matches.filter(match => ["heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)).length}</button></div><input value={invoiceSearch} onChange={event => { setInvoiceSearch(event.target.value); setStageBPage(1); }} placeholder="بحث بالاسم أو البيان أو رقم المستخدم..." className="min-w-52 flex-1 rounded-lg border px-3 py-2 text-xs" /><input value={transferSearch} onChange={event => { setTransferSearch(event.target.value); setStageBPage(1); }} placeholder="بحث في الحوالات..." className="min-w-52 flex-1 rounded-lg border px-3 py-2 text-xs" /><select value={invoiceSort} onChange={event => { setInvoiceSort(event.target.value as StageBSort); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="none">فواتير: بدون فرز</option><option value="name">فواتير: بالاسم</option><option value="amountDesc">فواتير: من الكبير للصغير</option><option value="amountAsc">فواتير: من الصغير للكبير</option></select><select value={transferSort} onChange={event => { setTransferSort(event.target.value as StageBSort); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="none">حوالات: بدون فرز</option><option value="name">حوالات: بالاسم</option><option value="amountDesc">حوالات: من الكبير للصغير</option><option value="amountAsc">حوالات: من الصغير للكبير</option></select><select value={stageBAccountFilter} onChange={event => { setStageBAccountFilter(event.target.value); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="all">كل الأنواع</option><option value="بنك فلسطين">بنك فلسطين</option><option value="محفظة تجارية">محفظة تجارية</option></select>      <button onClick={togglePage} className="rounded-lg border px-3 py-2 text-xs">{pageIsSelected ? "إلغاء تحديد الصفحة" : `تحديد الصفحة (${pagedMatches.length})`}</button><button onClick={confirmSelected} disabled={!selected.size} className="rounded-lg bg-blue-600 px-3 py-2 text-xs text-white disabled:opacity-40">حفظ المحدد ({selected.size})</button><button onClick={() => updateStatus(selected, "rejected")} disabled={!selected.size} className="rounded-lg bg-red-600 px-3 py-2 text-xs text-white disabled:opacity-40">رفض المحدد</button></div>
      <div className="flex items-start gap-4"><div className="min-w-0 flex-1 overflow-auto"><table className="w-full min-w-[1000px] text-sm"><thead className="sticky top-0 z-10 bg-slate-100 text-xs"><tr><th className="p-3">      <input type="checkbox" checked={pagedMatches.length > 0 && pagedMatches.every(match => selected.has(match.id))} onChange={toggleAll} /></th><th className="p-3 text-right">الفاتورة / المستخدم</th><th className="p-3 text-right">المبلغ</th><th className="p-3 text-right">نوع الحساب</th><th className="p-3 text-right">الحوالة</th><th className="p-3 text-right">مبلغ الحوالة</th><th className="p-3 text-right">نوع الحساب</th><th className="p-3 text-right">الحالة</th><th className="p-3 text-right">إجراءات</th></tr></thead>            <tbody>{pagedMatches.map(match => <tr ref={row => { rowRefs.current[match.id] = row; }} key={match.id} className="border-t align-top"><td className="p-3"><input type="checkbox" checked={selected.has(match.id)} onChange={() => setSelected(prev => { const next = new Set(prev); if (next.has(match.id)) next.delete(match.id); else next.add(match.id); return next; })} /></td><td className="p-3 font-medium">{match.invoice.name}      <div className="text-[10px] text-slate-500">المستخدم: {match.invoice.userId || "—"} · الزبون: {match.invoice.customerNumber || "—"} · الفاتورة: {match.invoice.invoiceNumber || "—"} · الساعة: {match.invoice.registrationTime || "—"}</div></td><td className="p-3 font-mono">{fmtNum(match.invoice.amount)}</td><td className="p-3">{match.invoice.accountType || "—"}</td><td className="p-3">{match.transfer?.description || "لا توجد حوالة مناسبة"}</td><td className="p-3 font-mono">{match.transfer ? fmtNum(match.transfer.amount) : "—"}</td><td className="p-3">{match.transfer?.accountType || "—"}</td><td className="p-3"><span className={match.status === "confirmed" ? "text-green-700" : match.status === "difference" ? "text-orange-700" : match.status === "visa" ? "text-teal-700" : "text-amber-700"}>{match.status === "confirmed" ? "مؤكد ومطابق" : match.status === "rejected" ? "مرفوض" : match.status === "heldInvoice" ? "معلق: فاتورة" : match.status === "heldTransfer" ? "معلق: حوالة" : match.status === "held" ? "معلق" : match.status === "visa" ? "فيزا" : "فرق يحتاج تدقيق"}</span><div className="text-[10px] text-slate-500">{match.reason} · {match.score}%</div></td><td className="p-3"><div className="flex flex-wrap gap-1"><button onClick={() => { const one = new Set([match.id]); if (match.transfer) saveAliases([match]); updateStatus(one, "confirmed"); }} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">تأكيد</button><button onClick={() => updateStatus(new Set([match.id]), "rejected")} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button>            <button onClick={() => openReplacement(match.id, "invoice")} className="rounded border border-emerald-200 bg-emerald-50 px-2 py-1 text-[10px] text-emerald-700">سحب كاشير</button><button onClick={() => openReplacement(match.id, "transfer")} className="rounded border border-blue-200 bg-blue-50 px-2 py-1 text-[10px] text-blue-700">سحب حوالة</button><button onClick={() => setDetailMatchId(detailMatchId === match.id ? null : match.id)} className="rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[10px] text-slate-700" title="تفاصيل">i تفاصيل</button><button onClick={() => updateStatus(new Set([match.id]), "heldInvoice")} className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-[10px] text-amber-700">علق فاتورة</button><button onClick={() => updateStatus(new Set([match.id]), "heldTransfer")} className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-[10px] text-amber-700">علق حوالة</button>      {stageAVisaNumber(match.invoice.originalName) && <button onClick={() => updateStatus(new Set([match.id]), "visa")} className="rounded border border-teal-200 bg-teal-50 px-2 py-1 text-[10px] text-teal-700">ترحيل للفيزا</button>}</div></td></tr>)}            </tbody></table></div>{drawerMatchId && (() => { const current = matches.find(match => match.id === drawerMatchId); if (!current) return null;             const transferCandidates = transferData.filter(item => item.id !== current.transfer?.id && (!drawerMinAmount || item.receivedAmount >= Number(drawerMinAmount)) && (!drawerMaxAmount || item.receivedAmount <= Number(drawerMaxAmount))).slice(0, 30); const invoiceCandidates = invoiceData.filter(item => item.id !== current.invoice.id && (!drawerMinAmount || item.paidAmount >= Number(drawerMinAmount)) && (!drawerMaxAmount || item.paidAmount <= Number(drawerMaxAmount))).slice(0, 30);       return <aside className="replacement-drawer sticky top-3 shrink-0 rounded-xl border bg-white p-3 shadow-lg"><div className="mb-2 flex items-center justify-between"><strong className="text-sm">بحث واستبدال</strong><button onClick={() => setDrawerMatchId(null)} className="text-xs text-slate-500">إغلاق</button></div>      <div className="space-y-3">      {drawerMode === "invoice" && <section><p className="mb-1 text-[10px] font-bold text-emerald-700">بحث وسحب من الكاشير</p><input value={drawerInvoiceSearch} onChange={event => setDrawerInvoiceSearch(event.target.value)} placeholder="ابحث باسم أو مبلغ الفاتورة..." className="mb-2 w-full rounded-lg border px-2 py-1.5 text-xs" /><div className="max-h-48 space-y-1 overflow-auto">{invoiceCandidates.filter(item => !drawerInvoiceSearch || `${item.name} ${item.amount}`.toLowerCase().includes(drawerInvoiceSearch.toLowerCase())).map(item =>       <button key={item.id} onClick={() => replaceInvoice(current, item)} className={`block w-full rounded border p-2 text-right text-[10px] ${suggestionClass(Math.abs(item.paidAmount - current.invoice.paidAmount), stageBNameMatch(item.name, current.invoice.name, aliases))}`}>{item.name} · {fmtNum(item.paidAmount)} · المستخدم {item.userId || "—"}</button>)}      </div></section>}      <section className={drawerMode === "transfer" ? "" : "hidden"}><p className="mb-1 text-[10px] font-bold text-blue-700">بحث وسحب من الحوالات</p><input value={drawerTransferSearch} onChange={event => setDrawerTransferSearch(event.target.value)} placeholder="ابحث باسم أو مبلغ الحوالة..." className="mb-2 w-full rounded-lg border px-2 py-1.5 text-xs" /><div className="max-h-48 space-y-1 overflow-auto">{transferCandidates.filter(item => !drawerTransferSearch || `${item.description} ${item.receivedAmount}`.toLowerCase().includes(drawerTransferSearch.toLowerCase())).map(item =>       <button key={item.id} onClick={() => replaceTransfer(current, item)} className={`block w-full rounded border p-2 text-right text-[10px] ${suggestionClass(Math.abs(item.receivedAmount - current.invoice.paidAmount), stageBNameMatch(current.invoice.name, item.description, aliases))}`}>{item.description} · {fmtNum(item.receivedAmount)} · {item.accountType || "—"}</button>)}</div></section></div>            {detailMatchId === current.id && <div className="mt-3 rounded bg-slate-50 p-2 text-[10px]"><strong>تفاصيل الفاتورة:</strong><div className="mt-1 space-y-1"><div>الساعة: {current.invoice.registrationTime || "—"}</div><div>رقم الزبون: {current.invoice.customerNumber || "—"}</div><div>رقم الفاتورة: {current.invoice.invoiceNumber || "—"}</div><div>رقم المستخدم: {current.invoice.userId || "—"}</div></div><pre className="mt-1 max-h-28 overflow-auto whitespace-pre-wrap">      {JSON.stringify(current.invoice.raw, null, 2)}</pre></div>}      </aside>; })()}</div><div className="flex items-center justify-center gap-3 border-t bg-slate-50 p-3 text-xs"><button onClick={() => setStageBPage(page => Math.max(1, page - 1))} disabled={currentStageBPage === 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button><span>صفحة {currentStageBPage} من {totalStageBPages} · عرض {pagedMatches.length} من {visibleMatches.length}</span><button onClick={() => setStageBPage(page => Math.min(totalStageBPages, page + 1))} disabled={currentStageBPage === totalStageBPages} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button></div>
    </div>}
  </div></div>;
}

interface StageCMatch {
  invoice: StageAResult;
  sonyName: string;
  sonyAmount: number;
  sonyVisa: string;
  authNumber: string;
  matched: boolean;
  issue: string;
}

interface StageDPending {
  description: string;
  amount: number;
  visaNumber: string;
  date: string;
}

function extractFourDigitVisa(value: string): string {
  const matches = value.match(/\d{4}(?!\d)/g);
  return matches?.[matches.length - 1] || "";
}

function StageDPlatform({ matches, pending, onBack }: { matches: StageCMatch[]; pending: StageDPending[]; onBack: () => void }) {
  const [file, setFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<Record<string, unknown>[]>([]);
  const [map, setMap] = useState({ description: "", amount: "", date: "" });
  const [error, setError] = useState<string | null>(null);
  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadFile = async (nextFile: File) => {
    try {
      const parsed = parseSheet(await readFileBuf(nextFile));
      setFile(nextFile); setHeaders(parsed.headers); setRawRows(parsed.rows);
      setMap({
        description: autoDetect(parsed.headers, ["مشتريات عمولة تجار", ...HINTS.desc]),
        amount: autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, HINTS.credit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        date: autoDetect(parsed.headers, HINTS.date)
      });
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };
  const pendingRows = rawRows
    .filter(row => String(row[map.description] ?? "").includes("مشتريات عمولة تجار"))
    .map(row => ({ description: String(row[map.description] ?? ""), amount: parseAmount(row[map.amount]), visaNumber: extractFourDigitVisa(String(row[map.description] ?? "")), date: String(row[map.date] ?? "") }));
  const rows = pendingRows.map(bank => {
    const match = matches.find(item => item.matched && item.sonyVisa === bank.visaNumber && Math.abs(item.invoice.amount - bank.amount) <= 0.01);
    return { bank, match };
  });
  const exportStageD = () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["بيان الحوالة المعلقة", "التاريخ", "مبلغ الحوالة", "رقم الفيزا", "اسم الزبون", "مبلغ الفاتورة", "رقم التفويض", "الحالة"],
      ...rows.map(({ bank, match }) => [bank.description, bank.date, bank.amount, bank.visaNumber, match?.invoice.name || "", match?.invoice.amount || "", match?.authNumber || "", match ? "مطابق" : "غير مطابق"])
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, "مطابقة D");
    XLSX.writeFile(workbook, "مطابقة_الفيزا_المرحلة_D.xlsx");
  };
  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900">
    <div className="mx-auto max-w-6xl space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة إلى C</button>
        <h1 className="text-xl font-black">D — مطابقة حوالات الغد مع مطابقة C</h1>
        <button onClick={exportStageD} disabled={!rows.length} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white disabled:opacity-40"><Download className="mr-1 inline h-4 w-4" />تصدير مطابقة D</button>
      </div>
      <div className="rounded-xl border border-indigo-200 bg-indigo-50 p-4 text-sm text-indigo-900">ارفع حوالات «مشتريات عمولة تجار» هنا. تتم المطابقة برقم الفيزا المكوّن من 4 أرقام والمبلغ، ويظهر رقم التفويض القادم من ملف سوني كاشير.</div>
      {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
      <div className="rounded-xl border border-indigo-200 bg-white p-4">
        <h2 className="mb-3 text-sm font-bold text-indigo-800">حوالات بنكية معلقة — مشتريات عمولة تجار</h2>
        <DropZone file={file} onFile={loadFile} onClear={() => { setFile(null); setHeaders([]); setRawRows([]); }} />
        {!!headers.length && <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
          <Sel label="بيان الحوالة" headers={headers} value={map.description} onChange={value => setMap(prev => ({ ...prev, description: value }))} />
          <Sel label="المبلغ" headers={headers} value={map.amount} onChange={value => setMap(prev => ({ ...prev, amount: value }))} />
          <Sel label="التاريخ" headers={headers} value={map.date} onChange={value => setMap(prev => ({ ...prev, date: value }))} />
        </div>}
        <p className="mt-2 text-xs text-slate-500">سيتم اعتماد الصفوف التي تحتوي على «مشتريات عمولة تجار» فقط: {pendingRows.length} حوالة.</p>
      </div>
      <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
        <table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right">بيان الحوالة</th><th className="p-3 text-right">المبلغ</th><th className="p-3 text-right">رقم الفيزا</th><th className="p-3 text-right">اسم الزبون</th><th className="p-3 text-right">رقم التفويض</th><th className="p-3 text-right">الحالة</th></tr></thead>
          <tbody>{rows.map(({ bank, match }) => <tr key={`${bank.description}-${bank.date}-${bank.amount}`} className="border-t"><td className="p-3">{bank.description}</td><td className="p-3 font-mono">{fmtNum(bank.amount)}</td><td className="p-3 font-mono">{bank.visaNumber || "—"}</td><td className="p-3">{match?.invoice.name || "—"}</td><td className="p-3 font-mono">{match?.authNumber || "—"}</td><td className={`p-3 font-semibold ${match ? "text-green-700" : "text-red-700"}`}>{match ? "مطابق" : "غير مطابق"}</td></tr>)}</tbody>
        </table>
        {!rows.length && <div className="p-10 text-center text-sm text-slate-500">لا توجد حوالات مرحّلة إلى D.</div>}
      </div>
    </div>
  </div>;
}

function StageCPlatform({ invoices, onBack, onTransferToD }: { invoices: StageAResult[]; onBack: () => void; onTransferToD?: (matches: StageCMatch[], pending: StageDPending[]) => void }) {
  const [file, setFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<Record<string, unknown>[]>([]);
  const [map, setMap] = useState({ name: "", amount: "", auth: "" });
  const [results, setResults] = useState<StageCMatch[]>([]);
  const [page, setPage] = useState(1);
  const [error, setError] = useState<string | null>(null);
  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadSonyFile = async (nextFile: File) => {
    try {
      const parsed = parseSheet(await readFileBuf(nextFile));
      setFile(nextFile); setHeaders(parsed.headers); setRows(parsed.rows);
      setMap({ name: autoDetect(parsed.headers, HINTS.name), amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]), auth: autoDetect(parsed.headers, HINTS.authNum) });
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };
  const runStageC = () => {
    if (!invoices.length) { setError("لا توجد فواتير فيزا مرحّلة من المرحلة A"); return; }
    if (!rows.length || !map.name || !map.amount) { setError("ارفع ملف سوني كاشير وحدد عمود البيان والمبلغ أولاً"); return; }
    const sonyRows = rows.map((row, index) => ({ index, name: String(row[map.name] ?? "").trim(), amount: parseAmount(row[map.amount]), visa: stageAVisaNumber(String(row[map.name] ?? "")), auth: String(row[map.auth] ?? "").trim(), used: false }));
    setResults(invoices.map(invoice => {
      const visa = stageAVisaNumber(invoice.originalName);
      const direct = sonyRows.find(row => !row.used && row.visa === visa && Math.abs(row.amount - invoice.amount) <= 0.01);
      if (direct) { direct.used = true; return { invoice, sonyName: direct.name, sonyAmount: direct.amount, sonyVisa: direct.visa, authNumber: direct.auth, matched: true, issue: "" }; }
      const group = sonyRows.filter(row => !row.used && row.visa === visa);
      const grouped = group.length > 1 ? group.reduce((sum, row) => sum + row.amount, 0) : 0;
      if (grouped && Math.abs(grouped - invoice.amount) <= 0.01) { group.forEach(row => { row.used = true; }); return { invoice, sonyName: group.map(row => row.name).join(" + "), sonyAmount: grouped, sonyVisa: visa, authNumber: group.map(row => row.auth).filter(Boolean).join(" + "), matched: true, issue: "مطابقة مجمعة" }; }
      return { invoice, sonyName: "", sonyAmount: 0, sonyVisa: visa, authNumber: "", matched: false, issue: visa ? "رقم الفيزا موجود لكن الرقم أو المبلغ غير موجود في سوني" : "لم يتم العثور على رقم فيزا من 4 أرقام" };
    }));
    setPage(1);
    setError(null);
  };
  const pageSize = 20;
  const totalPages = Math.max(1, Math.ceil(results.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pagedResults = results.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const exportStageC = () => { const workbook = XLSX.utils.book_new(); const sheet = XLSX.utils.aoa_to_sheet([["اسم الزبون", "رقم الفيزا", "مبلغ الفاتورة", "رقم التفويض", "بيان سوني", "مبلغ سوني", "الحالة", "ملاحظة"], ...results.map(result => [result.invoice.name, result.sonyVisa, result.invoice.amount, result.authNumber, result.sonyName, result.sonyAmount || "", result.matched ? "مطابق" : "غير مطابق", result.invoice.categoryIssue || result.issue])]); XLSX.utils.book_append_sheet(workbook, sheet, "مطابقة فيزا"); XLSX.writeFile(workbook, "مطابقة_الفيزا_مع_سوني.xlsx"); };
  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-6xl space-y-5">
    <div className="flex flex-wrap items-center justify-between gap-3">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة لمنصة الفيزا</button>
      <h1 className="text-xl font-black">C — مطابقة الفواتير مع سوني كاشير فيزا</h1>
      <div className="flex gap-2">
        <button onClick={runStageC} className="rounded-lg bg-amber-600 px-4 py-2 text-sm font-medium text-white"><Sparkles className="mr-1 inline h-4 w-4" />تشغيل المطابقة</button>
        {results.length > 0 && <button onClick={exportStageC} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white"><Download className="mr-1 inline h-4 w-4" />تصدير مطابقة فيزا</button>}
        {results.some(result => result.matched) && <button onClick={() => onTransferToD?.(results, [])} className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-medium text-white">ترحيل إلى D</button>}
      </div>
    </div>
    <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">تم ترحيل {invoices.length} فاتورة فيزا من A. تتم المطابقة برقم الفيزا والمبلغ، ويُحفظ رقم التفويض من سوني كاشير.</div>
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <h2 className="mb-3 text-sm font-bold">ملف سوني كاشير فيزا</h2>
      <DropZone file={file} onFile={loadSonyFile} onClear={() => { setFile(null); setHeaders([]); setRows([]); }} />
      {!!headers.length && <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
        <Sel label="البيان الذي يحتوي رقم الفيزا" headers={headers} value={map.name} onChange={value => setMap(prev => ({ ...prev, name: value }))} />
        <Sel label="المبلغ" headers={headers} value={map.amount} onChange={value => setMap(prev => ({ ...prev, amount: value }))} />
        <Sel label="رقم التفويض" headers={headers} value={map.auth} onChange={value => setMap(prev => ({ ...prev, auth: value }))} />
      </div>}
    </div>
    {!!results.length && <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="border-b bg-slate-50 px-4 py-3 text-sm font-bold">مطابق: {results.filter(result => result.matched).length} · غير مطابق: {results.filter(result => !result.matched).length}</div>
      <div className="overflow-x-auto"><table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right">اسم الزبون</th><th className="p-3 text-right">رقم الفيزا</th><th className="p-3 text-right">المبلغ</th><th className="p-3 text-right">رقم التفويض</th><th className="p-3 text-right">بيان سوني</th><th className="p-3 text-right">الحالة</th></tr></thead><tbody>
        {pagedResults.map(result => <tr key={`${result.invoice.source}-${result.invoice.id}`} className="border-t"><td className="p-3">{result.invoice.name}</td><td className="p-3 font-mono">{result.sonyVisa || "—"}</td><td className="p-3 font-mono">{fmtNum(result.invoice.amount)}</td><td className="p-3 font-mono">{result.authNumber || "—"}</td><td className="p-3">{result.sonyName || "—"}</td><td className={`p-3 font-semibold ${result.matched ? "text-green-700" : "text-red-700"}`}>{result.matched ? result.issue || "مطابق" : "غير مطابق"}</td></tr>)}
      </tbody></table></div>
    </div>}
  </div></div>;
}

function EmptyStageB({ onBack }: { onBack: () => void }) {
  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900">
    <div className="mx-auto flex min-h-[70vh] max-w-6xl items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-white">
      <div className="text-center">
        <div className="text-lg font-bold text-slate-700">منصة B</div>
        <p className="mt-2 text-sm text-slate-500">هذه المنصة فارغة حاليًا.</p>
        <button onClick={onBack} className="mt-4 rounded-lg bg-blue-600 px-4 py-2 text-sm text-white hover:bg-blue-700">العودة</button>
      </div>
    </div>
  </div>;
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function Reconciliation2({ onBack, onStageAToMain, onStageBToMain, visaInvoices = [], initialStage = "hub" }: { onBack: () => void; onStageAToMain?: (results: StageAResult[]) => void; onStageBToMain?: () => void; visaInvoices?: StageAResult[]; initialStage?: "hub" | "stageA" | "stageB" | "stageC" | "stageD" | "stageF" }) {
  const [stagePage, setStagePage] = useState<"hub" | "stageA" | "stageB" | "stageC" | "stageD" | "stageF">(initialStage);
  const [stageDMatches, setStageDMatches] = useState<StageCMatch[]>([]);
  const [stageDPending, setStageDPending] = useState<StageDPending[]>([]);
  const [stageBInvoices, setStageBInvoices] = useState<StageAResult[]>([]);
  const [stageCInvoices, setStageCInvoices] = useState<StageAResult[]>(visaInvoices);
  const [sessionLoaded, setSessionLoaded] = useState(false);

  // ─── File states ──────────────────────────────────────────────────────────
  // 1. كشف البنك (يُفرز منه مشتريات عمولة تجار تلقائياً لخانة "فيزا بنك")
  const [bankFile, setBankFile] = useState<File|null>(null);
  const [bankHeaders, setBankHeaders] = useState<string[]>([]);
  const [bankRowsRaw, setBankRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [bankMap, setBankMap] = useState({ date:"", desc:"", debit:"", credit:"", accountType:"", reference:"" });

  // 2. كشف الأستاذ (الاستاذ) — الفواتير الكاملة (~1050)
  const [azaFile, setAzaFile] = useState<File|null>(null);
  const [azaHeaders, setAzaHeaders] = useState<string[]>([]);
  const [azaRowsRaw, setAzaRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [azaMap, setAzaMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"" });

  // 3. كشف اليومي كاشير — حوالات يوم كامل (~1000)
  const [dailyFile, setDailyFile] = useState<File|null>(null);
  const [dailyHeaders, setDailyHeaders] = useState<string[]>([]);
  const [dailyRowsRaw, setDailyRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [dailyMap, setDailyMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"" });

  // 4. سوني كاشير فيزا — اسم الزبون + رقم 4 أرقام + مبلغ
  const [sonyFile, setSonyFile] = useState<File|null>(null);
  const [sonyHeaders, setSonyHeaders] = useState<string[]>([]);
  const [sonyRowsRaw, setSonyRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [sonyMap, setSonyMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", authNum:"" });

  // 5. رقم التفويض فيزا — ملف أرقام التفويض
  const [authFile, setAuthFile] = useState<File|null>(null);
  const [authHeaders, setAuthHeaders] = useState<string[]>([]);
  const [authRowsRaw, setAuthRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [authMap, setAuthMap] = useState({ name:"", authNum:"", amount:"", date:"" });

  // ─── Matched & pending state ──────────────────────────────────────────────
  const [matchedVisa, setMatchedVisa] = useState<MatchedVisa[]>([]);
  const [pendingVisaBank, setPendingVisaBank] = useState<PendingVisaBank[]>([]);
  const [rejectedPairs, setRejectedPairs] = useState<Set<string>>(new Set());
  const [tab, setTab] = useState<R2Tab>("overview");
  const [error, setError] = useState<string|null>(null);
  const [expandedKey, setExpandedKey] = useState<string|null>(null);
  const [visaSwaps, setVisaSwaps] = useState<Record<string, boolean>>({ bank: false, aza: false, daily: false, sony: false, auth: false });

  // ─── لصق إيصال SoftPOS وتحويله لملف سوني كاشير فيزا ──────────────────────
  const [softposText, setSoftposText] = useState("");

  // ─── Session restore ──────────────────────────────────────────────────────
  useEffect(() => {
    (async () => {
      const s = await storageGet<any>("recon2_session", null);
      if (s) {
        setBankHeaders(s.bankHeaders || []); setBankRowsRaw(s.bankRowsRaw || []);
        setBankMap({ date:"", desc:"", debit:"", credit:"", accountType:"", reference:"", ...(s.bankMap||{}) });
        setAzaHeaders(s.azaHeaders || []); setAzaRowsRaw(s.azaRowsRaw || []);
        setAzaMap({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"", ...(s.azaMap||{}) });
        setDailyHeaders(s.dailyHeaders || []); setDailyRowsRaw(s.dailyRowsRaw || []);
        setDailyMap({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"", ...(s.dailyMap||{}) });
        setSonyHeaders(s.sonyHeaders || []); setSonyRowsRaw(s.sonyRowsRaw || []);
        setSonyMap({ date:"", name:"", debit:"", credit:"", accountType:"", authNum:"", ...(s.sonyMap||{}) });
        setAuthHeaders(s.authHeaders || []); setAuthRowsRaw(s.authRowsRaw || []);
        setAuthMap({ name:"", authNum:"", amount:"", date:"", ...(s.authMap||{}) });
        setMatchedVisa(s.matchedVisa || []);
        setPendingVisaBank(s.pendingVisaBank || []);
        setRejectedPairs(new Set(s.rejectedPairs || []));
        if (s.visaSwaps) setVisaSwaps(s.visaSwaps);
      }
      setSessionLoaded(true);
    })();
  }, []);

  // ─── Auto-save ─────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!sessionLoaded) return;
    const t = setTimeout(() => {
      storageSet("recon2_session", {
        bankHeaders, bankRowsRaw, bankMap,
        azaHeaders, azaRowsRaw, azaMap,
        dailyHeaders, dailyRowsRaw, dailyMap,
        sonyHeaders, sonyRowsRaw, sonyMap,
        authHeaders, authRowsRaw, authMap,
        visaSwaps, matchedVisa, pendingVisaBank,
        rejectedPairs: Array.from(rejectedPairs),
      });
    }, 500);
    return () => clearTimeout(t);
  }, [sessionLoaded, bankHeaders, bankRowsRaw, bankMap,
      azaHeaders, azaRowsRaw, azaMap,
      dailyHeaders, dailyRowsRaw, dailyMap,
      sonyHeaders, sonyRowsRaw, sonyMap,
      authHeaders, authRowsRaw, authMap,
      visaSwaps, matchedVisa, pendingVisaBank, rejectedPairs]);

  // ─── File loaders ──────────────────────────────────────────────────────────
  const loadBank = async (f: File) => {
    setBankFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setBankHeaders(headers); setBankRowsRaw(rows);
      setBankMap({ date:autoDetect(headers,HINTS.date), desc:autoDetect(headers,HINTS.desc),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadAza = async (f: File) => {
    setAzaFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setAzaHeaders(headers); setAzaRowsRaw(rows);
      setAzaMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadDaily = async (f: File) => {
    setDailyFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setDailyHeaders(headers); setDailyRowsRaw(rows);
      setDailyMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadSony = async (f: File) => {
    setSonyFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setSonyHeaders(headers); setSonyRowsRaw(rows);
      setSonyMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), authNum:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadAuth = async (f: File) => {
    setAuthFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setAuthHeaders(headers); setAuthRowsRaw(rows);
      setAuthMap({ name:autoDetect(headers,HINTS.name), authNum:autoDetect(headers,HINTS.authNum),
        amount:autoDetect(headers,HINTS.debit), date:autoDetect(headers,HINTS.date) });
    } catch(e) { setError((e as Error).message); }
  };

  // ─── Parse bank rows + auto-separate مشتريات عمولة تجار ─────────────────────
  const isVisaBankDesc = useCallback((desc: string): boolean => {
    const visaBankKeywords = ["مشتريات عمولة تجار", "مشتريات تجار", "عمولة تجار", "مشتريات عمولة"];
    const d = desc.toLowerCase();
    return visaBankKeywords.some(kw => d.includes(kw.toLowerCase()));
  }, []);

  const parsedBank = useMemo((): BankRow2[] => {
    const swap = visaSwaps.bank;
    return bankRowsRaw.map((r, i) => {
      const desc = String(bankMap.desc ? r[bankMap.desc] : "").trim();
      const dRaw = Math.abs(toNum(bankMap.debit ? r[bankMap.debit] : 0));
      const cRaw = Math.abs(toNum(bankMap.credit ? r[bankMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let rawAmount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { rawAmount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { rawAmount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { rawAmount = debitRaw; type = "مدفوع"; }
      if (rawAmount === 0) return null;
      return {
        id: i,
        date: fmtDate(bankMap.date ? r[bankMap.date] : ""),
        description: desc,
        rawAmount,
        type,
        accountType: String(bankMap.accountType ? r[bankMap.accountType] : "").trim(),
        isVisaBank: isVisaBankDesc(desc),
        orig: r,
      };
    }).filter((r): r is BankRow2 => r !== null);
  }, [bankRowsRaw, bankMap, isVisaBankDesc, visaSwaps.bank]);

  // فيزا بنك = مشتريات عمولة تجار المفرزة تلقائياً
  const visaBankRows = useMemo(() => parsedBank.filter(b => b.isVisaBank), [parsedBank]);
  const regularBankRows = useMemo(() => parsedBank.filter(b => !b.isVisaBank), [parsedBank]);

  // ─── Parse أستاذ rows ──────────────────────────────────────────────────────
  const parsedAza = useMemo((): AzaRow[] => {
    const swap = visaSwaps.aza;
    return azaRowsRaw.map((r, i) => {
      const rawName = String(azaMap.name ? r[azaMap.name] : "").trim();
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(azaMap.debit ? r[azaMap.debit] : 0));
      const cRaw = Math.abs(toNum(azaMap.credit ? r[azaMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { amount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { amount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { amount = debitRaw; type = "مدفوع"; }
      if (amount === 0) return null;
      return {
        id: i, rawName, name, amount, type,
        accountType: String(azaMap.accountType ? r[azaMap.accountType] : "").trim(),
        date: fmtDate(azaMap.date ? r[azaMap.date] : ""),
        visaNumber: extractVisaNumber(rawName),
        orig: r,
      };
    }).filter((r): r is AzaRow => r !== null);
  }, [azaRowsRaw, azaMap, visaSwaps.aza]);

  // أستاذ فيزا = اللي فيها رقم فيزا (4 أرقام)
  const azaVisaRows = useMemo(() => parsedAza.filter(a => a.visaNumber), [parsedAza]);
  const azaRegularRows = useMemo(() => parsedAza.filter(a => !a.visaNumber), [parsedAza]);

  // ─── Parse daily cashier rows ──────────────────────────────────────────────
  const parsedDaily = useMemo((): AzaRow[] => {
    const swap = visaSwaps.daily;
    return dailyRowsRaw.map((r, i) => {
      const rawName = String(dailyMap.name ? r[dailyMap.name] : "").trim();
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(dailyMap.debit ? r[dailyMap.debit] : 0));
      const cRaw = Math.abs(toNum(dailyMap.credit ? r[dailyMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { amount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { amount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { amount = debitRaw; type = "مدفوع"; }
      if (amount === 0) return null;
      return {
        id: i, rawName, name, amount, type,
        accountType: String(dailyMap.accountType ? r[dailyMap.accountType] : "").trim(),
        date: fmtDate(dailyMap.date ? r[dailyMap.date] : ""),
        visaNumber: extractVisaNumber(rawName),
        orig: r,
      };
    }).filter((r): r is AzaRow => r !== null);
  }, [dailyRowsRaw, dailyMap, visaSwaps.daily]);

  // الفرق: أستاذ - يومي = الـ50 فاتورة الزيادة (مش موجودة باليومي)
  const azaOnlyRows = useMemo(() => {
    const dailyNames = new Set(parsedDaily.map(d => normName(d.name)));
    return azaRegularRows.filter(a => !dailyNames.has(normName(a.name)));
  }, [azaRegularRows, parsedDaily]);

  // ─── Parse سوني كاشير فيزا ─────────────────────────────────────────────────
  const parsedSony = useMemo((): SonyVisaRow[] => {
    const swap = visaSwaps.sony;
    return sonyRowsRaw.map((r, i) => {
      const rawName = String(sonyMap.name ? r[sonyMap.name] : "").trim();
      const visaNum = extractVisaNumber(rawName);
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(sonyMap.debit ? r[sonyMap.debit] : 0));
      const cRaw = Math.abs(toNum(sonyMap.credit ? r[sonyMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      if (debitRaw > 0) amount = debitRaw;
      else if (creditRaw > 0) amount = creditRaw;
      if (amount === 0 && !visaNum) return null;
      const authNum = sonyMap.authNum ? String(r[sonyMap.authNum] || "").trim() : null;
      return {
        id: i, rawName, name, amount, visaNumber: visaNum, authNumber: authNum || null,
        date: fmtDate(sonyMap.date ? r[sonyMap.date] : ""), orig: r,
      };
    }).filter((r): r is SonyVisaRow => r !== null);
  }, [sonyRowsRaw, sonyMap, visaSwaps.sony]);

  // ─── Parse رقم التفويض ─────────────────────────────────────────────────────
  const parsedAuth = useMemo(() => {
    return authRowsRaw.map((r, i) => {
      const name = String(authMap.name ? r[authMap.name] : "").trim();
      const authNum = String(authMap.authNum ? r[authMap.authNum] : "").trim();
      const amount = toNum(authMap.amount ? r[authMap.amount] : 0);
      return { id: i, name, authNumber: authNum, amount, orig: r };
    }).filter(a => a.authNumber || a.name);
  }, [authRowsRaw, authMap]);

  // ─── Matched IDs ───────────────────────────────────────────────────────────
  const matchedSonyIds = useMemo(() => new Set(matchedVisa.map(m => m.sonyId)), [matchedVisa]);
  const matchedAzaIds = useMemo(() => new Set(matchedVisa.map(m => m.azaId)), [matchedVisa]);

  // ─── Auto-match: سوني كاشير فيزا ↔ أستاذ فيزا ──────────────────────────────
  // المطابقة بالاسم + المبلغ + رقم الفيزا (إن وجد)
  const autoMatchResults = useMemo(() => {
    const results: Array<{
      sony: SonyVisaRow; aza: AzaRow; score: number; matchType: string;
      amountDiff: number; visaMatch: boolean;
    }> = [];
    const usedAza = new Set<number>(matchedAzaIds);
    const usedSony = new Set<number>(matchedSonyIds);

    parsedSony.forEach(s => {
      if (usedSony.has(s.id)) return;
      let best: { aza: AzaRow; score: number; matchType: string; amountDiff: number; visaMatch: boolean } | null = null;

      azaVisaRows.forEach(a => {
        if (usedAza.has(a.id)) return;
        const pairKey = `${s.id}-${a.id}`;
        if (rejectedPairs.has(pairKey)) return;

        const amtMatch = Math.abs(a.amount - s.amount) <= 0.01;
        const ms = advancedMatchCheck(s.name, a.name);
        let visaMatch = false;
        if (s.visaNumber && a.visaNumber) visaMatch = s.visaNumber === a.visaNumber;

        let score = 0;
        if (ms.isMatch && amtMatch) score = 0.95;
        else if (ms.isMatch && !amtMatch) score = 0.5 + nameSim(s.name, a.name) * 0.3;
        else if (!ms.isMatch && amtMatch) score = 0.2;
        if (visaMatch) score += 0.1;
        if (score < 0.15) return;

        if (!best || score > best.score) {
          best = { aza: a, score, matchType: ms.matchType, amountDiff: Math.abs(a.amount - s.amount), visaMatch };
        }
      });

      const selected = best as { aza: AzaRow; score: number; matchType: string; amountDiff: number; visaMatch: boolean } | null;
      if (selected && selected.score >= 0.15) {
        results.push({ sony: s, aza: selected.aza, score: selected.score, matchType: selected.matchType, amountDiff: selected.amountDiff, visaMatch: selected.visaMatch });
      }
    });

    return results.sort((a, b) => b.score - a.score);
  }, [parsedSony, azaVisaRows, rejectedPairs, matchedSonyIds, matchedAzaIds]);

  // ─── Unmatched سوني ─────────────────────────────────────────────────────────
  const unmatchedSony = useMemo(() => {
    const matchedIds = new Set<number>([...matchedSonyIds, ...autoMatchResults.map(r => r.sony.id)]);
    return parsedSony.filter(s => !matchedIds.has(s.id));
  }, [parsedSony, matchedSonyIds, autoMatchResults]);

  // ─── Unmatched أستاذ فيزا ──────────────────────────────────────────────────
  const unmatchedAzaVisa = useMemo(() => {
    const matchedIds = new Set<number>([...matchedAzaIds, ...autoMatchResults.map(r => r.aza.id)]);
    return azaVisaRows.filter(a => !matchedIds.has(a.id));
  }, [azaVisaRows, matchedAzaIds, autoMatchResults]);

  // ─── Handlers ──────────────────────────────────────────────────────────────
  const handleAcceptMatch = (sony: SonyVisaRow, aza: AzaRow, note?: string) => {
    const pairKey = `${sony.id}-${aza.id}`;
    if (matchedSonyIds.has(sony.id) || matchedAzaIds.has(aza.id)) return;
    const m: MatchedVisa = {
      id: `mv-${Date.now()}-${sony.id}-${aza.id}`,
      sonyId: sony.id, azaId: aza.id,
      sonyName: sony.name, azaName: aza.name,
      visaNumber: sony.visaNumber || aza.visaNumber || "",
      authNumber: sony.authNumber,
      amount: aza.amount,
      sonyRawName: sony.rawName, azaRawName: aza.rawName,
      matchedAt: new Date().toLocaleString("ar-SA"),
      note,
    };
    setMatchedVisa(prev => [...prev, m]);
    const next = new Set(rejectedPairs);
    next.delete(pairKey);
    setRejectedPairs(next);
    setExpandedKey(null);
  };

  const handleRejectMatch = (sonyId: number, azaId: number) => {
    const pairKey = `${sonyId}-${azaId}`;
    setRejectedPairs(prev => new Set(prev).add(pairKey));
    setExpandedKey(null);
  };

  const handleUnmatch = (m: MatchedVisa) => {
    if (!window.confirm("إلغاء هذه المطابقة؟")) return;
    setMatchedVisa(prev => prev.filter(x => x.id !== m.id));
  };

  const handleClearSaved = () => {
    if (!matchedVisa.length && !pendingVisaBank.length) return;
    if (!window.confirm("مسح كل المطابقات المحفوظة والمعلقات في منصة الفيزا؟")) return;
    setMatchedVisa([]);
    setPendingVisaBank([]);
    setRejectedPairs(new Set());
    setExpandedKey(null);
  };

  const handleClearSession = () => {
    if (!window.confirm("مسح جلسة منصة الفيزا كاملة؟ سيتم حذف الملفات والنتائج المحفوظة.")) return;
    setBankFile(null); setBankHeaders([]); setBankRowsRaw([]);
    setAzaFile(null); setAzaHeaders([]); setAzaRowsRaw([]);
    setDailyFile(null); setDailyHeaders([]); setDailyRowsRaw([]);
    setSonyFile(null); setSonyHeaders([]); setSonyRowsRaw([]);
    setAuthFile(null); setAuthHeaders([]); setAuthRowsRaw([]);
    setMatchedVisa([]); setPendingVisaBank([]); setRejectedPairs(new Set());
    setSoftposText(""); setError(null); setTab("overview");
    void storageSet("recon2_session", null);
  };

  const handleAcceptAllAuto = () => {
    if (!autoMatchResults.length) return;
    if (!window.confirm(`تأكيد ${autoMatchResults.length} مطابقة تلقائية؟`)) return;
    const newMatches: MatchedVisa[] = [];
    const usedAzaNow = new Set(matchedAzaIds);
    const usedSonyNow = new Set(matchedSonyIds);
    autoMatchResults.forEach(r => {
      if (usedSonyNow.has(r.sony.id) || usedAzaNow.has(r.aza.id)) return;
      usedSonyNow.add(r.sony.id);
      usedAzaNow.add(r.aza.id);
      newMatches.push({
        id: `mv-${Date.now()}-${r.sony.id}-${r.aza.id}`,
        sonyId: r.sony.id, azaId: r.aza.id,
        sonyName: r.sony.name, azaName: r.aza.name,
        visaNumber: r.sony.visaNumber || r.aza.visaNumber || "",
        authNumber: r.sony.authNumber,
        amount: r.aza.amount,
        sonyRawName: r.sony.rawName, azaRawName: r.aza.rawName,
        matchedAt: new Date().toLocaleString("ar-SA"),
        note: r.amountDiff > 0.01 ? `فرق مبلغ: ${fmtNum(r.amountDiff)}` : undefined,
      });
    });
    setMatchedVisa(prev => [...prev, ...newMatches]);
  };

  // ─── نقل مشتريات عمولة تجار للمعلقة (للمطابقة ثاني يوم) ──────────────────────
  const handleMoveVisaBankToPending = () => {
    if (!visaBankRows.length) return;
    if (!window.confirm(`نقل ${visaBankRows.length} حوالة "مشتريات عمولة تجار" للمعلقة؟\nستطابقها ثاني يوم مع سوني كاشير فيزا.`)) return;
    const items: PendingVisaBank[] = visaBankRows.map(b => ({
      id: `pb-${Date.now()}-${b.id}`,
      bankId: b.id,
      bankDesc: b.description,
      authNumber: null,
      amount: b.rawAmount,
      date: b.date,
      movedAt: new Date().toLocaleString("ar-SA"),
    }));
    setPendingVisaBank(prev => [...prev, ...items]);
  };

  // ─── مطابقة المعلقة مع سوني مؤكد ─────────────────────────────────────────────
  // ثاني يوم: نطابق pendingVisaBank (مشتريات البنك) مع matchedVisa (سوني مؤكد)
  // عبر رقم التفويض إن وجد، أو بالاسم + المبلغ
  const pendingMatchResults = useMemo(() => {
    if (!pendingVisaBank.length || !matchedVisa.length) return [];
    const results: Array<{ pending: PendingVisaBank; matched: MatchedVisa; score: number; reason: string }> = [];
    const usedMatched = new Set<string>();

    pendingVisaBank.forEach(p => {
      let best: { matched: MatchedVisa; score: number; reason: string } | null = null;
      matchedVisa.forEach(m => {
        if (usedMatched.has(m.id)) return;
        // مطابقة برقم التفويض إن وجد
        if (m.authNumber && p.authNumber && m.authNumber === p.authNumber) {
          if (!best || 1 > best.score) best = { matched: m, score: 1, reason: "رقم تفويض مطابق" };
          return;
        }
        // مطابقة بالاسم + المبلغ
        const amtMatch = Math.abs(p.amount - m.amount) <= 0.01;
        const ms = advancedMatchCheck(m.sonyName, p.bankDesc);
        if (ms.isMatch && amtMatch) {
          if (!best || 0.9 > best.score) best = { matched: m, score: 0.9, reason: "اسم ومبلغ متطابقان" };
        } else if (amtMatch) {
          if (!best || 0.3 > best.score) best = { matched: m, score: 0.3, reason: "مبلغ متطابق فقط" };
        }
      });
      const selected = best as { matched: MatchedVisa; score: number; reason: string } | null;
      if (selected) {
        results.push({ pending: p, matched: selected.matched, score: selected.score, reason: selected.reason });
        usedMatched.add(selected.matched.id);
      }
    });
    return results.sort((a, b) => b.score - a.score);
  }, [pendingVisaBank, matchedVisa]);

  // ─── لصق إيصال SoftPOS وتحويله لملف سوني كاشير فيزا ──────────────────────
  // النص بيحتوي على عدة حركات، كل حركة فيها:
  //   - رقم البطاقة (****-4351) → رقم الفيزا (4 أرقام)
  //   - المبلغ (ILS 40.00 أو بمبلغ 40.00 ILS)
  //   - رقم التفويض (236189)
  //   - التاريخ والوقت (2026-08-25 08:44 AM)
  //   - نوع البطاقة (Visa / MasterCard)
  // البيانات كلها بأسماء أرقام بس (بدون اسم زبون) — الاسم بيجي من ملف الأستاذ لاحقاً
  const handleParseSoftPOS = () => {
    const text = softposText.trim();
    if (!text) { setError("الصق نص إيصال SoftPOS أولاً"); return; }
    setError(null);

    // نقسم النص لحركات — كل حركة تبدأ بـ "حركة على بطاقة" أو "حركة"
    const lines = text.split("\n").map(l => l.trim());
    const parseReceiptAmount = (value: string) => {
      const normalized = value.replace(/[٫]/g, ".").replace(/[٬]/g, "").replace(/,/g, ".");
      return Number.parseFloat(normalized) || 0;
    };
    const merchant = lines.find(line => line.startsWith("SoftPOS Receipt") === false && !line.includes(":")) || "";
    const deviceNumber = lines.find(line => line.includes("رقم الجهاز"))?.match(/(\d+)\s*:\s*رقم الجهاز/)?.[1] || "";
    const outletNumber = lines.find(line => line.includes("رقم المخرج"))?.match(/(\d+)\s*:\s*رقم المخرج/)?.[1] || "";
    const summary = lines.find(line => line.includes("عدد حركات"));
    const declaredCount = summary?.match(/عدد حركات\s*=\s*(\d+)/)?.[1] || "";
    const declaredTotal = summary?.match(/مجموع\s*=\s*([\d٫.,]+)\s*ILS/i)?.[1] || "";
    const transactions: Array<{
      visaNumber: string; amount: number; authNumber: string;
      date: string; cardType: string; txId: string; merchant: string;
    }> = [];

    let current: Partial<{
      visaNumber: string; amount: number; authNumber: string;
      date: string; cardType: string; txId: string; merchant: string;
    }> = {};

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      // بداية حركة جديدة
      if (line.includes("حركة على بطاقة") || (line.includes("حركة") && line.includes("بمبلغ"))) {
        if (current.visaNumber || current.amount) {
          transactions.push(current as any);
        }
        current = {};
        const merchantMatch = line.match(/من المتجر\s+(.+)$/);
        if (merchantMatch) current.merchant = merchantMatch[1].trim();
        // استخراج رقم البطاقة من السطر: ****-4351
        const cardMatch = line.match(/\*{4}[-\s]*(\d{4})/);
        if (cardMatch) current.visaNumber = cardMatch[1];
        // استخراج المبلغ: بمبلغ 40.00 ILS أو بمبلغ 40٫00 ILS
        const amtMatch = line.match(/بمبلغ\s*([\d٫.,]+)\s*ILS/i) || line.match(/([\d٫.,]+)\s*ILS/i);
        if (amtMatch) current.amount = parseReceiptAmount(amtMatch[1]);
        // نوع البطاقة
        if (line.includes("Visa")) current.cardType = "Visa";
        else if (line.includes("MasterCard")) current.cardType = "MasterCard";
      }

      // رقم التفويض: "236189 رقم التفويض" أو "رقم التفويض 236189"
      if (line.includes("رقم التفويض")) {
        const m = line.match(/(\d{4,})/);
        if (m) current.authNumber = m[1];
      }

      // التاريخ والوقت: "2026-08-25 08:44 AM الوقت والتاريخ" أو العكس
      if (line.includes("الوقت والتاريخ") || line.includes("التاريخ والوقت")) {
        const m = line.match(/(\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)/i);
        if (m) current.date = m[1].trim();
      }

      // رقم تعريف وتسلسل الحركة: "13937 - 499538 رقم تعريف وتسلسل الحركة"
      if (line.includes("رقم تعريف") && line.includes("تسلسل")) {
        const m = line.match(/(\d+)\s*[-\s]\s*(\d+)/);
        if (m) current.txId = `${m[1]}-${m[2]}`;
      }

      // نوع ورقم البطاقة: "Visa / ****-4351 نوع ورقم البطاقة"
      if (line.includes("نوع ورقم البطاقة")) {
        const cardMatch = line.match(/\*{4}[-\s]*(\d{4})/);
        if (cardMatch && !current.visaNumber) current.visaNumber = cardMatch[1];
        if (line.includes("Visa")) current.cardType = "Visa";
        else if (line.includes("MasterCard")) current.cardType = "MasterCard";
      }

      // المبلغ في سطر مستقل: "ILS 40.00 المبلغ"
      if (line.includes("المبلغ") && !current.amount) {
        const m = line.match(/([\d٫.,]+)\s*ILS/i) || line.match(/ILS\s*([\d٫.,]+)/i);
        if (m) current.amount = parseReceiptAmount(m[1]);
      }
    }
    // آخر حركة
    if (current.visaNumber || current.amount) {
      transactions.push(current as any);
    }

    if (!transactions.length) {
      setError("لم يتم العثور على أي حركة في النص. تأكد إنك لصقت إيصال SoftPOS كامل.");
      return;
    }

    // إنشاء ملف Excel: سوني كاشير فيزا
    const wb = XLSX.utils.book_new();
    const rows = transactions.map((t, i) => [
      i + 1,
      t.merchant || merchant,
      deviceNumber,
      outletNumber,
      declaredCount,
      parseReceiptAmount(declaredTotal),
      t.visaNumber || "",
      t.authNumber || "",
      t.amount || 0,
      t.date || "",
      t.cardType || "",
      t.txId || "",
    ]);
    const ws = XLSX.utils.aoa_to_sheet([
      ["#", "اسم المتجر", "رقم الجهاز", "رقم المخرج", "عدد الحركات المعلن", "مجموع الحركات المعلن", "رقم الفيزا", "رقم التفويض", "المبلغ", "التاريخ والوقت", "نوع البطاقة", "رقم التعريف والتسلسل"],
      ...rows,
    ]);
    ws["!cols"] = [{wch:5},{wch:28},{wch:14},{wch:14},{wch:18},{wch:20},{wch:12},{wch:15},{wch:12},{wch:22},{wch:14},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "سوني كاشير فيزا");
    XLSX.writeFile(wb, "سوني_كاشير_فيزا.xlsx");
  };

  // ─── Export ─────────────────────────────────────────────────────────────────
  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    const add = (name: string, hdrs: string[], rows: unknown[][]) => {
      const ws = XLSX.utils.aoa_to_sheet([hdrs, ...rows]);
      ws["!cols"] = hdrs.map(() => ({ wch: 25 }));
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };

    add("💳 فيزا بنك (مشتريات تجار)", ["البيان","المبلغ","النوع","التاريخ"],
      visaBankRows.map(b => [b.description, b.rawAmount, b.type, b.date]));

    add("📋 أستاذ - فيزا", ["الاسم","رقم الفيزا","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaVisaRows.map(a => [a.name, a.visaNumber, a.amount, a.type, a.accountType, a.date]));

    add("📋 أستاذ - عادي", ["الاسم","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaRegularRows.map(a => [a.name, a.amount, a.type, a.accountType, a.date]));

    add("📋 أستاذ فقط (زيادة)", ["الاسم","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaOnlyRows.map(a => [a.name, a.amount, a.type, a.accountType, a.date]));

    add("🏪 سوني كاشير فيزا", ["الاسم","رقم الفيزا","رقم التفويض","المبلغ","التاريخ"],
      parsedSony.map(s => [s.name, s.visaNumber, s.authNumber, s.amount, s.date]));

    add("✅ مطابقات مؤكدة", ["اسم سوني","اسم أستاذ","رقم الفيزا","رقم التفويض","المبلغ","تاريخ المطابقة","ملاحظة"],
      matchedVisa.map(m => [m.sonyName, m.azaName, m.visaNumber, m.authNumber, m.amount, m.matchedAt, m.note || ""]));

    add("⏳ مشتريات معلقة (لثاني يوم)", ["البيان","المبلغ","التاريخ","تاريخ النقل"],
      pendingVisaBank.map(p => [p.bankDesc, p.amount, p.date, p.movedAt]));

    add("🔄 مطابقات المعلقة", ["بيان البنك","اسم سوني","المبلغ","سبب المطابقة"],
      pendingMatchResults.map(r => [r.pending.bankDesc, r.matched.sonyName, r.pending.amount, r.reason]));

    add("❌ سوني غير متطابق", ["الاسم","رقم الفيزا","رقم التفويض","المبلغ","التاريخ"],
      unmatchedSony.map(s => [s.name, s.visaNumber, s.authNumber, s.amount, s.date]));

    add("❌ أستاذ فيزا غير متطابق", ["الاسم","رقم الفيزا","المبلغ","النوع","التاريخ"],
      unmatchedAzaVisa.map(a => [a.name, a.visaNumber, a.amount, a.type, a.date]));

    XLSX.writeFile(wb, "تقرير_التسوية_2.xlsx");
  };

  // ─── Stats ──────────────────────────────────────────────────────────────────
  const stats = useMemo(() => ({
    bankTotal: parsedBank.length,
    visaBank: visaBankRows.length,
    azaTotal: parsedAza.length,
    azaVisa: azaVisaRows.length,
    azaRegular: azaRegularRows.length,
    azaOnly: azaOnlyRows.length,
    dailyTotal: parsedDaily.length,
    sonyTotal: parsedSony.length,
    matched: matchedVisa.length,
    pendingBank: pendingVisaBank.length,
    pendingMatched: pendingMatchResults.length,
    unmatchedSony: unmatchedSony.length,
    unmatchedAza: unmatchedAzaVisa.length,
    autoMatch: autoMatchResults.length,
  }), [parsedBank, visaBankRows, parsedAza, azaVisaRows, azaRegularRows, azaOnlyRows,
       parsedDaily, parsedSony, matchedVisa, pendingVisaBank, pendingMatchResults,
       unmatchedSony, unmatchedAzaVisa, autoMatchResults]);

  const canExport = parsedBank.length > 0 || parsedAza.length > 0 || parsedSony.length > 0;

  if (stagePage === "stageA") {
    return <StageAPlatform onBack={() => setStagePage("hub")} onTransferToB={results => { if (onStageAToMain) onStageAToMain(results); else { setStageBInvoices(results); setStagePage("stageB"); } }} onTransferVisaToC={results => { setStageCInvoices(results.filter(result => result.category === "فيزا")); setStagePage("stageC"); }} />;
  }
  if (stagePage === "stageB") {
    return <EmptyStageB onBack={() => setStagePage("hub")} />;
  }
  if (stagePage === "stageC") {
    return <StageCPlatform invoices={stageCInvoices} onBack={() => setStagePage("hub")} onTransferToD={(matches, pending) => { setStageDMatches(matches); setStageDPending(pending); setStagePage("stageD"); }} />;
  }
  if (stagePage === "stageD") {
    return <StageDPlatform matches={stageDMatches} pending={stageDPending} onBack={() => setStagePage("stageC")} />;
  }
  if (stagePage === "stageF") {
    return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-4xl space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-3"><button onClick={() => setStagePage("hub")} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-900"><ChevronLeft className="h-4 w-4" />العودة لمنصة الفيزا</button><h1 className="text-xl font-black">F — تحويل إيصال SoftPOS إلى Excel</h1></div>
      <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm"><div className="mb-3 flex items-center gap-2"><FileSpreadsheet className="h-5 w-5 text-slate-700" /><div><h2 className="text-sm font-bold">لصق إيصال SoftPOS</h2><p className="text-xs text-slate-500">الصق نص الإيصال لتحويل كل حركة إلى صف مستقل في ملف سوني كاشير فيزا</p></div></div><textarea value={softposText} onChange={e => setSoftposText(e.target.value)} placeholder="الصق نص إيصال SoftPOS هنا..." className="h-72 w-full resize-y rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs font-mono outline-none focus:ring-2 focus:ring-slate-200" />
      <div className="mt-3 flex flex-wrap gap-2"><button onClick={handleParseSoftPOS} disabled={!softposText.trim()} className="flex items-center gap-2 rounded-lg bg-slate-700 px-4 py-2 text-sm font-medium text-white hover:bg-slate-800 disabled:opacity-40"><Download className="h-4 w-4" />إنشاء ملف Excel</button>{softposText.trim() && <button onClick={() => setSoftposText("")} className="flex items-center gap-1 rounded-lg bg-slate-100 px-3 py-2 text-xs text-slate-600 hover:bg-slate-200"><X className="h-3.5 w-3.5" />مسح النص</button>}</div></div>
      <div className="rounded-xl border border-blue-200 bg-blue-50 p-4 text-xs leading-6 text-blue-900">يتم استخراج رقم البطاقة، المبلغ، رقم التفويض، الوقت والتاريخ، نوع البطاقة، ورقم تعريف وتسلسل الحركة، مع إضافة بيانات المتجر والجهاز والمخرج إلى ملف Excel.</div>
    </div></div>;
  }

  return (
    <div dir="rtl" className="min-h-screen bg-slate-50 text-slate-900 p-6">
      <div className="max-w-6xl mx-auto space-y-6">

        {/* Header */}
        <div className="relative overflow-hidden rounded-2xl border border-slate-200 bg-white px-5 py-6 text-center shadow-sm">
          <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center overflow-hidden rounded-2xl bg-white ring-4 ring-amber-100 shadow-lg">
            <img src="/taswiyah-logo.png" alt="Taswiyah Hub" className="h-full w-full object-contain p-1" />
          </div>
          <button onClick={onBack} className="absolute right-5 top-5 flex items-center gap-1.5 text-sm text-slate-500 transition-colors hover:text-indigo-600">
              <ChevronLeft className="w-4 h-4"/>المنصة الأولى
          </button>
           <div className="mb-1 text-xs font-bold tracking-[.18em] text-amber-600">TASWIYAH HUB</div>
           <h1 className="text-3xl font-black tracking-tight text-slate-900">محطتك الأولى لتسوية الحسابات بدقة.</h1>
           <p className="mx-auto mt-2 max-w-2xl text-sm text-slate-500">اختر المرحلة المناسبة لمعالجة الفواتير والحوالات والفيزا</p>
          <div className="absolute left-5 top-5 flex items-center gap-2">
          {canExport && (
            <button onClick={handleExport} className="absolute left-5 top-5 flex items-center gap-2 rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-green-700">
              <Download className="w-4 h-4"/>تصدير Excel
            </button>
          )}
          <div className="relative">
            <button onClick={e => e.currentTarget.nextElementSibling?.classList.toggle("hidden")} className="flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs font-medium text-red-700 hover:bg-red-100">
              <Trash2 className="h-3.5 w-3.5" />المحفوظات
            </button>
            <div className="absolute left-0 top-10 z-20 hidden w-56 rounded-xl border border-slate-200 bg-white p-1 text-right shadow-lg">
              <button onClick={handleClearSaved} className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-xs text-red-700 hover:bg-red-50"><Trash2 className="h-3.5 w-3.5" />مسح المطابقات والمعلقات</button>
              <button onClick={handleClearSession} className="flex w-full items-center gap-2 rounded-lg border-t border-slate-100 px-3 py-2 text-xs text-amber-700 hover:bg-amber-50"><RotateCcw className="h-3.5 w-3.5" />مسح الجلسة كاملة</button>
            </div>
          </div>
          </div>
        </div>

        {/* Visa workflow hub */}
        <section className="grid grid-cols-1 gap-4 md:grid-cols-2">
          {[
            { id: "stage-a", code: "A", title: "فرز الفواتير حسب قاعدة البيانات", description: "مطابقة اسم ومبلغ كل فاتورة مع قاعدة البيانات، ثم استخراج المطابقات ومشاكل أخرى.", color: "bg-blue-600", text: "text-blue-600", target: "stage-a" },
            { id: "stage-b", code: "B", title: "المطابقة الذكية للفواتير والحوالات", description: "رفع أربعة ملفات عند العمل المستقل، أو ملفي الحوالات عند ترحيل الفواتير من قاعدة البيانات.", color: "bg-emerald-600", text: "text-emerald-600", target: "r2-bank" },
            { id: "stage-c", code: "C", title: "مطابقة السوني كاشير (فيزا)", description: "مطابقة العمليات المعلقة واسترجاع أرقام التفويض مع بيانات الفيزا.", color: "bg-amber-600", text: "text-amber-600", target: "r2-sony" },
            { id: "stage-d", code: "D", title: "مطابقة الفيزا — المرحلة الثانية", description: "مطابقة أرقام الفيزا المؤكدة مع الحوالات البنكية المعلقة.", color: "bg-indigo-600", text: "text-indigo-600", target: "r2-auth" },
            { id: "stage-f", code: "F", title: "تحويل إيصال SoftPOS إلى Excel", description: "لصق نص الإيصال وتحويل الحركات وأرقام البطاقات والتفويض إلى ملف سوني جاهز.", color: "bg-slate-700", text: "text-slate-700", target: "r2-softpos" },
          ].map(stage => (
            <button key={stage.id} onClick={() => stage.id === "stage-a" ? setStagePage("stageA") : stage.id === "stage-b" ? (onStageBToMain ? onStageBToMain() : setStagePage("stageB")) : stage.id === "stage-c" ? setStagePage("stageC") : stage.id === "stage-d" ? setStagePage("stageD") : stage.id === "stage-f" ? setStagePage("stageF") : document.getElementById(stage.target)?.scrollIntoView({ behavior: "smooth", block: "start" })}
              className="group flex min-h-36 items-center gap-5 rounded-2xl border border-slate-200 bg-white p-5 text-right shadow-sm transition-all hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md">
              <span className={`flex h-14 w-14 shrink-0 items-center justify-center rounded-xl ${stage.color} text-xl font-black text-white shadow-sm`}>{stage.code}</span>
              <span className="min-w-0 flex-1">
                <span className={`mb-1 block text-base font-bold ${stage.text}`}>{stage.title}</span>
                <span className="block text-xs leading-6 text-slate-500">{stage.description}</span>
                <span className={`mt-2 block text-xs font-semibold ${stage.text}`}>فتح المرحلة ←</span>
              </span>
            </button>
          ))}
        </section>

        <div className="rounded-2xl border border-slate-200 bg-white px-5 py-4 shadow-sm">
          <div className="mb-3 flex items-center justify-between gap-3">
            <h2 className="text-sm font-bold text-slate-800">تسلسل حركة البيانات الصحيح</h2>
            <span className="rounded-full bg-slate-100 px-3 py-1 text-[11px] text-slate-500">النظام الحسابي</span>
          </div>
          <div className="flex flex-wrap items-center justify-center gap-2 text-xs font-medium">
            {["A — الفرز والتأكيد", "B — تحويل الفواتير المؤكدة", "C — تصفية الفيزا", "D — قائمة الغد"].map((label, index) => (
              <React.Fragment key={label}>
                <span className="rounded-lg border border-blue-100 bg-blue-50 px-3 py-2 text-blue-700">{label}</span>
                {index < 3 && <ChevronLeft className="h-4 w-4 text-slate-300" />}
              </React.Fragment>
            ))}
          </div>
        </div>

        {error && <div className="p-3 bg-red-100 text-red-800 rounded-lg text-sm">{error}</div>}

        {/* File uploads — 5 files */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* 1. كشف البنك */}
          <div id="r2-bank" className="bg-card border p-4 rounded-xl space-y-4 scroll-mt-5">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              كشف البنك
              {visaBankRows.length > 0 && (
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-teal-100 text-teal-700 border border-teal-300">
                  {visaBankRows.length} فيزا بنك
                </span>
              )}
            </h2>
            {bankHeaders.length > 0 && (
              <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" checked={visaSwaps.bank} onChange={event => setVisaSwaps(prev => ({ ...prev, bank: event.target.checked }))} className="rounded" />
                عكس المدين والدائن
              </label>
            )}
            <DropZone file={bankFile} onFile={loadBank} onClear={() => { setBankFile(null); setBankHeaders([]); setBankRowsRaw([]); }} />
            {bankHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={bankHeaders} value={bankMap.date} onChange={v => setBankMap(m => ({ ...m, date: v }))} />
                <Sel label="الإيضاحات / البيان" headers={bankHeaders} value={bankMap.desc} onChange={v => setBankMap(m => ({ ...m, desc: v }))} />
                <Sel label="المبالغ المدفوعة (Debit)" headers={bankHeaders} value={bankMap.debit} onChange={v => setBankMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المستلمة (Credit)" headers={bankHeaders} value={bankMap.credit} onChange={v => setBankMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب" headers={bankHeaders} value={bankMap.accountType} onChange={v => setBankMap(m => ({ ...m, accountType: v }))} />
                <Sel label="المرجع / رقم الحوالة" headers={bankHeaders} value={bankMap.reference} onChange={v => setBankMap(m => ({ ...m, reference: v }))} />
              </div>
            )}
            {visaBankRows.length > 0 && (
              <div className="flex items-center gap-2 text-xs">
                <span className="text-teal-700 font-medium">تم فرز {visaBankRows.length} حوالة "مشتريات عمولة تجار"</span>
                <button onClick={handleMoveVisaBankToPending}
                  className="flex items-center gap-1 px-3 py-1.5 bg-teal-600 text-white rounded-lg text-xs font-medium hover:bg-teal-700 transition-colors">
                  <Clock className="w-3.5 h-3.5"/>نقل للمعلقة (لثاني يوم)
                </button>
              </div>
            )}
          </div>

          {/* 2. كشف الأستاذ */}
          <div id="r2-aza" className="bg-card border p-4 rounded-xl space-y-4 scroll-mt-5">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              كشف الأستاذ
              {azaVisaRows.length > 0 && (
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-teal-100 text-teal-700 border border-teal-300">
                  {azaVisaRows.length} فيزا
                </span>
              )}
            </h2>
            {azaHeaders.length > 0 && (
              <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" checked={visaSwaps.aza} onChange={event => setVisaSwaps(prev => ({ ...prev, aza: event.target.checked }))} className="rounded" />
                عكس المدين والدائن
              </label>
            )}
            <DropZone file={azaFile} onFile={loadAza} onClear={() => { setAzaFile(null); setAzaHeaders([]); setAzaRowsRaw([]); }} />
            {azaHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={azaHeaders} value={azaMap.date} onChange={v => setAzaMap(m => ({ ...m, date: v }))} />
                <Sel label="البيان" headers={azaHeaders} value={azaMap.name} onChange={v => setAzaMap(m => ({ ...m, name: v }))} />
                <Sel label="المبالغ المدفوعة (Debit)" headers={azaHeaders} value={azaMap.debit} onChange={v => setAzaMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المستلمة (Credit)" headers={azaHeaders} value={azaMap.credit} onChange={v => setAzaMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب" headers={azaHeaders} value={azaMap.accountType} onChange={v => setAzaMap(m => ({ ...m, accountType: v }))} />
                <Sel label="المرجع / رقم الحوالة" headers={azaHeaders} value={azaMap.reference} onChange={v => setAzaMap(m => ({ ...m, reference: v }))} />
              </div>
            )}
          </div>

          {/* 3. كشف اليومي كاشير */}
          <div id="r2-daily" className="bg-card border p-4 rounded-xl space-y-4 scroll-mt-5">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              📅 كشف اليومي كاشير
              {parsedDaily.length > 0 && azaOnlyRows.length > 0 && (
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">
                  {azaOnlyRows.length} زيادة بالأستاذ
                </span>
              )}
            </h2>
            {dailyHeaders.length > 0 && (
              <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" checked={visaSwaps.daily} onChange={event => setVisaSwaps(prev => ({ ...prev, daily: event.target.checked }))} className="rounded" />
                عكس المدين والدائن
              </label>
            )}
            <DropZone file={dailyFile} onFile={loadDaily} onClear={() => { setDailyFile(null); setDailyHeaders([]); setDailyRowsRaw([]); }} />
            {dailyHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={dailyHeaders} value={dailyMap.date} onChange={v => setDailyMap(m => ({ ...m, date: v }))} />
                <Sel label="البيان" headers={dailyHeaders} value={dailyMap.name} onChange={v => setDailyMap(m => ({ ...m, name: v }))} />
                <Sel label="المبالغ المدفوعة (Debit)" headers={dailyHeaders} value={dailyMap.debit} onChange={v => setDailyMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المستلمة (Credit)" headers={dailyHeaders} value={dailyMap.credit} onChange={v => setDailyMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب" headers={dailyHeaders} value={dailyMap.accountType} onChange={v => setDailyMap(m => ({ ...m, accountType: v }))} />
                <Sel label="المرجع / رقم الحوالة" headers={dailyHeaders} value={dailyMap.reference} onChange={v => setDailyMap(m => ({ ...m, reference: v }))} />
              </div>
            )}
          </div>

          {/* 4. سوني كاشير فيزا */}
          <div id="r2-sony" className="bg-card border p-4 rounded-xl space-y-4 scroll-mt-5">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              🏪 سوني كاشير فيزا
              {parsedSony.length > 0 && (
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700 border border-indigo-300">
                  {parsedSony.length} فاتورة
                </span>
              )}
            </h2>
            {sonyHeaders.length > 0 && (
              <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" checked={visaSwaps.sony} onChange={event => setVisaSwaps(prev => ({ ...prev, sony: event.target.checked }))} className="rounded" />
                عكس المدين والدائن
              </label>
            )}
            <DropZone file={sonyFile} onFile={loadSony} onClear={() => { setSonyFile(null); setSonyHeaders([]); setSonyRowsRaw([]); }} />
            {sonyHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={sonyHeaders} value={sonyMap.date} onChange={v => setSonyMap(m => ({ ...m, date: v }))} />
                <Sel label="البيان (اسم الزبون)" headers={sonyHeaders} value={sonyMap.name} onChange={v => setSonyMap(m => ({ ...m, name: v }))} />
                <Sel label="المبالغ المدفوعة (Debit)" headers={sonyHeaders} value={sonyMap.debit} onChange={v => setSonyMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المستلمة (Credit)" headers={sonyHeaders} value={sonyMap.credit} onChange={v => setSonyMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب" headers={sonyHeaders} value={sonyMap.accountType} onChange={v => setSonyMap(m => ({ ...m, accountType: v }))} />
                <Sel label="رقم التفويض" headers={sonyHeaders} value={sonyMap.authNum} onChange={v => setSonyMap(m => ({ ...m, authNum: v }))} />
              </div>
            )}
          </div>

          {/* 5. رقم التفويض فيزا */}
          <div id="r2-auth" className="bg-card border p-4 rounded-xl space-y-4 lg:col-span-2 scroll-mt-5">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              🔢 رقم التفويض فيزا
              {parsedAuth.length > 0 && (
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-purple-100 text-purple-700 border border-purple-300">
                  {parsedAuth.length} سجل
                </span>
              )}
            </h2>
            {authHeaders.length > 0 && (
              <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
                <input type="checkbox" checked={visaSwaps.auth} onChange={event => setVisaSwaps(prev => ({ ...prev, auth: event.target.checked }))} className="rounded" />
                عكس المدين والدائن
              </label>
            )}
            <DropZone file={authFile} onFile={loadAuth} onClear={() => { setAuthFile(null); setAuthHeaders([]); setAuthRowsRaw([]); }} />
            {authHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="الاسم" headers={authHeaders} value={authMap.name} onChange={v => setAuthMap(m => ({ ...m, name: v }))} />
                <Sel label="رقم التفويض" headers={authHeaders} value={authMap.authNum} onChange={v => setAuthMap(m => ({ ...m, authNum: v }))} />
                <Sel label="المبلغ" headers={authHeaders} value={authMap.amount} onChange={v => setAuthMap(m => ({ ...m, amount: v }))} />
                <Sel label="التاريخ" headers={authHeaders} value={authMap.date} onChange={v => setAuthMap(m => ({ ...m, date: v }))} />
              </div>
            )}
          </div>
        </div>

        {/* SoftPOS is available on the dedicated F page. */}
        {false && <div id="r2-softpos" className="bg-card border p-4 rounded-xl space-y-3 scroll-mt-5">
          <div className="flex items-center gap-2 flex-wrap">
            <h2 className="font-semibold text-sm flex items-center gap-2">
              F — لصق إيصال SoftPOS وتحويله إلى Excel
            </h2>
            <span className="text-[10px] text-muted-foreground">الصق نص الإيصال وهنطلعلك ملف سوني كاشير فيزا جاهز</span>
          </div>
          <textarea
            value={softposText}
            onChange={e => setSoftposText(e.target.value)}
            placeholder="الصق هون نص إيصال SoftPOS كامل (اسم المتجر، الحركات، أرقام البطاقات، أرقام التفويض، المباليع...)"
            dir="rtl"
            className="w-full h-48 p-3 text-xs font-mono border border-border rounded-lg bg-input-background focus:outline-none focus:ring-1 focus:ring-ring resize-y"
          />
          <div className="flex items-center gap-2 flex-wrap">
            <button onClick={handleParseSoftPOS}
              disabled={!softposText.trim()}
              className="flex items-center gap-2 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm font-medium hover:bg-teal-700 disabled:opacity-40 transition-colors">
              <Download className="w-4 h-4"/>إنشاء ملف سوني كاشير فيزا
            </button>
            {softposText.trim() && (
              <button onClick={() => setSoftposText("")}
                className="flex items-center gap-1 px-3 py-2 bg-muted text-muted-foreground rounded-lg text-xs hover:bg-muted/80 transition-colors">
                <X className="w-3.5 h-3.5"/>مسح النص
              </button>
            )}
          </div>
        </div>}

        {/* Results */}
        {(parsedBank.length > 0 || parsedAza.length > 0 || parsedSony.length > 0) && (
          <div className="space-y-4">
            {/* Stats */}
            <div className="grid grid-cols-3 sm:grid-cols-7 gap-2">
              {([
                { id: "overview", label: "نظرة عامة", count: "", color: "text-blue-700" },
                { id: "bank", label: "فيزا بنك", count: stats.visaBank, color: "text-teal-700" },
                { id: "aza", label: "أستاذ", count: stats.azaTotal, color: "text-indigo-700" },
                { id: "sony", label: "🏪 سوني", count: stats.sonyTotal, color: "text-purple-700" },
                { id: "matched", label: "مؤكد", count: stats.matched, color: "text-green-700" },
                { id: "pendingBank", label: "معلقة بنك", count: stats.pendingBank, color: "text-amber-700" },
                { id: "unmatched", label: "غير متطابق", count: stats.unmatchedSony + stats.unmatchedAza, color: "text-red-700" },
              ] as const).map(t => (
                <button key={t.id} onClick={() => setTab(t.id as R2Tab)}
                  className={`p-2.5 rounded-xl border text-right transition-all ${tab === t.id ? "border-blue-500 bg-blue-50" : "bg-card hover:bg-muted/30"}`}>
                  <span className="text-[10px] text-muted-foreground block">{t.label}</span>
                  <span className={`text-base font-mono font-bold block mt-0.5 ${t.color}`}>{t.count}</span>
                </button>
              ))}
            </div>

            <div className="border rounded-xl bg-card overflow-hidden">

              {/* Overview tab */}
              {tab === "overview" && (
                <div className="p-6 space-y-4">
                  <h2 className="font-bold text-sm">نظرة عامة على الآلية الجديدة</h2>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="p-4 rounded-xl bg-teal-50 border border-teal-200">
                      <div className="text-xs text-teal-600 font-medium mb-1">كشف البنك</div>
                      <div className="text-2xl font-bold text-teal-800">{stats.bankTotal}</div>
                      <div className="text-xs text-teal-700 mt-1">فيزا بنك (مشتريات تجار): {stats.visaBank}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-indigo-50 border border-indigo-200">
                      <div className="text-xs text-indigo-600 font-medium mb-1">كشف الأستاذ</div>
                      <div className="text-2xl font-bold text-indigo-800">{stats.azaTotal}</div>
                      <div className="text-xs text-indigo-700 mt-1">فيزا: {stats.azaVisa} | عادي: {stats.azaRegular} | زيادة: {stats.azaOnly}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-purple-50 border border-purple-200">
                      <div className="text-xs text-purple-600 font-medium mb-1">🏪 سوني كاشير فيزا</div>
                      <div className="text-2xl font-bold text-purple-800">{stats.sonyTotal}</div>
                      <div className="text-xs text-purple-700 mt-1">مطابق: {stats.matched} | غير متطابق: {stats.unmatchedSony}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-amber-50 border border-amber-200">
                      <div className="text-xs text-amber-600 font-medium mb-1">مشتريات معلقة (لثاني يوم)</div>
                      <div className="text-2xl font-bold text-amber-800">{stats.pendingBank}</div>
                      <div className="text-xs text-amber-700 mt-1">مطابق مع سوني: {stats.pendingMatched}</div>
                    </div>
                  </div>

                  {autoMatchResults.length > 0 && (
                    <div className="p-4 rounded-xl bg-green-50 border border-green-200">
                      <div className="flex items-center justify-between flex-wrap gap-2">
                        <div>
                          <div className="text-sm font-bold text-green-800">🤖 مطابقات تلقائية مقترحة</div>
                          <div className="text-xs text-green-700 mt-0.5">{autoMatchResults.length} مطابقة (سوني ↔ أستاذ فيزا)</div>
                        </div>
                        <button onClick={handleAcceptAllAuto}
                          className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition-colors">
                          <Check className="w-4 h-4"/>تأكيد الكل
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Visa Bank tab */}
              {tab === "bank" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">التاريخ</th>
                  </tr></thead>
                  <tbody>
                    {visaBankRows.map(b => (
                      <tr key={b.id} className="border-t hover:bg-muted/20 transition-colors">
                        <td className="px-3 py-2.5 font-medium text-teal-700">{b.description}</td>
                        <td className="px-3 py-2.5 font-mono font-bold text-teal-700">{fmtNum(b.rawAmount)}</td>
                        <td className={`px-3 py-2.5 text-xs font-semibold ${b.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{b.type}</td>
                        <td className="px-3 py-2.5 text-xs text-muted-foreground">{b.date}</td>
                      </tr>
                    ))}
                    {!visaBankRows.length && <tr><td colSpan={4} className="py-10 text-center text-muted-foreground text-xs">لا توجد حوالات "مشتريات عمولة تجار". ارفع كشف البنك أولاً.</td></tr>}
                  </tbody>
                </table>
              )}

              {/* Aza tab */}
              {tab === "aza" && (
                <div>
                  <div className="flex items-center gap-2 px-4 py-2.5 bg-indigo-50/60 border-b border-indigo-100 flex-wrap">
                    <span className="text-xs font-medium text-indigo-800">أستاذ عادي: {azaRegularRows.length}</span>
                    <span className="text-xs text-indigo-600">| أستاذ فيزا: {azaVisaRows.length}</span>
                    {azaOnlyRows.length > 0 && (
                      <span className="text-xs px-2 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">{azaOnlyRows.length} زيادة بالأستاذ (مش باليومي)</span>
                    )}
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                    </tr></thead>
                    <tbody>
                      {parsedAza.map(a => (
                        <tr key={a.id} className={`border-t transition-colors ${a.visaNumber ? "bg-teal-50/40 hover:bg-teal-50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 font-medium">
                            {a.name}
                            {azaOnlyRows.some(o => o.id === a.id) && (
                              <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">زيادة</span>
                            )}
                          </td>
                          <td className="px-3 py-2.5">
                            {a.visaNumber ? <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{a.visaNumber}</span> : "—"}
                          </td>
                          <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(a.amount)}</td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${a.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{a.type}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.accountType || "—"}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.date}</td>
                        </tr>
                      ))}
                      {!parsedAza.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد بيانات. ارفع كشف الأستاذ أولاً.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Sony tab */}
              {tab === "sony" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">الاسم</th>
                    <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                    <th className="px-3 py-2.5 text-right">رقم التفويض</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">التاريخ</th>
                    <th className="px-3 py-2.5 text-right">الحالة</th>
                  </tr></thead>
                  <tbody>
                    {parsedSony.map(s => {
                      const isMatched = matchedSonyIds.has(s.id);
                      return (
                        <tr key={s.id} className={`border-t transition-colors ${isMatched ? "bg-green-50/50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 font-medium">{s.name}</td>
                          <td className="px-3 py-2.5">
                            {s.visaNumber ? <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{s.visaNumber}</span> : "—"}
                          </td>
                          <td className="px-3 py-2.5 text-xs font-mono">{s.authNumber || "—"}</td>
                          <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(s.amount)}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{s.date}</td>
                          <td className="px-3 py-2.5">
                            {isMatched ? (
                              <span className="text-green-600 bg-green-50 px-1.5 py-0.5 rounded text-xs border border-green-200 flex items-center gap-0.5 w-fit">
                                <Check className="w-2.5 h-2.5"/>مطابق
                              </span>
                            ) : (
                              <span className="text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded text-xs border border-amber-200">بانتظار</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                    {!parsedSony.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد بيانات. ارفع ملف سوني كاشير فيزا.</td></tr>}
                  </tbody>
                </table>
              )}

              {/* Matched tab — سوني ↔ أستاذ فيزا */}
              {tab === "matched" && (
                <div>
                  {autoMatchResults.length > 0 && (
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-green-50/60 border-b border-green-100 flex-wrap">
                      <span className="text-xs text-green-800">{autoMatchResults.length} مطابقة تلقائية مقترحة</span>
                      <button onClick={handleAcceptAllAuto}
                        className="mr-auto flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs font-medium hover:bg-green-700 transition-colors">
                        <Check className="w-3.5 h-3.5"/>تأكيد الكل
                      </button>
                    </div>
                  )}
                  {/* Auto-match suggestions */}
                  {autoMatchResults.map((r, i) => {
                    const key = `auto-${r.sony.id}-${r.aza.id}`;
                    const isExpanded = expandedKey === key;
                    return (
                      <React.Fragment key={i}>
                        <div className={`px-4 py-3 border-b transition-colors ${isExpanded ? "bg-blue-50" : "bg-blue-50/30 hover:bg-blue-50/60"}`}>
                          <div className="flex items-center gap-3 flex-wrap">
                            <div className="flex-1 min-w-0">
                              <div className="flex items-center gap-2 flex-wrap">
                                <span className="font-medium text-sm">{r.sony.name}</span>
                                {r.sony.visaNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 font-mono">{r.sony.visaNumber}</span>}
                                {r.visaMatch && <span className="text-[10px] px-1.5 py-0.5 rounded bg-green-100 text-green-700 border border-green-300">✓ فيزا مطابق</span>}
                                <span className="text-muted-foreground text-xs">↔</span>
                                <span className="text-sm text-muted-foreground">{r.aza.name}</span>
                                <span className="font-mono font-bold text-blue-700 text-sm">{fmtNum(r.aza.amount)}</span>
                                {r.amountDiff > 0.01 && <span className="text-[10px] text-orange-600">فرق: {fmtNum(r.amountDiff)}</span>}
                                <span className={`text-[10px] ${r.score * 100 >= 90 ? "text-green-600" : "text-amber-600"}`}>({(r.score * 100).toFixed(0)}%)</span>
                              </div>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              <button onClick={() => handleAcceptMatch(r.sony, r.aza, r.amountDiff > 0.01 ? `فرق مبلغ: ${fmtNum(r.amountDiff)}` : undefined)}
                                className="flex items-center gap-1 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                <Check className="w-3.5 h-3.5"/>تأكيد
                              </button>
                              <button onClick={() => handleRejectMatch(r.sony.id, r.aza.id)}
                                className="flex items-center gap-1 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>رفض
                              </button>
                              <button onClick={() => setExpandedKey(isExpanded ? null : key)}
                                className="p-1 text-muted-foreground hover:text-foreground">
                                <span className="text-base font-bold">⋮</span>
                              </button>
                            </div>
                          </div>
                          {isExpanded && (
                            <div className="mt-3 pt-3 border-t border-blue-200 text-xs space-y-1">
                              <div><strong>سوني:</strong> {r.sony.rawName} | مبلغ: {fmtNum(r.sony.amount)} | فيزا: {r.sony.visaNumber || "—"} | تفويض: {r.sony.authNumber || "—"}</div>
                              <div><strong>أستاذ:</strong> {r.aza.rawName} | مبلغ: {fmtNum(r.aza.amount)} | فيزا: {r.aza.visaNumber || "—"} | حساب: {r.aza.accountType || "—"}</div>
                              <div><strong>نوع المطابقة:</strong> {r.matchType} | <strong>التشابه:</strong> {(r.score * 100).toFixed(0)}%</div>
                            </div>
                          )}
                        </div>
                      </React.Fragment>
                    );
                  })}

                  {/* Confirmed matches */}
                  {matchedVisa.length > 0 && (
                    <div className="px-4 py-2 bg-green-50/40 border-b">
                      <span className="text-xs font-bold text-green-800">مطابقات مؤكدة ({matchedVisa.length})</span>
                    </div>
                  )}
                  {matchedVisa.map(m => (
                    <div key={m.id} className="px-4 py-3 border-b hover:bg-green-50/30 transition-colors">
                      <div className="flex items-center gap-3 flex-wrap">
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <Shield className="w-3.5 h-3.5 text-green-600"/>
                            <span className="font-medium text-sm">{m.sonyName}</span>
                            {m.visaNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 font-mono">{m.visaNumber}</span>}
                            {m.authNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 font-mono">تفويض: {m.authNumber}</span>}
                            <span className="text-muted-foreground text-xs">↔</span>
                            <span className="text-sm text-muted-foreground">{m.azaName}</span>
                            <span className="font-mono font-bold text-green-700 text-sm">{fmtNum(m.amount)}</span>
                            {m.note && <span className="text-[10px] text-orange-600">{m.note}</span>}
                          </div>
                        </div>
                        <button onClick={() => handleUnmatch(m)}
                          className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-600 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                          <X className="w-3.5 h-3.5"/>إلغاء
                        </button>
                      </div>
                    </div>
                  ))}

                  {!autoMatchResults.length && !matchedVisa.length && (
                    <div className="py-10 text-center text-muted-foreground text-xs">
                      لا توجد مطابقات بعد. ارفع ملف سوني كاشير فيزا وكشف الأستاذ لتظهر المطابقات التلقائية.
                    </div>
                  )}
                </div>
              )}

              {/* Pending Bank tab — مشتريات معلقة لثاني يوم */}
              {tab === "pendingBank" && (
                <div>
                  <div className="px-4 py-3 bg-amber-50/60 border-b border-amber-100">
                    <p className="text-xs text-amber-800">
                      هذه الحوالات "مشتريات عمولة تجار" من كشف البنك — تنتظر ثاني يوم للمطابقة مع سوني كاشير فيزا المؤكد.
                    </p>
                    {pendingMatchResults.length > 0 && (
                      <p className="text-xs text-green-700 mt-1">{pendingMatchResults.length} مطابقة محتملة مع سوني مؤكد!</p>
                    )}
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">بيان البنك</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">مطابقة محتملة</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {pendingVisaBank.map(p => {
                        const match = pendingMatchResults.find(r => r.pending.id === p.id);
                        return (
                          <tr key={p.id} className="border-t hover:bg-muted/20 transition-colors">
                            <td className="px-3 py-2.5 font-medium text-teal-700">{p.bankDesc}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-amber-700">{fmtNum(p.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{p.date}</td>
                            <td className="px-3 py-2.5 text-xs">
                              {match ? (
                                <span className="text-green-700 bg-green-50 px-1.5 py-0.5 rounded border border-green-200">
                                  {match.matched.sonyName} — {match.reason}
                                </span>
                              ) : (
                                <span className="text-muted-foreground">لا توجد</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => setPendingVisaBank(prev => prev.filter(x => x.id !== p.id))}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <Trash2 className="w-3.5 h-3.5"/>حذف
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                      {!pendingVisaBank.length && <tr><td colSpan={5} className="py-10 text-center text-muted-foreground text-xs">لا توجد حوالات معلقة. فرز كشف البنك وانقل "مشتريات تجار" للمعلقة.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Unmatched tab */}
              {tab === "unmatched" && (
                <div className="p-4 space-y-4">
                  <div>
                    <h3 className="font-semibold text-sm mb-2">سوني كاشير فيزا غير متطابق ({unmatchedSony.length})</h3>
                    <table className="w-full text-sm border rounded-lg overflow-hidden">
                      <thead><tr className="bg-muted text-xs">
                        <th className="px-3 py-2.5 text-right">الاسم</th>
                        <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                        <th className="px-3 py-2.5 text-right">رقم التفويض</th>
                        <th className="px-3 py-2.5 text-right">المبلغ</th>
                      </tr></thead>
                      <tbody>
                        {unmatchedSony.map(s => (
                          <tr key={s.id} className="border-t hover:bg-muted/20">
                            <td className="px-3 py-2.5 font-medium">{s.name}</td>
                            <td className="px-3 py-2.5">{s.visaNumber ? <span className="font-mono text-teal-700">{s.visaNumber}</span> : "—"}</td>
                            <td className="px-3 py-2.5 text-xs font-mono">{s.authNumber || "—"}</td>
                            <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(s.amount)}</td>
                          </tr>
                        ))}
                        {!unmatchedSony.length && <tr><td colSpan={4} className="py-6 text-center text-muted-foreground text-xs">كل سوني مطابق ✓</td></tr>}
                      </tbody>
                    </table>
                  </div>
                  <div>
                    <h3 className="font-semibold text-sm mb-2">أستاذ فيزا غير متطابق ({unmatchedAzaVisa.length})</h3>
                    <table className="w-full text-sm border rounded-lg overflow-hidden">
                      <thead><tr className="bg-muted text-xs">
                        <th className="px-3 py-2.5 text-right">الاسم</th>
                        <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                        <th className="px-3 py-2.5 text-right">المبلغ</th>
                        <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      </tr></thead>
                      <tbody>
                        {unmatchedAzaVisa.map(a => (
                          <tr key={a.id} className="border-t hover:bg-muted/20">
                            <td className="px-3 py-2.5 font-medium">{a.name}</td>
                            <td className="px-3 py-2.5">{a.visaNumber ? <span className="font-mono text-teal-700">{a.visaNumber}</span> : "—"}</td>
                            <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(a.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.accountType || "—"}</td>
                          </tr>
                        ))}
                        {!unmatchedAzaVisa.length && <tr><td colSpan={4} className="py-6 text-center text-muted-foreground text-xs">كل أستاذ فيزا مطابق ✓</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

            </div>
          </div>
        )}
      </div>
    </div>
  );
}
