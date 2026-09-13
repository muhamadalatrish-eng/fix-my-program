"use client"
/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars */
import React, { useState, useCallback, useMemo, useEffect, useRef, useDeferredValue } from "react";
import * as XLSX from "xlsx";
import Reconciliation2, { type StageAResult } from "./Reconciliation2";
import { Upload, FileSpreadsheet, Download, ChevronDown, X, Trash2, TriangleAlert as AlertTriangle, Check, Search, Link2, Link2Off, ChevronLeft, ChevronRight, CreditCard, Sparkles, Info, Save, Shield, FolderOpen, FolderPlus, RotateCcw, FolderMinus, GripVertical, Landmark, Users, BarChart3, ArrowUpRight } from "lucide-react";

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

function resolveDebitCredit(
  row: Record<string, unknown>,
  debitCol: string,
  creditCol: string,
  swap: boolean
): { debit: number; credit: number; rawAmount: number; type: "مدفوع" | "مستلم" } {
  const debitRaw = Math.abs(toNum(debitCol ? row[debitCol] : 0));
  const creditRaw = Math.abs(toNum(creditCol ? row[creditCol] : 0));
  const debit = swap ? creditRaw : debitRaw;
  const credit = swap ? debitRaw : creditRaw;

  let rawAmount = 0;
  let type: "مدفوع" | "مستلم" = "مستلم";

  if (debit > 0 && credit === 0) { rawAmount = debit; type = "مستلم"; }
  else if (credit > 0 && debit === 0) { rawAmount = credit; type = "مدفوع"; }
  else if (debit > 0 && credit > 0) { rawAmount = debit; type = "مستلم"; }

  return { debit, credit, rawAmount, type };
}

// ─── تخزين دائم ────────────────────────────────────────────────────────────────
interface AppStorage {
  get(key: string): Promise<{ value: string } | null>;
  set(key: string, value: string): Promise<void>;
  delete(key: string): Promise<void>;
  list(prefix: string): Promise<{ keys: string[] }>;
}

interface StorageWindow extends Window {
  storage?: AppStorage;
}

async function storageGet<T>(key: string, fallback: T): Promise<T> {
  try {
    const w = window as StorageWindow;
    if (w?.storage?.get) {
      const res = await w.storage.get(key);
      return res ? (JSON.parse(res.value) as T) : fallback;
    }
  } catch { /* المفتاح غير موجود أو خطأ - نكمل على localStorage */ }
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}
async function storageSet(key: string, value: unknown): Promise<void> {
  try {
    const w = window as StorageWindow;
    if (w?.storage?.set) {
      await w.storage.set(key, JSON.stringify(value));
      return;
    }
  } catch (e) { console.error("storage.set فشل، سيتم استخدام localStorage:", e); }
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { console.error("فشل الحفظ نهائياً:", e); }
}
async function storageDelete(key: string): Promise<void> {
  try {
    const w = window as StorageWindow;
    if (w?.storage?.delete) { await w.storage.delete(key); return; }
  } catch { /* تجاهل */ }
  try { localStorage.removeItem(key); } catch { /* تجاهل */ }
}

function normalizeIdSet(value: unknown): Set<number> {
  if (value instanceof Set) {
    return new Set(Array.from(value).map(toNum).filter(Number.isFinite));
  }
  if (Array.isArray(value)) {
    return new Set(value.map(toNum).filter(Number.isFinite));
  }
  if (value && typeof value === "object") {
    // Older sessions could contain a JSON-serialized Set, which becomes an
    // object instead of an iterable array. Object keys preserve any numeric ids.
    return new Set(Object.keys(value).map(toNum).filter(Number.isFinite));
  }
  return new Set();
}

// ─── Visa detection ────────────────────────────────────────────────────────────
const VISA_NUM_RE = new RegExp("[/-]\\s*(\\d{4})(?!\\d)");
function extractVisaNumber(rawName: string): string | null {
  const m = rawName.match(VISA_NUM_RE);
  return m ? m[1] : null;
}
function extractVisaName(rawName: string): string {
  return rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
}

// ─── Flags ────────────────────────────────────────────────────────────────────
const ARITHMETIC_RE = /(\d+(?:\.\d+)?)(\+(\d+(?:\.\d+)?))+/;
function evalSum(expr: string): number {
  return expr.split("+").reduce((acc, p) => acc + parseFloat(p.trim()), 0);
}
function parseCashierName(raw: string): { name: string; notes: string } {
  let notes = "";
  let name = raw.trim();
  const arithmeticMatch = name.match(ARITHMETIC_RE);
  if (arithmeticMatch) { notes += " " + arithmeticMatch[0]; name = name.replace(ARITHMETIC_RE, ""); }
  return { name: name.replace(/\s+/g, " ").trim(), notes: notes.trim() };
}
function parseNotes(notes: string) {
  const match = notes.match(ARITHMETIC_RE);
  const splitExpr = match ? match[0] : null;
  const matchAmount = splitExpr ? evalSum(splitExpr) : null;
  return { splitExpr, matchAmount };
}

// ─── Name matching ────────────────────────────────────────────────────────────
const AR_EN: Record<string, string> = {
  "ا":"a","أ":"a","إ":"a","آ":"a","ى":"a","ة":"h","ب":"b","ت":"t","ث":"th",
  "ج":"j","ح":"h","خ":"kh","د":"d","ذ":"th","ر":"r","ز":"z","س":"s","ش":"sh",
  "ص":"s","ض":"d","ط":"t","ظ":"z","ع":"a","غ":"g","ف":"f","ق":"q","ك":"k",
  "ل":"l","م":"m","ن":"n","ه":"h","و":"w","ي":"i","ئ":"i","ء":"","ّ":"",
  "َ":"","ُ":"","ِ":"","ْ":"","ً":"","ٌ":"","ٍ":"",
};
function normName(s: string) {
  const c = s.toLowerCase()
    .replace(/mahmoud/g, "mhmd")
    .replace(/raefet/g, "raft")
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

// ─── Types ────────────────────────────────────────────────────────────────────
interface BankRow {
  id: number; date: string; description: string;
  debit: number; credit: number; rawAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  ref: string;
  orig: Record<string,unknown>;
  _fromHeld?: boolean;
}
interface CashierRow {
  id: number; rawName: string; name: string; notes: string;
  splitExpr: string|null;
  debit: number; credit: number; amount: number; matchAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  ref: string;
  date: string; orig: Record<string,unknown>;
  _fromHeld?: boolean;
}
interface ManualMatchGroup {
  id: string;
  banks: BankRow[];
  cashiers: CashierRow[];
  note?: string;
}

// ─── Saved matches ──────────────────────────────────────────────────────────
interface SavedMatch {
  id: string;
  cashierId: number;
  bankId: number;
  cashierName: string;
  bankDesc: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  date: string;
  savedAt: string;
  note?: string;
  isAmountDiff?: boolean;
  isNameDiff?: boolean;
  isManual?: boolean;
  isAccountTypeDiff?: boolean;
  bankAccountType?: string;
  cashierAccountType?: string;
  sourceGroupId?: string;
  matchScore?: number;
  editorNotes?: string;
}

// ─── Visa items ──────────────────────────────────────────────────────────────
interface VisaItem {
  id: string;
  cashierId: number;
  rawName: string;
  name: string;
  visaNumber: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  date: string;
  movedAt: string;
  note?: string;
  source?: "pending" | "unmatched" | "saved" | "manual";
}

interface HeldItem {
  id: string;
  kind: "bank" | "cashier";
  refId: number;
  fileSessionId: number;
  data: BankRow | CashierRow;
  heldAt: string;
  note?: string;
}

let heldReturnSequence = 0;
function nextHeldRowId(): number {
  heldReturnSequence = (heldReturnSequence + 1) % 1000;
  return -(Date.now() * 1000 + heldReturnSequence);
}

function normAccountType(s: string): string {
  return (s || "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}
function accountTypesDiffer(bankType: string, cashierType: string): boolean {
  const a = normAccountType(bankType), b = normAccountType(cashierType);
  if (!a || !b) return false;
  return a !== b;
}

// ─── Advanced match check ───────────────────────────────────────────────────
function advancedMatchCheckSingle(cashierName: string, bankDesc: string): {
  isMatch: boolean;
  isApprox: boolean;
  matchType: "exact" | "firstSecond" | "fourthName" | "typo" | "none";
  matchedTokens?: string[];
} {
  const cT = nameTokens(cashierName);
  const bT = nameTokens(bankDesc);

  if (!cT.length || !bT.length) return { isMatch: false, isApprox: false, matchType: "none" };

  const squished = (arr: string[]) => arr.slice().sort().join("");
  if (squished(cT) === squished(bT)) {
    return { isMatch: true, isApprox: false, matchType: "exact" };
  }

  if (cT.length >= 2 && bT.length >= 2) {
    const firstMatch = wordTypoMatch(cT[0], bT[0]);
    const lastMatch = wordTypoMatch(cT[cT.length - 1], bT[bT.length - 1]);
    if (firstMatch && lastMatch) {
      return {
        isMatch: true,
        isApprox: cT[0] !== bT[0] || cT[cT.length - 1] !== bT[bT.length - 1],
        matchType: "firstSecond",
        matchedTokens: [cT[0], cT[cT.length - 1]]
      };
    }
  }

  if (cT.length >= 2 && bT.length >= 2) {
    const firstMatch = wordTypoMatch(cT[0], bT[0]);
    const secondMatch = wordTypoMatch(cT[1], bT[1]);
    if (firstMatch && secondMatch) {
      const exactFirst = cT[0] === bT[0];
      const exactSecond = cT[1] === bT[1];
      return {
        isMatch: true,
        isApprox: !(exactFirst && exactSecond),
        matchType: "firstSecond",
        matchedTokens: [cT[0], cT[1]]
      };
    }
  }

  if (cT.length >= 4 && bT.length >= 4) {
    if (wordTypoMatch(cT[3], bT[3])) {
      return {
        isMatch: true,
        isApprox: cT[3] !== bT[3],
        matchType: "fourthName",
        matchedTokens: [cT[3]]
      };
    }
  }

  if (cT.length === 1 || bT.length === 1) {
    const shortSide = cT.length === 1 ? cT : bT;
    const longSide = cT.length === 1 ? bT : cT;
    if (wordTypoMatch(shortSide[0], longSide[0])) {
      return {
        isMatch: true,
        isApprox: shortSide[0] !== longSide[0],
        matchType: "typo",
        matchedTokens: [shortSide[0]]
      };
    }
  }

  for (const ct of cT) {
    for (const bt of bT) {
      if (wordTypoMatch(ct, bt)) {
        return {
          isMatch: true,
          isApprox: true,
          matchType: "typo",
          matchedTokens: [ct, bt]
        };
      }
    }
  }

  return { isMatch: false, isApprox: false, matchType: "none" };
}

function advancedMatchCheck(cashierName: string, bankDesc: string): {
  isMatch: boolean;
  isApprox: boolean;
  matchType: "exact" | "firstSecond" | "fourthName" | "typo" | "none";
  matchedTokens?: string[];
} {
  const bankSegments = bankDesc.split("/").map(segment => segment.trim()).filter(Boolean);
  if (bankSegments.length <= 1) return advancedMatchCheckSingle(cashierName, bankDesc);
  const results = bankSegments.map(segment => advancedMatchCheckSingle(cashierName, segment));
  return results
    .filter(result => result.isMatch)
    .sort((a, b) => Number(a.isApprox) - Number(b.isApprox))[0] ||
    { isMatch: false, isApprox: false, matchType: "none" };
}

function smartNameKey(value: string): string {
  const tokens = nameTokens(value);
  return tokens.length >= 2 ? `${tokens[0]} ${tokens[tokens.length - 1]}` : tokens.join(" ");
}

function smartAliasTarget(cashierName: string, bankDescription: string): string {
  return bankDescription.split("/")
    .map(segment => segment.trim())
    .filter(Boolean)
    .sort((a, b) => nameSim(cashierName, b) - nameSim(cashierName, a))[0] || bankDescription;
}

// ─── Reconcile ───────────────────────────────────────────��──────────────────
function reconcile(
  bank: BankRow[],
  cashier: CashierRow[],
  manualGroups: ManualMatchGroup[],
  savedMatches: SavedMatch[],
  rejectedPairs: Set<string>,
  visaCashierIds: Set<number>,
  jawwalPayCashierIds: Set<number>,
  mahmoudWalletCashierIds: Set<number>,
  amountTolerancePercent: number = 0.5,
  aliases: Record<string, string> = {},
): MatchResult[] {
  const results: MatchResult[] = [];
  const usedBank = new Set<number>();
  const usedCashier = new Set<number>();

  const savedKeys = new Set(savedMatches.map(s => `${s.cashierId}-${s.bankId}`));
  const savedCashierIds = new Set(savedMatches.map(s => s.cashierId));
  const savedBankIds = new Set(savedMatches.map(s => s.bankId));

  savedMatches.forEach(sm => {
    const bankRow = bank.find(b => b.id === sm.bankId);
    const cashierRow = cashier.find(c => c.id === sm.cashierId);
    if (bankRow && cashierRow) {
      results.push({ type: "saved", bank: bankRow, cashier: cashierRow, savedMatch: sm });
      usedBank.add(bankRow.id);
      usedCashier.add(cashierRow.id);
    }
  });
  cashier.forEach(c => {
    if (usedCashier.has(c.id) || savedCashierIds.has(c.id)) return;
    if (jawwalPayCashierIds.has(c.id)) {
      results.push({ type: "jawwalPay", cashier: c });
      usedCashier.add(c.id);
      return;
    }
    if (mahmoudWalletCashierIds.has(c.id)) {
      results.push({ type: "mahmoudWallet", cashier: c });
      usedCashier.add(c.id);
      return;
    }
    if (visaCashierIds.has(c.id)) {
      results.push({ type:"visa", cashier:c });
      usedCashier.add(c.id);
    }
  });

  const candidates: Array<{ cashier: CashierRow; bank: BankRow; score: number; isApprox: boolean; matchType: string }> = [];
  cashier.forEach(c => {
    if (usedCashier.has(c.id) || savedCashierIds.has(c.id) || visaCashierIds.has(c.id) || jawwalPayCashierIds.has(c.id) || mahmoudWalletCashierIds.has(c.id)) return;
    bank.forEach(b => {
      if (usedBank.has(b.id) || savedBankIds.has(b.id) || c.type !== b.type) return;
      const pairKey = `${c.id}-${b.id}`;
      if (rejectedPairs.has(pairKey) || savedKeys.has(pairKey)) return;
      const amountTolerance = Math.max(0.01, c.matchAmount * amountTolerancePercent / 100);
      if (Math.abs(b.rawAmount - c.matchAmount) > amountTolerance) return;
      const aliasMatch = aliases[smartNameKey(c.name)] && b.description.split("/").some(segment => smartNameKey(segment) === aliases[smartNameKey(c.name)]);
      const ms = aliasMatch ? { isMatch: true, isApprox: false, matchType: "exact" as const } : advancedMatchCheck(c.name, b.description);
      const similarity = nameSim(c.name, b.description);
      if (!ms.isMatch && similarity < 0.25) return;
      const priority = ms.isMatch ? (ms.matchType === "exact" ? 4 : ms.matchType === "firstSecond" ? 3 : ms.matchType === "fourthName" ? 2 : 1) : 0;
      candidates.push({ cashier: c, bank: b, score: priority + similarity, isApprox: !ms.isMatch || ms.isApprox, matchType: ms.isMatch ? ms.matchType : "closest-name" });
    });
  });
  candidates.sort((a, b) => b.score - a.score);
  candidates.forEach(candidate => {
    if (usedCashier.has(candidate.cashier.id) || usedBank.has(candidate.bank.id)) return;
    results.push({
      type: "pending",
      bank: candidate.bank,
      cashier: candidate.cashier,
      isApprox: candidate.isApprox,
      isExactAmount: true,
      matchType: candidate.matchType,
      amountDiff: Math.abs(candidate.bank.rawAmount - candidate.cashier.matchAmount),
      matchScore: Math.round(Math.min(1, candidate.score / 4.5) * 100),
      accountTypeDiff: accountTypesDiffer(candidate.bank.accountType, candidate.cashier.accountType)
    });
    usedBank.add(candidate.bank.id);
    usedCashier.add(candidate.cashier.id);
  });

  const remC = cashier.filter(c => !usedCashier.has(c.id) && !savedCashierIds.has(c.id) && !visaCashierIds.has(c.id) && !jawwalPayCashierIds.has(c.id) && !mahmoudWalletCashierIds.has(c.id));
  const remB = bank.filter(b => !usedBank.has(b.id) && !savedBankIds.has(b.id));

  const getBase = (name:string) => normName(name).split(" ").filter(t=>t.length>2&&t!=="al").slice(0,2).join(" ");
  const groups: Record<string,{cashiers:CashierRow[];banks:BankRow[]}> = {};
  remC.forEach(c => {
    const k = getBase(c.name);
    if (!k) return;
    if (!groups[k]) groups[k] = {cashiers:[], banks:[]};
    groups[k].cashiers.push(c);
  });
  remB.forEach(b => {
    for (const k of Object.keys(groups)) {
      const nb = normName(b.description), kT = k.split(" ");
      if (kT.every(kt => nb.includes(kt))) {
        groups[k].banks.push(b);
        break;
      }
    }
  });

  Object.values(groups).forEach(g => {
    if (!g.cashiers.length || !g.banks.length) return;

    const types = new Set(g.cashiers.map(c => c.type));
    if (types.size > 1) return;

    const totC = g.cashiers.reduce((s,c) => s + c.matchAmount, 0);
    const totB = g.banks.reduce((s,b) => s + b.rawAmount, 0);
    if (Math.abs(totC - totB) > 0.05) return;

    const pairCount = Math.min(g.cashiers.length, g.banks.length);
    for (let i = 0; i < pairCount; i++) {
      const c = g.cashiers[i];
      const b = g.banks[i];
      const pairKey = `${c.id}-${b.id}`;
      if (!rejectedPairs.has(pairKey) && !savedKeys.has(pairKey)) {
        results.push({
          type: "pending",
          bank: b,
          cashier: c,
          isApprox: false,
          isBundled: true,
          isExactAmount: true,
          matchType: "bundle",
          amountDiff: 0
        });
        usedCashier.add(c.id);
        usedBank.add(b.id);
      }
    }
  });

  cashier.forEach(c => {
    if (usedCashier.has(c.id) || savedCashierIds.has(c.id) || visaCashierIds.has(c.id) || jawwalPayCashierIds.has(c.id) || mahmoudWalletCashierIds.has(c.id)) return;
    const hasAmt = bank.some(b => !usedBank.has(b.id) && !savedBankIds.has(b.id) &&
      b.type === c.type && Math.abs(b.rawAmount - c.matchAmount) <= 0.01);
    results.push({
      type:"unmatchedCashier",
      cashier:c,
      reason: hasAmt ? "اختلاف في الاسم" : "غير موجودة في البنك"
    });
  });

  bank.forEach(b => {
    if (usedBank.has(b.id) || savedBankIds.has(b.id)) return;
    const hasAmt = cashier.some(c => !usedCashier.has(c.id) && !savedCashierIds.has(c.id) && !visaCashierIds.has(c.id) && !jawwalPayCashierIds.has(c.id) && !mahmoudWalletCashierIds.has(c.id) &&
      c.type === b.type && Math.abs(c.matchAmount - b.rawAmount) <= 0.01);
    results.push({
      type:"unmatchedBank",
      bank:b,
      reason: hasAmt ? "اختلاف في الاسم" : "الحوالة غير موجودة في الكاشير"
    });
  });

  return results;
}

// ─── MatchResult type ────────────────────────────────────────────────────────
type MatchResult =
  | { type:"saved"; bank:BankRow; cashier:CashierRow; savedMatch:SavedMatch }
  | { type:"pending"; bank:BankRow; cashier:CashierRow; isApprox:boolean; isBundled?:boolean; isExactAmount:boolean; matchType?:string; amountDiff?:number; accountTypeDiff?:boolean; matchScore?:number }
  | { type:"manualGroup"; group:ManualMatchGroup }
  | { type:"visa"; cashier:CashierRow }
  | { type:"jawwalPay"; cashier:CashierRow }
  | { type:"mahmoudWallet"; cashier:CashierRow }
  | { type:"unmatchedCashier"; cashier:CashierRow; reason:"اختلاف في الاسم"|"غير موجودة في البنك" }
  | { type:"unmatchedBank"; bank:BankRow; reason:"اختلاف في الاسم"|"الحوالة غير موجودة في الكاشير" };

type TabId = "saved"|"pending"|"manual"|"visa"|"jawwalPay"|"mahmoudWallet"|"held"|"uCashier"|"uBank";

function scoreCandidate(
  cashierRow: CashierRow,
  b: BankRow,
  claimedNames: Set<string>,
  aliases: Record<string, string> = {}
): { score:number; matchType:string; reason:string } | null {
  const aliasMatch = aliases[smartNameKey(cashierRow.name)] && b.description.split("/").some(segment => smartNameKey(segment) === aliases[smartNameKey(cashierRow.name)]);
  const ms = aliasMatch ? { isMatch: true, isApprox: false, matchType: "exact" } : advancedMatchCheck(cashierRow.name, b.description);
  const diff = Math.abs(b.rawAmount - cashierRow.matchAmount);
  const amtMatch = diff <= 0.01;

  if (ms.isMatch && amtMatch) {
    return { score: 0.95, matchType: ms.matchType, reason: "الاسم والمبلغ متطابقان" };
  }
  if (ms.isMatch && !amtMatch) {
    const score = 0.5 + nameSim(cashierRow.name, b.description) * 0.4;
    return { score, matchType: ms.matchType, reason: "الاسم متطابق، المبلغ مختلف" };
  }
  if (!ms.isMatch && amtMatch) {
    const bankNameKey = normName(b.description);
    if (bankNameKey && claimedNames.has(bankNameKey)) return null;
    return { score: 0.2, matchType: "amount_only", reason: "المبلغ متطابق فقط (الاسم مختلف)" };
  }
  return null;
}

// ─── Suggestions ──────────────────────────────────────────────────────────────
function getSuggestions(
  cashierRow: CashierRow,
  availableBank: BankRow[],
  rejectedPairs: Set<string>,
  savedKeys: Set<string>,
  claimedNames: Set<string>,
  ownerMap: Map<number, {cashierId:number; score:number}>,
  aliases: Record<string, string> = {}
): Array<{bank:BankRow; score:number; matchType:string; amountDiff:number; reason:string}> {
  return availableBank
    .filter(b => {
      const pairKey = `${cashierRow.id}-${b.id}`;
      if (cashierRow.type !== b.type) return false;
      return !rejectedPairs.has(pairKey) && !savedKeys.has(pairKey);
    })
    .map(b => {
      const sc = scoreCandidate(cashierRow, b, claimedNames, aliases);
      if (!sc) return null;
      const owner = ownerMap.get(b.id);
      if (owner && owner.cashierId !== cashierRow.id) return null;
      const diff = Math.abs(b.rawAmount - cashierRow.matchAmount);
      return { bank: b, score: sc.score, matchType: sc.matchType, amountDiff: diff, reason: sc.reason };
    })
    .filter((x): x is {bank:BankRow; score:number; matchType:string; amountDiff:number; reason:string} => x !== null)
    .sort((a, b) => b.score - a.score)
    .slice(0, 5);
}

// ─── Export ───────────────────────────────────────────────────────────────────
function sanitizeSheetName(name: string, used: Set<string>): string {
  const clean = name.replace(/[\\/?*[\]:]/g, "-").slice(0, 31).trim() || "شيت";
  let final = clean;
  let i = 2;
  while (used.has(final)) {
    const suffix = ` ${i}`;
    final = clean.slice(0, 31 - suffix.length) + suffix;
    i++;
  }
  used.add(final);
  return final;
}

function matchChannelLabel(bank: BankRow, cashier: CashierRow): string {
  const bankChannel = bank.accountType || "غير محدد";
  const cashierChannel = cashier.accountType || "غير محدد";
  return bankChannel === cashierChannel ? bankChannel : `${cashierChannel} ← ${bankChannel}`;
}

function matchedChannelRows(results: MatchResult[]) {
  return results
    .filter((row): row is Extract<MatchResult, { type: "saved" }> => row.type === "saved")
    .map(row => ({
      ...row,
      channel: matchChannelLabel(row.bank, row.cashier),
      bankNet: row.bank.debit - row.bank.credit,
      cashierNet: row.cashier.debit - row.cashier.credit,
    }));
}

function financialIssueRows(results: MatchResult[]) {
  return matchedChannelRows(results).filter(row =>
    Math.abs(row.bank.rawAmount - row.cashier.matchAmount) > 0.01
  );
}

function doExport(
  results: MatchResult[],
  savedMatches: SavedMatch[],
  visaItems: VisaItem[],
  heldItems: HeldItem[],
  bankRows: BankRow[] = [],
  cashierRows: CashierRow[] = [],
  resumeData?: ResumeData,
) {
  const wb = XLSX.utils.book_new();
  const usedSheetNames = new Set<string>();
  const add=(name:string,hdrs:string[],rows:unknown[][])=>{
    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    ws["!cols"]=hdrs.map(()=>({wch:25}));
    XLSX.utils.book_append_sheet(wb,ws,sanitizeSheetName(name, usedSheetNames));
  };

  const bankTotal = bankRows.reduce((sum, row) => sum + row.rawAmount, 0);
  const cashierTotal = cashierRows.reduce((sum, row) => sum + row.matchAmount, 0);
  const matchedResults = results.filter((row): row is Extract<MatchResult, { type: "saved" }> => row.type === "saved");
  const matchedBankTotal = matchedResults.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  const matchedCashierTotal = matchedResults.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const categorySummary = [
    ["إجمالي الكشوف", bankTotal, cashierTotal, bankTotal - cashierTotal, bankRows.length, cashierRows.length],
    ["المطابقات المؤكدة", matchedBankTotal, matchedCashierTotal, matchedBankTotal - matchedCashierTotal, matchedResults.length, matchedResults.length],
    ["فيزا", 0, results.filter(row => row.type === "visa").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "visa").length],
    ["محفظة محمود", 0, results.filter(row => row.type === "mahmoudWallet").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "mahmoudWallet").length],
    ["جوال بي", 0, results.filter(row => row.type === "jawwalPay").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "jawwalPay").length],
    ["فواتير بلا حوالات", 0, results.filter(row => row.type === "unmatchedCashier").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "unmatchedCashier").length],
    ["حوالات بلا فواتير", results.filter(row => row.type === "unmatchedBank").reduce((sum, row) => sum + row.bank.rawAmount, 0), 0, 0, results.filter(row => row.type === "unmatchedBank").length, 0],
    ["معلقات بنك", heldItems.filter(item => item.kind === "bank").reduce((sum, item) => sum + (item.data as BankRow).rawAmount, 0), 0, 0, heldItems.filter(item => item.kind === "bank").length, 0],
    ["معلقات كاشير", 0, heldItems.filter(item => item.kind === "cashier").reduce((sum, item) => sum + (item.data as CashierRow).matchAmount, 0), 0, 0, heldItems.filter(item => item.kind === "cashier").length],
  ];
  add("تقرير المقارنة",
    ["البند","من البنك","من الكاشير","الفرق (بنك - كاشير)","عدد البنك","عدد الكاشير"],
    categorySummary.map(row => row.map((value, index) => typeof value === "number" && index > 0 && index < 4 ? fmtNum(value) : value)));
  const unmatchedFinancialRows = [
    ["فواتير بلا حوالات · مدين", results.filter(r => r.type === "unmatchedCashier" && r.cashier.debit > 0).length, results.filter(r => r.type === "unmatchedCashier").reduce((sum, r) => sum + r.cashier.debit, 0)],
    ["فواتير بلا حوالات · دائن", results.filter(r => r.type === "unmatchedCashier" && r.cashier.credit > 0).length, results.filter(r => r.type === "unmatchedCashier").reduce((sum, r) => sum + r.cashier.credit, 0)],
    ["حوالات بلا فواتير · مدين", results.filter(r => r.type === "unmatchedBank" && r.bank.debit > 0).length, results.filter(r => r.type === "unmatchedBank").reduce((sum, r) => sum + r.bank.debit, 0)],
    ["حوالات بلا فواتير · دائن", results.filter(r => r.type === "unmatchedBank" && r.bank.credit > 0).length, results.filter(r => r.type === "unmatchedBank").reduce((sum, r) => sum + r.bank.credit, 0)],
  ];
  add("غير المطابقات مدين ودائن", ["الفئة","العدد","الإجمالي بدون طرح الدائن"], unmatchedFinancialRows);

  const savedRows = results.filter(r=>r.type==="saved").map((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    return [
      r.cashier.name, r.bank.description, sm.type, r.bank.rawAmount, r.cashier.amount,
      sm.isAmountDiff ? "نعم" : "لا", sm.isNameDiff ? "نعم" : "لا",
      r.bank.accountType || "—", r.cashier.accountType || "—",
      sm.isAccountTypeDiff ? "نعم" : "لا", sm.isManual ? "يدوي" : "تلقائي",
      sm.note || "", sm.matchScore ?? "", sm.editorNotes || "",
      r.bank.date, r.cashier.date, sm.savedAt
    ];
  });
  add("✅ المطابقات (المؤكدة)",
    ["البيان (الكاشير)","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","اختلاف مبلغ","اختلاف اسم","نوع حساب البنك","نوع حساب الكاشير","اختلاف نوع حساب","نوع المطابقة","ملاحظة","درجة التطابق","ملاحظات المحرر","ت.البنك","ت.الكاشير","تاريخ الحفظ"],
    savedRows);

  const matchedDetails = matchedChannelRows(results);
  add("تفصيل المطابقات والقنوات",
    ["القناة في المطابقة","فاتورة الكاشير","حوالة البنك","كاشير مدين","كاشير دائن","كاشير صافي","بنك مدين","بنك دائن","بنك صافي","فرق المطابقة","نوع المطابقة","تاريخ الحفظ"],
    matchedDetails.map(row => [
      row.channel, row.cashier.name, row.bank.description,
      row.cashier.debit, row.cashier.credit, row.cashierNet,
      row.bank.debit, row.bank.credit, row.bankNet,
      row.bankNet - row.cashierNet, row.savedMatch.isManual ? "يدوي" : "تلقائي", row.savedMatch.savedAt,
    ]));
  const matchedChannelTotals = Array.from(new Set(matchedDetails.map(row => row.channel))).map(channel => {
    const rows = matchedDetails.filter(row => row.channel === channel);
    const sum = (pick: (row: typeof rows[number]) => number) => rows.reduce((total, row) => total + pick(row), 0);
    return [channel, rows.length, sum(row => row.bank.debit), sum(row => row.bank.credit), sum(row => row.bankNet), sum(row => row.cashier.debit), sum(row => row.cashier.credit), sum(row => row.cashierNet), sum(row => row.bankNet - row.cashierNet)];
  });
  add("إجماليات المطابقة حسب القناة",
    ["القناة","عدد المطابقات","البنك مدين","البنك دائن","البنك صافي","الفواتير مدين","الفواتير دائن","الفواتير صافي","فرق القناة"],
    matchedChannelTotals);
  const financialIssues = financialIssueRows(results);
  add("الفواتير ذات المشاكل المالية",
    ["القناة","الفاتورة","الحوالة","مبلغ الفاتورة","الواصل الفعلي بالبنك","فرق الفاتورة عن البنك","الحالة","الفاتورة مدين","الفاتورة دائن","البنك مدين","البنك دائن","التاريخ"],
    financialIssues.map(row => [
      row.channel, row.cashier.name, row.bank.description, row.cashier.matchAmount, row.bank.rawAmount,
      row.cashier.matchAmount - row.bank.rawAmount, "خطأ — البنك هو الصحيح", row.cashier.debit, row.cashier.credit,
      row.bank.debit, row.bank.credit, row.cashier.date,
    ]));
  const financialExpected = financialIssues.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const financialActual = financialIssues.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  add("إجمالي المشاكل المالية",
    ["عدد الفواتير","إجمالي المفروض حسب الفواتير","إجمالي الواصل الفعلي بالبنك","فرق الفاتورة عن البنك"],
    [[financialIssues.length, financialExpected, financialActual, financialExpected - financialActual]]);

  const needsReview = results.filter(r=>r.type==="saved").filter((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    return sm.isAmountDiff || sm.isAccountTypeDiff;
  });
  const amountDiffRows = needsReview.map((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    const diff = r.bank.rawAmount - r.cashier.amount;
    const issues = [sm.isAmountDiff ? "الفاتورة مختلفة عن الواصل البنكي" : null, sm.isAccountTypeDiff ? "فرق نوع حساب" : null].filter(Boolean).join(" + ");
    return [r.cashier.name, r.bank.description, sm.type, r.bank.rawAmount, r.cashier.amount,
      sm.isAmountDiff ? fmtNum(-diff) : "0", r.bank.accountType || "—", r.cashier.accountType || "—",
      issues, sm.isManual ? "يدوي" : "تلقائي", sm.note || "", r.bank.date, r.cashier.date];
  });
  add("⚠️ فروقات تحتاج تدقيق",
    ["البيان (الكاشير)","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","الفرق (بنك - كاشير)","نوع حساب البنك","نوع حساب الكاشير","نوع المشكلة","نوع المطابقة","ملاحظة","ت.البنك","ت.الكاشير"],
    amountDiffRows);

  const pendingRows = results.filter(r=>r.type==="pending").map((r:any) => [
    r.cashier.name, r.bank.description, r.cashier.type, r.bank.rawAmount, r.cashier.amount,
    r.isApprox ? "تقريبي" : "ممتاز", r.matchType || "عادي", r.amountDiff ? fmtNum(r.amountDiff) : "0",
    r.bank.accountType || "—", r.cashier.accountType || "—", r.accountTypeDiff ? "نعم" : "لا",
    r.isBundled ? "تجميع" : "فردي", r.bank.date, r.cashier.date
  ]);
  add("⏳ مطابقات منتظرة (غير محفوظة)",
    ["البيان (الكاشير)","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","الحالة","نوع المطابقة","فرق المبلغ","نوع حساب البنك","نوع حساب الكاشير","اختلاف نوع حساب","نوع التجميع","ت.البنك","ت.الكاشير"],
    pendingRows);

  const manualRows = results
    .filter((r): r is Extract<MatchResult, { type: "saved" }> => r.type === "saved" && Boolean(r.savedMatch?.isManual))
    .map((r) => {
      const sm = r.savedMatch;
      return [r.cashier.name, r.bank.description, sm.type, r.bank.rawAmount, r.cashier.amount, r.bank.date, r.cashier.date, sm.note || ""];
    });
  add("🛠️ مطابقات يدوية (محفوظة)",
    ["البيان (الكاشير)","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","ت.البنك","ت.الكاشير","ملاحظة"],
    manualRows);

  const visaRows = visaItems.map(v => [v.name, v.visaNumber, v.type, v.amount, v.date, v.note || "", v.movedAt]);
  add("💳 الفيزا", ["الاسم","رقم الفيزا","النوع","المبلغ","التاريخ","ملاحظة","تاريخ النقل"], visaRows);
  const jawwalRows = results.filter(r => r.type === "jawwalPay").map((r: any) => [
    r.cashier.name, r.cashier.rawName, r.cashier.type, r.cashier.amount,
    r.cashier.accountType || "—", r.cashier.date
  ]);
  add("جوال بي", ["الاسم","البيان الأصلي","النوع","المبلغ","نوع الحساب","التاريخ"], jawwalRows);
  const mahmoudWalletRows = results.filter(r => r.type === "mahmoudWallet").map((r: any) => [
    r.cashier.name, r.cashier.rawName, r.cashier.type, r.cashier.amount,
    r.cashier.accountType || "—", r.cashier.date
  ]);
  add("محفظة محمود", ["الاسم","البيان الأصلي","النوع","المبلغ","نوع الحساب","التاريخ"], mahmoudWalletRows);

  const uCashier = results.filter(r=>r.type==="unmatchedCashier").map((r:any) => [
    r.cashier.name, r.cashier.type, r.cashier.amount, r.cashier.accountType || "—", r.cashier.date, r.reason]);
  add("فواتير لم يجد لها حوالات مطابقة", ["البيان","النوع","المبلغ","نوع الحساب المسجل","اسم المستخدم","التاريخ","السبب"], uCashier.map((row, index) => {
    const item = results.filter(r => r.type === "unmatchedCashier")[index] as any;
    return [...row.slice(0, 3), row[3], item?.cashier?.rawName || "—", row[4], row[5]];
  }));

  const uBank = results.filter(r=>r.type==="unmatchedBank").map((r:any) => [
    r.bank.description, r.bank.type, r.bank.rawAmount, r.bank.accountType || "—", r.bank.date, r.reason]);
  add("حوالات لم يجد لها فواتير مسجلة", ["البيان","النوع","المبلغ","نوع الحساب المسجل","التاريخ","السبب"], uBank);

  const heldRows = heldItems.map(h => {
    const isBank = h.kind === "bank";
    const label = isBank ? (h.data as BankRow).description : (h.data as CashierRow).name;
    const amt = isBank ? (h.data as BankRow).rawAmount : (h.data as CashierRow).amount;
    return [ isBank ? "بنك" : "كاشير", label, h.data.type, amt, h.data.accountType || "—", h.heldAt, h.note || "" ];
  });
  add("🗂️ المعلقات", ["المصدر","البيان","النوع","المبلغ","نوع الحساب","تاريخ التعليق","ملاحظة"], heldRows);

  if (resumeData) {
    try { embedResumeSheet(wb, resumeData); } catch (e) { console.error("تعذّر تضمين بيانات الاستكمال بالملف:", e); }
  }
  XLSX.writeFile(wb,"تقرير_التسوية.xlsx");
}

// ─── استكمال العمل من الملف المُصدَّر ──────────────────────────────────────────
type ResumeData = {
  bankHeaders: string[];
  bankRowsRaw: Record<string,unknown>[];
  bankMap: { date:string; desc:string; debit:string; credit:string; accountType?:string };
  bankSwap: boolean;
  bankFileSessionId?: number;
  cashHeaders: string[];
  cashRowsRaw: Record<string,unknown>[];
  cashMap: { date:string; name:string; debit:string; credit:string; accountType?:string };
  cashSwap: boolean;
  cashFileSessionId?: number;
  manualGroups: ManualMatchGroup[];
  savedMatches: SavedMatch[];
  rejectedPairs: string[];
  visaItems: VisaItem[];
  heldItems: HeldItem[];
  returnedHeldBank?: BankRow[];
  returnedHeldCashier?: BankRow[] | CashierRow[];
};

const RESUME_SHEET_NAME = "⚙️ بيانات_الاستكمال";
const RESUME_CHUNK_SIZE = 25000;

function embedResumeSheet(wb: XLSX.WorkBook, data: ResumeData) {
  const json = JSON.stringify(data);
  const MAX_JSON_SIZE = 4_000_000;
  if (json.length > MAX_JSON_SIZE) {
    console.warn(`بيانات الجلسة كبيرة جداً (${json.length} حرف) لتضمينها بالملف - تم تجاهل خاصية الاستكمال لهالمرة.`);
    return;
  }
  const chunks: string[] = [];
  for (let i = 0; i < json.length; i += RESUME_CHUNK_SIZE) {
    chunks.push(json.slice(i, i + RESUME_CHUNK_SIZE));
  }
  const rows: string[][] = [
    ["⚠️ هذا الشيت خاص بالبرنامج - لا تعدّله أو تحذفه. هو اللي بيخلي الملف قابل للاستكمال لاحقاً عبر زر (استيراد ملف سابق)."],
    ...chunks.map(c => [c])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, RESUME_SHEET_NAME);
}

async function extractResumeData(file: File): Promise<ResumeData | null> {
  try {
    const buf = await readFileBuf(file);
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames.find(n => n === RESUME_SHEET_NAME || n.includes("استكمال"));
    if (!sheetName) return null;
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
    const json = rows.slice(1).map(r => (r?.[0] ?? "")).join("");
    if (!json.trim()) return null;
    return JSON.parse(json) as ResumeData;
  } catch {
    return null;
  }
}

function autoDetect(headers: string[], hints: string[]): string {
  const scores=headers.map(h=>({h,s:hints.reduce((a,hint)=>a+(h.toLowerCase().includes(hint.toLowerCase())?1:0),0)})).sort((a,b)=>b.s-a.s);
  return scores[0]?.s>0?scores[0].h:"";
}
const HINTS={
  date:   ["date","تاريخ","value","posting"],
  desc:   ["description","narrative","detail","إيضاح","بيان","وصف","narr","payee","particular"],
  debit:  ["debit","مدين","مستلم","مبالغ مستلمة","received","deposit","cr","دخول","إيداع"],
  credit: ["credit","دائن","مدفوع","مبالغ مدفوعة","paid","withdrawal","dr","خروج","سحب"],
  name:   ["name","اسم","customer","client","employee","موظف","عميل","الزبون","بيان","وصف","narrative","detail","إيضاح"],
  accountType: ["account type","account_type","نوع الحساب","نوع حساب","نوع","حساب","account"],
  ref:    ["reference","ref","transaction","trx","id","reference no","reference number","مرجع","رقم المرجع","رقم الحركة","رقم الحوالة","transaction id","txn","voucher","سند","رقم"],
};

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
      <span className="text-sm">اسحب ملف الكشف أو انقر هنا</span>
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

// ─── Manual Workbench ──────────────────────────────────────────────────────
type SortMode = "amount_desc" | "amount_asc" | "name";

function ListSection({
  title, icon, rows, selectedIds, onToggle, search, onSearch, sort, onSort, onHold
}: {
  title:string; icon:string; rows:(BankRow|CashierRow)[]; selectedIds:Set<number>;
  onToggle:(id:number)=>void; search:string; onSearch:(s:string)=>void;
  sort:SortMode; onSort:(s:SortMode)=>void;
  onHold:(row:BankRow|CashierRow)=>void;
}) {
  return (
    <div className="space-y-2">
      <div className="flex items-center gap-2 flex-wrap">
        <span className="font-semibold text-sm">{icon} {title}</span>
        <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded-full">{rows.length}</span>
        <div className="flex items-center gap-1.5 mr-auto">
          <div className="relative">
            <Search className="absolute right-2 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-muted-foreground"/>
            <input value={search} onChange={e=>onSearch(e.target.value)} placeholder="بحث بالبيان أو المبلغ..."
              className="pl-2 pr-7 py-1.5 text-xs border border-border rounded-lg bg-input-background focus:outline-none w-48 focus:ring-1 focus:ring-ring"/>
          </div>
          <select value={sort} onChange={e=>onSort(e.target.value as SortMode)}
            className="text-xs border border-border rounded-lg px-2 py-1.5 bg-input-background focus:outline-none">
            <option value="amount_desc">المبلغ: الأكبر أولاً</option>
            <option value="amount_asc">المبلغ: الأصغر أولاً</option>
            <option value="name">ترتيب: الاسم</option>
          </select>
        </div>
      </div>
      <div className="border border-border rounded-xl overflow-hidden">
        <div className="max-h-56 overflow-y-auto">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-muted z-10"><tr>
              <th className="px-3 py-2.5 w-8"></th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">البيان</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">النوع</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">المبلغ</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">التاريخ</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">تعليق</th>
            </tr></thead>
            <tbody>
              {rows.map(r=>{
                const isBank="rawAmount" in r;
                const id=r.id, selected=selectedIds.has(id);
                const amt=isBank?(r as BankRow).rawAmount:(r as CashierRow).amount;
                const label=isBank?(r as BankRow).description:(r as CashierRow).name;
                const type=isBank?(r as BankRow).type:(r as CashierRow).type;
                const date=isBank?(r as BankRow).date:(r as CashierRow).date;
                const typeColor = type === "مدفوع" ? "text-red-600" : "text-green-600";
                const isFromHeld = !!(r as any)._fromHeld;
                return (
                  <tr key={id} onClick={()=>onToggle(id)}
                    className={`border-t border-border cursor-pointer transition-colors ${selected?"bg-blue-50 border-r-2 border-r-blue-500":isFromHeld?"bg-purple-50/60 hover:bg-purple-50":"hover:bg-muted/30"}`}>
                    <td className="px-3 py-2.5">
                      <div className={`w-4 h-4 rounded border-2 flex items-center justify-center transition-colors ${selected?"bg-blue-600 border-blue-600":"border-border"}`}>
                        {selected&&<Check className="w-2.5 h-2.5 text-white"/>}
                      </div>
                    </td>
                    <td className="px-3 py-2.5 font-medium">
                      {label}
                      {isFromHeld && (
                        <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 align-middle">📦 من ملف سابق</span>
                      )}
                    </td>
                    <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{type}</td>
                    <td className={`px-3 py-2.5 font-mono font-semibold ${isBank?"text-blue-700":"text-green-700"}`}>{fmtNum(amt)}</td>
                    <td className="px-3 py-2.5 text-xs text-muted-foreground">{date}</td>
                    <td className="px-3 py-2.5">
                      <button
                        onClick={(e)=>{ e.stopPropagation(); onHold(r); }}
                        title="علّق هذا العنصر لمطابقته لاحقاً"
                        className="flex items-center gap-1 px-2 py-1 rounded-lg text-[10px] font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors">
                        <CreditCard className="w-3 h-3"/>علّق
                      </button>
                    </td>
                  </tr>
                );
              })}
              {rows.length===0&&<tr><td colSpan={6} className="px-3 py-8 text-center text-muted-foreground text-xs">لا توجد سجلات</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ManualWorkbench({
  bankRows, cashierRows, manualGroups, onAddGroup, onRemoveGroup, onBack,
  savedCashierIds, savedBankIds, onHoldBank, onHoldCashier
}: {
  bankRows: BankRow[]; cashierRows: CashierRow[];
  manualGroups: ManualMatchGroup[];
  onAddGroup: (g:ManualMatchGroup)=>void;
  onRemoveGroup: (id:string)=>void;
  onBack: ()=>void;
  savedCashierIds: Set<number>;
  savedBankIds: Set<number>;
  onHoldBank: (b:BankRow)=>void;
  onHoldCashier: (c:CashierRow)=>void;
}) {
  const [bankSearch, setBankSearch] = useState("");
  const [cashierSearch, setCashierSearch] = useState("");
  const [bankSort, setBankSort] = useState<SortMode>("amount_desc");
  const [cashierSort, setCashierSort] = useState<SortMode>("amount_desc");
  const [selectedBankIds, setSelectedBankIds] = useState<Set<number>>(new Set());
  const [selectedCashierIds, setSelectedCashierIds] = useState<Set<number>>(new Set());
  const [noteInput, setNoteInput] = useState("");

  const matchedBankIds = useMemo(()=>new Set(manualGroups.flatMap(g=>g.banks.map(b=>b.id))),[manualGroups]);
  const matchedCashierIds = useMemo(()=>new Set(manualGroups.flatMap(g=>g.cashiers.map(c=>c.id))),[manualGroups]);

  const filteredBanks = useMemo(()=>{
    let rows=bankRows.filter(b => !matchedBankIds.has(b.id) && !savedBankIds.has(b.id));
    if (bankSearch) { const s=bankSearch.toLowerCase(); rows=rows.filter(b=>b.description.toLowerCase().includes(s)||String(b.rawAmount).includes(s)); }
    rows=[...rows].sort((a,b)=>{
      if (bankSort==="name") return a.description.localeCompare(b.description,"ar");
      return bankSort==="amount_asc" ? a.rawAmount-b.rawAmount : b.rawAmount-a.rawAmount;
    });
    return rows;
  },[bankRows,bankSearch,bankSort,matchedBankIds,savedBankIds]);

  const filteredCashiers = useMemo(()=>{
    let rows=cashierRows.filter(c => !matchedCashierIds.has(c.id) && !savedCashierIds.has(c.id));
    if (cashierSearch) { const s=cashierSearch.toLowerCase(); rows=rows.filter(c=>c.name.toLowerCase().includes(s)||c.rawName.toLowerCase().includes(s)||String(c.amount).includes(s)); }
    rows=[...rows].sort((a,b)=>{
      if (cashierSort==="name") return a.name.localeCompare(b.name,"ar");
      return cashierSort==="amount_asc" ? a.amount-b.amount : b.amount-a.amount;
    });
    return rows;
  },[cashierRows,cashierSearch,cashierSort,matchedCashierIds,savedCashierIds]);

  const toggleBank=(id:number)=>setSelectedBankIds(p=>{
    const n=new Set(p);
    if (n.has(id)) n.delete(id); else n.add(id);
    return n;
  });
  const toggleCashier=(id:number)=>setSelectedCashierIds(p=>{
    const n=new Set(p);
    if (n.has(id)) n.delete(id); else n.add(id);
    return n;
  });

  const handleMatch=()=>{
    if (!selectedBankIds.size||!selectedCashierIds.size) return;
    const banks=bankRows.filter(b=>selectedBankIds.has(b.id));
    const cashiers=cashierRows.filter(c=>selectedCashierIds.has(c.id));
    onAddGroup({ id:`grp-${Date.now()}`, banks, cashiers, note: noteInput.trim() || undefined });
    setSelectedBankIds(new Set());
    setSelectedCashierIds(new Set());
    setNoteInput("");
  };

  const selBTotal=bankRows.filter(b=>selectedBankIds.has(b.id)).reduce((s,b)=>s+b.rawAmount,0);
  const selCTotal=cashierRows.filter(c=>selectedCashierIds.has(c.id)).reduce((s,c)=>s+c.amount,0);
  const amtMatch=selectedBankIds.size>0&&selectedCashierIds.size>0&&Math.abs(selBTotal-selCTotal)<=0.01;

  useEffect(()=>{
    if (selectedBankIds.size>0 && selectedCashierIds.size>0 && !amtMatch && !noteInput) {
      setNoteInput(`اختلاف مبلغ: ${fmtNum(Math.abs(selBTotal-selCTotal))}`);
    }
  }, [selectedBankIds.size, selectedCashierIds.size, amtMatch, noteInput, selBTotal, selCTotal]);

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground">
      <div className="sticky top-0 z-20 bg-card border-b border-border px-6 py-3.5 flex items-center gap-3 shadow-sm">
        <button onClick={onBack} className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors">
          <ChevronLeft className="w-4 h-4"/>رجوع
        </button>
        <div className="w-px h-4 bg-border"/>
        <h1 className="font-bold text-sm">المطابقة اليدوية المتقدمة</h1>
        {selectedBankIds.size>0||selectedCashierIds.size>0?(
          <div className="mr-auto flex items-center gap-3 flex-wrap">
            <div className="text-xs">
              <span className="text-blue-700 font-semibold">{selectedBankIds.size} بنك</span>
              <span className="text-muted-foreground mx-1">↔</span>
              <span className="text-green-700 font-semibold">{selectedCashierIds.size} كاشير</span>
              {selectedBankIds.size>0&&selectedCashierIds.size>0&&(
                <span className={`mr-2 px-1.5 py-0.5 rounded text-xs font-medium ${amtMatch?"bg-green-100 text-green-700":"bg-amber-100 text-amber-700"}`}>
                  {amtMatch?"المبالغ متطابقة ✓":`فرق: ${fmtNum(Math.abs(selBTotal-selCTotal))}`}
                </span>
              )}
            </div>
            {selectedBankIds.size>0&&selectedCashierIds.size>0&&(
              <input value={noteInput} onChange={e=>setNoteInput(e.target.value)}
                placeholder="سبب الربط (اختياري): اختلاف مبلغ، تحويل من الزوجة..."
                className="text-xs border border-border rounded-lg px-2.5 py-1.5 bg-input-background focus:outline-none w-64 focus:ring-1 focus:ring-ring"/>
            )}
            <button onClick={handleMatch} disabled={!selectedBankIds.size||!selectedCashierIds.size}
              className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 disabled:opacity-40 transition-colors">
              <Save className="w-4 h-4"/>حفظ المطابقة (تؤكد فوراً)
            </button>
            <button onClick={()=>{setSelectedBankIds(new Set());setSelectedCashierIds(new Set());setNoteInput("");}}
              className="p-2 text-muted-foreground hover:bg-muted rounded-lg transition-colors">
              <X className="w-4 h-4"/>
            </button>
          </div>
        ):(
          <div className="mr-auto text-xs text-muted-foreground">اختر من كلا القائمتين ثم اضغط "حفظ المطابقة"</div>
        )}
      </div>

      <div className="max-w-5xl mx-auto px-6 py-6 space-y-6">
        {bankRows.length===0&&cashierRows.length===0&&(
          <div className="text-center py-16 text-muted-foreground">
            <Info className="w-10 h-10 mx-auto mb-3 opacity-40"/>
            <p>لا توجد بيانات محملة. ارجع وأضف ملفات البنك والكاشير أولاً.</p>
          </div>
        )}

        <ListSection title="كشف البنك" icon="" rows={filteredBanks} selectedIds={selectedBankIds} 
          onToggle={toggleBank} search={bankSearch} onSearch={setBankSearch}
          sort={bankSort} onSort={setBankSort}
          onHold={(row)=>{ onHoldBank(row as BankRow); setSelectedBankIds(p=>{const n=new Set(p);n.delete(row.id);return n;}); }}/>

        <ListSection title="كشف الكاشير" icon="" rows={filteredCashiers} selectedIds={selectedCashierIds}
          onToggle={toggleCashier} search={cashierSearch} onSearch={setCashierSearch}
          sort={cashierSort} onSort={setCashierSort}
          onHold={(row)=>{ onHoldCashier(row as CashierRow); setSelectedCashierIds(p=>{const n=new Set(p);n.delete(row.id);return n;}); }}/>

        {(selectedBankIds.size>0&&selectedCashierIds.size>0)&&(
          <button onClick={handleMatch}
            className="w-full py-3.5 rounded-xl font-medium flex items-center justify-center gap-2 text-sm transition-colors bg-green-600 hover:bg-green-700 text-white">
            <Save className="w-5 h-5"/>
            حفظ {selectedBankIds.size} حوالة بنكية ↔ {selectedCashierIds.size} فاتورة كاشier (تؤكد فوراً)
            {!amtMatch&&<span className="text-xs opacity-80">(مجاميع مختلفة)</span>}
          </button>
        )}

        {manualGroups.length>0&&(
          <div className="space-y-3">
            <h2 className="font-semibold text-sm">المجموعات المربوطة يدوياً (سيتم حفظها فوراً)</h2>
            <div className="text-xs text-muted-foreground">المطابقة اليدوية تذهب للمؤكدة مباشرة</div>
            <div className="space-y-2">
              {manualGroups.map(g=>{
                const gDiff = g.banks.reduce((s,b)=>s+b.rawAmount,0) - g.cashiers.reduce((s,c)=>s+c.amount,0);
                const hasDiff = Math.abs(gDiff) > 0.01;
                return (
                <div key={g.id} className="border border-green-200 bg-green-50/60 rounded-xl p-4">
                  <div className="flex items-start gap-3">
                    <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <div className="text-xs font-medium text-muted-foreground mb-2">البنك ({g.banks.length})</div>
                        {g.banks.map(b=>(
                          <div key={b.id} className="flex items-center gap-2 text-sm mb-1">
                            <span className="font-medium truncate">{b.description}</span>
                            <span className={`text-xs font-semibold ${b.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{b.type}</span>
                            <span className="text-blue-700 font-mono text-xs shrink-0">{fmtNum(b.rawAmount)}</span>
                          </div>
                        ))}
                        <div className="text-xs text-blue-700 font-semibold mt-1">المجموع: {fmtNum(g.banks.reduce((s,b)=>s+b.rawAmount,0))}</div>
                      </div>
                      <div>
                        <div className="text-xs font-medium text-muted-foreground mb-2">الكاشير ({g.cashiers.length})</div>
                        {g.cashiers.map(c=>(
                          <div key={c.id} className="flex items-center gap-2 text-sm mb-1">
                            <span className="font-medium">{c.name}</span>
                            <span className={`text-xs font-semibold ${c.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{c.type}</span>
                            <span className="text-green-700 font-mono text-xs shrink-0">{fmtNum(c.amount)}</span>
                          </div>
                        ))}
                        <div className="text-xs text-green-700 font-semibold mt-1">المجموع: {fmtNum(g.cashiers.reduce((s,c)=>s+c.amount,0))}</div>
                      </div>
                    </div>
                    <button onClick={()=>onRemoveGroup(g.id)} className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors shrink-0">
                      <Trash2 className="w-4 h-4"/>
                    </button>
                  </div>
                  <div className="mt-3 pt-3 border-t border-green-200/70 flex items-center gap-2 flex-wrap">
                    <span className="text-xs font-medium px-2 py-1 rounded-lg bg-green-100 text-green-800 flex items-center gap-1">
                      <Check className="w-3 h-3"/>سيتم حفظها في المؤكدة
                    </span>
                    {(hasDiff||g.note)&&(
                      <>
                        {hasDiff&&(
                          <span className="text-xs font-medium px-2 py-1 rounded-lg bg-orange-100 text-orange-800 flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3"/>فرق مبلغ {fmtNum(Math.abs(gDiff))}
                          </span>
                        )}
                        {g.note&&(
                          <span className="text-xs px-2 py-1 rounded-lg bg-white border border-amber-300 text-amber-800">📝 {g.note}</span>
                        )}
                      </>
                    )}
                  </div>
                </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── مساعد المطابقة ────────────────────────────────────────────────────────
function MatchAssistant({
  cashierRows, bankRows, rejectedPairs, onMatch, onHold, onHoldBank, onReject, onBack
}: {
  cashierRows: CashierRow[];
  bankRows: BankRow[];
  rejectedPairs: Set<string>;
  onMatch: (c: CashierRow, b: BankRow) => void;
  onHold: (c: CashierRow) => void;
  onHoldBank: (b: BankRow) => void;
  onReject: (c: CashierRow, b: BankRow) => void;
  onBack: () => void;
}) {
  const [skipped, setSkipped] = useState<Set<number>>(new Set());
  const [search, setSearch] = useState("");
  const deferredSearch = useDeferredValue(search);

  const queue = useMemo(() => {
    const rest = cashierRows.filter(c => !skipped.has(c.id));
    const later = cashierRows.filter(c => skipped.has(c.id));
    return [...rest, ...later];
  }, [cashierRows, skipped]);

  const current = queue[0] ?? null;

  useEffect(() => { setSearch(""); }, [current?.id]);

  const candidates = useMemo(() => {
    if (!current) return [];
    let list = bankRows.filter(b => {
      if (b.type !== current.type) return false;
      const pairKey = `${current.id}-${b.id}`;
      return !rejectedPairs.has(pairKey);
    });
    if (deferredSearch.trim()) {
      const s = deferredSearch.trim().toLowerCase();
      list = list.filter(b => b.description.toLowerCase().includes(s) || String(b.rawAmount).includes(s));
    }
    return list
      .map(b => ({
        bank: b,
        diff: Math.abs(b.rawAmount - current.matchAmount),
        exact: Math.abs(b.rawAmount - current.matchAmount) <= 0.01,
        accountDiff: accountTypesDiffer(b.accountType, current.accountType),
        sim: nameSim(current.name, b.description)
      }))
      .sort((a, b) => a.diff - b.diff)
      .slice(0, 40);
  }, [current, bankRows, deferredSearch, rejectedPairs]);

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground">
      <div className="sticky top-0 z-20 bg-card border-b border-border px-6 py-3.5 flex items-center gap-3 shadow-sm flex-wrap">
        <button onClick={onBack} className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors">
          <ChevronLeft className="w-4 h-4"/>رجوع
        </button>
        <div className="w-px h-4 bg-border"/>
        <h1 className="font-bold text-sm">مساعد المطابقة</h1>
        <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded-full">{queue.length} فاتورة متبقية</span>
      </div>

      <div className="max-w-4xl mx-auto px-6 py-6 space-y-5">
        {!current ? (
          <div className="text-center py-20 text-muted-foreground">
            <Check className="w-12 h-12 mx-auto mb-3 text-green-500"/>
            <p className="text-sm font-medium">ممتاز! ما في فواتير كاشير قيد الانتظار حالياً 🎉</p>
          </div>
        ) : (
          <>
            <div className="border-2 border-indigo-300 bg-indigo-50/50 rounded-2xl p-5">
              <div className="flex items-center justify-between gap-3 flex-wrap">
                <div>
                  <div className="text-xs text-indigo-600 font-medium mb-1">فاتورة الكاشير الحالية</div>
                  <div className="text-lg font-bold">{current.name}</div>
                  <div className="flex items-center gap-2 mt-1.5 flex-wrap">
                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${current.type === "مدفوع" ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>{current.type}</span>
                    <span className="text-lg font-mono font-bold text-indigo-700">{fmtNum(current.matchAmount)}</span>
                    {current.accountType && (
                      <span className="text-xs px-2 py-0.5 rounded-full bg-purple-100 text-purple-700 border border-purple-200">{current.accountType}</span>
                    )}
                    <span className="text-xs text-muted-foreground">{current.date}</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => setSkipped(prev => new Set(prev).add(current.id))}
                    className="flex items-center gap-1.5 px-3 py-2 bg-card border border-border rounded-lg text-xs font-medium hover:bg-muted/40 transition-colors">
                    ⏭️ تخطي
                  </button>
                  <button onClick={() => onHold(current)}
                    className="flex items-center gap-1.5 px-3 py-2 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-xs font-medium hover:bg-purple-100 transition-colors">
                    <CreditCard className="w-3.5 h-3.5"/>نقل للمعلقات
                  </button>
                </div>
              </div>
            </div>

            <div className="relative">
              <Search className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground"/>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="بحث بالبيان أو المبلغ داخل الحوالات..."
                className="w-full pl-3 pr-9 py-2.5 text-sm border border-border rounded-lg bg-input-background focus:outline-none focus:ring-1 focus:ring-ring"/>
            </div>

            <div className="space-y-2">
              <p className="text-xs font-medium text-muted-foreground">اضغط على بيان الحوالة لمطابقتها فوراً، أو استخدم "مش هاي"/"علّق" لإبعاد حوالة غير مرتبطة:</p>
              <div className="border border-border rounded-xl overflow-hidden divide-y divide-border">
                {candidates.map(({bank, diff, exact, accountDiff, sim}) => (
                  <div key={bank.id} className="w-full px-4 py-3 hover:bg-indigo-50/40 transition-colors flex items-center gap-3">
                    <button onClick={() => onMatch(current, bank)} className="flex-1 min-w-0 text-right">
                      <div className="font-medium truncate">{bank.description}</div>
                      <div className="flex items-center gap-2 mt-1 flex-wrap">
                        <span className="font-mono font-bold text-blue-700">{fmtNum(bank.rawAmount)}</span>
                        {exact ? (
                          <span className="text-[10px] px-1.5 py-0.5 rounded bg-green-100 text-green-700 border border-green-300">✓ نفس المبلغ</span>
                        ) : (
                          <span className="text-[10px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">فرق مبلغ: {fmtNum(diff)}</span>
                        )}
                        {bank.accountType && (
                          <span className={`text-[10px] px-1.5 py-0.5 rounded border ${accountDiff ? "bg-orange-100 text-orange-700 border-orange-300" : "bg-purple-50 text-purple-700 border-purple-200"}`}>
                            {bank.accountType}{accountDiff ? " (مختلف)" : ""}
                          </span>
                        )}
                        <span className="text-[10px] text-muted-foreground">تشابه اسم: {(sim*100).toFixed(0)}٪</span>
                        <span className="text-[10px] text-muted-foreground">{bank.date}</span>
                      </div>
                    </button>
                    <div className="flex items-center gap-1 shrink-0">
                      <button onClick={() => onReject(current, bank)}
                        title="مش هاي — أبعدها عن هذه الفاتورة بس خلّيها متاحة لغيرها"
                        className="px-2 py-1.5 bg-gray-50 text-gray-500 border border-gray-200 rounded-lg text-[10px] font-medium hover:bg-gray-100 transition-colors">
                        ✕ مش هاي
                      </button>
                      <button onClick={() => onHoldBank(bank)}
                        title="علّق هذه الحوالة كلياً"
                        className="px-2 py-1.5 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-[10px] font-medium hover:bg-purple-100 transition-colors flex items-center gap-1">
                        <CreditCard className="w-3 h-3"/>علّق
                      </button>
                    </div>
                  </div>
                ))}
                {!candidates.length && (
                  <p className="px-4 py-8 text-center text-xs text-muted-foreground">لا توجد حوالات بنفس نوع الحركة، أو كلها استُبعدت بالبحث/الرفض.</p>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── Projects ──────────────────────────────────────────────────────────────
interface SavedProject {
  id: string;
  name: string;
  savedAt: string;
  bankHeaders: string[];
  bankRowsRaw: Record<string,unknown>[];
  bankMap: { date:string; desc:string; debit:string; credit:string; accountType?:string };
  bankSwap: boolean;
  bankFileSessionId?: number;
  cashHeaders: string[];
  cashRowsRaw: Record<string,unknown>[];
  cashMap: { date:string; name:string; debit:string; credit:string; accountType?:string };
  cashSwap: boolean;
  cashFileSessionId?: number;
  manualGroups: ManualMatchGroup[];
  savedMatches: SavedMatch[];
  rejectedPairs: string[];
  visaItems: VisaItem[];
  heldItems: HeldItem[];
  returnedHeldBank?: BankRow[];
  returnedHeldCashier?: CashierRow[];
}

async function loadProjectsList(): Promise<SavedProject[]> {
  return storageGet<SavedProject[]>("recon_projects", []);
}
async function persistProjectsList(list: SavedProject[]): Promise<void> {
  await storageSet("recon_projects", list);
}

function ProjectsBar({
  onSave, onLoad, onDelete
}: {
  onSave: (name:string)=>void;
  onLoad: (p:SavedProject)=>void;
  onDelete: (id:string)=>Promise<void>;
}) {
  const [open, setOpen] = useState(false);
  const [projects, setProjects] = useState<SavedProject[]>([]);
  const [loading, setLoading] = useState(false);

  const refresh = async () => {
    setLoading(true);
    const list = await loadProjectsList();
    setProjects(list);
    setLoading(false);
  };

  useEffect(() => { refresh(); }, []);

  return (
    <div className="relative">
      <div className="flex items-center gap-2">
        <button
          onClick={() => {
            const name = window.prompt("اسم المشروع (مثال: تسوية 17-07-2026):", "");
            if (name && name.trim()) { onSave(name.trim()); setTimeout(refresh, 300); }
          }}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors">
          <FolderPlus className="w-4 h-4"/>حفظ كمشروع
        </button>
        <button
          onClick={() => { refresh(); setOpen(o=>!o); }}
          className="flex items-center gap-2 px-4 py-2 bg-card border border-border rounded-lg text-sm font-medium hover:bg-muted/40 transition-colors">
          <FolderOpen className="w-4 h-4"/>المشاريع المحفوظة ({projects.length})
        </button>
      </div>
      {open && (
        <div className="absolute z-30 mt-2 w-96 bg-card border border-border rounded-xl shadow-lg p-2 max-h-80 overflow-y-auto">
          {loading && <p className="text-xs text-muted-foreground p-3 text-center">جارِ التحميل...</p>}
          {!loading && projects.length === 0 && (
            <p className="text-xs text-muted-foreground p-3 text-center">لا توجد مشاريع محفوظة بعد</p>
          )}
          {projects.map(p=>(
            <div key={p.id} className="flex items-center justify-between gap-2 p-2.5 rounded-lg hover:bg-muted/30">
              <div className="min-w-0">
                <div className="text-sm font-medium truncate">{p.name}</div>
                <div className="text-[10px] text-muted-foreground">{p.savedAt} · {p.savedMatches.length} مؤكدة</div>
              </div>
              <div className="flex items-center gap-1 shrink-0">
                <button onClick={()=>{ onLoad(p); setOpen(false); }}
                  className="px-2.5 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-lg text-xs hover:bg-blue-100 transition-colors">
                  فتح
                </button>
                <button onClick={async ()=>{
                    if (window.confirm(`حذف مشروع "${p.name}" نهائياً؟`)) { await onDelete(p.id); refresh(); }
                  }}
                  className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors">
                  <Trash2 className="w-3.5 h-3.5"/>
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ConfidenceBadge({ score }: { score?: number }) {
  if (score == null) return <span className="text-xs text-muted-foreground">—</span>;
  const tone = score > 90 ? "bg-green-100 text-green-700 border-green-200" : score >= 70 ? "bg-amber-100 text-amber-700 border-amber-200" : "bg-red-100 text-red-700 border-red-200";
  return <span className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[11px] font-semibold ${tone}`}>{score}% ثقة</span>;
}

function FilterBar({
  search, onSearch, type, onType, minAmount, onMinAmount, maxAmount, onMaxAmount
}: {
  search: string; onSearch: (value: string) => void;
  type: "all" | "مدفوع" | "مستلم"; onType: (value: "all" | "مدفوع" | "مستلم") => void;
  minAmount: string; onMinAmount: (value: string) => void;
  maxAmount: string; onMaxAmount: (value: string) => void;
}) {
  return (
    <div className="flex flex-wrap items-center gap-2 border-b bg-slate-50/80 px-4 py-3">
      <div className="relative min-w-52 flex-1">
        <Search className="absolute right-2.5 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
        <input value={search} onChange={e => onSearch(e.target.value)} placeholder="ابحث بالاسم أو البيان..." className="w-full rounded-lg border bg-white py-2 pl-3 pr-9 text-xs outline-none focus:ring-2 focus:ring-blue-200" />
      </div>
      <select value={type} onChange={e => onType(e.target.value as "all" | "مدفوع" | "مستلم")} className="rounded-lg border bg-white px-3 py-2 text-xs outline-none">
        <option value="all">كل الأنواع</option><option value="مدفوع">مدفوع</option><option value="مستلم">مستلم</option>
      </select>
      <input value={minAmount} onChange={e => onMinAmount(e.target.value)} type="number" min="0" placeholder="من مبلغ" className="w-24 rounded-lg border bg-white px-2.5 py-2 text-xs outline-none" />
      <input value={maxAmount} onChange={e => onMaxAmount(e.target.value)} type="number" min="0" placeholder="إلى مبلغ" className="w-24 rounded-lg border bg-white px-2.5 py-2 text-xs outline-none" />
    </div>
  );
}

// ─── لوحة السحب والاستبدال (تظهر جمب الجدول وتزيحه) ──────────────────────────
function MatchDrawer({
  mode, cashier, oldBank, nameSearch, amountFrom, amountTo, banks, cashiers, draggedItemId,
  onNameSearch, onAmountFrom, onAmountTo, onDragStart, onDragEnd, onReplaceBank, onReplaceCashier, onClose
}: {
  mode: "bank" | "cashier"; cashier: CashierRow; oldBank: BankRow;
  nameSearch: string; amountFrom: string; amountTo: string;
  banks: BankRow[]; cashiers: CashierRow[]; draggedItemId: number | null;
  onNameSearch: (v:string)=>void; onAmountFrom: (v:string)=>void; onAmountTo: (v:string)=>void;
  onDragStart: (id:number)=>void; onDragEnd: ()=>void;
  onReplaceBank: (b:BankRow)=>void; onReplaceCashier: (c:CashierRow)=>void; onClose: ()=>void;
}) {
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const dragState = useRef<{ startX: number; startY: number; originX: number; originY: number } | null>(null);
  const handleDragMove = (event: PointerEvent) => {
    const current = dragState.current;
    if (!current) return;
    setPosition({
      x: current.originX + event.clientX - current.startX,
      y: current.originY + event.clientY - current.startY
    });
  };
  const stopDragging = () => {
    dragState.current = null;
    window.removeEventListener("pointermove", handleDragMove);
    window.removeEventListener("pointerup", stopDragging);
  };
  const startDragging = (event: React.PointerEvent<HTMLDivElement>) => {
    if ((event.target as HTMLElement).closest("button, input, select")) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    dragState.current = { startX: event.clientX, startY: event.clientY, originX: position.x, originY: position.y };
    window.addEventListener("pointermove", handleDragMove);
    window.addEventListener("pointerup", stopDragging);
  };
  return (
    <div
      className="fixed z-50 w-[550px] max-w-[95vw] h-[70vh] max-h-[calc(100vh-2rem)] bg-white border border-border rounded-xl shadow-2xl flex flex-col overflow-hidden isolate"
      style={{ top: "15vh", right: "1rem", transform: `translate(${position.x}px, ${position.y}px)` }}
      dir="rtl"
    >
      <div
        onPointerDown={startDragging}
        className={`flex cursor-move items-center justify-between gap-3 px-4 py-3.5 text-white shrink-0 select-none ${mode === "bank" ? "bg-gradient-to-l from-blue-600 to-indigo-600" : "bg-gradient-to-l from-emerald-600 to-teal-600"}`}
      >
        <div className="min-w-0 flex items-center gap-2.5">
          <span className="flex items-center justify-center w-9 h-9 rounded-xl bg-white/15 shrink-0">
            {mode === "bank" ? <Landmark className="w-5 h-5"/> : <Users className="w-5 h-5"/>}
          </span>
          <div className="min-w-0">
            <div className="text-[11px] text-white/80">{mode === "bank" ? "قائمة سحب الحوالات" : "قائمة سحب فواتير الكاشير"} · اسحب من هنا لتحريك اللوحة</div>
            <div className="max-w-[min(480px,42vw)] whitespace-normal break-words text-sm font-bold leading-6">{mode === "bank" ? cashier.name : oldBank.description}</div>
          </div>
        </div>
        <button onClick={onClose} className="p-1.5 hover:bg-white/15 rounded-lg transition-colors shrink-0"><X className="w-4 h-4"/></button>
      </div>

      <div className="px-4 py-3 border-b border-border bg-slate-50 space-y-1.5 shrink-0">
        <div className="flex items-center gap-2 flex-wrap text-xs">
          <span className={`font-semibold px-2 py-0.5 rounded-full ${(mode === "bank" ? cashier.type : oldBank.type) === "مدفوع" ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>
            {mode === "bank" ? cashier.type : oldBank.type}
          </span>
          <span className="font-mono font-bold text-indigo-700">{fmtNum(mode === "bank" ? cashier.matchAmount : oldBank.rawAmount)}</span>
        </div>
        <div className="text-[11px] text-muted-foreground">
          {mode === "bank"
            ? <>سيتم استبدال الحوالة الحالية: <span className="text-foreground font-medium">{oldBank.description}</span> ({fmtNum(oldBank.rawAmount)})</>
            : <>سيتم استبدال الفاتورة الحالية: <span className="text-foreground font-medium">{cashier.name}</span> ({fmtNum(cashier.matchAmount)})</>}
        </div>
      </div>

      <div className="px-4 py-3 border-b border-border space-y-2 bg-white shrink-0">
        <div className="relative">
          <Search className="absolute right-2.5 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground"/>
          <input value={nameSearch} onChange={e => onNameSearch(e.target.value)}
            placeholder={mode === "bank" ? "بحث ببيان الحوالة..." : "بحث باسم الفاتورة..."}
            className="w-full pl-3 pr-9 py-2 text-sm border border-border rounded-lg bg-input-background focus:outline-none focus:ring-1 focus:ring-ring"/>
        </div>
        <div className="flex items-center gap-2">
          <input type="number" value={amountFrom} onChange={e => onAmountFrom(e.target.value)} placeholder="من مبلغ"
            className="flex-1 rounded-lg border border-border bg-input-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring"/>
          <span className="text-muted-foreground text-xs">إلى</span>
          <input type="number" value={amountTo} onChange={e => onAmountTo(e.target.value)} placeholder="إلى مبلغ"
            className="flex-1 rounded-lg border border-border bg-input-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring"/>
        </div>
        <div className="text-[10px] text-muted-foreground leading-relaxed">
          اسحب {mode === "bank" ? "الحوالة" : "الفاتورة"} من القائمة وأفلتها فوق خانة {mode === "bank" ? "«بيان البنك»" : "«البيان (الكاشير)»"} بالجدول جنب هاي اللوحة، أو اضغط زر "استبدال".
        </div>
      </div>

      <div className="flex-1 overflow-y-auto bg-white">
        {mode === "bank" ? (
          banks.length === 0 ? (
            <div className="flex flex-col items-center justify-center gap-2 p-10 text-center text-xs text-muted-foreground"><Search className="w-8 h-8 opacity-30"/>لا توجد حوالات تطابق البحث.</div>
          ) : banks.map(b => {
            const diff = Math.abs(b.rawAmount - cashier.matchAmount);
            const exact = diff <= 0.01;
            const sim = nameSim(cashier.name, b.description);
            const isDragged = draggedItemId === b.id;
            return (
              <div key={b.id} draggable onDragStart={() => onDragStart(b.id)} onDragEnd={onDragEnd}
                className={`group flex items-center gap-3 px-4 py-3 border-b border-border last:border-b-0 cursor-grab active:cursor-grabbing hover:bg-blue-50 transition-colors ${isDragged ? "opacity-40" : ""}`}>
                <GripVertical className="w-4 h-4 text-muted-foreground/40 shrink-0"/>
                <div className="flex-1 min-w-0">
                  <div className="whitespace-normal break-words text-sm font-medium leading-6">{b.description}</div>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    <span className="font-mono font-bold text-blue-700 text-sm">{fmtNum(b.rawAmount)}</span>
                    {exact ? <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-green-100 text-green-700 border border-green-300">نفس المبلغ</span>
                      : <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">فرق: {fmtNum(diff)}</span>}
                    <span className="text-[10px] text-muted-foreground">تشابه: {(sim*100).toFixed(0)}%</span>
                    {b.accountType && <span className="text-[10px] text-muted-foreground">| {b.accountType}</span>}
                  </div>
                </div>
                <button onClick={() => onReplaceBank(b)} className="px-2.5 py-1.5 bg-blue-600 text-white rounded-lg text-xs font-medium hover:bg-blue-700 transition-colors shrink-0">استبدال</button>
              </div>
            );
          })
        ) : (
          cashiers.length === 0 ? (
            <div className="flex flex-col items-center justify-center gap-2 p-10 text-center text-xs text-muted-foreground"><Search className="w-8 h-8 opacity-30"/>لا توجد فواتير كاشير تطابق البحث.</div>
          ) : cashiers.map(c => {
            const diff = Math.abs(c.matchAmount - oldBank.rawAmount);
            const exact = diff <= 0.01;
            const sim = nameSim(oldBank.description, c.name);
            const isDragged = draggedItemId === c.id;
            return (
              <div key={c.id} draggable onDragStart={() => onDragStart(c.id)} onDragEnd={onDragEnd}
                className={`group flex items-center gap-3 px-4 py-3 border-b border-border last:border-b-0 cursor-grab active:cursor-grabbing hover:bg-emerald-50 transition-colors ${isDragged ? "opacity-40" : ""}`}>
                <GripVertical className="w-4 h-4 text-muted-foreground/40 shrink-0"/>
                <div className="flex-1 min-w-0">
                  <div className="whitespace-normal break-words text-sm font-medium leading-6">{c.name}</div>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    <span className="font-mono font-bold text-emerald-700 text-sm">{fmtNum(c.matchAmount)}</span>
                    {exact ? <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-green-100 text-green-700 border border-green-300">نفس المبلغ</span>
                      : <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">فرق: {fmtNum(diff)}</span>}
                    <span className="text-[10px] text-muted-foreground">تشابه: {(sim*100).toFixed(0)}%</span>
                    {c.accountType && <span className="text-[10px] text-muted-foreground">| {c.accountType}</span>}
                  </div>
                </div>
                <button onClick={() => onReplaceCashier(c)} className="px-2.5 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-medium hover:bg-emerald-700 transition-colors shrink-0">استبدال</button>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
function ReportPage({ bank, cashier, results, onBack }: { bank: BankRow[]; cashier: CashierRow[]; results: MatchResult[] | null; onBack: () => void }) {
  const [visibleMatchedCount, setVisibleMatchedCount] = useState(5);
  const netTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.debit - row.credit, 0);
  const bankTotal = netTotal(bank);
  const cashierTotal = netTotal(cashier);
  const difference = bankTotal - cashierTotal;
  const reportResults = results ?? [];
  const matchedRows = reportResults.filter(row => row.type === "saved");
  const reviewRows = reportResults.filter(row => row.type === "pending");
  const unmatchedCashierRows = reportResults.filter(row => row.type === "unmatchedCashier");
  const unmatchedBankRows = reportResults.filter(row => row.type === "unmatchedBank");
  const matched = matchedRows.length;
  const debitTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.debit, 0);
  const creditTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.credit, 0);
  const matchedDetails = matchedChannelRows(reportResults);
  const financialIssues = financialIssueRows(reportResults);
  const financialExpected = financialIssues.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const financialActual = financialIssues.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  const matchedBankNet = matchedDetails.reduce((sum, row) => sum + row.bankNet, 0);
  const matchedCashierNet = matchedDetails.reduce((sum, row) => sum + row.cashierNet, 0);
  const beforeMatchDifference = difference;
  const adjustedCashierTotal = cashierTotal - matchedCashierNet + matchedBankNet;
  const afterMatchDifference = bankTotal - adjustedCashierTotal;
  const unmatchedCashierDebit = unmatchedCashierRows.reduce((sum, row) => sum + row.cashier.debit, 0);
  const unmatchedCashierCredit = unmatchedCashierRows.reduce((sum, row) => sum + row.cashier.credit, 0);
  const unmatchedBankDebit = unmatchedBankRows.reduce((sum, row) => sum + row.bank.debit, 0);
  const unmatchedBankCredit = unmatchedBankRows.reduce((sum, row) => sum + row.bank.credit, 0);
  const savedBankIds = new Set(matchedRows.map(row => row.bank.id));
  const savedCashierIds = new Set(matchedRows.map(row => row.cashier.id));
  const reviewBankIds = new Set(reviewRows.map(row => row.bank.id));
  const reviewCashierIds = new Set(reviewRows.map(row => row.cashier.id));
  const visaCashierIds = new Set(reportResults.filter(row => row.type === "visa").map(row => row.cashier.id));
  const jawwalCashierIds = new Set(reportResults.filter(row => row.type === "jawwalPay").map(row => row.cashier.id));
  const channelDefinitions = [
    { label: "بنك فلسطين", matcher: (row: BankRow | CashierRow) => row.accountType === "بنك فلسطين" },
    { label: "بال بي", matcher: (row: BankRow | CashierRow) => row.accountType === "بال بي" },
    { label: "جوال بي", matcher: (row: BankRow | CashierRow) => row.accountType === "جوال بي" },
    {
      label: "فيزا",
      matcher: (row: BankRow | CashierRow) =>
        ("id" in row && visaCashierIds.has(row.id)) ||
        ("accountType" in row && /فيزا|visa/i.test(row.accountType || "")),
    },
  ];
  const channelSummaries = channelDefinitions.map(channel => {
    const channelBank = bank.filter(channel.matcher);
    const channelCashier = cashier.filter(channel.matcher);
    const channelCashierIds = new Set(channelCashier.map(row => row.id));
    const isVisaChannel = channel.label === "فيزا";
    const confirmed = matchedRows.filter(row =>
      isVisaChannel
        ? visaCashierIds.has(row.cashier.id)
        : channelCashierIds.has(row.cashier.id) || channelBank.some(bankRow => bankRow.id === row.bank.id)
    ).length;
    const needsReview = reviewRows.filter(row =>
      isVisaChannel
        ? visaCashierIds.has(row.cashier.id)
        : channelCashierIds.has(row.cashier.id) || channelBank.some(bankRow => bankRow.id === row.bank.id)
    ).length;
    const noTransfer = unmatchedCashierRows.filter(row => channelCashierIds.has(row.cashier.id)).length;
    const noInvoice = unmatchedBankRows.filter(row => channelBank.some(bankRow => bankRow.id === row.bank.id)).length;
    return {
      ...channel,
      bank: channelBank,
      cashier: channelCashier,
      bankTotal: netTotal(channelBank),
      cashierTotal: netTotal(channelCashier),
      difference: netTotal(channelBank) - netTotal(channelCashier),
      debitBank: debitTotal(channelBank),
      creditBank: creditTotal(channelBank),
      debitCashier: debitTotal(channelCashier),
      creditCashier: creditTotal(channelCashier),
      confirmed,
      needsReview,
      noTransfer,
      noInvoice,
    };
  });
  const totals = {
    debitBank: debitTotal(bank),
    creditBank: creditTotal(bank),
    debitCashier: debitTotal(cashier),
    creditCashier: creditTotal(cashier),
  };
  const categories = [
    { label: "المطابقات المؤكدة", value: matched, amount: matchedRows.reduce((sum, row) => sum + row.bank.rawAmount, 0), color: "#fbb313" },
    { label: "مطابقات تحتاج مراجعة", value: reviewRows.length, amount: reviewRows.reduce((sum, row) => sum + row.bank.rawAmount, 0), color: "#1554ad" },
    { label: "فواتير بلا حوالات · مدين", value: unmatchedCashierRows.filter(row => row.cashier.debit > 0).length, amount: unmatchedCashierDebit, color: "#df2828" },
    { label: "فواتير بلا حوالات · دائن", value: unmatchedCashierRows.filter(row => row.cashier.credit > 0).length, amount: unmatchedCashierCredit, color: "#ef6c52" },
    { label: "حوالات بلا فواتير · مدين", value: unmatchedBankRows.filter(row => row.bank.debit > 0).length, amount: unmatchedBankDebit, color: "#9250c8" },
    { label: "حوالات بلا فواتير · دائن", value: unmatchedBankRows.filter(row => row.bank.credit > 0).length, amount: unmatchedBankCredit, color: "#b477d1" },
    { label: "فيزا", value: reportResults.filter(row => row.type === "visa").length, amount: reportResults.filter(row => row.type === "visa").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#19a8bf" },
    { label: "محفظة محمود", value: reportResults.filter(row => row.type === "mahmoudWallet").length, amount: reportResults.filter(row => row.type === "mahmoudWallet").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#f58a12" },
    { label: "جوال بي", value: reportResults.filter(row => row.type === "jawwalPay").length, amount: reportResults.filter(row => row.type === "jawwalPay").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#19a8bf" },
  ];
  const maxAmount = Math.max(...categories.map(category => category.amount), 1);
  return (
    <div dir="rtl" className="report-page min-h-[100dvh] bg-background p-4 sm:p-8">
      <div className="mx-auto max-w-6xl space-y-6">
        <header className="flex flex-wrap items-end justify-between gap-4">
          <div>
            <div className="mb-2 flex items-center gap-2 text-xs font-semibold text-primary"><BarChart3 className="h-4 w-4" /> مركز التقارير</div>
            <h1 className="text-3xl font-bold tracking-tight">تقرير التسوية</h1>
            <p className="mt-2 text-sm text-muted-foreground">صورة واحدة لما سُجّل، وما بقي يحتاج قراراً.</p>
          </div>
          <div className="flex flex-wrap items-center gap-2 report-actions">
            <button onClick={() => window.print()} className="flex items-center gap-2 rounded-xl border border-primary/30 bg-primary/5 px-4 py-2.5 text-sm font-semibold text-primary hover:bg-primary/10">
              <Download className="h-4 w-4" /> تنزيل PDF
            </button>
            <button onClick={onBack} className="flex items-center gap-2 rounded-xl border bg-card px-4 py-2.5 text-sm font-semibold hover:bg-muted"><ArrowUpRight className="h-4 w-4 rotate-180" /> العودة إلى مساحة العمل</button>
          </div>
        </header>
        {!bank.length && !cashier.length ? (
          <div className="rounded-2xl border border-dashed bg-card p-12 text-center">
            <BarChart3 className="mx-auto mb-4 h-10 w-10 text-muted-foreground/50" />
            <h2 className="font-bold">لا يوجد كشف محفوظ بعد</h2>
            <p className="mt-2 text-sm text-muted-foreground">ارفع كشفي الكاشير والبنك ثم شغّل المطابقة لعرض التقرير.</p>
            <button onClick={onBack} className="mt-6 rounded-xl bg-primary px-5 py-2.5 text-sm font-bold text-primary-foreground">بدء جلسة جديدة</button>
          </div>
        ) : (
          <>
            <div className="grid gap-4 md:grid-cols-3">
              <div className="rounded-2xl border bg-card p-5"><div className="text-xs font-semibold text-muted-foreground">إجمالي الكاشير</div><div className="mt-3 font-mono text-2xl font-bold">{fmtNum(cashierTotal)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 h-1 rounded-full bg-primary/15"><div className="h-full w-full rounded-full bg-primary" /></div></div>
              <div className="rounded-2xl border bg-card p-5"><div className="text-xs font-semibold text-muted-foreground">إجمالي الحوالات البنكية</div><div className="mt-3 font-mono text-2xl font-bold">{fmtNum(bankTotal)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 h-1 rounded-full bg-[hsl(43_96%_52%/.18)]"><div className="h-full w-[88%] rounded-full bg-[hsl(43_96%_52%)]" /></div></div>
              <div className={`rounded-2xl border p-5 ${Math.abs(difference) < .01 ? "border-emerald-200 bg-emerald-50/60" : "border-amber-200 bg-amber-50/70"}`}><div className="text-xs font-semibold text-muted-foreground">الفرق بين السجلين</div><div className="mt-3 font-mono text-2xl font-bold">{difference > 0 ? "+" : ""}{fmtNum(difference)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 text-xs font-semibold">{Math.abs(difference) < .01 ? "السجلان متوازنان" : "يحتاج إلى مراجعة الفروقات"}</div></div>
            </div>
            <section className="rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-4">
                <h2 className="font-bold">تفصيل الدائن والمدين</h2>
                <p className="mt-1 text-xs text-muted-foreground">المدين يُجمع، والدائن يُطرح منه لحساب الصافي. وتبقى قيم المدين والدائن معروضة بشكل مستقل.</p>
              </div>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                {[
                  { label: "الكاشير · مدين", value: totals.debitCashier, tone: "text-emerald-700" },
                  { label: "الكاشير · دائن", value: totals.creditCashier, tone: "text-red-600" },
                  { label: "الحوالات · مدين", value: totals.debitBank, tone: "text-blue-700" },
                  { label: "الحوالات · دائن", value: totals.creditBank, tone: "text-orange-600" },
                ].map(item => (
                  <div key={item.label} className="rounded-xl border bg-muted/20 p-4">
                    <div className="text-xs font-semibold text-muted-foreground">{item.label}</div>
                    <div className={`mt-2 font-mono text-xl font-bold ${item.tone}`}>{fmtNum(item.value)} <span className="text-xs font-normal">₪</span></div>
                  </div>
                ))}
              </div>
            </section>
            <section className="report-channel-section report-page-break-after rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-4">
                <h2 className="font-bold">تفصيل القنوات المالية</h2>
                <p className="mt-1 text-xs text-muted-foreground">لكل قناة: إجمالي الكاشير والبنك، الفرق، الدائن والمدين، وحالة المطابقات.</p>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[980px] text-xs">
                  <thead className="bg-muted/60 text-muted-foreground">
                    <tr>
                      <th className="px-3 py-3 text-right font-semibold">القناة</th>
                      <th className="px-3 py-3 text-right font-semibold">الكاشير</th>
                      <th className="px-3 py-3 text-right font-semibold">البنك</th>
                      <th className="px-3 py-3 text-right font-semibold">الفرق</th>
                      <th className="px-3 py-3 text-right font-semibold">كاشير: مدين / دائن</th>
                      <th className="px-3 py-3 text-right font-semibold">بنك: مدين / دائن</th>
                      <th className="px-3 py-3 text-right font-semibold">مؤكدة</th>
                      <th className="px-3 py-3 text-right font-semibold">تحتاج مراجعة</th>
                      <th className="px-3 py-3 text-right font-semibold">بلا حوالات</th>
                      <th className="px-3 py-3 text-right font-semibold">بلا فواتير</th>
                    </tr>
                  </thead>
                  <tbody>
                    {channelSummaries.map(channel => (
                      <tr key={channel.label} className="border-t border-border">
                        <td className="px-3 py-3 font-bold">{channel.label}</td>
                        <td className="px-3 py-3 font-mono text-emerald-700">{fmtNum(channel.cashierTotal)}</td>
                        <td className="px-3 py-3 font-mono text-blue-700">{fmtNum(channel.bankTotal)}</td>
                        <td className={`px-3 py-3 font-mono font-bold ${Math.abs(channel.difference) < .01 ? "text-emerald-700" : "text-amber-700"}`}>{fmtNum(channel.difference)}</td>
                        <td className="px-3 py-3 font-mono">{fmtNum(channel.debitCashier)} / {fmtNum(channel.creditCashier)}</td>
                        <td className="px-3 py-3 font-mono">{fmtNum(channel.debitBank)} / {fmtNum(channel.creditBank)}</td>
                        <td className="px-3 py-3 font-semibold text-emerald-700">{channel.confirmed}</td>
                        <td className="px-3 py-3 font-semibold text-amber-700">{channel.needsReview}</td>
                        <td className="px-3 py-3 font-semibold text-red-600">{channel.noTransfer}</td>
                        <td className="px-3 py-3 font-semibold text-violet-700">{channel.noInvoice}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
            <section className="rounded-2xl border border-blue-100 bg-blue-50/60 p-5 text-sm text-blue-950">
              <h2 className="font-bold">طريقة قراءة التقرير</h2>
              <div className="mt-3 grid gap-2 text-xs leading-6 text-blue-900 sm:grid-cols-2">
                <p><strong>الدائن والمدين:</strong> القيم مأخوذة من الأعمدة الأصلية بعد تطبيق خيار عكس المدين والدائن إن كان مفعّلاً؛ لا نحتاج إلى كتابة إشارة سالبة للدائن.</p>
                <p><strong>الصافي:</strong> المدين ناقص الدائن. مثال: 50 مدين و30 دائن = 20 صافي، وليس 80.</p>
                <p><strong>الفرق:</strong> صافي البنك ناقص صافي الكاشير؛ الرقم الموجب يعني أن صافي الحوالات أعلى.</p>
                <p><strong>تحتاج مراجعة:</strong> مطابقات مقترحة لم تُعتمد بعد ويمكن مراجعتها من مساعد المطابقة.</p>
                <p><strong>بلا حوالات / بلا فواتير:</strong> فواتير لا يقابلها تحويل، أو تحويلات لا تقابلها فاتورة في الكشف الآخر.</p>
              </div>
            </section>
            <div className="report-page-break-after grid gap-6 lg:grid-cols-[1.35fr_.65fr]">
              <section className="rounded-2xl border bg-card p-5 sm:p-6">
                <div className="mb-6 flex items-center justify-between"><div><h2 className="font-bold">توزيع الفروقات والنتائج</h2><p className="mt-1 text-xs text-muted-foreground">القيمة المالية لكل فئة في الجلسة الحالية</p></div><span className="rounded-full bg-muted px-3 py-1 text-xs">{categories.reduce((sum, item) => sum + item.value, 0)} بند</span></div>
                <div className="space-y-5">
                  {categories.map(category => <div key={category.label}><div className="mb-2 flex items-center justify-between gap-3 text-sm"><span className="font-semibold">{category.label}</span><span className="font-mono text-xs text-muted-foreground">{fmtNum(category.amount)} ₪ · {category.value}</span></div><div className="h-3 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full" style={{ width: `${Math.max(category.value ? 4 : 0, category.amount / maxAmount * 100)}%`, backgroundColor: category.color }} /></div></div>)}
                </div>
              </section>
              <section className="rounded-2xl border bg-[hsl(216_48%_15%)] p-6 text-white">
                <div className="text-xs font-semibold text-white/55">ملخص القرار</div><h2 className="mt-3 text-2xl font-bold leading-relaxed">كل رقم له مكان واضح.</h2>
                <p className="mt-3 text-sm leading-7 text-white/65">راجع العناصر المعلّقة أولاً، ثم صدّر ملف التسوية ليبقى أثر القرار مع الفريق.</p>
                <div className="mt-8 border-t border-white/10 pt-5"><div className="flex justify-between text-xs text-white/55"><span>نسبة المطابقة المؤكدة</span><span className="font-mono text-[hsl(43_96%_52%)]">{cashier.length ? Math.round(matched / cashier.length * 100) : 0}%</span></div><div className="mt-3 h-2 rounded-full bg-white/10"><div className="h-full rounded-full bg-[hsl(43_96%_52%)]" style={{ width: `${cashier.length ? matched / cashier.length * 100 : 0}%` }} /></div></div>
              </section>
            </div>
            <section className="report-page-three rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-8 border-b border-border pb-6">
                <div className="mb-4">
                  <h2 className="font-bold">الفواتير التي فيها مشاكل مالية</h2>
                  <p className="mt-1 text-xs text-muted-foreground">البنك هو المرجع الصحيح: نقارن مبلغ الفاتورة بالواصل الفعلي في البنك، ونبيّن الفاتورة الصحيحة أو الخاطئة.</p>
                </div>
                <div className="mb-4 grid gap-3 sm:grid-cols-3">
                  <div className="rounded-xl border bg-muted/20 p-4"><div className="text-xs font-semibold text-muted-foreground">إجمالي الفواتير للمقارنة</div><div className="mt-2 font-mono text-xl font-bold text-blue-700">{fmtNum(financialExpected)} ₪</div></div>
                  <div className="rounded-xl border bg-muted/20 p-4"><div className="text-xs font-semibold text-muted-foreground">إجمالي الواصل الفعلي بالبنك</div><div className="mt-2 font-mono text-xl font-bold text-emerald-700">{fmtNum(financialActual)} ₪</div></div>
                  <div className="rounded-xl border border-red-200 bg-red-50/60 p-4"><div className="text-xs font-semibold text-muted-foreground">فرق الفاتورة عن البنك</div><div className="mt-2 font-mono text-xl font-bold text-red-600">{fmtNum(financialExpected - financialActual)} ₪</div></div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full min-w-[1100px] text-xs">
                    <thead className="bg-red-50 text-muted-foreground"><tr>
                      <th className="px-3 py-3 text-right">القناة</th><th className="px-3 py-3 text-right">الفاتورة</th><th className="px-3 py-3 text-right">الحوالة</th><th className="px-3 py-3 text-right">مبلغ الفاتورة</th><th className="px-3 py-3 text-right">الواصل بالبنك (المرجع)</th><th className="px-3 py-3 text-right">فرق الفاتورة عن البنك</th><th className="px-3 py-3 text-right">الحالة</th><th className="px-3 py-3 text-right">الفاتورة: مدين / دائن</th><th className="px-3 py-3 text-right">البنك: مدين / دائن</th>
                    </tr></thead>
                    <tbody>{financialIssues.map(row => <tr key={`financial-top-${row.savedMatch.id}`} className="border-t border-border">
                      <td className="px-3 py-3 font-bold">{row.channel}</td><td className="px-3 py-3">{row.cashier.name}</td><td className="px-3 py-3">{row.bank.description}</td><td className="px-3 py-3 font-mono text-blue-700">{fmtNum(row.cashier.matchAmount)}</td><td className="px-3 py-3 font-mono text-emerald-700">{fmtNum(row.bank.rawAmount)}</td>                                            <td className="px-3 py-3 font-mono font-bold text-red-600">{fmtNum(row.cashier.matchAmount - row.bank.rawAmount)}</td><td className="px-3 py-3 font-semibold text-red-600">خطأ — البنك هو الصحيح</td><td className="px-3 py-3 font-mono">{fmtNum(row.cashier.debit)} / {fmtNum(row.cashier.credit)}</td><td className="px-3 py-3 font-mono">{fmtNum(row.bank.debit)} / {fmtNum(row.bank.credit)}</td>
                    </tr>)}</tbody>
                  </table>
                  {!financialIssues.length && <p className="py-5 text-center text-xs text-emerald-700">لا توجد فواتير مؤكدة فيها فرق مالي.</p>}
                </div>
              </div>
              <div className="mb-4">
                <h2 className="font-bold">الفرق قبل المطابقة وبعدها</h2>
                <p className="mt-1 text-xs text-muted-foreground">الفرق محسوب من صافي المدين ناقص الدائن، مع عزل المطابقات المؤكدة فقط.</p>
              </div>
              <div className="mb-6 grid gap-3 sm:grid-cols-2">
                <div className="rounded-xl border border-amber-200 bg-amber-50/60 p-4">
                  <div className="text-xs font-semibold text-muted-foreground">الفرق قبل المطابقة</div>
                  <div className="mt-2 font-mono text-2xl font-bold text-amber-700">{fmtNum(beforeMatchDifference)} ₪</div>
                </div>
                <div className="rounded-xl border border-emerald-200 bg-emerald-50/60 p-4">
                  <div className="text-xs font-semibold text-muted-foreground">الفرق بعد المطابقة المؤكدة</div>
                  <div className="mt-2 font-mono text-2xl font-bold text-emerald-700">{fmtNum(afterMatchDifference)} ₪</div>
                </div>
              </div>
              <h3 className="mb-3 font-bold">إجماليات المطابقات حسب القناة</h3>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[760px] text-xs">
                  <thead className="bg-muted/60 text-muted-foreground"><tr>
                    <th className="px-3 py-3 text-right">القناة</th><th className="px-3 py-3 text-right">عدد المطابقات</th>
                    <th className="px-3 py-3 text-right">البنك: مدين / دائن / صافي</th><th className="px-3 py-3 text-right">الفواتير: مدين / دائن / صافي</th><th className="px-3 py-3 text-right">فرق القناة</th>
                  </tr></thead>
                  <tbody>{Array.from(new Set(matchedDetails.map(row => row.channel))).map(channel => {
                    const rows = matchedDetails.filter(row => row.channel === channel);
                    const sum = (pick: (row: typeof rows[number]) => number) => rows.reduce((total, row) => total + pick(row), 0);
                    return <tr key={channel} className="border-t border-border">
                      <td className="px-3 py-3 font-bold">{channel}</td><td className="px-3 py-3">{rows.length}</td>
                      <td className="px-3 py-3 font-mono">{fmtNum(sum(row => row.bank.debit))} / {fmtNum(sum(row => row.bank.credit))} / {fmtNum(sum(row => row.bankNet))}</td>
                      <td className="px-3 py-3 font-mono">{fmtNum(sum(row => row.cashier.debit))} / {fmtNum(sum(row => row.cashier.credit))} / {fmtNum(sum(row => row.cashierNet))}</td>
                      <td className="px-3 py-3 font-mono font-bold">{fmtNum(sum(row => row.bankNet - row.cashierNet))}</td>
                    </tr>;
                  })}</tbody>
                </table>
              </div>
              <h3 className="mb-3 mt-6 font-bold">تفصيل كل مطابقة مؤكدة</h3>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[1100px] text-xs">
                  <thead className="bg-muted/60 text-muted-foreground"><tr>
                    <th className="px-3 py-3 text-right">القناة</th><th className="px-3 py-3 text-right">فاتورة الكاشير</th><th className="px-3 py-3 text-right">حوالة البنك</th>
                    <th className="px-3 py-3 text-right">كاشير مدين</th><th className="px-3 py-3 text-right">كاشير دائن</th><th className="px-3 py-3 text-right">بنك مدين</th><th className="px-3 py-3 text-right">بنك دائن</th><th className="px-3 py-3 text-right">فرق المطابقة</th>
                  </tr></thead>
                  <tbody>{matchedDetails.map((row, index) => <tr key={row.savedMatch.id} className={`border-t border-border ${index >= visibleMatchedCount ? "report-collapsed-row" : ""}`}>
                    <td className="px-3 py-3 font-bold">{row.channel}</td><td className="px-3 py-3">{row.cashier.name}</td><td className="px-3 py-3">{row.bank.description}</td>
                    <td className="px-3 py-3 font-mono">{fmtNum(row.cashier.debit)}</td><td className="px-3 py-3 font-mono">{fmtNum(row.cashier.credit)}</td>
                    <td className="px-3 py-3 font-mono">{fmtNum(row.bank.debit)}</td><td className="px-3 py-3 font-mono">{fmtNum(row.bank.credit)}</td><td className="px-3 py-3 font-mono font-bold">{fmtNum(row.bankNet - row.cashierNet)}</td>
                  </tr>)}</tbody>
                </table>
                {matchedDetails.length > 5 && (
                  <div className="report-match-controls mt-3 flex items-center justify-center gap-2">
                    {visibleMatchedCount < matchedDetails.length && <button onClick={() => setVisibleMatchedCount(count => Math.min(count + 5, matchedDetails.length))} className="rounded-lg border border-primary/30 bg-primary/5 px-4 py-2 text-xs font-semibold text-primary hover:bg-primary/10">عرض 5 إضافية ({Math.min(visibleMatchedCount + 5, matchedDetails.length)} من {matchedDetails.length})</button>}
                    {visibleMatchedCount > 5 && <button onClick={() => setVisibleMatchedCount(5)} className="rounded-lg border px-4 py-2 text-xs font-semibold text-muted-foreground hover:bg-muted">إظهار أول 5 فقط</button>}
                  </div>
                )}
              </div>
            </section>
          </>
        )}
      </div>
    </div>
  );
}

type PageId = "main" | "manual" | "assist" | "report" | "recon2" | "recon2StageA" | "recon2Visa";

export default function App({ initialPage = "recon2" }: { initialPage?: PageId } = {}) {
  const [page, setPage] = useState<PageId>(initialPage);
  const [mainMode, setMainMode] = useState<"default" | "stageB">("default");
  const [sessionLoaded, setSessionLoaded] = useState(false);
  const [showDeleteMenu, setShowDeleteMenu] = useState(false);
  const [sessionRestoredNotice, setSessionRestoredNotice] = useState(false);

  const [bankFile, setBankFile]   = useState<File|null>(null);
  const [bankHeaders, setBankH]   = useState<string[]>([]);
  const [bankRowsRaw, setBankRows] = useState<Record<string,unknown>[]>([]);
  const [bankMap, setBankMap]     = useState({ date:"", desc:"", debit:"", credit:"", accountType:"", ref:"" });
  const [bankSwap, setBankSwap]   = useState(false);
  const [bankFileSessionId, setBankFileSessionId] = useState(0);

  const [cashFile, setCashFile]   = useState<File|null>(null);
  const [cashHeaders, setCashH]   = useState<string[]>([]);
  const [cashRowsRaw, setCashRows] = useState<Record<string,unknown>[]>([]);
  const [cashMap, setCashMap]     = useState({ date:"", name:"", debit:"", credit:"", accountType:"", ref:"" });
  const [cashSwap, setCashSwap]   = useState(false);
  const [cashFileSessionId, setCashFileSessionId] = useState(0);
  const [stageAInvoices, setStageAInvoices] = useState<StageAResult[]>([]);
  const [walletBankFile, setWalletBankFile] = useState<File|null>(null);
  const [walletBankHeaders, setWalletBankHeaders] = useState<string[]>([]);
  const [walletBankRows, setWalletBankRows] = useState<Record<string,unknown>[]>([]);
  const [walletBankMap, setWalletBankMap] = useState({ date:"", desc:"", debit:"", credit:"", ref:"" });
  const [walletBankSwap, setWalletBankSwap] = useState(false);
  const [walletCashFile, setWalletCashFile] = useState<File|null>(null);
  const [walletCashHeaders, setWalletCashHeaders] = useState<string[]>([]);
  const [walletCashRows, setWalletCashRows] = useState<Record<string,unknown>[]>([]);
  const [walletCashMap, setWalletCashMap] = useState({ date:"", name:"", debit:"", credit:"", ref:"" });
  const [walletCashSwap, setWalletCashSwap] = useState(false);
  const [jawwalBankFile, setJawwalBankFile] = useState<File|null>(null);
  const [jawwalBankHeaders, setJawwalBankHeaders] = useState<string[]>([]);
  const [jawwalBankRows, setJawwalBankRows] = useState<Record<string,unknown>[]>([]);
  const [jawwalBankMap, setJawwalBankMap] = useState({ date:"", desc:"", debit:"", credit:"", ref:"" });
  const [jawwalBankSwap, setJawwalBankSwap] = useState(false);
  const [jawwalCashFile, setJawwalCashFile] = useState<File|null>(null);
  const [jawwalCashHeaders, setJawwalCashHeaders] = useState<string[]>([]);
  const [jawwalCashRows, setJawwalCashRows] = useState<Record<string,unknown>[]>([]);
  const [jawwalCashMap, setJawwalCashMap] = useState({ date:"", name:"", debit:"", credit:"", ref:"" });
  const [jawwalCashSwap, setJawwalCashSwap] = useState(false);

  // ─── File snapshots for restore functionality ──────────────────────────────
  interface FileSnapshot {
    label: string;
    headers: string[];
    rows: Record<string, unknown>[];
    savedAt: string;
  }
  const [bankFileSnapshots, setBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [cashFileSnapshots, setCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [walletBankFileSnapshots, setWalletBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [walletCashFileSnapshots, setWalletCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [jawwalBankFileSnapshots, setJawwalBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [jawwalCashFileSnapshots, setJawwalCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [showBankSnapshots, setShowBankSnapshots] = useState(false);
  const [showCashSnapshots, setShowCashSnapshots] = useState(false);
  const [showWalletBankSnapshots, setShowWalletBankSnapshots] = useState(false);
  const [showWalletCashSnapshots, setShowWalletCashSnapshots] = useState(false);
  const [showJawwalBankSnapshots, setShowJawwalBankSnapshots] = useState(false);
  const [showJawwalCashSnapshots, setShowJawwalCashSnapshots] = useState(false);

  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState<string|null>(null);
  const [results, setResults] = useState<MatchResult[]|null>(null);
  const [tab, setTab]         = useState<TabId>("saved");

  const [manualGroups, setManualGroups]     = useState<ManualMatchGroup[]>([]);
  const [savedMatches, setSavedMatches]     = useState<SavedMatch[]>([]);
  const [nameAliases, setNameAliases]       = useState<Record<string, string>>({});
  const [rejectedPairs, setRejected]        = useState<Set<string>>(new Set());
  const [expandedMatchKey, setExpandedMatchKey] = useState<string|null>(null);
  const [expandedUnmatchedCashier, setExpandedUnmatchedCashier] = useState<number|null>(null);
  const [selectedPendingKeys, setSelectedPendingKeys] = useState<Set<string>>(new Set());
  const [selectedSavedIds, setSelectedSavedIds] = useState<Set<string>>(new Set());

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [drawerMode, setDrawerMode] = useState<"bank" | "cashier">("bank");
  const [drawerCashier, setDrawerCashier] = useState<CashierRow | null>(null);
  const [drawerOldBank, setDrawerOldBank] = useState<BankRow | null>(null);
  const [drawerNameSearch, setDrawerNameSearch] = useState("");
  const [drawerAmountFrom, setDrawerAmountFrom] = useState("");
  const [drawerAmountTo, setDrawerAmountTo] = useState("");
  const [draggedItemId, setDraggedItemId] = useState<number | null>(null);
  const [dropTargetKey, setDropTargetKey] = useState<string | null>(null);
  const [pendingPage, setPendingPage] = useState(1);
  const PENDING_PAGE_SIZE = 20;
  const pendingRowRefs = useRef<Record<string, HTMLTableRowElement | null>>({});
  const [amountTolerancePercent, setAmountTolerancePercent] = useState(0.5);
  const [showSettings, setShowSettings] = useState(false);
  const [resultSearch, setResultSearch] = useState("");
  const [resultType, setResultType] = useState<"all" | "مدفوع" | "مستلم">("all");
  const [minAmount, setMinAmount] = useState("");
  const [maxAmount, setMaxAmount] = useState("");
  const [toast, setToast] = useState<string | null>(null);

  const [visaItems, setVisaItems] = useState<VisaItem[]>([]);

  const [heldItems, setHeldItems] = useState<HeldItem[]>([]);
  const [returnedHeldBank, setReturnedHeldBank] = useState<BankRow[]>([]);
  const [returnedHeldCashier, setReturnedHeldCashier] = useState<CashierRow[]>([]);
  const [rejectedSpecialCashierIds, setRejectedSpecialCashierIds] = useState<Set<number>>(new Set());
  const skipNextAutoReconcile = useRef(false);

  useEffect(() => {
    (async () => {
      const session = await storageGet<any>("current_session", null);
      if (session) {
        setBankH(session.bankHeaders || []);
        setBankRows(session.bankRowsRaw || []);
        setBankMap({ date:"", desc:"", debit:"", credit:"", accountType:"", ...(session.bankMap||{}) });
        setBankSwap(!!session.bankSwap);
        setBankFileSessionId(session.bankFileSessionId ?? 0);
        setCashH(session.cashHeaders || []);
        setCashRows(session.cashRowsRaw || []);
        setCashMap({ date:"", name:"", debit:"", credit:"", accountType:"", ...(session.cashMap||{}) });
        setCashSwap(!!session.cashSwap);
        setCashFileSessionId(session.cashFileSessionId ?? 0);
        setStageAInvoices(session.stageAInvoices || []);
        setWalletBankHeaders(session.walletBankHeaders || []);
        setWalletBankRows(session.walletBankRows || []);
        setWalletBankMap({ date:"", desc:"", debit:"", credit:"", ref:"", ...(session.walletBankMap || {}) });
        setWalletBankSwap(!!session.walletBankSwap);
        setWalletCashHeaders(session.walletCashHeaders || []);
        setWalletCashRows(session.walletCashRows || []);
        setWalletCashMap({ date:"", name:"", debit:"", credit:"", ref:"", ...(session.walletCashMap || {}) });
        setWalletCashSwap(!!session.walletCashSwap);
        setJawwalBankHeaders(session.jawwalBankHeaders || []);
        setJawwalBankRows(session.jawwalBankRows || []);
        setJawwalBankMap({ date:"", desc:"", debit:"", credit:"", ref:"", ...(session.jawwalBankMap || {}) });
        setJawwalBankSwap(!!session.jawwalBankSwap);
        setJawwalCashHeaders(session.jawwalCashHeaders || []);
        setJawwalCashRows(session.jawwalCashRows || []);
        setJawwalCashMap({ date:"", name:"", debit:"", credit:"", ref:"", ...(session.jawwalCashMap || {}) });
        setJawwalCashSwap(!!session.jawwalCashSwap);
        setBankFileSnapshots(session.bankFileSnapshots || []);
        setCashFileSnapshots(session.cashFileSnapshots || []);
        setWalletBankFileSnapshots(session.walletBankFileSnapshots || []);
        setWalletCashFileSnapshots(session.walletCashFileSnapshots || []);
        setJawwalBankFileSnapshots(session.jawwalBankFileSnapshots || []);
        setJawwalCashFileSnapshots(session.jawwalCashFileSnapshots || []);
        setManualGroups(session.manualGroups || []);
        setSavedMatches(session.savedMatches || []);
        setNameAliases(session.nameAliases || {});
        setRejected(new Set(session.rejectedPairs || []));
        setVisaItems(session.visaItems || []);
        setHeldItems(session.heldItems || []);
        setReturnedHeldBank(session.returnedHeldBank || []);
        setReturnedHeldCashier(session.returnedHeldCashier || []);
        setRejectedSpecialCashierIds(normalizeIdSet(session.rejectedSpecialCashierIds));
        if ((session.bankRowsRaw?.length || session.savedMatches?.length)) {
          setSessionRestoredNotice(true);
        }
      }
      setSessionLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!sessionLoaded) return;
    const t = setTimeout(() => {
      storageSet("current_session", {
        bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
        cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId, stageAInvoices,
        walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap,
        walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap,
        jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap,
        jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap,
        bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
        jawwalBankFileSnapshots, jawwalCashFileSnapshots,
        manualGroups, savedMatches, nameAliases,
        rejectedPairs: Array.from(rejectedPairs),
        visaItems, heldItems, returnedHeldBank, returnedHeldCashier,
        rejectedSpecialCashierIds: Array.from(rejectedSpecialCashierIds)
      });
    }, 500);
    return () => clearTimeout(t);
  }, [
    sessionLoaded, bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
    cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId, stageAInvoices,
    walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap,
    walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap,
    jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap,
    jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap,
    bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
    jawwalBankFileSnapshots, jawwalCashFileSnapshots,
    manualGroups, savedMatches, nameAliases, rejectedPairs,
    visaItems, heldItems, returnedHeldBank, returnedHeldCashier, rejectedSpecialCashierIds
  ]);

  const loadBank=async(f:File)=>{
    setBankFile(f);setError(null);setResults(null);
    try {
      const {headers,rows}=parseSheet(await readFileBuf(f));
      setBankH(headers);setBankRows(rows);
      setBankMap({date:autoDetect(headers,HINTS.date),desc:autoDetect(headers,HINTS.desc),debit:autoDetect(headers,HINTS.debit),credit:autoDetect(headers,HINTS.credit),accountType:autoDetect(headers,HINTS.accountType),ref:autoDetect(headers,HINTS.ref)});
      setBankFileSessionId(id => id + 1);
      setBankFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch(e){setError((e as Error).message);}
  };
  const loadCash=async(f:File)=>{
    setCashFile(f);setError(null);setResults(null);
    setRejectedSpecialCashierIds(new Set());
    try {
      const {headers,rows}=parseSheet(await readFileBuf(f));
      setCashH(headers);setCashRows(rows);
      setCashMap({date:autoDetect(headers,HINTS.date),name:autoDetect(headers,HINTS.name),debit:autoDetect(headers,HINTS.debit),credit:autoDetect(headers,HINTS.credit),accountType:autoDetect(headers,HINTS.accountType),ref:autoDetect(headers,HINTS.ref)});
      setCashFileSessionId(id => id + 1);
      setCashFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch(e){setError((e as Error).message);}
  };
  const loadWalletBank = async (f: File) => {
    setWalletBankFile(f); setError(null); setResults(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setWalletBankHeaders(headers); setWalletBankRows(rows);
      setWalletBankMap({ date:autoDetect(headers,HINTS.date), desc:autoDetect(headers,HINTS.desc), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref) });
      setWalletBankFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadWalletCash = async (f: File) => {
    setWalletCashFile(f); setError(null); setResults(null);
    setRejectedSpecialCashierIds(new Set());
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setWalletCashHeaders(headers); setWalletCashRows(rows);
      setWalletCashMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref) });
      setWalletCashFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadJawwalBank = async (f: File) => {
    setJawwalBankFile(f); setError(null); setResults(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setJawwalBankHeaders(headers); setJawwalBankRows(rows);
      setJawwalBankMap({ date:autoDetect(headers,HINTS.date), desc:autoDetect(headers,HINTS.desc), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref) });
      setJawwalBankFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadJawwalCash = async (f: File) => {
    setJawwalCashFile(f); setError(null); setResults(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setJawwalCashHeaders(headers); setJawwalCashRows(rows);
      setJawwalCashMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref) });
      setJawwalCashFileSnapshots(prev => [{ label: f.name, headers, rows, savedAt: new Date().toLocaleString("ar-SA") }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };

  const parsedBank = useMemo(():BankRow[]=>{
    return bankRowsRaw.map((r,i)=>{
      const { debit, credit, rawAmount, type } = resolveDebitCredit(r, bankMap.debit, bankMap.credit, bankSwap);
      if (rawAmount === 0) return null;
      return {
        id:i, date:fmtDate(bankMap.date?r[bankMap.date]:""),
        description:String(bankMap.desc?r[bankMap.desc]:"").trim(),
        debit, credit, rawAmount, type,
        accountType: "بنك فلسطين",
        ref: String((bankMap as any).ref ? r[(bankMap as any).ref] : "").trim(),
        orig:r
      };
    }).filter((r): r is BankRow => r !== null);
  },[bankRowsRaw,bankMap,bankSwap]);

  const parseWalletBankRows = useMemo(():BankRow[] => walletBankRows.map((r,i) => {
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, walletBankMap.debit, walletBankMap.credit, walletBankSwap);
    if (!rawAmount) return null;
    return { id: 1000000 + i, date:fmtDate(walletBankMap.date ? r[walletBankMap.date] : ""), description:String(walletBankMap.desc ? r[walletBankMap.desc] : "").trim(), debit, credit, rawAmount, type, accountType:"بال بي", ref:String(walletBankMap.ref ? r[walletBankMap.ref] : "").trim(), orig:r };
  }).filter((r): r is BankRow => r !== null), [walletBankRows, walletBankMap, walletBankSwap]);

  const parsedCashier = useMemo(():CashierRow[]=>{
    const rows = cashRowsRaw.map((r,i)=>{
      const rawName=String(cashMap.name?r[cashMap.name]:"").trim();
      const {name,notes}=parseCashierName(rawName);
      const {splitExpr,matchAmount:ma}=parseNotes(notes);
      const { debit, credit, rawAmount, type } = resolveDebitCredit(r, cashMap.debit, cashMap.credit, cashSwap);
      const amount = rawAmount;
      if (amount === 0 && !ma) return null;
      return {
        id:i, rawName, name, notes, splitExpr, debit, credit, amount,
        matchAmount: ma ?? amount, type,
        accountType: "بنك فلسطين",
        ref: String((cashMap as any).ref ? r[(cashMap as any).ref] : "").trim(),
        date:fmtDate(cashMap.date?r[cashMap.date]:""), orig:r
      };
    }).filter((r): r is CashierRow => r !== null);

    return rows;
  },[cashRowsRaw,cashMap,cashSwap]);

  const parsedWalletCashier = useMemo(():CashierRow[] => walletCashRows.map((r,i) => {
    const rawName = String(walletCashMap.name ? r[walletCashMap.name] : "").trim();
    const { name, notes } = parseCashierName(rawName);
    const { splitExpr, matchAmount: ma } = parseNotes(notes);
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, walletCashMap.debit, walletCashMap.credit, walletCashSwap);
    const amount = rawAmount || ma || 0;
    if (!amount) return null;
    return { id: 2000000 + i, rawName, name, notes, splitExpr, debit, credit, amount, matchAmount:ma ?? amount, type, accountType:"بال بي", ref:String(walletCashMap.ref ? r[walletCashMap.ref] : "").trim(), date:fmtDate(walletCashMap.date ? r[walletCashMap.date] : ""), orig:r };
  }).filter((r): r is CashierRow => r !== null), [walletCashRows, walletCashMap, walletCashSwap]);

  const parseJawwalBankRows = useMemo(():BankRow[] => jawwalBankRows.map((r,i) => {
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, jawwalBankMap.debit, jawwalBankMap.credit, jawwalBankSwap);
    if (!rawAmount) return null;
    return { id: 4000000 + i, date:fmtDate(jawwalBankMap.date ? r[jawwalBankMap.date] : ""), description:String(jawwalBankMap.desc ? r[jawwalBankMap.desc] : "").trim(), debit, credit, rawAmount, type, accountType:"جوال بي", ref:String(jawwalBankMap.ref ? r[jawwalBankMap.ref] : "").trim(), orig:r };
  }).filter((r): r is BankRow => r !== null), [jawwalBankRows, jawwalBankMap, jawwalBankSwap]);

  const parsedJawwalCashier = useMemo(():CashierRow[] => jawwalCashRows.map((r,i) => {
    const rawName = String(jawwalCashMap.name ? r[jawwalCashMap.name] : "").trim();
    const { name, notes } = parseCashierName(rawName);
    const { splitExpr, matchAmount: ma } = parseNotes(notes);
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, jawwalCashMap.debit, jawwalCashMap.credit, jawwalCashSwap);
    const amount = rawAmount || ma || 0;
    if (!amount) return null;
    return { id: 5000000 + i, rawName, name, notes, splitExpr, debit, credit, amount, matchAmount:ma ?? amount, type, accountType:"جوال بي", ref:String(jawwalCashMap.ref ? r[jawwalCashMap.ref] : "").trim(), date:fmtDate(jawwalCashMap.date ? r[jawwalCashMap.date] : ""), orig:r };
  }).filter((r): r is CashierRow => r !== null), [jawwalCashRows, jawwalCashMap, jawwalCashSwap]);

  const heldBankKeys = useMemo(() => new Set(heldItems.filter(h=>h.kind==="bank").map(h => `${h.fileSessionId}:${h.refId}`)), [heldItems]);
  const heldCashierKeys = useMemo(() => new Set(heldItems.filter(h=>h.kind==="cashier").map(h => `${h.fileSessionId}:${h.refId}`)), [heldItems]);

  const activeBank = useMemo(
    () => [...parsedBank.filter(b => !heldBankKeys.has(`${bankFileSessionId}:${b.id}`)), ...parseWalletBankRows, ...parseJawwalBankRows, ...returnedHeldBank],
    [parsedBank, parseWalletBankRows, parseJawwalBankRows, heldBankKeys, bankFileSessionId, returnedHeldBank]
  );
  const activeCashier = useMemo(
    () => [...stageAInvoices.map((invoice, index) => ({
      id: 3000000 + index, rawName: invoice.originalName, name: invoice.name, notes: invoice.issue, splitExpr: "",
      debit: invoice.amount, credit: 0, amount: invoice.amount, matchAmount: invoice.amount,
      type: "مدفوع" as const, accountType: invoice.source === "بال بي" ? "بال بي" : "بنك فلسطين",
      ref: invoice.invoiceNumber, date: invoice.registrationTime, orig: invoice.invoice
    })), ...parsedCashier.filter(c => !heldCashierKeys.has(`${cashFileSessionId}:${c.id}`)), ...parsedWalletCashier, ...parsedJawwalCashier, ...returnedHeldCashier],
    [stageAInvoices, parsedCashier, parsedWalletCashier, parsedJawwalCashier, heldCashierKeys, cashFileSessionId, returnedHeldCashier]
  );

  const savedKeys = useMemo(() => new Set(savedMatches.map(s => `${s.cashierId}-${s.bankId}`)), [savedMatches]);
  const savedCashierIds = useMemo(() => new Set(savedMatches.map(s => s.cashierId)), [savedMatches]);
  const savedBankIds = useMemo(() => new Set(savedMatches.map(s => s.bankId)), [savedMatches]);
  const pendingCashierIds = useMemo(
    () => new Set((results?.filter(r => r.type === "pending") as any[] ?? []).map((r:any) => r.cashier.id as number)),
    [results]
  );
  const pendingBankIds = useMemo(
    () => new Set((results?.filter(r => r.type === "pending") as any[] ?? []).map((r:any) => r.bank.id as number)),
    [results]
  );
  const claimedNames = useMemo(() => {
    const s = new Set<string>();
    savedMatches.forEach(sm => {
      const cn = normName(sm.cashierName);
      const bn = normName(sm.bankDesc);
      if (cn) s.add(cn);
      if (bn) s.add(bn);
    });
    return s;
  }, [savedMatches]);

  const visaCashierIds = useMemo(() => new Set(visaItems.map(v => v.cashierId)), [visaItems]);
  const jawwalPayCashierIds = useMemo(() => new Set(
    activeCashier.filter(c => !rejectedSpecialCashierIds.has(c.id) && /جوال\s*بي|jawwal\s*pay/i.test(`${c.rawName} ${c.name}`)).map(c => c.id)
  ), [activeCashier, rejectedSpecialCashierIds]);
  const mahmoudWalletCashierIds = useMemo(() => new Set(
    activeCashier.filter(c => !rejectedSpecialCashierIds.has(c.id) && /محفظة\s+محمود/i.test(`${c.rawName} ${c.name}`)).map(c => c.id)
  ), [activeCashier, rejectedSpecialCashierIds]);
  const specialCashierIds = useMemo(() => new Set([...jawwalPayCashierIds, ...mahmoudWalletCashierIds]), [jawwalPayCashierIds, mahmoudWalletCashierIds]);

  const savedRows = useMemo(
    () => (results?.filter(r => r.type === "saved") as any[] ?? []).sort((a,b) => b.bank.rawAmount - a.bank.rawAmount),
    [results]
  );
  const pendingRows = useMemo(
    () => (results?.filter(r => r.type === "pending") as any[] ?? []).sort((a,b) => b.bank.rawAmount - a.bank.rawAmount),
    [results]
  );
  const visaRows = useMemo(() => results?.filter(r => r.type === "visa") as any[] ?? [], [results]);
  const uCashierRows = useMemo(
    () => (results?.filter(r => r.type === "unmatchedCashier") as any[] ?? []).sort((a,b) => b.cashier.amount - a.cashier.amount),
    [results]
  );
  const uBankRows = useMemo(
    () => (results?.filter(r => r.type === "unmatchedBank") as any[] ?? []).sort((a,b) => b.bank.rawAmount - a.bank.rawAmount),
    [results]
  );

  const unmatchedBankForSuggestions = useMemo(() =>
    results?.filter(r => r.type === "unmatchedBank").map((r: any) => r.bank as BankRow) ?? []
  , [results]);
  const suggestionOwnerMap = useMemo(() => {
    const map = new Map<number, {cashierId:number; score:number}>();
    uCashierRows.forEach((r:any) => {
      const cashierRow = r.cashier as CashierRow;
      unmatchedBankForSuggestions.forEach(b => {
        if (cashierRow.type !== b.type) return;
        const pairKey = `${cashierRow.id}-${b.id}`;
        if (rejectedPairs.has(pairKey) || savedKeys.has(pairKey)) return;
        const sc = scoreCandidate(cashierRow, b, claimedNames, nameAliases);
        if (!sc) return;
        const cur = map.get(b.id);
        if (!cur || sc.score > cur.score) map.set(b.id, { cashierId: cashierRow.id, score: sc.score });
      });
    });
    return map;
  }, [uCashierRows, unmatchedBankForSuggestions, rejectedPairs, savedKeys, claimedNames, nameAliases]);

  const rerun = useCallback(() => {
    if (!activeBank.length || !activeCashier.length) return;
    const res = reconcile(activeBank, activeCashier, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases);
    setResults(res);
  }, [activeBank, activeCashier, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases]);

  const run = async() => {
    setLoading(true); setError(null);
    try {
      const res = reconcile(activeBank, activeCashier, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases);
      setResults(res);
      setTab("saved");
      setToast(`تمت مطابقة ${res.filter(r => r.type === "pending" || r.type === "saved").length} سجل`);
    } catch(e){setError((e as Error).message);}
    finally{setLoading(false);}
  };

  const handleMoveToVisa = (cashierRow: CashierRow, source?: "pending" | "unmatched" | "saved" | "manual") => {
    const visaNum = extractVisaNumber(cashierRow.rawName);
    if (!visaNum) {
      setToast("هذه الفاتورة لا تحتوي على رقم فيزا من 4 أرقام بعد / أو -.");
      return;
    }
    if (visaItems.some(v => v.cashierId === cashierRow.id)) {
      setToast("هذه الفاتورة موجودة بالفعل في قائمة الفيزا.");
      return;
    }
    const cleanName = extractVisaName(cashierRow.rawName);
    const item: VisaItem = {
      id: `visa-${Date.now()}-${cashierRow.id}`, cashierId: cashierRow.id,
      rawName: cashierRow.rawName, name: cleanName, visaNumber: visaNum,
      amount: cashierRow.amount, type: cashierRow.type, date: cashierRow.date,
      movedAt: new Date().toLocaleString("ar-SA"), source
    };
    setVisaItems(prev => [...prev, item]);
    if (source === "saved") {
      setSavedMatches(prev => prev.filter(s => s.cashierId !== cashierRow.id));
    }
    setExpandedMatchKey(null);
    setExpandedUnmatchedCashier(null);
  };

  const handleRemoveVisaItem = (id: string) => {
    setVisaItems(prev => prev.filter(v => v.id !== id));
  };

  const handleRejectSpecialCashier = (cashierRow: CashierRow) => {
    setRejectedSpecialCashierIds(prev => {
      const next = new Set(prev);
      next.add(cashierRow.id);
      return next;
    });
    setToast("تم ��فض التصنيف وإرجاع الفاتورة إلى المطابقة.");
  };

  const handleSaveMatch = (
    cashierRow: CashierRow,
    bankRow: BankRow,
    note?: string,
    options?: { force?: boolean }
  ): SavedMatch | null => {
    const pairKey = `${cashierRow.id}-${bankRow.id}`;
    if (savedKeys.has(pairKey) && !options?.force) return null;
    skipNextAutoReconcile.current = true;

    const isAmountDiff = Math.abs(bankRow.rawAmount - cashierRow.amount) > 0.01;
    const ms = advancedMatchCheck(cashierRow.name, bankRow.description);
    const isNameDiff = ms.isApprox || ms.matchType === "none";
    const isAccountTypeDiff = accountTypesDiffer(bankRow.accountType, cashierRow.accountType);

    const autoNote = note || [
      isAmountDiff ? `اختلاف مبلغ: ${fmtNum(Math.abs(bankRow.rawAmount - cashierRow.amount))}` : "",
      isNameDiff ? `اختلاف اسم: ${cashierRow.name} ↔ ${bankRow.description}` : "",
      isAccountTypeDiff ? `نوع الحساب مختلف: بنك (${bankRow.accountType || "—"}) ↔ كاشير (${cashierRow.accountType || "—"})` : ""
    ].filter(Boolean).join(" | ") || undefined;

    const newSaved: SavedMatch = {
      id: `saved-${Date.now()}-${cashierRow.id}-${bankRow.id}`,
      cashierId: cashierRow.id, bankId: bankRow.id,
      cashierName: cashierRow.name, bankDesc: bankRow.description,
      amount: bankRow.rawAmount, type: cashierRow.type,
      date: new Date().toLocaleDateString("ar-SA"),
      savedAt: new Date().toLocaleString("ar-SA"),
      note: autoNote, isAmountDiff, isNameDiff, isManual: false, isAccountTypeDiff,
      matchScore: Math.round(nameSim(cashierRow.name, bankRow.description) * 100),
      editorNotes: note,
      bankAccountType: bankRow.accountType, cashierAccountType: cashierRow.accountType
    };

    setSavedMatches(prev => [...prev, newSaved]);
    setResults(prev => {
      if (!prev) return prev;
      const savedResult: MatchResult = { type: "saved", cashier: cashierRow, bank: bankRow, savedMatch: newSaved };
      const next = prev.filter(result => {
        const resultCashierId = "cashier" in result ? result.cashier.id : null;
        const resultBankId = "bank" in result ? result.bank.id : null;
        return resultCashierId !== cashierRow.id && resultBankId !== bankRow.id;
      });
      return [savedResult, ...next];
    });
    if (isNameDiff) {
      setNameAliases(prev => ({ ...prev, [smartNameKey(cashierRow.name)]: smartNameKey(smartAliasTarget(cashierRow.name, bankRow.description)) }));
    }
    const nextRejected = new Set(rejectedPairs);
    nextRejected.delete(pairKey);
    setRejected(nextRejected);
    setExpandedMatchKey(null);
    return newSaved;
  };

  const finishReplacement = (
    oldCashier: CashierRow,
    oldBank: BankRow,
    newCashier: CashierRow,
    newBank: BankRow,
    newSaved: SavedMatch
  ) => {
    const oldPairKey = `${oldCashier.id}-${oldBank.id}`;
    const newPairKey = `${newCashier.id}-${newBank.id}`;

    // A selected row owns both sides of the match. Remove any previous
    // saved match using either the selected bank or the selected cashier.
    setSavedMatches(prev => prev.filter(saved =>
      saved.id === newSaved.id ||
      (saved.cashierId !== newCashier.id && saved.bankId !== newBank.id)
    ));

    setRejected(prev => {
      const next = new Set(prev);
      next.add(oldPairKey);
      next.delete(newPairKey);
      return next;
    });

    // Replace the old row immediately and remove the same bank/cashier from
    // every other visible result, including another pending suggestion.
    setResults(prev => {
      if (!prev) return prev;
      const replacement: MatchResult = {
        type: "saved",
        cashier: newCashier,
        bank: newBank,
        savedMatch: newSaved
      };
      let inserted = false;
      const next: MatchResult[] = [];

      for (const result of prev) {
        const resultCashierId = "cashier" in result ? result.cashier.id : null;
        const resultBankId = "bank" in result ? result.bank.id : null;
        const isOldRow = resultCashierId === oldCashier.id && resultBankId === oldBank.id;
        const usesClaimedCashier = resultCashierId === newCashier.id;
        const usesClaimedBank = resultBankId === newBank.id;

        if (usesClaimedCashier || usesClaimedBank) {
          if (!inserted && isOldRow) {
            next.push(replacement);
            inserted = true;
          }
          continue;
        }
        next.push(result);
      }

      if (!inserted) next.push(replacement);
      return next;
    });
    // Replacement changes ownership of rows, so let the normal reconciliation
    // pass rebuild the remaining suggestions immediately.
    skipNextAutoReconcile.current = false;
  };

  const handleSaveMatchesBulk = (pairs: Array<{cashier: CashierRow; bank: BankRow}>) => {
    if (!pairs.length) return;
    const newSaved: SavedMatch[] = [];
    const usedKeysNow = new Set(savedKeys);
    pairs.forEach(({cashier: cashierRow, bank: bankRow}) => {
      const pairKey = `${cashierRow.id}-${bankRow.id}`;
      if (usedKeysNow.has(pairKey)) return;
      usedKeysNow.add(pairKey);

      const isAmountDiff = Math.abs(bankRow.rawAmount - cashierRow.amount) > 0.01;
      const ms = advancedMatchCheck(cashierRow.name, bankRow.description);
      const isNameDiff = ms.isApprox || ms.matchType === "none";
      const isAccountTypeDiff = accountTypesDiffer(bankRow.accountType, cashierRow.accountType);
      const autoNote = [
        isAmountDiff ? `اختلاف مبلغ: ${fmtNum(Math.abs(bankRow.rawAmount - cashierRow.amount))}` : "",
        isNameDiff ? `اختلاف اسم: ${cashierRow.name} ↔ ${bankRow.description}` : "",
        isAccountTypeDiff ? `نوع الحساب مختلف: بنك (${bankRow.accountType || "—"}) ↔ كاشير (${cashierRow.accountType || "—"})` : ""
      ].filter(Boolean).join(" | ") || undefined;

      newSaved.push({
        id: `saved-${Date.now()}-${cashierRow.id}-${bankRow.id}`,
        cashierId: cashierRow.id, bankId: bankRow.id,
        cashierName: cashierRow.name, bankDesc: bankRow.description,
        amount: bankRow.rawAmount, type: cashierRow.type,
        date: new Date().toLocaleDateString("ar-SA"),
        savedAt: new Date().toLocaleString("ar-SA"),
        note: autoNote, isAmountDiff, isNameDiff, isManual: false, isAccountTypeDiff,
        matchScore: Math.round(nameSim(cashierRow.name, bankRow.description) * 100),
        editorNotes: undefined,
        bankAccountType: bankRow.accountType, cashierAccountType: cashierRow.accountType
      });
    });
    if (!newSaved.length) return;
    skipNextAutoReconcile.current = true;
    setSavedMatches(prev => [...prev, ...newSaved]);
    const groupCashierIds = new Set(pairs.map(({ cashier }) => cashier.id));
    const groupBankIds = new Set(pairs.map(({ bank }) => bank.id));
    setResults(prev => {
      if (!prev) return prev;
      const savedResults: MatchResult[] = newSaved.flatMap(saved => {
        const pair = pairs.find(({ cashier, bank }) =>
          cashier.id === saved.cashierId && bank.id === saved.bankId
        );
        return pair ? [{ type: "saved" as const, cashier: pair.cashier, bank: pair.bank, savedMatch: saved }] : [];
      });
      const next = prev.filter(result => {
        const resultCashierId = "cashier" in result ? result.cashier.id : null;
        const resultBankId = "bank" in result ? result.bank.id : null;
        return !groupCashierIds.has(resultCashierId ?? -1) && !groupBankIds.has(resultBankId ?? -1);
      });
      return [...savedResults, ...next];
    });
    const aliasUpdates: Record<string, string> = {};
    pairs.forEach(({ cashier, bank }) => {
      const ms = advancedMatchCheck(cashier.name, bank.description);
      if (ms.isApprox || ms.matchType === "none") {
        aliasUpdates[smartNameKey(cashier.name)] = smartNameKey(smartAliasTarget(cashier.name, bank.description));
      }
    });
    if (Object.keys(aliasUpdates).length) setNameAliases(prev => ({ ...prev, ...aliasUpdates }));
    setToast(`تم تأكيد ${newSaved.length} مطابقة`);
    setSelectedPendingKeys(new Set());
  };

  const handleUnsaveMatch = (savedMatch: SavedMatch) => {
    if (!window.confirm("هل أنت متأكد من إلغاء حفظ هذه المطابقة؟")) return;
    if (savedMatch.sourceGroupId) {
      setSavedMatches(prev => prev.filter(s => s.sourceGroupId !== savedMatch.sourceGroupId));
      setManualGroups(prev => prev.filter(g => g.id !== savedMatch.sourceGroupId));
    } else {
      setSavedMatches(prev => prev.filter(s => s.id !== savedMatch.id));
    }
    setSelectedSavedIds(prev => {
      const next = new Set(prev);
      next.delete(savedMatch.id);
      return next;
    });
  };

  const handleUnsaveSelected = () => {
    const selected = new Set(selectedSavedIds);
    if (!selected.size) return;
    const selectedMatches = savedMatches.filter(match => selected.has(match.id));
    if (!window.confirm(`إلغاء تأكيد ${selectedMatches.length} مطابقة محددة؟`)) return;
    const groupIds = new Set(selectedMatches.flatMap(match => match.sourceGroupId ? [match.sourceGroupId] : []));
    setSavedMatches(prev => prev.filter(match =>
      !selected.has(match.id) && (!match.sourceGroupId || !groupIds.has(match.sourceGroupId))
    ));
    if (groupIds.size) {
      setManualGroups(prev => prev.filter(group => !groupIds.has(group.id)));
    }
    setSelectedSavedIds(new Set());
    setExpandedMatchKey(null);
    setToast(`تم إلغاء تأكيد ${selectedMatches.length} مطابقة`);
  };

  const handleRejectMatch = (cashierRow: CashierRow, bankRow: BankRow) => {
    const pairKey = `${cashierRow.id}-${bankRow.id}`;
    const next = new Set(rejectedPairs);
    next.add(pairKey);
    setRejected(next);
    setExpandedMatchKey(null);
    setSavedMatches(prev => prev.filter(s => `${s.cashierId}-${s.bankId}` !== pairKey));
    setSelectedPendingKeys(prev => { const n = new Set(prev); n.delete(pairKey); return n; });
  };

  const handleReplaceBank = (cashierRow: CashierRow, oldBank: BankRow, newBank: BankRow) => {
    const newSaved = handleSaveMatch(cashierRow, newBank, undefined, { force: true });
    if (!newSaved) {
      setToast("هذه الحوالة مرتبطة مسبقاً ولا يمكن استخدامها للاستبدال.");
      return;
    }
    finishReplacement(cashierRow, oldBank, cashierRow, newBank, newSaved);
    closeDrawer();
  };

  const drawerBanks = useMemo(() => {
    if (drawerMode !== "bank" || !drawerCashier) return [];
    return activeBank
      .filter(b => {
        if (b.type !== drawerCashier.type) return false;
        const pairKey = `${drawerCashier.id}-${b.id}`;
        if (savedBankIds.has(b.id) && !savedKeys.has(pairKey)) return false;
        if (b.id === drawerOldBank?.id) return false;
        const nameMatch = !drawerNameSearch.trim() ||
          b.description.toLowerCase().includes(drawerNameSearch.trim().toLowerCase());
        const amt = b.rawAmount;
        const fromOk = !drawerAmountFrom || amt >= Number(drawerAmountFrom);
        const toOk = !drawerAmountTo || amt <= Number(drawerAmountTo);
        return nameMatch && fromOk && toOk;
      })
      .sort((a, b) => {
        const diffA = Math.abs(a.rawAmount - drawerCashier.matchAmount);
        const diffB = Math.abs(b.rawAmount - drawerCashier.matchAmount);
        if (diffA !== diffB) return diffA - diffB;
        return nameSim(drawerCashier.name, b.description) - nameSim(drawerCashier.name, a.description);
      })
      .slice(0, 50);
  }, [drawerMode, drawerCashier, drawerNameSearch, drawerAmountFrom, drawerAmountTo, activeBank, savedBankIds, savedKeys, drawerOldBank]);

  const drawerCashiers = useMemo(() => {
    if (drawerMode !== "cashier" || !drawerOldBank) return [];
    return activeCashier
      .filter(c => {
        if (c.type !== drawerOldBank.type) return false;
        if (visaCashierIds.has(c.id)) return false;
        const pairKey = `${c.id}-${drawerOldBank.id}`;
        if (savedCashierIds.has(c.id) && !savedKeys.has(pairKey)) return false;
        if (c.id === drawerCashier?.id) return false;
        const nameMatch = !drawerNameSearch.trim() ||
          c.name.toLowerCase().includes(drawerNameSearch.trim().toLowerCase());
        const amt = c.matchAmount;
        const fromOk = !drawerAmountFrom || amt >= Number(drawerAmountFrom);
        const toOk = !drawerAmountTo || amt <= Number(drawerAmountTo);
        return nameMatch && fromOk && toOk;
      })
      .sort((a, b) => {
        const diffA = Math.abs(a.matchAmount - drawerOldBank.rawAmount);
        const diffB = Math.abs(b.matchAmount - drawerOldBank.rawAmount);
        if (diffA !== diffB) return diffA - diffB;
        return nameSim(drawerOldBank.description, b.name) - nameSim(drawerOldBank.description, a.name);
      })
      .slice(0, 50);
  }, [drawerMode, drawerOldBank, drawerCashier, drawerNameSearch, drawerAmountFrom, drawerAmountTo, activeCashier, savedCashierIds, savedKeys, visaCashierIds]);

  const openDrawer = (cashier: CashierRow, oldBank: BankRow, mode: "bank" | "cashier" = "bank") => {
    setDrawerMode(mode);
    setDrawerCashier(cashier);
    setDrawerOldBank(oldBank);
    setDrawerNameSearch("");
    setDrawerAmountFrom("");
    setDrawerAmountTo("");
    setDrawerOpen(true);
    const pairKey = `${cashier.id}-${oldBank.id}`;
    window.setTimeout(() => {
      pendingRowRefs.current[pairKey]?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
    }, 50);
  };

  const closeDrawer = () => {
    setDrawerOpen(false);
    setDrawerCashier(null);
    setDrawerOldBank(null);
    setDrawerNameSearch("");
    setDrawerAmountFrom("");
    setDrawerAmountTo("");
    setDraggedItemId(null);
    setDropTargetKey(null);
  };

  useEffect(() => {
    if (!drawerOpen) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") closeDrawer(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [drawerOpen]);

  const handleDropBank = (cashier: CashierRow, oldBank: BankRow, newBankId: number) => {
    const newBank = activeBank.find(b => b.id === newBankId);
    if (!newBank) return;
    handleReplaceBank(cashier, oldBank, newBank);
  };

  const handleReplaceCashier = (bankRow: BankRow, oldCashier: CashierRow, newCashier: CashierRow) => {
    const newSaved = handleSaveMatch(newCashier, bankRow, undefined, { force: true });
    if (!newSaved) {
      setToast("هذه الفاتورة مرتبطة مسبقاً ولا يمكن استخدامها للاستبدال.");
      return;
    }
    finishReplacement(oldCashier, bankRow, newCashier, bankRow, newSaved);
    closeDrawer();
  };

  const handleDropCashier = (bankRow: BankRow, oldCashier: CashierRow, newCashierId: number) => {
    const newCashier = activeCashier.find(c => c.id === newCashierId);
    if (!newCashier) return;
    handleReplaceCashier(bankRow, oldCashier, newCashier);
  };

  const handleAcceptSuggestion = (cashierRow: CashierRow, bankRow: BankRow) => {
    handleSaveMatch(cashierRow, bankRow);
    setExpandedUnmatchedCashier(null);
  };

  const handleRejectSuggestion = (cashierRow: CashierRow, bankRow: BankRow) => {
    const pairKey = `${cashierRow.id}-${bankRow.id}`;
    const next = new Set(rejectedPairs);
    next.add(pairKey);
    setRejected(next);
    setExpandedUnmatchedCashier(cashierRow.id);
  };

  const handleAddGroup = (g: ManualMatchGroup) => {
    const newSaved: SavedMatch[] = [];
    g.banks.forEach(b => {
      g.cashiers.forEach(c => {
        const pairKey = `${c.id}-${b.id}`;
        if (savedKeys.has(pairKey)) return;
        newSaved.push({
          id: `manual-${g.id}-${c.id}-${b.id}`,
          cashierId: c.id, bankId: b.id,
          cashierName: c.name, bankDesc: b.description,
          amount: b.rawAmount, type: c.type,
          date: new Date().toLocaleDateString("ar-SA"),
          savedAt: new Date().toLocaleString("ar-SA"),
          note: g.note || "مطابقة يدوية",
          isAmountDiff: Math.abs(b.rawAmount - c.amount) > 0.01,
          isNameDiff: true, isManual: true,
          isAccountTypeDiff: accountTypesDiffer(b.accountType, c.accountType),
          bankAccountType: b.accountType, cashierAccountType: c.accountType,
          sourceGroupId: g.id,
          matchScore: Math.round(nameSim(c.name, b.description) * 100)
        });
      });
    });
    setSavedMatches(prev => [...prev, ...newSaved]);
    const groupPairKeys = new Set(newSaved.map(s => `${s.cashierId}-${s.bankId}`));
    const groupCashierIds = new Set(newSaved.map(s => s.cashierId));
    const groupBankIds = new Set(newSaved.map(s => s.bankId));
    setResults(prev => prev
  ? prev.map(result => {
      if (result.type !== "pending") return result;
      const saved = newSaved.find(item => item.cashierId === result.cashier.id && item.bankId === result.bank.id);
      return saved ? { type: "saved", cashier: result.cashier, bank: result.bank, savedMatch: saved } : result;
    })
  : prev);
    setManualGroups(prev => [...prev, g]);
    setSelectedPendingKeys(prev => { const n = new Set(prev); groupPairKeys.forEach(k => n.delete(k)); return n; });
  };

  const handleRemoveGroup = (id: string) => {
    setSavedMatches(prev => prev.filter(s => s.sourceGroupId !== id));
    setManualGroups(prev => prev.filter(g => g.id !== id));
  };

  const handleHoldCashier = (c: CashierRow, note?: string) => {
    if (heldItems.some(h => h.kind === "cashier" && h.refId === c.id && h.fileSessionId === cashFileSessionId)) {
      setToast("هذه الفاتورة موجودة بالفعل في المعلقات.");
      return;
    }
    if ((c as any)._fromHeld) {
      setReturnedHeldCashier(prev => prev.filter(x => x.id !== c.id));
    }
    const item: HeldItem = {
      id: `held-c-${Date.now()}-${c.id}`, kind: "cashier", refId: c.id,
      fileSessionId: cashFileSessionId, data: c,
      heldAt: new Date().toLocaleString("ar-SA"), note
    };
    setHeldItems(prev => [...prev, item]);
    setExpandedMatchKey(null);
    setExpandedUnmatchedCashier(null);
  };
  const handleHoldBank = (b: BankRow, note?: string) => {
    if (heldItems.some(h => h.kind === "bank" && h.refId === b.id && h.fileSessionId === bankFileSessionId)) {
      setToast("هذه الحوالة موجودة بالفعل في المعلقات.");
      return;
    }
    if ((b as any)._fromHeld) {
      setReturnedHeldBank(prev => prev.filter(x => x.id !== b.id));
    }
    const item: HeldItem = {
      id: `held-b-${Date.now()}-${b.id}`, kind: "bank", refId: b.id,
      fileSessionId: bankFileSessionId, data: b,
      heldAt: new Date().toLocaleString("ar-SA"), note
    };
    setHeldItems(prev => [...prev, item]);
    setExpandedMatchKey(null);
  };
  const handleUnhold = (h: HeldItem) => {
    const uniqueId = nextHeldRowId();
    if (h.kind === "bank") {
      const row: BankRow = { ...(h.data as BankRow), id: uniqueId, _fromHeld: true };
      setReturnedHeldBank(prev => [...prev, row]);
    } else {
      const row: CashierRow = { ...(h.data as CashierRow), id: uniqueId, _fromHeld: true };
      setReturnedHeldCashier(prev => [...prev, row]);
    }
    setHeldItems(prev => prev.filter(x => x.id !== h.id));
  };
  const handleDeleteHeld = (id: string) => {
    if (!window.confirm("حذف هذا العنصر نهائياً من المعلقات؟")) return;
    setHeldItems(prev => prev.filter(h => h.id !== id));
  };

  useEffect(() => {
    if (skipNextAutoReconcile.current) {
      skipNextAutoReconcile.current = false;
      return;
    }
    if (activeBank.length && activeCashier.length) {
      rerun();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [manualGroups, savedMatches, rejectedPairs, visaItems, jawwalPayCashierIds, mahmoudWalletCashierIds, rejectedSpecialCashierIds, heldItems, activeBank, activeCashier, amountTolerancePercent]);

  const stats = useMemo(() => {
    if (!results) return null;
    return {
      saved: results.filter(r => r.type === "saved").length,
      pending: results.filter(r => r.type === "pending").length,
      manual: results.filter((r:any) => r.type === "saved" && (r.savedMatch as SavedMatch).isManual).length,
      visa: visaItems.length,
      jawwalPay: results.filter(r => r.type === "jawwalPay").length,
      mahmoudWallet: results.filter(r => r.type === "mahmoudWallet").length,
      held: heldItems.length,
      uCashier: results.filter(r => r.type === "unmatchedCashier").length,
      uBank: results.filter(r => r.type === "unmatchedBank").length,
    };
  }, [results, visaItems, heldItems]);

  const canRun = bankRowsRaw.length > 0 && cashRowsRaw.length > 0 && bankMap.desc && cashMap.name && (cashMap.debit || cashMap.credit);

  const filterResult = (row: any) => {
    const cashier = row.cashier as CashierRow | undefined;
    const bank = row.bank as BankRow | undefined;
    const text = `${cashier?.name || ""} ${bank?.description || ""}`.toLowerCase();
    const amount = bank?.rawAmount ?? cashier?.amount ?? 0;
    return (!resultSearch.trim() || text.includes(resultSearch.trim().toLowerCase())) &&
      (resultType === "all" || cashier?.type === resultType || bank?.type === resultType) &&
      (!minAmount || amount >= Number(minAmount)) && (!maxAmount || amount <= Number(maxAmount));
  };
  const visibleSavedRows = savedRows.filter(filterResult);
  const visiblePendingRows = pendingRows.filter(filterResult);
  const pagedPendingRows = useMemo(() => {
    const totalPages = Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE));
    const page = Math.min(pendingPage, totalPages);
    const start = (page - 1) * PENDING_PAGE_SIZE;
    return visiblePendingRows.slice(start, start + PENDING_PAGE_SIZE);
  }, [visiblePendingRows, pendingPage]);
  const visibleUCashierRows = uCashierRows.filter(filterResult);
  const visibleUBankRows = uBankRows.filter(filterResult);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(null), 2800);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => { setPendingPage(1); }, [resultSearch, resultType, minAmount, maxAmount]);

  const saveProject = async (name: string) => {
    const projects = await loadProjectsList();
    const project: SavedProject = {
      id: `proj-${Date.now()}`, name,
      savedAt: new Date().toLocaleString("ar-SA"),
      bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
      cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId,
      manualGroups, savedMatches,
      rejectedPairs: Array.from(rejectedPairs),
      visaItems, heldItems, returnedHeldBank, returnedHeldCashier
    };
    await persistProjectsList([...projects, project]);
    setToast(`تم حفظ المشروع "${name}" ✓`);
  };

  const loadProject = (p: SavedProject) => {
    setBankH(p.bankHeaders); setBankRows(p.bankRowsRaw); setBankMap({ date: p.bankMap?.date ?? "", desc: p.bankMap?.desc ?? "", debit: p.bankMap?.debit ?? "", credit: p.bankMap?.credit ?? "", accountType: p.bankMap?.accountType ?? "", ref: (p.bankMap as any)?.ref ?? "" }); setBankSwap(p.bankSwap ?? false);
    setBankFileSessionId(p.bankFileSessionId ?? 0);
    setCashH(p.cashHeaders); setCashRows(p.cashRowsRaw); setCashMap({ date: p.cashMap?.date ?? "", name: p.cashMap?.name ?? "", debit: p.cashMap?.debit ?? "", credit: p.cashMap?.credit ?? "", accountType: p.cashMap?.accountType ?? "", ref: (p.cashMap as any)?.ref ?? "" }); setCashSwap(p.cashSwap ?? false);
    setCashFileSessionId(p.cashFileSessionId ?? 0);
    setManualGroups(p.manualGroups); setSavedMatches(p.savedMatches);
    setRejected(new Set(p.rejectedPairs));
    setVisaItems(p.visaItems ?? []);
    setHeldItems(p.heldItems ?? []);
    setReturnedHeldBank((p as any).returnedHeldBank ?? []);
    setReturnedHeldCashier((p as any).returnedHeldCashier ?? []);
    setBankFile(null); setCashFile(null);
    setResults(null);
  };

  const deleteProject = async (id: string) => {
    const projects = (await loadProjectsList()).filter(x => x.id !== id);
    await persistProjectsList(projects);
  };

  const handleClearSavedMatches = () => {
    if (!savedMatches.length) return;
    if (!window.confirm(`مسح جميع المطابقات المؤكدة (${savedMatches.length})؟`)) return;
    setSavedMatches([]);
    setManualGroups([]);
    setSelectedSavedIds(new Set());
    setExpandedMatchKey(null);
    setShowDeleteMenu(false);
    setToast("تم إلغاء جميع المطابقات المؤكدة، بما فيها المطابقات اليدوية");
  };

  const handleClearSession = () => {
    if (!window.confirm("مسح الجلسة الحالية كاملة؟\nهاد بيحذف: الملفات (بكل المنصات)، الملفات المحفوظة للاسترجاع، المطابقات، الفيزا، المعلقات، والرفوض.\nالمشاريع المحفوظة رح تضل محفوظة.")) return;
    setBankH([]); setBankRows([]); setBankMap({date:"",desc:"",debit:"",credit:"",accountType:"",ref:""}); setBankSwap(false); setBankFile(null); setBankFileSessionId(0);
    setCashH([]); setCashRows([]); setCashMap({date:"",name:"",debit:"",credit:"",accountType:"",ref:""}); setCashSwap(false); setCashFile(null); setCashFileSessionId(0); setStageAInvoices([]);
    setWalletBankFile(null); setWalletBankHeaders([]); setWalletBankRows([]); setWalletBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setWalletBankSwap(false);
    setWalletCashFile(null); setWalletCashHeaders([]); setWalletCashRows([]); setWalletCashMap({date:"",name:"",debit:"",credit:"",ref:""}); setWalletCashSwap(false);
    setJawwalBankFile(null); setJawwalBankHeaders([]); setJawwalBankRows([]); setJawwalBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setJawwalBankSwap(false);
    setJawwalCashFile(null); setJawwalCashHeaders([]); setJawwalCashRows([]); setJawwalCashMap({date:"",name:"",debit:"",credit:"",ref:""}); setJawwalCashSwap(false);
    setBankFileSnapshots([]); setCashFileSnapshots([]);
    setWalletBankFileSnapshots([]); setWalletCashFileSnapshots([]);
    setJawwalBankFileSnapshots([]); setJawwalCashFileSnapshots([]);
    setManualGroups([]); setSavedMatches([]); setRejected(new Set());
    setVisaItems([]); setHeldItems([]); setReturnedHeldBank([]); setReturnedHeldCashier([]);
    setRejectedSpecialCashierIds(new Set());
    setResults(null); setTab("saved");
    setSessionRestoredNotice(false);
    setShowDeleteMenu(false);
  };

  const handleClearPlatform = (plat: "bank" | "cash" | "walletBank" | "walletCash" | "jawwalBank" | "jawwalCash") => {
    if (!window.confirm("مسح ملف هذه المنصة والملفات المحفوظة للاسترجاع الخاصة بها فقط؟\nباقي المنصات والمطابقات المؤكدة رح تضل كما هي.")) return;
    if (plat === "bank") { setBankFile(null); setBankH([]); setBankRows([]); setBankMap({date:"",desc:"",debit:"",credit:"",accountType:"",ref:""}); setBankSwap(false); setBankFileSessionId(0); setBankFileSnapshots([]); }
    if (plat === "cash") { setCashFile(null); setCashH([]); setCashRows([]); setCashMap({date:"",name:"",debit:"",credit:"",accountType:"",ref:""}); setCashSwap(false); setCashFileSessionId(0); setCashFileSnapshots([]); }
    if (plat === "walletBank") { setWalletBankFile(null); setWalletBankHeaders([]); setWalletBankRows([]); setWalletBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setWalletBankSwap(false); setWalletBankFileSnapshots([]); }
    if (plat === "walletCash") { setWalletCashFile(null); setWalletCashHeaders([]); setWalletCashRows([]); setWalletCashMap({date:"",name:"",debit:"",credit:"",ref:""}); setWalletCashSwap(false); setWalletCashFileSnapshots([]); }
    if (plat === "jawwalBank") { setJawwalBankFile(null); setJawwalBankHeaders([]); setJawwalBankRows([]); setJawwalBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setJawwalBankSwap(false); setJawwalBankFileSnapshots([]); }
    if (plat === "jawwalCash") { setJawwalCashFile(null); setJawwalCashHeaders([]); setJawwalCashRows([]); setJawwalCashMap({date:"",name:"",debit:"",credit:"",ref:""}); setJawwalCashSwap(false); setJawwalCashFileSnapshots([]); }
    setResults(null);
    setToast("تم مسح بيانات هذه المنصة ✓");
  };

  const handleDeleteAllProjects = async () => {
    const projects = await loadProjectsList();
    if (!projects.length) return;
    if (!window.confirm(`حذف كل المشاريع المحفوظة (${projects.length})؟\nما رح تقدر ترجعها!`)) return;
    await persistProjectsList([]);
    setShowDeleteMenu(false);
  };

  const buildResumeData = (): ResumeData => ({
    bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
    cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId,
    manualGroups, savedMatches,
    rejectedPairs: Array.from(rejectedPairs),
    visaItems, heldItems, returnedHeldBank, returnedHeldCashier
  });

  const loadResumeData = (p: ResumeData) => {
    setBankH(p.bankHeaders || []); setBankRows(p.bankRowsRaw || []);
    setBankMap({ date: p.bankMap?.date ?? "", desc: p.bankMap?.desc ?? "", debit: p.bankMap?.debit ?? "", credit: p.bankMap?.credit ?? "", accountType: p.bankMap?.accountType ?? "", ref: (p.bankMap as any)?.ref ?? "" }); setBankSwap(!!p.bankSwap);
    setBankFileSessionId(p.bankFileSessionId ?? 0);
    setCashH(p.cashHeaders || []); setCashRows(p.cashRowsRaw || []);
    setCashMap({ date: p.cashMap?.date ?? "", name: p.cashMap?.name ?? "", debit: p.cashMap?.debit ?? "", credit: p.cashMap?.credit ?? "", accountType: p.cashMap?.accountType ?? "", ref: (p.cashMap as any)?.ref ?? "" }); setCashSwap(!!p.cashSwap);
    setCashFileSessionId(p.cashFileSessionId ?? 0);
    setManualGroups(p.manualGroups || []); setSavedMatches(p.savedMatches || []);
    setRejected(new Set(p.rejectedPairs || []));
    setVisaItems(p.visaItems || []);
    setHeldItems(p.heldItems || []);
    setReturnedHeldBank((p.returnedHeldBank as BankRow[]) ?? []);
    setReturnedHeldCashier((p.returnedHeldCashier as CashierRow[]) ?? []);
    setBankFile(null); setCashFile(null);
    setResults(null);
  };

  const handleImportResumeFile = async (f: File) => {
    setError(null);
    const data = await extractResumeData(f);
    if (!data) {
      setError("هذا الملف مش ملف تصدير من البرنامج (أو الشيت الخاص بالاستكمال غير موجود فيه). تأكد إنه نفس ملف الإكسل اللي حمّلته من هون.");
      return;
    }
    loadResumeData(data);
    setToast("تم استرجاع كل شغلك من الملف ✓");
  };

  if (page === "manual") {
    const manualCashierRows = activeCashier.filter(c => !visaCashierIds.has(c.id) && !specialCashierIds.has(c.id));
    const manualBankRows = activeBank;
    return (
      <ManualWorkbench
        bankRows={manualBankRows}
        cashierRows={manualCashierRows}
        manualGroups={manualGroups}
        onAddGroup={handleAddGroup}
        onRemoveGroup={handleRemoveGroup}
        onBack={() => setPage("main")}
        savedCashierIds={savedCashierIds}
        savedBankIds={savedBankIds}
        onHoldBank={handleHoldBank}
        onHoldCashier={handleHoldCashier}
      />
    );
  }

  if (page === "report") {
    return <ReportPage bank={activeBank} cashier={activeCashier} results={results} onBack={() => setPage("main")} />;
  }

  if (page === "assist") {
    return (
      <MatchAssistant
        cashierRows={uCashierRows.map((r:any) => r.cashier as CashierRow)}
        bankRows={uBankRows.map((r:any) => r.bank as BankRow)}
        rejectedPairs={rejectedPairs}
        onMatch={(c, b) => handleSaveMatch(c, b)}
        onHold={(c) => handleHoldCashier(c)}
        onHoldBank={(b) => handleHoldBank(b)}
        onReject={(c, b) => handleRejectMatch(c, b)}
        onBack={() => setPage("main")}
      />
    );
  }

  if (page === "recon2") {
    return <Reconciliation2 onBack={() => { setMainMode("default"); setPage("main"); }} onStageBToMain={() => { setMainMode("stageB"); setPage("main"); }} onStageAToMain={results => {
      setStageAInvoices(results.filter(result => result.matched && result.category !== "فيزا"));
      setPage("main");
    }} />;
  }
  if (page === "recon2StageA") {
    return <Reconciliation2 initialStage="stageA" onBack={() => { setMainMode("default"); setPage("main"); }} onStageBToMain={() => { setMainMode("stageB"); setPage("main"); }} onStageAToMain={results => {
      setStageAInvoices(results.filter(result => result.matched && result.category !== "فيزا"));
      setPage("main");
    }} />;
  }
  if (page === "recon2Visa") {
    const visaInvoices: StageAResult[] = visaItems.map(item => ({
      id: item.cashierId,
      source: "المنصة الرئيسية",
      name: item.name,
      originalName: item.rawName,
      amount: item.amount,
      databaseName: "",
      databaseAmount: 0,
      userId: "",
      registrationTime: item.date,
      customerNumber: "",
      invoiceNumber: item.visaNumber,
      accountType: "",
      invoice: { ...item },
      matched: true,
      issue: item.note || "",
      category: "فيزا",
      categoryIssue: ""
    }));
    return <Reconciliation2 initialStage="stageC" visaInvoices={visaInvoices} onBack={() => setPage("main")} />;
  }

  const togglePendingSelect = (key: string) => {
    setSelectedPendingKeys(prev => {
      const n = new Set(prev);
      if (n.has(key)) n.delete(key); else n.add(key);
      return n;
    });
  };
  const allPendingSelected = pendingRows.length > 0 && pendingRows.every((r:any) => selectedPendingKeys.has(`${r.cashier.id}-${r.bank.id}`));
  const togglePendingSelectAll = () => {
    if (allPendingSelected) { setSelectedPendingKeys(new Set()); return; }
    setSelectedPendingKeys(new Set(pendingRows.map((r:any) => `${r.cashier.id}-${r.bank.id}`)));
  };
  const handleSaveSelectedPending = () => {
    const pairs = pendingRows
      .filter((r:any) => selectedPendingKeys.has(`${r.cashier.id}-${r.bank.id}`))
      .map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow }));
    handleSaveMatchesBulk(pairs);
  };

  const handleRetryMatch = () => {
    if (!pendingRows.length) {
      setToast("ما في مطابقات منتظرة لإعادة فرزها.");
      return;
    }

    if (!window.confirm(
      `إعادة فرز دقيقة لـ ${pendingRows.length} مطابقة منتظرة.\n` +
      `رح نفك كل الروابط الحالية ونعيد المطابقة من الصفر بسعر تنافسي لكل فاتورة وحوالة.\n` +
      `النتيجة بتضل بقائمة "منتظر" — ما رح تنتقل للمؤكدة إلا إذا ضغطتي حفظ.`
    )) return;

    const pendingCashierIds = new Set(pendingRows.map((r:any) => r.cashier.id as number));
    const pendingBankIds = new Set(pendingRows.map((r:any) => r.bank.id as number));

    setRejected(prev => {
      const next = new Set<string>();
      prev.forEach(key => {
        const [cidStr, bidStr] = key.split("-");
        const cid = parseInt(cidStr);
        const bid = parseInt(bidStr);
        if (!pendingCashierIds.has(cid) && !pendingBankIds.has(bid)) {
          next.add(key);
        }
      });
      return next;
    });

    setToast(`تمت إعادة الفرز التنافسي لـ ${pendingRows.length} مطابقة منتظرة`);
  };

  const handleMoveToManual = (pairs: Array<{ cashier: CashierRow; bank: BankRow }>) => {
    if (!pairs.length) return;
    setRejected(prev => {
      const next = new Set(prev);
      pairs.forEach(({ cashier, bank }) => next.add(`${cashier.id}-${bank.id}`));
      return next;
    });
    setSelectedPendingKeys(new Set());
    setPage("manual");
  };
  const handleMoveSelectedPendingToManual = () => {
    const pairs = pendingRows
      .filter((r:any) => selectedPendingKeys.has(`${r.cashier.id}-${r.bank.id}`))
      .map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow }));
    handleMoveToManual(pairs);
  };
  const handleMoveAllPendingToManual = () => {
    if (!pendingRows.length) return;
    if (!window.confirm(`نقل كل المطابقات المنتظرة (${pendingRows.length}) للمطابقة اليدوية؟`)) return;
    handleMoveToManual(pendingRows.map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow })));
  };

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground p-6">
      <div className="max-w-6xl mx-auto space-y-5">

        {toast && <div className="fixed bottom-5 left-5 z-50 flex items-center gap-2 rounded-lg bg-slate-900 px-4 py-3 text-sm text-white shadow-xl"><Check className="h-4 w-4 text-green-300" />{toast}</div>}

        <div className="flex flex-wrap items-center gap-3 justify-between">
          <div><h1 className="text-xl font-bold">منصة التسويةوالمطابقة الذكية</h1><p className="mt-1 text-xs text-muted-foreground">مراجعة مالية أسرع، بقرارات قابلة للتتبع</p></div>
          <div className="flex items-center gap-2 flex-wrap">
          {mainMode === "stageB" && <button onClick={() => { setMainMode("default"); setPage("recon2"); }}
            className="flex items-center gap-2 rounded-lg bg-teal-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-teal-700">
            <CreditCard className="h-4 w-4"/>العودة إلى منصة التسوية والمراحل
          </button>}
          {mainMode === "default" && <>
              <button onClick={() => setPage("recon2StageA")} className="rounded-lg border bg-white px-3 py-2 text-xs font-medium hover:bg-slate-50">Stage A · آلي</button>
              <button onClick={() => setPage("assist")} className="rounded-lg border bg-white px-3 py-2 text-xs font-medium hover:bg-slate-50">Stage B · مراجعة</button>
            </>}
            <button onClick={() => setShowSettings(v => !v)} className="rounded-lg border bg-white px-3 py-2 text-xs font-medium hover:bg-slate-50">الإعدادات</button>
            <ProjectsBar onSave={saveProject} onLoad={loadProject} onDelete={deleteProject} />
             <button onClick={() => setPage("assist")}
              className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors">
               مساعد المطابقة
            </button>
            <button onClick={() => setPage("manual")}
              className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition-colors">
              <Save className="w-4 h-4"/>المطابقة اليدوية (تؤكد فوراً)
            </button>
            {mainMode === "default" && <button onClick={() => setPage("recon2")}
              className="flex items-center gap-2 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm font-medium hover:bg-teal-700 transition-colors">
              <CreditCard className="w-4 h-4"/>منصة التسوية والمراحل
            </button>}
            <div className="relative">
              <button onClick={() => setShowDeleteMenu(v => !v)}
                className="px-3 py-2 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs flex items-center gap-1 hover:bg-red-100 transition-colors">
                <Trash2 className="w-3.5 h-3.5"/>أدوات الحذف <ChevronDown className="w-3 h-3"/>
              </button>
              {showDeleteMenu && (
                <>
                  <div className="fixed inset-0 z-10" onClick={() => setShowDeleteMenu(false)} />
                  <div className="absolute left-0 mt-1 w-64 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs">
                    <button onClick={handleClearSavedMatches}
                      disabled={!savedMatches.length}
                      className="w-full text-right px-4 py-2.5 hover:bg-red-50 flex items-center gap-2 disabled:opacity-40 transition-colors">
                      <Trash2 className="w-3.5 h-3.5 text-red-600"/>
                      <span>مسح المطابقات المؤكدة</span>
                      <span className="text-muted-foreground mr-auto">{savedMatches.length}</span>
                    </button>
                    <button onClick={handleClearSession}
                      className="w-full text-right px-4 py-2.5 hover:bg-amber-50 flex items-center gap-2 transition-colors border-t border-border">
                      <RotateCcw className="w-3.5 h-3.5 text-amber-600"/>
                      <span>مسح الجلسة الحالية كاملة</span>
                    </button>
                    <button onClick={handleDeleteAllProjects}
                      className="w-full text-right px-4 py-2.5 hover:bg-red-50 flex items-center gap-2 transition-colors border-t border-border">
                      <FolderMinus className="w-3.5 h-3.5 text-red-600"/>
                      <span>حذف كل المشاريع المحفوظة</span>
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        {sessionRestoredNotice && (
          <div className="flex items-center gap-3 bg-green-50 border border-green-200 rounded-xl px-4 py-3">
            <Shield className="w-5 h-5 text-green-600 shrink-0"/>
            <p className="flex-1 text-sm font-medium text-green-800">
              تم استرجاع شغلك من آخر جلسة تلقائياً ✓ (الملفات، المطابقات، الفيزا، المعلقات)
            </p>
            <button onClick={() => setSessionRestoredNotice(false)} className="p-1 text-green-500 hover:text-green-800 transition-colors">
              <X className="w-4 h-4"/>
            </button>
          </div>
        )}

        {visaItems.length > 0 && (
          <div className="flex items-start gap-3 bg-teal-50 border border-teal-200 rounded-xl px-4 py-3">
            <CreditCard className="w-5 h-5 text-teal-600 shrink-0 mt-0.5"/>
            <div className="flex-1">
              <p className="text-sm font-semibold text-teal-800">{visaItems.length} عنصر منقول للفيزا</p>
              <p className="text-xs text-teal-600 mt-0.5">اضغط على تبويب "الفيزا" بالأسفل لمراجعتها</p>
            </div>
            <button onClick={() => setPage("recon2Visa")} className="shrink-0 rounded-lg bg-teal-600 px-3 py-2 text-xs font-medium text-white hover:bg-teal-700">
              ترحيل على C
            </button>
          </div>
        )}

        {error && <div className="p-3 bg-red-100 text-red-800 rounded-lg text-sm">{error}</div>}

        {showSettings && <div className="rounded-xl border border-blue-100 bg-blue-50/60 p-4">
          <div className="flex flex-wrap items-center gap-3">
            <div><p className="text-sm font-semibold">إعدادات المطابقة</p><p className="text-xs text-muted-foreground">تسامح رسوم التحويل يطبق على المبلغ النسبي</p></div>
            <label className="mr-auto flex items-center gap-2 text-xs">نسبة التسامح
              <input type="number" min="0" max="10" step="0.1" value={amountTolerancePercent} onChange={e => setAmountTolerancePercent(Math.min(10, Math.max(0, Number(e.target.value) || 0)))} className="w-20 rounded-lg border bg-white px-2 py-1.5 text-center" />%
            </label>
          </div>
        </div>}

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات بنك فلسطين</h2>
              {(bankHeaders.length > 0 || bankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("bank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={bankFile} onFile={loadBank} onClear={() => { setBankFile(null); setBankH([]); setBankRows([]); setResults(null); }} />
            {bankHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={bankHeaders} value={bankMap.date} onChange={v => setBankMap(m => ({ ...m, date: v }))} />
                <Sel label="الإيضاحات / البيان" headers={bankHeaders} value={bankMap.desc} onChange={v => setBankMap(m => ({ ...m, desc: v }))} />
                <Sel label="المبالغ المستلمة (مدين)" headers={bankHeaders} value={bankMap.debit} onChange={v => setBankMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المدفوعة (دائن)" headers={bankHeaders} value={bankMap.credit} onChange={v => setBankMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب (بنك فلسطين / محفظة...)" headers={bankHeaders} value={bankMap.accountType} onChange={v => setBankMap(m => ({ ...m, accountType: v }))} />
                <Sel label="رقم المرجع / الحوالة" headers={bankHeaders} value={bankMap.ref} onChange={v => setBankMap(m => ({ ...m, ref: v }))} />
              </div>
            )}
            {bankHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={bankSwap} onChange={e=>setBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {bankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({bankFileSnapshots.length})
                </button>
                {showBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {bankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setBankH(s.headers); setBankRows(s.rows);
                          setBankMap({ date:autoDetect(s.headers,HINTS.date), desc:autoDetect(s.headers,HINTS.desc), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), accountType:autoDetect(s.headers,HINTS.accountType), ref:autoDetect(s.headers,HINTS.ref) });
                          setBankFileSessionId(id => id + 1);
                          setShowBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير بنك فلسطين</h2>
              {(cashHeaders.length > 0 || cashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("cash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={cashFile} onFile={loadCash} onClear={() => { setCashFile(null); setCashH([]); setCashRows([]); setResults(null); }} />
            {cashHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={cashHeaders} value={cashMap.date} onChange={v => setCashMap(m => ({ ...m, date: v }))} />
                <Sel label="البيان" headers={cashHeaders} value={cashMap.name} onChange={v => setCashMap(m => ({ ...m, name: v }))} />
                <Sel label="المبالغ المستلمة (مدين)" headers={cashHeaders} value={cashMap.debit} onChange={v => setCashMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المدفوعة (دائن)" headers={cashHeaders} value={cashMap.credit} onChange={v => setCashMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب (بنك فلسطين / محفظة...)" headers={cashHeaders} value={cashMap.accountType} onChange={v => setCashMap(m => ({ ...m, accountType: v }))} />
                <Sel label="رقم المرجع / الحوالة" headers={cashHeaders} value={cashMap.ref} onChange={v => setCashMap(m => ({ ...m, ref: v }))} />
              </div>
            )}
            {cashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={cashSwap} onChange={e=>setCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {cashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({cashFileSnapshots.length})
                </button>
                {showCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {cashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setCashH(s.headers); setCashRows(s.rows);
                          setCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetect(s.headers,HINTS.name), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), accountType:autoDetect(s.headers,HINTS.accountType), ref:autoDetect(s.headers,HINTS.ref) });
                          setCashFileSessionId(id => id + 1);
                          setShowCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
          <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات بال بي</h2>
              {(walletBankHeaders.length > 0 || walletBankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("walletBank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={walletBankFile} onFile={loadWalletBank} onClear={() => { setWalletBankFile(null); setWalletBankHeaders([]); setWalletBankRows([]); setResults(null); }} />
            {walletBankHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={walletBankHeaders} value={walletBankMap.date} onChange={v => setWalletBankMap(m => ({ ...m, date:v }))} />
              <Sel label="الإيضاحات / البيان" headers={walletBankHeaders} value={walletBankMap.desc} onChange={v => setWalletBankMap(m => ({ ...m, desc:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={walletBankHeaders} value={walletBankMap.debit} onChange={v => setWalletBankMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={walletBankHeaders} value={walletBankMap.credit} onChange={v => setWalletBankMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={walletBankHeaders} value={walletBankMap.ref} onChange={v => setWalletBankMap(m => ({ ...m, ref:v }))} />
            </div>}
            {walletBankHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={walletBankSwap} onChange={e=>setWalletBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {walletBankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowWalletBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({walletBankFileSnapshots.length})
                </button>
                {showWalletBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowWalletBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {walletBankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setWalletBankHeaders(s.headers); setWalletBankRows(s.rows);
                          setWalletBankMap({ date:autoDetect(s.headers,HINTS.date), desc:autoDetect(s.headers,HINTS.desc), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref) });
                          setShowWalletBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير بال بي</h2>
              {(walletCashHeaders.length > 0 || walletCashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("walletCash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={walletCashFile} onFile={loadWalletCash} onClear={() => { setWalletCashFile(null); setWalletCashHeaders([]); setWalletCashRows([]); setResults(null); }} />
            {walletCashHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={walletCashHeaders} value={walletCashMap.date} onChange={v => setWalletCashMap(m => ({ ...m, date:v }))} />
              <Sel label="البيان" headers={walletCashHeaders} value={walletCashMap.name} onChange={v => setWalletCashMap(m => ({ ...m, name:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={walletCashHeaders} value={walletCashMap.debit} onChange={v => setWalletCashMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={walletCashHeaders} value={walletCashMap.credit} onChange={v => setWalletCashMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={walletCashHeaders} value={walletCashMap.ref} onChange={v => setWalletCashMap(m => ({ ...m, ref:v }))} />
            </div>}
            {walletCashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={walletCashSwap} onChange={e=>setWalletCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {walletCashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowWalletCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({walletCashFileSnapshots.length})
                </button>
                {showWalletCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowWalletCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {walletCashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setWalletCashHeaders(s.headers); setWalletCashRows(s.rows);
                          setWalletCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetect(s.headers,HINTS.name), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref) });
                          setShowWalletCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات جوال بي</h2>
              {(jawwalBankHeaders.length > 0 || jawwalBankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("jawwalBank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={jawwalBankFile} onFile={loadJawwalBank} onClear={() => { setJawwalBankFile(null); setJawwalBankHeaders([]); setJawwalBankRows([]); setResults(null); }} />
            {jawwalBankHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={jawwalBankHeaders} value={jawwalBankMap.date} onChange={v => setJawwalBankMap(m => ({ ...m, date:v }))} />
              <Sel label="الإيضاحات / البيان" headers={jawwalBankHeaders} value={jawwalBankMap.desc} onChange={v => setJawwalBankMap(m => ({ ...m, desc:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={jawwalBankHeaders} value={jawwalBankMap.debit} onChange={v => setJawwalBankMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={jawwalBankHeaders} value={jawwalBankMap.credit} onChange={v => setJawwalBankMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={jawwalBankHeaders} value={jawwalBankMap.ref} onChange={v => setJawwalBankMap(m => ({ ...m, ref:v }))} />
            </div>}
            {jawwalBankHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={jawwalBankSwap} onChange={e=>setJawwalBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {jawwalBankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowJawwalBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({jawwalBankFileSnapshots.length})
                </button>
                {showJawwalBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowJawwalBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {jawwalBankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setJawwalBankHeaders(s.headers); setJawwalBankRows(s.rows);
                          setJawwalBankMap({ date:autoDetect(s.headers,HINTS.date), desc:autoDetect(s.headers,HINTS.desc), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref) });
                          setShowJawwalBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير جوال بي</h2>
              {(jawwalCashHeaders.length > 0 || jawwalCashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("jawwalCash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={jawwalCashFile} onFile={loadJawwalCash} onClear={() => { setJawwalCashFile(null); setJawwalCashHeaders([]); setJawwalCashRows([]); setResults(null); }} />
            {jawwalCashHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={jawwalCashHeaders} value={jawwalCashMap.date} onChange={v => setJawwalCashMap(m => ({ ...m, date:v }))} />
              <Sel label="البيان" headers={jawwalCashHeaders} value={jawwalCashMap.name} onChange={v => setJawwalCashMap(m => ({ ...m, name:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={jawwalCashHeaders} value={jawwalCashMap.debit} onChange={v => setJawwalCashMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={jawwalCashHeaders} value={jawwalCashMap.credit} onChange={v => setJawwalCashMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={jawwalCashHeaders} value={jawwalCashMap.ref} onChange={v => setJawwalCashMap(m => ({ ...m, ref:v }))} />
            </div>}
            {jawwalCashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={jawwalCashSwap} onChange={e=>setJawwalCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {jawwalCashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowJawwalCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({jawwalCashFileSnapshots.length})
                </button>
                {showJawwalCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowJawwalCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {jawwalCashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setJawwalCashHeaders(s.headers); setJawwalCashRows(s.rows);
                          setJawwalCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetect(s.headers,HINTS.name), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref) });
                          setShowJawwalCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
        </div>

        <div className="flex gap-3 flex-wrap">
          <button onClick={run} disabled={!canRun || loading}
            className="px-6 py-2.5 bg-blue-600 text-white rounded-lg disabled:opacity-40 font-medium hover:bg-blue-700 transition-colors flex items-center gap-2">
             {loading ? <><span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />جاري المعالجة...</> : "تشغيل المطابقة"}
          </button>
          {results && (
            <button onClick={() => {
                try {
                  doExport(results, savedMatches, visaItems, heldItems, activeBank, activeCashier, buildResumeData());
                } catch (e) {
                  console.error("فشل تصدير الملف:", e);
                  setError("صار خطأ أثناء تصدير الملف. جرّب تاني.");
                }
              }}
              className="px-4 py-2.5 border rounded-lg flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white border-transparent transition-colors font-medium text-sm">
              <Download className="w-4 h-4" />تصدير Excel
            </button>
          )}
          <label className="px-4 py-2.5 border border-teal-300 bg-teal-50 text-teal-700 rounded-lg flex items-center gap-1.5 hover:bg-teal-100 transition-colors font-medium text-sm cursor-pointer">
            <Upload className="w-4 h-4" />استيراد ملف سابق
            <input type="file" accept=".xlsx,.xls" className="hidden"
              onChange={e => { const f = e.target.files?.[0]; if (f) handleImportResumeFile(f); e.target.value = ""; }}/>
          </label>
        </div>

        {results && stats && (
          <div className="space-y-4">
            <div className="grid grid-cols-4 sm:grid-cols-8 gap-2">
               {([
                 { id: "saved", label: "مؤكد", count: stats.saved, color: "text-green-700" },
                 { id: "pending", label: "منتظر", count: stats.pending, color: "text-blue-700" },
                 { id: "manual", label: "يدوي", count: stats.manual, color: "text-orange-700" },
                 { id: "visa", label: "فيزا", count: stats.visa, color: "text-cyan-700" },
                 { id: "jawwalPay", label: "جوال بي", count: stats.jawwalPay, color: "text-cyan-700" },
                 { id: "mahmoudWallet", label: "محفظة محمود", count: stats.mahmoudWallet, color: "text-violet-700" },
                 { id: "held", label: "معلّقة", count: stats.held, color: "text-orange-700" },
                 { id: "uCashier", label: "فواتير بلا حوالات مطابقة", count: stats.uCashier, color: "text-red-700" },
                 { id: "uBank", label: "حوالات بلا فواتير مسجلة", count: stats.uBank, color: "text-red-700" },
              ] as const).map(t => (
                <button key={t.id} onClick={() => setTab(t.id as TabId)}
                  className={`min-h-[62px] rounded-2xl border px-3 py-2 text-right transition-all ${tab === t.id ? "border-blue-500 bg-blue-50/70" : "bg-card hover:bg-muted/30"}`}>
                  <span className="block truncate text-[10px] leading-4 text-muted-foreground">{t.label}</span>
                  <span className={`mt-0.5 block font-mono text-base font-bold leading-5 ${t.color}`}>{t.count}</span>
                </button>
              ))}
            </div>

            <div className="border rounded-xl bg-card overflow-hidden">
              <FilterBar search={resultSearch} onSearch={setResultSearch} type={resultType} onType={setResultType} minAmount={minAmount} onMinAmount={setMinAmount} maxAmount={maxAmount} onMaxAmount={setMaxAmount} />

              {tab === "saved" && (
                <>
                <div className="flex flex-wrap items-center gap-2 border-b bg-card px-3 py-2.5">
                  <button
                    type="button"
                    onClick={() => setSelectedSavedIds(new Set(visibleSavedRows.map(row => (row.savedMatch as SavedMatch).id)))}
                    className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs text-blue-700 hover:bg-blue-100"
                  >
                    تحديد الكل
                  </button>
                  <button
                    type="button"
                    onClick={handleUnsaveSelected}
                    disabled={!selectedSavedIds.size}
                    className="flex items-center gap-1.5 rounded-lg border border-amber-200 bg-amber-50 px-3 py-1.5 text-xs text-amber-700 hover:bg-amber-100 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <Link2Off className="h-3.5 w-3.5" /> إلغاء تأكيد المحدد ({selectedSavedIds.size})
                  </button>
                </div>
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="w-10 px-2 py-2.5 text-center">
                      <input
                        type="checkbox"
                        checked={visibleSavedRows.length > 0 && visibleSavedRows.every(row => selectedSavedIds.has((row.savedMatch as SavedMatch).id))}
                        onChange={event => setSelectedSavedIds(event.target.checked
                          ? new Set(visibleSavedRows.map(row => (row.savedMatch as SavedMatch).id))
                          : new Set())}
                        aria-label="تحديد المطابقات المؤكدة"
                      />
                    </th>
                    <th className="px-3 py-2.5 text-right">البيان (الكاشير)</th>
                    <th className="px-3 py-2.5 text-right">بيان البنك</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع المطابقة</th>
                    <th className="px-3 py-2.5 text-right">الحالة</th>
                    <th className="px-3 py-2.5 text-right">الثقة</th>
                    <th className="px-3 py-2.5 w-10"></th>
                  </tr></thead>
                  <tbody>
                    {visibleSavedRows.map((r: any, i: number) => {
                      const sm = r.savedMatch as SavedMatch;
                      const key = sm.id;
                      const isExpanded = expandedMatchKey === key;
                      const typeColor = sm.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const matchTypeLabel = sm.isManual ? "🖐️ يدوي" : "🤖 تلقائي";
                      const isFromHeld = !!r.cashier._fromHeld || !!r.bank._fromHeld;
                      const visaNum = extractVisaNumber(r.cashier.rawName);
                      return (
                        <React.Fragment key={i}>
                          <tr className={`border-t transition-colors ${isExpanded ? "bg-green-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : "hover:bg-muted/20"}`}>
                            <td className="px-2 py-2.5 text-center">
                              <input
                                type="checkbox"
                                checked={selectedSavedIds.has(sm.id)}
                                onChange={event => setSelectedSavedIds(prev => {
                                  const next = new Set(prev);
                                  if (event.target.checked) next.add(sm.id); else next.delete(sm.id);
                                  return next;
                                })}
                                aria-label={`تحديد مطابقة ${r.cashier.name}`}
                              />
                            </td>
                            <td className="px-3 py-2.5 font-medium flex items-center gap-1.5">
                              <Shield className="w-3.5 h-3.5 text-green-600" />
                              {r.cashier.name}
                              {isFromHeld && <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                              {visaNum && <span className="text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300">فيزا {visaNum}</span>}
                            </td>
                            <td className="px-3 py-2.5 text-muted-foreground text-xs">
                              {r.bank.description}
                              {sm.bankAccountType && <div className="text-[10px] text-muted-foreground">{sm.bankAccountType}</div>}
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{sm.type}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-green-700">
                              {fmtNum(r.bank.rawAmount)}
                              {sm.isAmountDiff && <div className="text-[10px] text-orange-600 font-normal">كاشير: {fmtNum(r.cashier.amount)}</div>}
                            </td>
                            <td className="px-3 py-2.5 text-xs">{matchTypeLabel}</td>
                            <td className="px-3 py-2.5"><ConfidenceBadge score={sm.matchScore} /></td>
                            <td className="px-3 py-2.5">
                              <div className="space-y-0.5">
                                <span className="text-green-600 bg-green-50 px-1.5 py-0.5 rounded text-xs border border-green-200 flex items-center gap-0.5 w-fit">
                                  <Check className="w-2.5 h-2.5"/>✓ محفوظ
                                </span>
                                {sm.isAmountDiff && <div className="text-[10px] text-amber-600">⚠️ اختلاف مبلغ</div>}
                                {sm.isNameDiff && <div className="text-[10px] text-amber-600">⚠️ اختلاف اسم</div>}
                                {sm.isAccountTypeDiff && <div className="text-[10px] text-orange-700 font-medium">🏦 {sm.bankAccountType || "؟"} ↔ {sm.cashierAccountType || "؟"}</div>}
                                {sm.note && <div className="text-[10px] text-gray-500">📝 {sm.note}</div>}
                              </div>
                            </td>
                            <td className="px-3 py-2.5 text-center">
                              <button onClick={() => setExpandedMatchKey(isExpanded ? null : key)}
                                className={`p-1 rounded transition-colors ${isExpanded ? "text-blue-600 bg-blue-100" : "text-muted-foreground hover:text-foreground hover:bg-muted"}`}>
                                <span className="text-base leading-none font-bold">⋮</span>
                              </button>
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr className="bg-green-50 border-t border-green-200">
                              <td colSpan={8} className="px-4 py-3">
                                <div className="flex items-center gap-3 flex-wrap">
                                  <span className="text-xs text-green-700 font-medium">{r.cashier.name} ↔ {r.bank.description.slice(0, 30)}...</span>
                                  <button onClick={() => handleUnsaveMatch(sm)}
                                    className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-600 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                                    <Link2Off className="w-3.5 h-3.5"/>إلغاء الحفظ
                                  </button>
                                  <button onClick={() => handleRejectMatch(r.cashier, r.bank)}
                                    className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                    <X className="w-3.5 h-3.5"/>رفض
                                  </button>
                                  {visaNum && (
                                    <button onClick={() => handleMoveToVisa(r.cashier, "saved")}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-teal-600 text-white rounded-lg text-xs hover:bg-teal-700 transition-colors">
                                      <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                    </button>
                                  )}
                                  <span className="text-xs text-muted-foreground">{sm.type}: {fmtNum(r.bank.rawAmount)}</span>
                                  <button onClick={() => setExpandedMatchKey(null)} className="mr-auto p-1 text-muted-foreground hover:text-foreground"><X className="w-3.5 h-3.5"/></button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                    {!visibleSavedRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد مطابقات مؤكدة تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
                </>
              )}

              {tab === "pending" && (
                <div className="flex gap-4">
                <div className="flex-1 min-w-0 overflow-hidden">
                  {pendingRows.length > 0 && (
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-indigo-50/60 border-b border-indigo-100 flex-wrap">
                      <span className="text-xs text-indigo-800">إعادة المطابقة بتفحص المنتظرة، تفك ربط الضعيفة وتعيد فرزها</span>
                      <button onClick={handleRetryMatch}
                        className="mr-auto flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 text-white rounded-lg text-xs font-medium hover:bg-indigo-700 transition-colors">
                        🔄 إعادة المطابقة
                      </button>
                    </div>
                  )}
                  {pendingRows.length > 0 && (
                    <>
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-blue-50/60 border-b border-blue-100 flex-wrap">
                      <label className="flex items-center gap-1.5 text-xs font-medium text-blue-800 cursor-pointer">
                        <input type="checkbox" checked={allPendingSelected} onChange={togglePendingSelectAll} className="rounded"/>
                        تحديد الكل
                      </label>
                      <span className="text-xs text-muted-foreground">{selectedPendingKeys.size} محدد من {pendingRows.length}</span>
                      <div className="mr-auto flex items-center gap-2 flex-wrap">
                        <button onClick={handleMoveSelectedPendingToManual} disabled={!selectedPendingKeys.size}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs font-medium hover:bg-amber-100 disabled:opacity-40 transition-colors">
                          نقل المحدد لليدوي ({selectedPendingKeys.size})
                        </button>
                        <button onClick={handleMoveAllPendingToManual}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs font-medium hover:bg-amber-100 transition-colors">
                          نقل الكل لليدوي ({pendingRows.length})
                        </button>
                        <button onClick={handleSaveSelectedPending} disabled={!selectedPendingKeys.size}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs font-medium hover:bg-green-700 disabled:opacity-40 transition-colors">
                          <Save className="w-3.5 h-3.5"/>حفظ المحدد ({selectedPendingKeys.size})
                        </button>
                      </div>
                    </div>
                    <div className="px-4 py-1.5 bg-indigo-50/50 border-b border-indigo-100 text-[11px] text-indigo-700">
                      💡 اضغط "بحث وسحب" بجانب بيان البنك لاستبدال الحوالة، أو بجانب اسم الكاشير لاستبدال الفاتورة — وبعدين اسحب العنصر من اللوحة وأفلته فوق نفس الخانة بالجدول.
                    </div>
                    </>
                  )}
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 w-8"></th>
                      <th className="px-3 py-2.5 text-right">البيان (الكاشير)</th>
                      <th className="px-3 py-2.5 text-right">بيان البنك</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع المطابقة</th>
                      <th className="px-3 py-2.5 text-right">الحالة</th>
                      <th className="px-3 py-2.5 w-10"></th>
                    </tr></thead>
                    <tbody>
                      {pagedPendingRows.map((r: any, i: number) => {
                        const key = `${r.cashier.id}-${r.bank.id}`;
                        const isExpanded = expandedMatchKey === key;
                        const isSelected = selectedPendingKeys.has(key);
                        const isDropTarget = dropTargetKey === key;
                        const typeColor = r.cashier.type === "مدفوع" ? "text-red-600" : "text-green-600";
                        const matchTypeLabel = r.matchType === "firstSecond" ? "👤 الأول+الثاني" :
                                              r.matchType === "fourthName" ? "👤 الرابع" :
                                              r.matchType === "exact" ? "تطابق تام" :
                                              r.matchType === "bundle" ? "تجميع" :
                                              r.matchType === "typo" ? "تقريبي" : "غير محدد";
                        const isFromHeld = !!r.cashier._fromHeld || !!r.bank._fromHeld;
                        const visaNum = extractVisaNumber(r.cashier.rawName);
                        return (
                          <React.Fragment key={i}>
                            <tr ref={row => { pendingRowRefs.current[key] = row; }} className={`border-t transition-colors ${isDropTarget ? (drawerMode === "bank" ? "bg-blue-50 ring-2 ring-blue-400 ring-inset" : "bg-emerald-50 ring-2 ring-emerald-400 ring-inset") : isSelected ? "bg-blue-50/70" : isExpanded ? "bg-blue-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : r.isApprox ? "bg-amber-50/30 hover:bg-amber-50/60" : "hover:bg-muted/20"}`}>
                              <td className="px-3 py-2.5">
                                <div onClick={() => togglePendingSelect(key)}
                                  className={`w-4 h-4 rounded border-2 flex items-center justify-center cursor-pointer transition-colors ${isSelected?"bg-blue-600 border-blue-600":"border-border"}`}>
                                  {isSelected && <Check className="w-2.5 h-2.5 text-white"/>}
                                </div>
                              </td>
                              <td
                                className={`px-3 py-2.5 font-medium transition-colors ${isDropTarget && drawerMode === "cashier" ? "bg-emerald-100" : ""}`}
                                onDragOver={(e) => { if (draggedItemId !== null && drawerMode === "cashier") { e.preventDefault(); setDropTargetKey(key); } }}
                                onDragLeave={() => { if (dropTargetKey === key) setDropTargetKey(null); }}
                                onDrop={(e) => { e.preventDefault(); if (draggedItemId !== null && drawerMode === "cashier") { handleDropCashier(r.bank, r.cashier, draggedItemId); } setDraggedItemId(null); }}
                              >
                                <div className="flex items-center gap-1.5">
                                  <span className="flex-1">
                                    {r.cashier.name}
                                    {r.cashier.accountType && <div className="text-[10px] text-muted-foreground font-normal">{r.cashier.accountType}</div>}
                                  </span>
                                  {isFromHeld && <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 shrink-0">📦 من ملف سابق</span>}
                                  {visaNum && <span className="text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 shrink-0">فيزا {visaNum}</span>}
                                  <button
                                    onClick={() => openDrawer(r.cashier, r.bank, "cashier")}
                                    title="افتح قائمة فواتير الكاشير للبحث والسحب"
                                    className="px-1.5 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded text-[10px] font-medium hover:bg-emerald-100 transition-colors flex items-center gap-0.5 shrink-0"
                                  >
                                    <Search className="w-2.5 h-2.5"/>بحث وسحب
                                  </button>
                                </div>
                              </td>
                              <td
                                className={`px-3 py-2.5 text-muted-foreground text-xs transition-colors ${isDropTarget && drawerMode === "bank" ? "bg-blue-100" : ""}`}
                                onDragOver={(e) => { if (draggedItemId !== null && drawerMode === "bank") { e.preventDefault(); setDropTargetKey(key); } }}
                                onDragLeave={() => { if (dropTargetKey === key) setDropTargetKey(null); }}
                                onDrop={(e) => { e.preventDefault(); if (draggedItemId !== null && drawerMode === "bank") { handleDropBank(r.cashier, r.bank, draggedItemId); } setDraggedItemId(null); }}
                              >
                                <div className="flex items-center gap-1.5">
                                  <span className="flex-1">{r.bank.description}</span>
                                  {r.bank.accountType && <span className="text-[10px] text-muted-foreground">{r.bank.accountType}</span>}
                                  <button
                                    onClick={() => openDrawer(r.cashier, r.bank, "bank")}
                                    title="افتح قائمة الحوالات للبحث والسحب"
                                    className="px-1.5 py-0.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded text-[10px] font-medium hover:bg-indigo-100 transition-colors flex items-center gap-0.5 shrink-0"
                                  >
                                    <Search className="w-2.5 h-2.5"/>بحث وسحب
                                  </button>
                                </div>
                              </td>
                              <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.cashier.type}</td>
                              <td className="px-3 py-2.5 font-mono font-bold text-blue-700">{fmtNum(r.bank.rawAmount)}</td>
                              <td className="px-3 py-2.5 text-xs">{matchTypeLabel}</td>
                              <td className="px-3 py-2.5">
                                <div className="flex flex-col gap-1">
                                  {r.isBundled ? (
                                    <span className="text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded text-xs border border-blue-200">تجميع</span>
                                  ) : r.isApprox ? (
                                    <span className="text-amber-700 bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded text-xs flex items-center gap-0.5 w-fit">
                                      <AlertTriangle className="w-2.5 h-2.5"/>اسم تقريبي
                                    </span>
                                  ) : (
                                    <span className="text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded text-xs">آلي ✓</span>
                                  )}
                                  <ConfidenceBadge score={Math.round(Math.min(1, nameSim(r.cashier.name, r.bank.description) + (r.isApprox ? 0.25 : 0.45)) * 100)} />
                                  {r.accountTypeDiff && (
                                    <span className="text-orange-700 bg-orange-50 border border-orange-200 px-1.5 py-0.5 rounded text-xs flex items-center gap-0.5 w-fit">
                                      <AlertTriangle className="w-2.5 h-2.5"/>{r.bank.accountType || "؟"} ↔ {r.cashier.accountType || "؟"}
                                    </span>
                                  )}
                                  <div className="flex gap-1 flex-wrap">
                                    <button onClick={() => handleSaveMatch(r.cashier, r.bank)}
                                      className="px-2 py-0.5 bg-green-600 text-white rounded text-[10px] hover:bg-green-700 transition-colors flex items-center gap-0.5">
                                      <Save className="w-2.5 h-2.5"/>حفظ
                                    </button>
                                    <button onClick={() => handleRejectMatch(r.cashier, r.bank)}
                                      className="px-2 py-0.5 bg-red-50 text-red-600 border border-red-200 rounded text-[10px] hover:bg-red-100 transition-colors flex items-center gap-0.5">
                                      <X className="w-2.5 h-2.5"/>رفض
                                    </button>
                                    <button onClick={() => handleHoldCashier(r.cashier)}
                                      title="علّق فاتورة الكاشير لمطابقتها لاحقاً"
                                      className="px-2 py-0.5 bg-purple-50 text-purple-700 border border-purple-200 rounded text-[10px] hover:bg-purple-100 transition-colors flex items-center gap-0.5">
                                      <CreditCard className="w-2.5 h-2.5"/>علّق كاشير
                                    </button>
                                    <button onClick={() => handleHoldBank(r.bank)}
                                      title="علّق حوالة البنك لمطابقتها لاحقاً"
                                      className="px-2 py-0.5 bg-purple-50 text-purple-700 border border-purple-200 rounded text-[10px] hover:bg-purple-100 transition-colors flex items-center gap-0.5">
                                      <CreditCard className="w-2.5 h-2.5"/>علّق بنك
                                    </button>
                                    {visaNum && (
                                      <button onClick={() => handleMoveToVisa(r.cashier, "pending")}
                                        title="نقل هذه الفاتورة للفيزا"
                                        className="px-2 py-0.5 bg-teal-600 text-white rounded text-[10px] hover:bg-teal-700 transition-colors flex items-center gap-0.5">
                                        <CreditCard className="w-2.5 h-2.5"/>نقل للفيزا
                                      </button>
                                    )}
                                  </div>
                                </div>
                              </td>
                              <td className="px-3 py-2.5 text-center">
                                <button onClick={() => setExpandedMatchKey(isExpanded ? null : key)}
                                  className={`p-1 rounded transition-colors ${isExpanded ? "text-blue-600 bg-blue-100" : "text-muted-foreground hover:text-foreground hover:bg-muted"}`}>
                                  <span className="text-base leading-none font-bold">⋮</span>
                                </button>
                              </td>
                            </tr>
                            {isExpanded && (
                              <tr className="bg-blue-50 border-t border-blue-200">
                                <td colSpan={8} className="px-4 py-3">
                                  <div className="flex items-center gap-3 flex-wrap">
                                    <span className="text-xs text-blue-700 font-medium">{r.cashier.name} ↔ {r.bank.description.slice(0, 30)}...</span>
                                    <button onClick={() => handleSaveMatch(r.cashier, r.bank)}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                      <Save className="w-3.5 h-3.5"/>حفظ المطابقة
                                    </button>
                                    <button onClick={() => handleRejectMatch(r.cashier, r.bank)}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                      <X className="w-3.5 h-3.5"/>رفض
                                    </button>
                                    {visaNum && (
                                      <button onClick={() => handleMoveToVisa(r.cashier, "pending")}
                                        className="flex items-center gap-1.5 px-3 py-1.5 bg-teal-600 text-white rounded-lg text-xs hover:bg-teal-700 transition-colors">
                                        <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                      </button>
                                    )}
                                    <span className="text-xs text-muted-foreground">{r.cashier.type}: {fmtNum(r.bank.rawAmount)}</span>
                                    <button onClick={() => setExpandedMatchKey(null)} className="mr-auto p-1 text-muted-foreground hover:text-foreground"><X className="w-3.5 h-3.5"/></button>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                      {!pagedPendingRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد مطابقات منتظرة تطابق الفلاتر</td></tr>}
                    </tbody>
                  </table>
                  {visiblePendingRows.length > PENDING_PAGE_SIZE && (
                    <div className="flex items-center justify-center gap-3 px-4 py-3 border-t border-border bg-muted/30 text-xs">
                      <button onClick={() => setPendingPage(p => Math.max(1, p - 1))} disabled={pendingPage <= 1}
                        className="px-3 py-1.5 rounded-lg border border-border bg-card disabled:opacity-40 hover:bg-muted/50 transition-colors flex items-center gap-1">
                        <ChevronRight className="w-3.5 h-3.5"/>السابق
                      </button>
                      <span className="text-muted-foreground">
                        صفحة {pendingPage} من {Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE))} · عرض {pagedPendingRows.length} من {visiblePendingRows.length}
                      </span>
                      <button onClick={() => setPendingPage(p => Math.min(Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE)), p + 1))}
                        disabled={pendingPage >= Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE)}
                        className="px-3 py-1.5 rounded-lg border border-border bg-card disabled:opacity-40 hover:bg-muted/50 transition-colors flex items-center gap-1">
                        التالي<ChevronLeft className="w-3.5 h-3.5"/>
                      </button>
                    </div>
                  )}
                </div>
                {drawerOpen && drawerCashier && drawerOldBank && (
                  <MatchDrawer
                    mode={drawerMode}
                    cashier={drawerCashier}
                    oldBank={drawerOldBank}
                    nameSearch={drawerNameSearch}
                    amountFrom={drawerAmountFrom}
                    amountTo={drawerAmountTo}
                    banks={drawerBanks}
                    cashiers={drawerCashiers}
                    draggedItemId={draggedItemId}
                    onNameSearch={setDrawerNameSearch}
                    onAmountFrom={setDrawerAmountFrom}
                    onAmountTo={setDrawerAmountTo}
                    onDragStart={setDraggedItemId}
                    onDragEnd={() => { setDraggedItemId(null); setDropTargetKey(null); }}
                    onReplaceBank={(b) => handleReplaceBank(drawerCashier, drawerOldBank, b)}
                    onReplaceCashier={(c) => handleReplaceCashier(drawerOldBank, drawerCashier, c)}
                    onClose={closeDrawer}
                  />
                )}
              </div>
              )}

              {tab === "manual" && (
                <div className="p-4 text-center text-muted-foreground text-xs">
                  <p>المطابقات اليدوية تذهب مباشرة إلى "المؤكدة"</p>
                  <p className="mt-2">جميع المطابقات اليدوية محفوظة بشكل نهائي</p>
                </div>
              )}

              {tab === "visa" && (
                <div>
                  <div className="px-4 py-3 bg-teal-50/60 border-b border-teal-100">
                    <p className="text-xs text-teal-800">
                      هذه الخانة خاصة بالفواتير التي تحتوي على رقم فيزا (4 أرقام بعد / أو -).
                      تقدر تنقل أي فاتورة من "المنتظر" أو "غير المتطابق" أو "المؤكد" لهونا بضغطة زر "نقل للفيزا".
                    </p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">ملاحظة</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {visaItems.map(v => {
                        const typeColor = v.type === "مدفوع" ? "text-red-600" : "text-green-600";
                        return (
                          <tr key={v.id} className="border-t hover:bg-muted/20 transition-colors">
                            <td className="px-3 py-2.5 font-medium">
                              {v.name}
                              <div className="text-[10px] text-muted-foreground font-normal">{v.rawName}</div>
                            </td>
                            <td className="px-3 py-2.5">
                              <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{v.visaNumber}</span>
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{v.type}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-teal-700">{fmtNum(v.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{v.date}</td>
                            <td className="px-3 py-2.5 text-xs text-gray-500">
                              {v.note || "—"}
                              <div className="text-[10px] text-muted-foreground">{v.movedAt}</div>
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => handleRemoveVisaItem(v.id)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>إزالة من الفيزا
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                      {!visaItems.length && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">
                        لا توجد عناصر فيزا حالياً. نقل أي فاتورة تحتوي على رقم (4 أرقام بعد / أو -) من تبويب "منتظر" أو "كاشير غير متطابق" بضغط زر "نقل للفيزا".
                      </td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "jawwalPay" && (
                <div>
                  <div className="border-b border-cyan-100 bg-cyan-50/60 px-4 py-3">
                    <p className="text-xs text-cyan-800">هذه الخانة خاصة بكل الفواتير التي تحتوي على «جوال» أو «جوال بي». لا تدخل هذه الفواتير في المطابقة.</p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">البيان الأصلي</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {results.filter(r => r.type === "jawwalPay").map((r: any) => (
                        <tr key={r.cashier.id} className="border-t">
                          <td className="px-3 py-2.5 font-medium">{r.cashier.name}</td>
                          <td className="px-3 py-2.5">{r.cashier.rawName}</td>
                          <td className="px-3 py-2.5">{r.cashier.type}</td>
                          <td className="px-3 py-2.5 font-mono font-bold text-cyan-700">{fmtNum(r.cashier.amount)}</td>
                          <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                          <td className="px-3 py-2.5">{r.cashier.date || "—"}</td>
                          <td className="px-3 py-2.5">
                            <button onClick={() => handleRejectSpecialCashier(r.cashier)}
                              className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                              <X className="w-3.5 h-3.5"/>رفض وإرجاع للمنتظر
                            </button>
                          </td>
                        </tr>
                      ))}
                      {!results.some(r => r.type === "jawwalPay") && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد فواتير جوال بي.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "mahmoudWallet" && (
                <div>
                  <div className="border-b border-indigo-100 bg-indigo-50/60 px-4 py-3">
                    <p className="text-xs text-indigo-800">هذه الخانة خاصة بكل الفواتير التي تحتوي على «محفظة محمود». لا تدخل هذه الفواتير في المطابقة.</p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">البيان الأصلي</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {results.filter(r => r.type === "mahmoudWallet").map((r: any) => (
                        <tr key={r.cashier.id} className="border-t">
                          <td className="px-3 py-2.5 font-medium">{r.cashier.name}</td>
                          <td className="px-3 py-2.5">{r.cashier.rawName}</td>
                          <td className="px-3 py-2.5">{r.cashier.type}</td>
                          <td className="px-3 py-2.5 font-mono font-bold text-indigo-700">{fmtNum(r.cashier.amount)}</td>
                          <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                          <td className="px-3 py-2.5">{r.cashier.date || "—"}</td>
                          <td className="px-3 py-2.5">
                            <button onClick={() => handleRejectSpecialCashier(r.cashier)}
                              className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                              <X className="w-3.5 h-3.5"/>رفض وإرجاع للمنتظر
                            </button>
                          </td>
                        </tr>
                      ))}
                      {!results.some(r => r.type === "mahmoudWallet") && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد فواتير محفظة محمود.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "held" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع (مدفوع/مستلم)</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">تاريخ التعليق</th>
                    <th className="px-3 py-2.5 text-right">ملاحظة</th>
                    <th className="px-3 py-2.5 text-right">إجراء</th>
                  </tr></thead>
                  <tbody>
                    {heldItems.map(h => {
                      const isBank = h.kind === "bank";
                      const label = isBank ? (h.data as BankRow).description : (h.data as CashierRow).name;
                      const amt = isBank ? (h.data as BankRow).rawAmount : (h.data as CashierRow).amount;
                      const type = h.data.type;
                      const typeColor = type === "مدفوع" ? "text-red-600" : "text-green-600";
                      return (
                        <tr key={h.id} className="border-t hover:bg-muted/20 transition-colors">
                          <td className="px-3 py-2.5">
                            <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${isBank ? "bg-blue-100 text-blue-800" : "bg-green-100 text-green-800"}`}>
                              {isBank ? "بنك" : "كاشير"}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 font-medium">{label}</td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{type}</td>
                          <td className="px-3 py-2.5 font-mono font-semibold text-purple-700">{fmtNum(amt)}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{h.heldAt}</td>
                          <td className="px-3 py-2.5 text-xs text-gray-500">{h.note || "—"}</td>
                          <td className="px-3 py-2.5">
                            <div className="flex gap-1">
                              <button onClick={() => handleUnhold(h)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                <Link2 className="w-3.5 h-3.5"/>استرجاع للمطابقة
                              </button>
                              <button onClick={() => handleDeleteHeld(h.id)}
                                className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors">
                                <Trash2 className="w-3.5 h-3.5"/>
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {!heldItems.length && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد عناصر معلّقة حالياً. علّق أي حوالة أو فاتورة من تبويب "منتظر" أو "كاشير/بنك غير متطابق" لمطابقتها لاحقاً.</td></tr>}
                  </tbody>
                </table>
              )}

              {tab === "uCashier" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع الحساب المسجل</th>
                    <th className="px-3 py-2.5 text-right">اسم المستخدم</th>
                    <th className="px-3 py-2.5 text-right">السبب</th>
                    <th className="px-3 py-2.5 text-right">اقتراح</th>
                    <th className="px-3 py-2.5 text-right">إجراء</th>
                  </tr></thead>
                  <tbody>
                    {visibleUCashierRows.map((r: any, i: number) => {
                      const isExpanded = expandedUnmatchedCashier === r.cashier.id;
                      const typeColor = r.cashier.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const rejectedForThisCashier = new Set(
                        Array.from(rejectedPairs).filter(key => key.startsWith(`${r.cashier.id}-`)).map(key => parseInt(key.split('-')[1]))
                      );
                      const availableBanks = unmatchedBankForSuggestions.filter(
                        b => !rejectedForThisCashier.has(b.id) && !savedBankIds.has(b.id) && b.type === r.cashier.type
                      );
                      const suggestions = isExpanded ? getSuggestions(r.cashier, availableBanks, rejectedPairs, savedKeys, claimedNames, suggestionOwnerMap, nameAliases) : [];
                      const isFromHeld = !!r.cashier._fromHeld;
                      const visaNum = extractVisaNumber(r.cashier.rawName);
                      return (
                        <React.Fragment key={i}>
                          <tr className={`border-t transition-colors ${isExpanded ? "bg-blue-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : r.reason === "اختلاف في الاسم" ? "bg-amber-50/30 hover:bg-amber-50/60" : "hover:bg-muted/20"}`}>
                            <td className="px-3 py-2.5 font-medium">
                              {r.cashier.name}
                              {isFromHeld && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                              {visaNum && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300">فيزا {visaNum}</span>}
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.cashier.type}</td>
                            <td className="px-3 py-2.5 font-mono font-semibold text-red-700">{fmtNum(r.cashier.amount)}</td>
                            <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                            <td className="px-3 py-2.5">{r.cashier.rawName || "—"}</td>
                            <td className="px-3 py-2.5">
                              <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${r.reason === "اختلاف في الاسم" ? "bg-amber-100 text-amber-800" : "bg-red-100 text-red-800"}`}>{r.reason}</span>
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => setExpandedUnmatchedCashier(isExpanded ? null : r.cashier.id)}
                                className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${isExpanded ? "bg-blue-100 text-blue-700" : "bg-muted hover:bg-muted/80 text-muted-foreground"}`}>
                                <Sparkles className="w-3.5 h-3.5"/>{isExpanded ? "إخفاء" : "اقتراح بديل"}
                              </button>
                            </td>
                            <td className="px-3 py-2.5">
                              <div className="flex flex-col gap-1">
                                <button onClick={() => handleHoldCashier(r.cashier)}
                                  title="علّق هذه الفاتورة لمطابقتها لاحقاً"
                                  className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors w-fit">
                                  <CreditCard className="w-3.5 h-3.5"/>علّق
                                </button>
                                {visaNum && (
                                  <button onClick={() => handleMoveToVisa(r.cashier, "unmatched")}
                                    title="نقل هذه الفاتورة للفيزا"
                                    className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-teal-600 text-white border border-teal-600 hover:bg-teal-700 transition-colors w-fit">
                                    <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr className="bg-blue-50 border-t border-blue-100">
                              <td colSpan={6} className="px-4 py-3">
                                {suggestions.length === 0 ? (
                                  <p className="text-xs text-muted-foreground">
                                    {availableBanks.length === 0 ? "جميع الحوالات البنكية المتاحة تم رفضها أو حفظها لهذا الكاشير" : "لم يتم إيجاد اقتراحات مشابهة من البنك (لا الاسم ولا المبلغ متطابقان)."}
                                  </p>
                                ) : (
                                  <div className="space-y-2">
                                    <p className="text-xs font-medium text-blue-700 mb-2">أقرب الحوالات البنكية المشابهة (مرتبة حسب قوة الاقتراح):</p>
                                    {suggestions.map(({ bank, score, matchType, amountDiff, reason }) => {
                                      const matchTypeLabel = matchType === "firstSecond" ? "👤 الأول+الثاني" :
                                                            matchType === "fourthName" ? "👤 الرابع" :
                                                            matchType === "exact" ? "تطابق تام" :
                                                            matchType === "typo" ? "✏️ اسم تقريبي" :
                                                            matchType === "amount_only" ? "مبلغ فقط (ضعيف)" : "غير محدد";
                                      const weak = matchType === "amount_only";
                                      return (
                                        <div key={bank.id} className={`flex items-center gap-3 p-2.5 bg-white rounded-lg border ${weak ? "border-gray-200" : "border-blue-200"}`}>
                                          <div className="flex-1 text-xs">
                                            <span className="font-medium">{bank.description}</span>
                                            <span className={`text-xs font-semibold mr-2 ${bank.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{bank.type}</span>
                                            <span className="text-muted-foreground mr-2">{fmtNum(bank.rawAmount)}</span>
                                            <span className={`mr-1 ${weak ? "text-gray-400" : "text-blue-500"}`}>({(score * 100).toFixed(0)}% تطابق)</span>
                                            <span className="text-gray-400 mr-1">| {matchTypeLabel}</span>
                                            {amountDiff > 0.01 && <span className="text-orange-500 mr-1">| فرق: {fmtNum(amountDiff)}</span>}
                                            <div className="text-[10px] text-gray-400 mt-0.5">{reason}</div>
                                          </div>
                                          <button onClick={() => handleAcceptSuggestion(r.cashier, bank)}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                            <Check className="w-3.5 h-3.5"/>قبول
                                          </button>
                                          <button onClick={() => handleRejectSuggestion(r.cashier, bank)}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                            <X className="w-3.5 h-3.5"/>رفض
                                          </button>
                                          <button onClick={() => handleHoldBank(bank)}
                                            title="علّق حوالة البنك هذه لمطابقتها لاحقاً"
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-xs hover:bg-purple-100 transition-colors">
                                            <CreditCard className="w-3.5 h-3.5"/>علّق
                                          </button>
                                        </div>
                                      );
                                    })}
                                    <p className="text-xs text-muted-foreground mt-1">
                                      أو <button onClick={() => { setPage("manual"); }} className="text-blue-600 hover:underline">اذهب للمطابقة اليدوية</button> لربط يدوي كامل
                                    </p>
                                  </div>
                                )}
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                    {!visibleUCashierRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد سجلات تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
              )}

              {tab === "uBank" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع الحساب المسجل</th>
                    <th className="px-3 py-2.5 text-right">السبب</th>
                    <th className="px-3 py-2.5 text-right">تعليق</th>
                  </tr></thead>
                  <tbody>
                    {visibleUBankRows.map((r: any, i: number) => {
                      const typeColor = r.bank.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const isFromHeld = !!r.bank._fromHeld;
                      return (
                        <tr key={i} className={`border-t transition-colors ${isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 text-muted-foreground">
                            {r.bank.description}
                            {isFromHeld && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                          </td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.bank.type}</td>
                          <td className="px-3 py-2.5 font-mono font-semibold text-orange-700">{fmtNum(r.bank.rawAmount)}</td>
                          <td className="px-3 py-2.5">{r.bank.accountType || "—"}</td>
                          <td className="px-3 py-2.5">
                            <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${r.reason === "اختلاف في الاسم" ? "bg-amber-100 text-amber-800" : "bg-gray-100 text-gray-700"}`}>{r.reason}</span>
                          </td>
                          <td className="px-3 py-2.5">
                            <button onClick={() => handleHoldBank(r.bank)}
                              title="علّق هذه الحوالة لمطابقتها لاحقاً"
                              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors">
                              <CreditCard className="w-3.5 h-3.5"/>علّق
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                    {!visibleUBankRows.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد سجلات تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
