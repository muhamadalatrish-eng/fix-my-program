export type SafetyAuditKind = "blocked" | "safe";

export interface SafetyAuditEntry {
  id: string;
  kind: SafetyAuditKind;
  occurredAt: string;
  screen: string;
  action: string;
  detail?: string;
  reviewedAt?: string;
}

export const SAFETY_LOCK_KEY = "recon_safety_lock";
export const SAFETY_AUDIT_KEY = "recon_safety_audit";
export const SAFETY_AUDIT_LIMIT = 5000;

export function readSafetyAudit(): SafetyAuditEntry[] {
  try {
    const value = JSON.parse(localStorage.getItem(SAFETY_AUDIT_KEY) || "[]");
    if (!Array.isArray(value)) return [];
    return value.filter((entry): entry is SafetyAuditEntry =>
      entry && typeof entry.id === "string" &&
      (entry.kind === "blocked" || entry.kind === "safe") &&
      typeof entry.occurredAt === "string" &&
      typeof entry.screen === "string" &&
      typeof entry.action === "string"
    );
  } catch {
    return [];
  }
}

export function writeSafetyAudit(entries: SafetyAuditEntry[]): void {
  try {
    localStorage.setItem(SAFETY_AUDIT_KEY, JSON.stringify(entries));
  } catch (error) {
    console.error("تعذر حفظ سجل خط الأمان:", error);
  }
}

export function verifySafetyPassword(password: string): boolean {
  try {
    const raw = localStorage.getItem("recon_credentials");
    const credentials = raw ? JSON.parse(raw) as { password?: unknown } : null;
    const expected = typeof credentials?.password === "string" ? credentials.password : "admin123";
    return password === expected;
  } catch {
    return password === "admin123";
  }
}
