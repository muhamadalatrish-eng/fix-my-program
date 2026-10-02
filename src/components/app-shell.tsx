"use client"

import { useEffect, useRef, useState } from "react"
import { ChevronLeft } from "lucide-react"
import { LoginScreen } from "./login-screen"
import { AppSidebar, type ViewId } from "./app-sidebar"
import { ProfileView } from "./profile-view"
import { SettingsView } from "./settings-view"
import { SafetyLockOverlay, type BlockedAction } from "./safety-lock-overlay"
import App from "./reconciliation-app"
import { AccountStatementProcessor } from "./account-statement-processor"
import { readSafetyAudit, SAFETY_AUDIT_LIMIT, SAFETY_LOCK_KEY, type SafetyAuditEntry, verifySafetyPassword, writeSafetyAudit } from "../lib/safety-audit"

const VIEW_LABELS: Record<ViewId, string> = {
  recon: "أدوات التسوية",
  assistant: "مساعد المطابقة",
  manual: "المطابقة اليدوية",
  report: "تقرير الفروقات",
  splitTool: "تجزئة الحوالات",
  batchCutoff: "تجميع فترة التعطل",
  clearing: "تسكير السندات",
  processor: "معالج الكشوفات",
  profile: "الملف الشخصي",
  settings: "الإعدادات",
}

export function AppShell() {
  const [ready, setReady] = useState(false)
  const [username, setUsername] = useState<string | null>(null)
  const [view, setView] = useState<ViewId>("recon")
  const [collapsed, setCollapsed] = useState(false)
  const [hovering, setHovering] = useState(false)
  const [safetyLocked, setSafetyLocked] = useState(false)
  const [auditEntries, setAuditEntries] = useState<SafetyAuditEntry[]>([])
  const auditEntriesRef = useRef<SafetyAuditEntry[]>([])

  const appendAuditEntries = (entries: Array<Omit<SafetyAuditEntry, "id" | "occurredAt">>) => {
    if (!entries.length) return
    const stamped = entries.map(entry => ({
      ...entry,
      id: `audit-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`,
      occurredAt: new Date().toISOString(),
    }))
    const combined = [...stamped, ...auditEntriesRef.current]
    const pending = combined.filter(entry => entry.kind === "blocked" && !entry.reviewedAt)
    const recent = combined
      .filter(entry => entry.kind !== "blocked" || !!entry.reviewedAt)
      .slice(0, Math.max(0, SAFETY_AUDIT_LIMIT - pending.length))
    const next = [...pending, ...recent].sort((a, b) => b.occurredAt.localeCompare(a.occurredAt))
    auditEntriesRef.current = next
    setAuditEntries(next)
    writeSafetyAudit(next)
  }

  const recordBlockedAction = (attempt: BlockedAction) => {
    appendAuditEntries([{ ...attempt, kind: "blocked" }])
  }

  const recordSafeActivities = (activities: string[], screen: string) => {
    appendAuditEntries(activities.map(detail => ({
      kind: "safe" as const,
      screen,
      action: "مطابقة آمنة محفوظة",
      detail,
    })))
  }

  useEffect(() => {
    try {
      const raw = localStorage.getItem("recon_auth")
      if (raw) setUsername((JSON.parse(raw) as { username: string }).username)
    } catch {
      /* ignore */
    }
    try {
      setCollapsed(localStorage.getItem("recon_sidebar_collapsed") === "1")
    } catch {
      /* ignore */
    }
    const savedAudit = readSafetyAudit()
    auditEntriesRef.current = savedAudit
    setAuditEntries(savedAudit)
    try {
      setSafetyLocked(localStorage.getItem(SAFETY_LOCK_KEY) === "1")
    } catch {
      setSafetyLocked(false)
    }
    setReady(true)
  }, [])

  useEffect(() => {
    if (!auditEntries.some(entry => entry.kind === "blocked" && !entry.reviewedAt)) return
    const remindBeforeClose = (event: BeforeUnloadEvent) => {
      event.preventDefault()
      event.returnValue = "توجد حركات غير آمنة لم تتم مراجعتها."
    }
    window.addEventListener("beforeunload", remindBeforeClose)
    return () => window.removeEventListener("beforeunload", remindBeforeClose)
  }, [auditEntries])

  const activateSafetyLock = () => {
    try {
      localStorage.setItem(SAFETY_LOCK_KEY, "1")
    } catch {
      /* Keep the in-memory lock active if storage is unavailable. */
    }
    setSafetyLocked(true)
  }

  const unlockSafety = (password: string) => {
    if (!verifySafetyPassword(password)) {
      recordBlockedAction({ screen: "خط الأمان", action: "محاولة فك القفل", detail: "كلمة المرور غير صحيحة." })
      return false
    }
    try {
      localStorage.removeItem(SAFETY_LOCK_KEY)
    } catch {
      /* The in-memory lock can still be removed for this session. */
    }
    setSafetyLocked(false)
    return true
  }

  const markAuditReviewed = () => {
    const reviewedAt = new Date().toISOString()
    const next = auditEntriesRef.current.map(entry =>
      entry.kind === "blocked" && !entry.reviewedAt ? { ...entry, reviewedAt } : entry
    )
    auditEntriesRef.current = next
    setAuditEntries(next)
    writeSafetyAudit(next)
  }

  const toggleCollapsed = () => {
    setCollapsed((prev) => {
      const next = !prev
      try {
        localStorage.setItem("recon_sidebar_collapsed", next ? "1" : "0")
      } catch {
        /* ignore */
      }
      if (next) setHovering(false)
      return next
    })
  }

  const handleLogout = () => {
    localStorage.removeItem("recon_auth")
    setUsername(null)
    setView("recon")
  }

  if (!ready) return null
  if (!username) return <LoginScreen onLogin={setUsername} />

  const sidebarVisible = !collapsed || hovering
  const pendingAuditCount = auditEntries.filter(entry => entry.kind === "blocked" && !entry.reviewedAt).length
  const onSafeActivities = (activities: string[]) => recordSafeActivities(activities, VIEW_LABELS[view])

  return (
    <div dir="rtl" className="app-shell flex h-[100dvh] max-h-[100dvh] overflow-hidden bg-background">
      {collapsed && (
        <>
          <div className="fixed inset-y-0 right-0 z-30 w-4" onMouseEnter={() => setHovering(true)} />
          {!hovering && (
            <button
              type="button"
              onClick={() => setHovering(true)}
              aria-label="إظهار القائمة"
              title="إظهار القائمة"
              className="fixed right-0 top-1/2 z-30 -translate-y-1/2 rounded-l-lg border border-sidebar-border bg-sidebar px-1 py-3 text-sidebar-foreground shadow-md transition-opacity hover:opacity-100"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
          )}
        </>
      )}
      <div
        onMouseEnter={() => collapsed && setHovering(true)}
        onMouseLeave={() => collapsed && setHovering(false)}
        className={collapsed
          ? `fixed inset-y-0 right-0 z-40 transition-transform duration-200 ${sidebarVisible ? "translate-x-0" : "translate-x-full"}`
          : "relative shrink-0"}
      >
        <AppSidebar
          active={view}
          onNavigate={setView}
          username={username}
          onLogout={handleLogout}
          collapsed={collapsed}
          onToggleCollapse={toggleCollapsed}
        />
      </div>
      <div className="recon-scope flex-1 overflow-y-auto">
        {view === "recon" && <App key="recon" initialPage="recon2" onSafeActivities={onSafeActivities} />}
        {view === "assistant" && <App key="assistant" initialPage="assist" onSafeActivities={onSafeActivities} />}
        {view === "manual" && <App key="manual" initialPage="manual" onSafeActivities={onSafeActivities} />}
        {view === "report" && <App key="report" initialPage="report" onSafeActivities={onSafeActivities} />}
        {view === "splitTool" && <App key="splitTool" initialPage="splitTool" onSafeActivities={onSafeActivities} />}
        {view === "batchCutoff" && <App key="batchCutoff" initialPage="batchCutoff" onSafeActivities={onSafeActivities} />}
        {view === "clearing" && <App key="clearing" initialPage="clearing" onSafeActivities={onSafeActivities} />}
        {view === "processor" && <AccountStatementProcessor />}
        {view === "profile" && <ProfileView username={username} onUsernameChange={setUsername} />}
        {view === "settings" && <SettingsView
          safetyLocked={safetyLocked}
          auditEntries={auditEntries}
          onActivateSafetyLock={activateSafetyLock}
          onMarkAuditReviewed={markAuditReviewed}
        />}
      </div>
      {pendingAuditCount > 0 && !safetyLocked && (
        <button
          type="button"
          onClick={() => setView("settings")}
          className="fixed left-4 top-4 z-50 flex items-center gap-2 rounded-lg border border-red-300 bg-red-50 px-3 py-2 text-xs font-semibold text-red-800 shadow-lg hover:bg-red-100"
        >
          <span className="h-2 w-2 rounded-full bg-red-600" />
          {pendingAuditCount} محاولة غير آمنة بانتظار المراجعة
        </button>
      )}
      {safetyLocked && <SafetyLockOverlay activeScreen={VIEW_LABELS[view]} onAttempt={recordBlockedAction} onUnlock={unlockSafety} />}
    </div>
  )
}
