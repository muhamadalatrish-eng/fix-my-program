"use client"

import { Trash2, Database } from "lucide-react"
import { useState } from "react"

export function SettingsView() {
  const [cleared, setCleared] = useState(false)

  const clearReconData = () => {
    const keep = new Set(["recon_auth", "recon_credentials", "recon_sidebar_collapsed"])
    try {
      Object.keys(localStorage)
        .filter((k) => !keep.has(k))
        .forEach((k) => localStorage.removeItem(k))
      indexedDB.deleteDatabase("recon_store")
    } catch {
      /* ignore */
    }
    setCleared(true)
    setTimeout(() => setCleared(false), 2500)
  }

  return (
    <div dir="rtl" className="mx-auto min-h-[100dvh] max-w-3xl p-4 sm:p-8">
      <header className="mb-6">
        <div className="mb-2 text-xs font-semibold text-primary">التحكم والخصوصية</div>
        <h1 className="text-3xl font-bold">الإعدادات</h1>
        <p className="mt-2 text-sm text-muted-foreground">إدارة البيانات المحفوظة في هذا المتصفح</p>
      </header>

      <div className="space-y-4">
          <section className="rounded-2xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-3 flex items-center gap-2 text-sm font-semibold">
            <Database className="h-4 w-4" />
            البيانات المحفوظة
          </h2>
          <p className="mb-4 text-xs text-muted-foreground">
            تحذف عمليات التسوية والملفات المحفوظة محلياً في هذا المتصفح. لن يؤثر ذلك على بيانات الدخول.
          </p>
          <button
            onClick={clearReconData}
            className="flex items-center gap-2 rounded-lg border border-destructive/30 px-4 py-2.5 text-sm font-medium text-destructive transition-colors hover:bg-destructive/10"
          >
            <Trash2 className="h-4 w-4" />
            {cleared ? "تم مسح البيانات" : "مسح بيانات التسوية"}
          </button>
        </section>
      </div>
    </div>
  )
}
