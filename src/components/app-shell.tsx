"use client"

import { useEffect, useState } from "react"
import { ChevronLeft } from "lucide-react"
import { LoginScreen } from "./login-screen"
import { AppSidebar, type ViewId } from "./app-sidebar"
import { ProfileView } from "./profile-view"
import { SettingsView } from "./settings-view"
import App from "./reconciliation-app"

export function AppShell() {
  const [ready, setReady] = useState(false)
  const [username, setUsername] = useState<string | null>(null)
  const [view, setView] = useState<ViewId>("recon")
  const [collapsed, setCollapsed] = useState(false)
  const [hovering, setHovering] = useState(false)

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
    setReady(true)
  }, [])

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

  return (
    <div dir="rtl" className="app-shell flex h-[100dvh] max-h-[100dvh] overflow-hidden bg-background">
      {collapsed && (
        <>
          <div
            className="fixed inset-y-0 right-0 z-30 w-4"
            onMouseEnter={() => setHovering(true)}
          />
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
        className={
          collapsed
            ? `fixed inset-y-0 right-0 z-40 transition-transform duration-200 ${
                sidebarVisible ? "translate-x-0" : "translate-x-full"
              }`
            : "relative shrink-0"
        }
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
        {view === "recon" && <App key="recon" initialPage="recon2" />}
        {view === "assistant" && <App key="assistant" initialPage="assist" />}
        {view === "manual" && <App key="manual" initialPage="manual" />}
        {view === "report" && <App key="report" initialPage="report" />}
        {view === "profile" && <ProfileView username={username} onUsernameChange={setUsername} />}
        {view === "settings" && <SettingsView />}
      </div>
    </div>
  )
}
