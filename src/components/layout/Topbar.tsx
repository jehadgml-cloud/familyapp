'use client'

import { useState, useEffect } from 'react'
import Link from 'next/link'
import { useRouter } from 'next/navigation'
import { Menu, X, LayoutDashboard, FileSpreadsheet, Settings, LogOut } from 'lucide-react'
import { createClient } from '@/lib/supabase/client'

export default function Topbar() {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)
  const [profile, setProfile] = useState<{ full_name: string; role: string } | null>(null)
  const router = useRouter()
  const supabase = createClient()

  useEffect(() => {
    async function loadProfile() {
      const { data: { user } } = await supabase.auth.getUser()
      if (user) {
        const { data } = await supabase
          .from('profiles')
          .select('full_name, role')
          .eq('id', user.id)
          .single()
        if (data) setProfile(data)
      }
    }
    loadProfile()
  }, [supabase])

  const handleLogout = async () => {
    await supabase.auth.signOut()
    router.push('/login')
    router.refresh()
  }

  const toggleMobileMenu = () => setMobileMenuOpen(!mobileMenuOpen)

  const menuItems = [
    { name: 'لوحة التحكم', href: '/dashboard', icon: LayoutDashboard },
    { name: 'التقارير والكشوفات', href: '/reports', icon: FileSpreadsheet },
  ]

  const canManageUsers = profile?.role === 'Super Administrator' || profile?.role === 'Administrator'

  return (
    <header className="sticky top-0 z-30 flex h-16 w-full items-center justify-between border-b border-slate-100 bg-white px-6 lg:px-8">
      {/* Page Title / Info */}
      <div className="flex items-center gap-3">
        <h2 className="text-sm font-bold text-slate-800 hidden md:block">
          بوابة العشيرة الإلكترونية لإدارة وتثبيت سندات التحصيل
        </h2>
      </div>

      {/* Hamburger Switcher on Mobile/Tablet */}
      <div className="flex lg:hidden">
        <button
          onClick={toggleMobileMenu}
          className="rounded-lg p-2 text-slate-600 hover:bg-slate-50 hover:text-slate-900"
          aria-label="Toggle Menu"
        >
          {mobileMenuOpen ? <X className="h-6 w-6" /> : <Menu className="h-6 w-6" />}
        </button>
      </div>

      {/* Desktop user welcome tag */}
      {profile && (
        <span className="hidden lg:inline-block text-xs font-bold text-slate-600 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-100">
           {profile.full_name} ({profile.role})
        </span>
      )}

      {/* Mobile Drawer Overlay */}
      {mobileMenuOpen && (
        <div className="absolute left-0 right-0 top-16 z-50 flex flex-col bg-white border-b border-slate-100 px-6 py-4 shadow-lg lg:hidden space-y-4">
          <nav className="flex flex-col space-y-1">
            {menuItems.map((item) => {
              const Icon = item.icon
              return (
                <Link
                  key={item.href}
                  href={item.href}
                  onClick={() => setMobileMenuOpen(false)}
                  className="flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-50 hover:text-emerald-500 transition-all text-right"
                >
                  <Icon className="h-5 w-5" />
                  <span>{item.name}</span>
                </Link>
              )
            })}

            {canManageUsers && (
              <Link
                href="/settings"
                onClick={() => setMobileMenuOpen(false)}
                className="flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-50 hover:text-emerald-500 transition-all text-right"
              >
                <Settings className="h-5 w-5" />
                <span>إدارة المستخدمين</span>
              </Link>
            )}
          </nav>

          <div className="border-t border-slate-50 pt-4 flex flex-col gap-3">
            {profile && (
              <div className="px-4 text-right">
                <p className="text-xs font-bold text-slate-800">{profile.full_name}</p>
                <p className="text-[10px] text-emerald-600 font-semibold mt-0.5">{profile.role}</p>
              </div>
            )}
            <button
              onClick={handleLogout}
              className="flex w-full items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold text-red-600 hover:bg-red-50 transition-all text-right"
            >
              <LogOut className="h-5 w-5" />
              <span>تسجيل الخروج</span>
            </button>
          </div>
        </div>
      )}
    </header>
  )
}
