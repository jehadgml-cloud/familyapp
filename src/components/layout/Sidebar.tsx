'use client'

import Link from 'next/link'
import { usePathname, useRouter } from 'next/navigation'
import { LayoutDashboard, FileSpreadsheet, Settings, LogOut, Users, UserCheck } from 'lucide-react'
import { createClient } from '@/lib/supabase/client'
import { useEffect, useState } from 'react'

export default function Sidebar() {
  const pathname = usePathname()
  const router = useRouter()
  const supabase = createClient()
  const [profile, setProfile] = useState<{ full_name: string; role: string } | null>(null)

  useEffect(() => {
    async function loadProfile() {
      const { data: { user } } = await supabase.auth.getUser()
      if (user) {
        const { data } = await supabase
          .from('profiles')
          .select('full_name, role')
          .eq('id', user.id)
          .single()
        if (data) {
          setProfile(data)
        }
      }
    }
    loadProfile()
  }, [supabase])

  const handleLogout = async () => {
    await supabase.auth.signOut()
    router.push('/login')
    router.refresh()
  }

  const menuItems = [
    { name: 'لوحة التحكم', href: '/dashboard', icon: LayoutDashboard },
    { name: 'التقارير والكشوفات', href: '/reports', icon: FileSpreadsheet },
  ]

  // Add settings link for Administative roles
  const canManageUsers = profile?.role === 'Super Administrator' || profile?.role === 'Administrator'

  return (
    <aside className="fixed right-0 top-0 z-40 hidden h-screen w-64 border-l border-slate-100 bg-white p-5 lg:flex lg:flex-col justify-between">
      <div className="space-y-6">
        <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-emerald-50 text-emerald-600 font-bold">
            ع.ا
          </div>
          <div>
            <h1 className="text-sm font-extrabold text-slate-800">صندوق عائلة اطفيحة</h1>
            {profile && (
              <span className="inline-block text-[10px] bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded font-bold mt-1">
                {profile.role}
              </span>
            )}
          </div>
        </div>

        <nav className="space-y-1">
          {menuItems.map((item) => {
            const Icon = item.icon
            const isActive = pathname === item.href
            return (
              <Link
                key={item.href}
                href={item.href}
                className={`flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition-all ${
                  isActive
                    ? 'bg-emerald-500 text-white'
                    : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
                }`}
              >
                <Icon className="h-5 w-5" />
                <span>{item.name}</span>
              </Link>
            )
          })}

          {canManageUsers && (
            <Link
              href="/settings"
              className={`flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition-all ${
                pathname === '/settings'
                  ? 'bg-emerald-500 text-white'
                  : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
              }`}
            >
              <Settings className="h-5 w-5" />
              <span>إدارة المستخدمين</span>
            </Link>
          )}
        </nav>
      </div>

      <div className="space-y-4">
        {profile && (
          <div className="rounded-xl bg-slate-50 p-3 text-right">
            <p className="text-xs font-bold text-slate-800">{profile.full_name}</p>
            <p className="text-[10px] text-slate-500 mt-0.5">مرحباً بك</p>
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
    </aside>
  )
}
