'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'
import Image from 'next/image'

export default function LoginPage() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [errorMsg, setErrorMsg] = useState('')
  const [loading, setLoading] = useState(false)
  const router = useRouter()
  const supabase = createClient()

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setErrorMsg('')

    try {
      const { data, error } = await supabase.auth.signInWithPassword({
        email,
        password,
      })

      if (error) {
        setErrorMsg(error.message === 'Invalid login credentials' ? 'خطأ في البريد الإلكتروني أو كلمة المرور' : error.message)
      } else {
        // Redirection handled by middleware, but enforce locally:
        router.push('/dashboard')
        router.refresh()
      }
    } catch (err: any) {
      setErrorMsg('حدث خطأ غير متوقع. يرجى المحاولة لاحقاً.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-50 px-4 py-12 sm:px-6 lg:px-8">
      <div className="w-full max-w-md space-y-8 bg-white p-8 rounded-2xl shadow-premium border border-slate-100">
        <div className="flex flex-col items-center justify-center text-center">
          <div className="relative w-24 h-24 mb-4 overflow-hidden rounded-full border-2 border-emerald-500">
            {/* Fallback to simple styled Arabic initials if image is missing */}
            <div className="flex items-center justify-center w-full h-full bg-emerald-50 text-emerald-600 font-bold text-2xl">
              C.F
            </div>
          </div>
          <h2 className="text-2xl font-extrabold text-slate-900 font-cairo">صندوق عائلة آل اطفيحة</h2>
          <p className="mt-2 text-sm text-slate-500 font-cairo">بوابة التسجيل للجنة المالية والتدقيق</p>
        </div>

        <form className="mt-8 space-y-6" onSubmit={handleLogin}>
          {errorMsg && (
            <div className="p-3 text-sm text-red-600 bg-red-50 border border-red-100 rounded-lg text-right font-cairo">
              ⚠️ {errorMsg}
            </div>
          )}

          <div className="space-y-4 rounded-md shadow-sm">
            <div>
              <label htmlFor="email-address" className="block text-sm font-semibold text-slate-700 text-right mb-1">
                البريد الإلكتروني
              </label>
              <input
                id="email-address"
                name="email"
                type="email"
                autoComplete="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="relative block w-full rounded-xl border border-slate-200 px-4 py-3 text-slate-900 placeholder-slate-400 focus:z-10 focus:border-emerald-500 focus:outline-none focus:ring-emerald-500 text-right text-sm"
                placeholder="mail@example.com"
              />
            </div>

            <div>
              <label htmlFor="password" className="block text-sm font-semibold text-slate-700 text-right mb-1">
                كلمة المرور
              </label>
              <input
                id="password"
                name="password"
                type="password"
                autoComplete="current-password"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="relative block w-full rounded-xl border border-slate-200 px-4 py-3 text-slate-900 placeholder-slate-400 focus:z-10 focus:border-emerald-500 focus:outline-none focus:ring-emerald-500 text-right text-sm"
                placeholder="••••••••"
              />
            </div>
          </div>

          <div className="flex items-center justify-between text-sm">
            <a href="#" className="font-semibold text-emerald-600 hover:text-emerald-500">
              نسيت كلمة المرور؟
            </a>
          </div>

          <div>
            <button
              type="submit"
              disabled={loading}
              className="group relative flex w-full justify-center rounded-xl bg-emerald-500 py-3 text-sm font-bold text-white hover:bg-emerald-600 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:ring-offset-2 transition-all disabled:opacity-50"
            >
              {loading ? 'جاري التحقق...' : 'تسجيل الدخول'}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
