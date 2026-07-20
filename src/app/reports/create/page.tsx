'use client'

import { useState, useEffect } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter } from 'next/navigation'
import MainWrapper from '@/components/layout/MainWrapper'
import { ArrowRight, Save } from 'lucide-react'
import Link from 'next/link'

// Quick static dump of members matching default Palestinian family ledger starting parameters
const INITIAL_FAMILY_MEMBERS = [
  { id: 1, name: 'علاء نبيل يوسف اطفيحه', payments: {} },
  { id: 2, name: 'يوسف علاء نبيل اطفيحه', payments: {} },
  { id: 3, name: 'احمد علاء نبيل اطفيحه', payments: {} },
  { id: 4, name: 'محمد علاء نبيل اطفيحه', payments: {} },
  { id: 5, name: 'ايوب علاء نبيل اطفيحه', payments: {} },
  { id: 6, name: 'ادم علاء نبيل اطفيحه', payments: {} },
  { id: 7, name: 'سليمان نبيل يوسف اطفيحه', payments: {} },
  { id: 8, name: 'نبيل سليمان نبيل اطفيحه', payments: {} },
  { id: 9, name: 'اسماعيل نبيل يوسف اطفيحه', payments: {} },
  { id: 10, name: 'يوسف سفيان يوسف اطفيحه', payments: {} }
]

export default function CreateReportPage() {
  const [title, setTitle] = useState('')
  const [type, setType] = useState('كشف التحصيل الشهري')
  const [period, setPeriod] = useState('')
  const [loading, setLoading] = useState(false)
  const [errorMsg, setErrorMsg] = useState('')
  const [profile, setProfile] = useState<any>(null)
  
  const router = useRouter()
  const supabase = createClient()

  useEffect(() => {
    async function loadProfile() {
      const { data: { user } } = await supabase.auth.getUser()
      if (user) {
        const { data } = await supabase
          .from('profiles')
          .select('*')
          .eq('id', user.id)
          .single()
        setProfile(data)
      }
    }
    loadProfile()
  }, [supabase])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setErrorMsg('')

    if (!profile) {
      setErrorMsg('غير مصرح لك بإنشاء تقارير. يرجى تسجيل الدخول مجدداً.')
      setLoading(false)
      return
    }

    try {
      const uniqueNum = `ATF-${new Date().getFullYear()}-${type === 'كشف التحصيل الشهري' ? 'FIN' : 'ADM'}-${Math.floor(1000 + Math.random() * 9000)}`;
      
      const reportPayload = {
        report_number: uniqueNum,
        report_title: title,
        report_type: type,
        department_id: profile.department_id || '811be074-ceee-40ba-81df-85e656ad4e81', // default fallback to finance
        reporting_period: period,
        status: 'Draft',
        report_data: {
          members: INITIAL_FAMILY_MEMBERS,
          expenses: [],
          month_state: { [period]: false } // false implies unpaid/incomplete by default
        },
        created_by: profile.id,
        version_number: 1,
      }

      const { data, error } = await supabase
        .from('reports')
        .insert(reportPayload)
        .select()
        .single()

      if (error) throw error

      if (data) {
        router.push(`/reports/${data.id}/edit`)
      }
    } catch (err: any) {
      setErrorMsg(err.message || 'فشلت عملية إنشاء المسودة. تفقد البيانات المدخلة.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <MainWrapper>
      <div className="flex flex-col gap-6 text-right">
        {/* Back Link */}
        <div className="flex items-center justify-between">
          <Link href="/reports" className="flex items-center gap-1.5 text-slate-500 font-bold hover:text-slate-800 text-sm">
            <ArrowRight className="h-4 w-4" />
            <span>العودة للأرشيف</span>
          </Link>
          <h1 className="text-xl font-bold text-slate-800 font-cairo">إنشاء كشف جديد</h1>
        </div>

        {/* Wizard Form */}
        <div className="bg-white p-6 md:p-8 rounded-2xl border border-slate-100 shadow-premium max-w-xl mx-auto w-full">
          {errorMsg && (
            <div className="p-3 text-sm text-red-600 bg-red-50 border border-red-100 rounded-lg text-right mb-4">
              ⚠️ {errorMsg}
            </div>
          )}

          <form onSubmit={handleSubmit} className="space-y-6">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">
                عنوان الكشف / التقرير
              </label>
              <input
                type="text"
                required
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="مثال: كشف تحصيل عائلة آل اطفيحة لشهر أبريل"
                className="w-full text-right px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-emerald-500"
              />
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">
                نوع المستند المالي
              </label>
              <select
                value={type}
                onChange={(e) => setType(e.target.value)}
                className="w-full text-right px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-emerald-500"
              >
                <option value="كشف التحصيل الشهري">كشف التحصيل الشهري (صندوق الجباية)</option>
                <option value="بيان المصاريف">بيان المصاريف (الإنفاق العام)</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">
                الفترة المالية التابعة
              </label>
              <input
                type="text"
                required
                value={period}
                onChange={(e) => setPeriod(e.target.value)}
                placeholder="مثال: أبريل 2026"
                className="w-full text-right px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-emerald-500"
              />
            </div>

            <div>
              <button
                type="submit"
                disabled={loading}
                className="flex w-full items-center justify-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-white font-bold py-3 px-5 rounded-xl text-sm transition-all shadow-premium border-0 disabled:opacity-50"
              >
                <Save className="h-4 w-4" />
                <span>{loading ? 'جاري بدء المسودة...' : '💾 إنشاء مسودة والذهاب للمحرر'}</span>
              </button>
            </div>
          </form>
        </div>
      </div>
    </MainWrapper>
  )
}
