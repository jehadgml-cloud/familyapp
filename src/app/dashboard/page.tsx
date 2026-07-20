'use client'

import { useEffect, useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import MainWrapper from '@/components/layout/MainWrapper'
import { 
  FileText, CheckCircle2, AlertCircle, FileEdit, Clock, 
  TrendingUp, Users, Building, ShieldCheck 
} from 'lucide-react'
import Link from 'next/link'

interface DashboardStats {
  total: number
  approved: number
  pending: number
  drafts: number
  rejected: number
}

export default function DashboardPage() {
  const [stats, setStats] = useState<DashboardStats>({
    total: 0,
    approved: 0,
    pending: 0,
    drafts: 0,
    rejected: 0,
  })
  const [loading, setLoading] = useState(true)
  const [recentActivities, setRecentActivities] = useState<any[]>([])
  const supabase = createClient()

  useEffect(() => {
    async function loadDashboardStats() {
      try {
        // Query reports metadata
        const { data: reports, error } = await supabase
          .from('reports')
          .select('status, id, report_number, report_title, created_at')

        if (error) throw error

        if (reports) {
          const fetchedStats = reports.reduce(
            (acc: DashboardStats, curr: any) => {
              acc.total++
              if (curr.status === 'Approved') acc.approved++
              else if (curr.status === 'Under Review' || curr.status === 'Submitted') acc.pending++
              else if (curr.status === 'Draft') acc.drafts++
              else if (curr.status === 'Rejected' || curr.status === 'Returned for Correction') acc.rejected++
              return acc
            },
            { total: 0, approved: 0, pending: 0, drafts: 0, rejected: 0 }
          )
          setStats(fetchedStats)
          setRecentActivities(reports.slice(0, 5))
        }
      } catch (err) {
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    loadDashboardStats()
  }, [supabase])

  const kpiItems = [
    { label: 'إجمالي التقارير', value: stats.total, color: 'text-blue-600 bg-blue-50 border-blue-100', icon: FileText },
    { label: 'المعتمدة والمدققة', value: stats.approved, color: 'text-emerald-600 bg-emerald-50 border-emerald-100', icon: CheckCircle2 },
    { label: 'قيد المراجعة والاعتماد', value: stats.pending, color: 'text-amber-600 bg-amber-50 border-amber-100', icon: Clock },
    { label: 'المسودات المؤقتة', value: stats.drafts, color: 'text-slate-600 bg-slate-100 border-slate-200', icon: FileEdit },
    { label: 'المرفوضة أو المعادة', value: stats.rejected, color: 'text-red-600 bg-red-50 border-red-100', icon: AlertCircle },
  ]

  return (
    <MainWrapper>
      <div className="flex flex-col gap-6 text-right">
        {/* Welcome header banner */}
        <div className="rounded-2xl bg-gradient-to-r from-emerald-500 to-teal-700 p-6 md:p-8 text-white shadow-premium">
          <h1 className="text-2xl font-bold font-cairo">أهلاً بك في البوابة الإلكترونية لإدارة ومتابعة الصندوق</h1>
          <p className="mt-2 text-sm text-emerald-50 opacity-90 font-cairo">
            تتيح لك المنصة تسجيل سندات التحصيل الشهري للمشتركين، تسجيل الصرفيات العامة، وإجراء دورات المراجعة والاعتماد.
          </p>
        </div>

        {loading ? (
          <div className="flex justify-center items-center py-20">
            <div className="h-10 w-10 animate-spin rounded-full border-4 border-slate-200 border-t-emerald-500"></div>
          </div>
        ) : (
          <>
            {/* KPI Cards Grid */}
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-5">
              {kpiItems.map((item, idx) => {
                const Icon = item.icon
                return (
                  <div key={idx} className={`p-5 rounded-2xl border bg-white shadow-premium flex flex-col justify-between gap-3 text-right`}>
                    <div className="flex items-center justify-between">
                      <span className={`p-2.5 rounded-xl border ${item.color}`}>
                        <Icon className="h-5 w-5" />
                      </span>
                      <span className="text-2xl font-black text-slate-800">{item.value}</span>
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-slate-500">{item.label}</p>
                    </div>
                  </div>
                )
              })}
            </div>

            {/* Quick Actions and Recent Stream */}
            <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
              {/* Quick Actions Panel */}
              <div className="lg:col-span-1 rounded-2xl bg-white border border-slate-100 p-5 shadow-premium text-right">
                <h3 className="text-sm font-bold text-slate-800 border-b border-slate-50 pb-3 mb-4">إجراءات سريعة</h3>
                <div className="flex flex-col gap-3">
                  <Link href="/reports/create" className="flex items-center justify-center p-3 rounded-xl border border-emerald-500 text-emerald-600 hover:bg-emerald-50 transition-all font-bold text-sm">
                    🆕 إنشاء تقرير أو كشف مالي جديد
                  </Link>
                  <Link href="/reports" className="flex items-center justify-center p-3 rounded-xl bg-slate-900 text-white hover:bg-slate-800 transition-all font-bold text-sm">
                    📑 تصفح الأرشيف والكشوف السابقة
                  </Link>
                </div>
              </div>

              {/* Recent Active Reports list */}
              <div className="lg:col-span-2 rounded-2xl bg-white border border-slate-100 p-5 shadow-premium text-right">
                <h3 className="text-sm font-bold text-slate-800 border-b border-slate-50 pb-3 mb-4">أحدث النشاطات والكشوفات</h3>
                
                {recentActivities.length === 0 ? (
                  <p className="text-sm text-slate-500 py-6 text-center">لا توجد تغييرات أو تقارير مضافة حالياً.</p>
                ) : (
                  <div className="divide-y divide-slate-50">
                    {recentActivities.map((report) => (
                      <div key={report.id} className="py-3 flex justify-between items-center text-sm">
                        <span className="text-xs text-slate-400">
                          {new Date(report.created_at).toLocaleDateString('ar-EG', { day: 'numeric', month: 'long', year: 'numeric' })}
                        </span>
                        <div className="flex items-center gap-3">
                          <Link href={`/reports/${report.id}`} className="font-bold text-slate-700 hover:text-emerald-500">
                            {report.report_title} ({report.report_number})
                          </Link>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </MainWrapper>
  )
}
