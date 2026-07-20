'use client'

import { useEffect, useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import MainWrapper from '@/components/layout/MainWrapper'
import Link from 'next/link'
import { PlusCircle, Search, Eye, Filter, ChevronLeft, ChevronRight } from 'lucide-react'

export default function ReportsListPage() {
  const [reports, setReports] = useState<any[]>([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')
  const [statusFilter, setStatusFilter] = useState('all')
  const [typeFilter, setTypeFilter] = useState('all')
  const [profile, setProfile] = useState<any>(null)
  
  // Pagination
  const [currentPage, setCurrentPage] = useState(1)
  const itemsPerPage = 8

  const supabase = createClient()

  useEffect(() => {
    async function loadInfo() {
      const { data: { user } } = await supabase.auth.getUser()
      if (user) {
        const { data: prof } = await supabase
          .from('profiles')
          .select('*')
          .eq('id', user.id)
          .single()
        setProfile(prof)
      }
    }
    loadInfo()
  }, [supabase])

  useEffect(() => {
    async function loadReports() {
      setLoading(true)
      try {
        let query = supabase
          .from('reports')
          .select('*, departments(department_name_ar)')
          .eq('is_deleted', false)
          .order('created_at', { ascending: false })

        const { data, error } = await query
        if (error) throw error
        setReports(data || [])
      } catch (err) {
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    loadReports()
  }, [supabase])

  // Client side filters
  const filteredReports = reports.filter((item) => {
    const matchesSearch = 
      item.report_title?.toLowerCase().includes(search.toLowerCase()) ||
      item.report_number?.toLowerCase().includes(search.toLowerCase()) ||
      item.reporting_period?.toLowerCase().includes(search.toLowerCase())

    const matchesStatus = statusFilter === 'all' || item.status === statusFilter
    const matchesType = typeFilter === 'all' || item.report_type === typeFilter

    return matchesSearch && matchesStatus && matchesType
  })

  // Pagination bounds
  const totalPages = Math.ceil(filteredReports.length / itemsPerPage)
  const paginatedReports = filteredReports.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  )

  const getStatusBadgeClass = (status: string) => {
    switch (status) {
      case 'Approved': return 'bg-emerald-50 text-emerald-700 border-emerald-100'
      case 'Under Review':
      case 'Submitted': return 'bg-amber-50 text-amber-700 border-amber-100'
      case 'Rejected':
      case 'Returned for Correction': return 'bg-red-50 text-red-700 border-red-100'
      default: return 'bg-slate-100 text-slate-700 border-slate-200'
    }
  };

  const getStatusLabelAr = (status: string) => {
    switch (status) {
      case 'Draft': return 'مسودة'
      case 'Submitted': return 'مرفوع'
      case 'Under Review': return 'تحت التدقيق'
      case 'Returned for Correction': return 'معاد للتصحيح'
      case 'Reviewed': return 'مُدقق'
      case 'Approved': return 'معتمد'
      case 'Rejected': return 'مرفوض'
      case 'Archived': return 'مؤرشف'
      default: return status
    }
  }

  return (
    <MainWrapper>
      <div className="flex flex-col gap-6 text-right">
        {/* Header toolbar */}
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div>
            <h1 className="text-xl font-bold text-slate-800 font-cairo">أرشيف التقارير والتحصيلات</h1>
            <p className="text-xs text-slate-500 mt-1 font-cairo">عرض ومراجعة كافة كشوف سندات الاستلام والمصاريف للجنة المالية</p>
          </div>
          <Link
            href="/reports/create"
            className="flex items-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-white font-bold py-3 px-5 rounded-xl text-sm transition-all shadow-premium border-0 w-fit self-end md:self-auto"
          >
            <PlusCircle className="h-5 w-5" />
            <span>إنشاء كشف جديد</span>
          </Link>
        </div>

        {/* Filter board */}
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-premium flex flex-wrap gap-4 items-center">
          <div className="flex-1 min-w-[240px] relative">
            <input
              type="text"
              placeholder="ابحث بـ رقم الكشف أو العنوان أو الفترة..."
              value={search}
              onChange={(e) => { setSearch(e.target.value); setCurrentPage(1); }}
              className="w-full pr-10 pl-4 py-2.5 rounded-xl border border-slate-200 text-right text-sm placeholder-slate-400 focus:outline-none focus:border-emerald-500"
            />
            <Search className="absolute right-3.5 top-3.5 h-4 w-4 text-slate-400" />
          </div>

          <div className="w-full sm:w-auto">
            <select
              value={statusFilter}
              onChange={(e) => { setStatusFilter(e.target.value); setCurrentPage(1); }}
              className="w-full sm:w-44 px-3 py-2.5 rounded-xl border border-slate-200 text-right text-sm focus:outline-none focus:border-emerald-500"
            >
              <option value="all">كل الحالات</option>
              <option value="Draft">مسودة</option>
              <option value="Submitted">مرفوع للمراجعة</option>
              <option value="Under Review">قيد التدقيق</option>
              <option value="Approved">معتمد</option>
              <option value="Returned for Correction">معاد للتصحيح</option>
            </select>
          </div>

          <div className="w-full sm:w-auto">
            <select
              value={typeFilter}
              onChange={(e) => { setTypeFilter(e.target.value); setCurrentPage(1); }}
              className="w-full sm:w-44 px-3 py-2.5 rounded-xl border border-slate-200 text-right text-sm focus:outline-none focus:border-emerald-500"
            >
              <option value="all">كل الأنواع</option>
              <option value="كشف التحصيل الشهري">كشف التحصيل الشهري</option>
              <option value="بيان المصاريف">بيان المصاريف</option>
            </select>
          </div>
        </div>

        {/* Reports Content List */}
        {loading ? (
          <div className="flex justify-center items-center py-20">
            <div className="h-10 w-10 animate-spin rounded-full border-4 border-slate-200 border-t-emerald-500"></div>
          </div>
        ) : paginatedReports.length === 0 ? (
          <div className="bg-white border rounded-2xl p-16 text-center shadow-premium">
            <p className="text-slate-400 md:text-base text-sm font-semibold">لا تتوفر أي تقارير مطابقة لمعايير البحث حالياً.</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {paginatedReports.map((report) => (
              <div key={report.id} className="bg-white rounded-2xl border border-slate-100 p-5 shadow-premium flex flex-col justify-between h-56 text-right relative hover:border-emerald-200 transition-all">
                <div>
                  <span className={`inline-block text-[10px] font-bold border px-2 py-0.5 rounded-md mb-2.5 ${getStatusBadgeClass(report.status)}`}>
                    {getStatusLabelAr(report.status)}
                  </span>
                  <h3 className="font-extrabold text-slate-800 text-base line-clamp-1">{report.report_title}</h3>
                  <p className="text-xs text-slate-400 font-semibold mt-1">كود: {report.report_number}</p>
                </div>

                <div className="border-t border-slate-50 pt-3 mt-4 flex items-center justify-between">
                  <div>
                    <span className="text-[10px] text-slate-400 block font-semibold">الفترة المالية</span>
                    <span className="text-xs font-bold text-slate-700">{report.reporting_period}</span>
                  </div>
                  <Link
                    href={`/reports/${report.id}`}
                    className="flex h-9 w-9 items-center justify-center rounded-xl bg-slate-50 border border-slate-100 hover:bg-emerald-50 hover:text-emerald-600 transition-all font-bold text-slate-600"
                  >
                    <Eye className="h-4 w-4" />
                  </Link>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Pagination Row */}
        {!loading && totalPages > 1 && (
          <div className="flex items-center justify-center gap-4 mt-6">
            <button
              onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="flex items-center gap-1 px-4 py-2 border rounded-xl bg-white text-slate-600 font-bold hover:bg-slate-50 disabled:opacity-40"
            >
              <ChevronRight className="h-4 w-4" />
              <span>التالي</span>
            </button>
            <span className="text-sm font-bold text-slate-600">صفحة {currentPage} من {totalPages}</span>
            <button
              onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
              className="flex items-center gap-1 px-4 py-2 border rounded-xl bg-white text-slate-600 font-bold hover:bg-slate-50 disabled:opacity-40"
            >
              <span>السابق</span>
              <ChevronLeft className="h-4 w-4" />
            </button>
          </div>
        )}
      </div>
    </MainWrapper>
  )
}
