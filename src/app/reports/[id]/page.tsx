'use client'

import { useState, useEffect } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter, useParams } from 'next/navigation'
import MainWrapper from '@/components/layout/MainWrapper'
import { 
  ArrowRight, Edit, Printer, FileDown, 
  MessageSquare, ShieldCheck, CheckCircle2, XCircle, RefreshCw 
} from 'lucide-react'
import Link from 'next/link'

export default function ReportDetailPage() {
  const params = useParams()
  const router = useRouter()
  const reportId = params.id as string
  const supabase = createClient()

  // State
  const [report, setReport] = useState<any>(null)
  const [reportData, setReportData] = useState<any>({ members: [], expenses: [] })
  const [profile, setProfile] = useState<any>(null)
  const [loading, setLoading] = useState(true)
  
  // Comments
  const [comments, setComments] = useState<any[]>([])
  const [newComment, setNewComment] = useState('')
  const [addingComment, setAddingComment] = useState(false)

  // Review states
  const [rejectionReason, setRejectionReason] = useState('')
  const [showRejectForm, setShowRejectForm] = useState(false)

  // Signature Block
  const [showSignatureModal, setShowSignatureModal] = useState(false)
  const [signatureData, setSignatureData] = useState('')

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

  useEffect(() => {
    async function loadReportInfo() {
      setLoading(true)
      try {
        const { data, error } = await supabase
          .from('reports')
          .select('*, departments(department_name_ar)')
          .eq('id', reportId)
          .single()

        if (error) throw error

        if (data) {
          setReport(data)
          setReportData(data.report_data)
        }

        // Fetch comments
        const { data: comms } = await supabase
          .from('report_comments')
          .select('*, profiles(full_name, role)')
          .eq('report_id', reportId)
          .order('created_at', { ascending: true })
        setComments(comms || [])

      } catch (err) {
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    loadReportInfo()
  }, [reportId, supabase])

  const submitComment = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!newComment.trim() || addingComment || !profile) return
    setAddingComment(true)

    try {
      const { data, error } = await supabase
        .from('report_comments')
        .insert({
          report_id: reportId,
          comment: newComment.trim(),
          comment_type: 'general',
          created_by: profile.id,
        })
        .select('*, profiles(full_name, role)')
        .single()

      if (error) throw error

      setComments([...comments, data])
      setNewComment('')
    } catch (err: any) {
      alert(`خطأ في حفظ التعليق: ${err.message}`)
    } finally {
      setAddingComment(false)
    }
  }

  // Workflow actions
  const changeReportStatus = async (newStatus: string, reason?: string) => {
    if (!confirm(`هل أنت متأكد من تغيير حالة التقرير إلى [${newStatus}]؟`)) return

    try {
      const updatePayload: any = { status: newStatus }
      
      if (newStatus === 'Approved') {
        updatePayload.approved_by = profile.id
        updatePayload.approved_at = new Date().toISOString()
      } else if (newStatus === 'Reviewed') {
        updatePayload.reviewed_by = profile.id
        updatePayload.reviewed_at = new Date().toISOString()
      } else if (newStatus === 'Returned for Correction' || newStatus === 'Rejected') {
        updatePayload.rejection_reason = reason || rejectionReason
      }

      const { error } = await supabase
        .from('reports')
        .update(updatePayload)
        .eq('id', reportId)

      if (error) throw error

      // Reload
      router.refresh()
      window.location.reload()
    } catch (err: any) {
      alert(`خطأ في تحديث دورة الاعتماد: ${err.message}`)
    }
  }

  // Electronic Signature
  const saveElectronicSignature = async () => {
    if (!signatureData) return
    try {
      // Append signature to comments or system settings
      await supabase
        .from('report_comments')
        .insert({
          report_id: reportId,
          comment: `إمضاء وتوقيع إلكتروني رسمي معتمد من: ${profile.full_name}`,
          comment_type: 'general',
          created_by: profile.id
        })
      
      alert('تم إمضاء وتوقيع التقرير إلكترونياً بنجاح!')
      setShowSignatureModal(false)
      window.location.reload()
    } catch (err) {
      console.error(err)
    }
  }

  const triggerPrint = () => {
    window.print()
  }

  const period = report?.reporting_period || ''
  const totalPaidMembers = reportData.members?.filter((m: any) => !!m.payments?.[period]).length || 0
  const totalCollected = totalPaidMembers * 10
  const totalExpenses = reportData.expenses?.reduce((acc: number, curr: any) => acc + curr.amount, 0) || 0
  const netBalance = totalCollected - totalExpenses

  // Permissions checks
  const isCreatorState = report?.created_by === profile?.id && (report?.status === 'Draft' || report?.status === 'Returned for Correction')
  const isReviewer = profile?.role === 'Reviewer' || profile?.role === 'Super Administrator'
  const isApprover = profile?.role === 'Approver' || profile?.role === 'Super Administrator'
  
  if (loading) {
    return (
      <MainWrapper>
        <div className="flex justify-center items-center py-40">
          <div className="h-10 w-10 animate-spin rounded-full border-4 border-slate-200 border-t-emerald-500"></div>
        </div>
      </MainWrapper>
    )
  }

  return (
    <MainWrapper>
      <div className="flex flex-col gap-6 text-right no-print">
        {/* Detail page navbar heading */}
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b pb-4 border-slate-100">
          <div>
            <Link href="/reports" className="flex items-center gap-1.5 text-slate-500 font-bold hover:text-slate-800 text-sm mb-2">
              <ArrowRight className="h-4 w-4" />
              <span>العودة للأرشيف</span>
            </Link>
            <h1 className="text-xl font-bold text-slate-800 font-cairo">{report.report_title}</h1>
          </div>

          <div className="flex flex-wrap gap-2 text-xs">
            {isCreatorState && (
              <Link href={`/reports/${reportId}/edit`} className="flex items-center gap-1.5 bg-emerald-500 hover:bg-emerald-600 text-white font-bold px-4 py-2.5 rounded-xl transition-all">
                <Edit className="h-4 w-4" />
                <span>تعديل الكشف</span>
              </Link>
            )}

            <button onClick={triggerPrint} className="flex items-center gap-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-4 py-2.5 rounded-xl border border-slate-200 transition-all font-cairo">
              <Printer className="h-4 w-4" />
              <span>طباعة A4</span>
            </button>

            {report.status === 'Approved' && (
              <button 
                onClick={() => setShowSignatureModal(true)}
                className="flex items-center gap-1.5 bg-indigo-500 hover:bg-indigo-600 text-white font-bold px-4 py-2.5 rounded-xl transition-all font-cairo"
              >
                ✍️ توقيع الكتروني
              </button>
            )}
          </div>
        </div>

        {/* Workflow Approval panel */}
        {report.status !== 'Approved' && (isReviewer || isApprover) && (
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 flex flex-col gap-4 text-right">
            <div className="flex items-center gap-2">
              <ShieldCheck className="h-5 w-5 text-emerald-600" />
              <h3 className="text-sm font-bold text-slate-800 font-cairo">لوحة تدقيق واعتماد التقرير</h3>
            </div>
            
            <div className="flex flex-wrap gap-2">
              {report.status === 'Submitted' && isReviewer && (
                <>
                  <button onClick={() => changeReportStatus('Under Review')} className="bg-amber-500 hover:bg-amber-600 text-white font-bold px-4 py-2 rounded-xl text-xs">
                    🔍 بدء التدقيق والتحقق
                  </button>
                </>
              )}

              {report.status === 'Under Review' && isReviewer && (
                <>
                  <button onClick={() => changeReportStatus('Approved')} className="bg-emerald-500 hover:bg-emerald-600 text-white font-bold px-4 py-2 rounded-xl text-xs">
                    ✅ اعتماد الكشف المالي
                  </button>
                  <button onClick={() => setShowRejectForm(true)} className="bg-red-500 hover:bg-red-600 text-white font-bold px-4 py-2 rounded-xl text-xs font-cairo">
                    ❌ إعادة للتصحيح
                  </button>
                </>
              )}
            </div>

            {showRejectForm && (
              <div className="space-y-3 bg-white p-4 rounded-xl border border-slate-200 mt-2">
                <label className="block text-xs font-bold text-slate-700">ملاحظات وسبب الإعادة للمعد</label>
                <textarea
                  value={rejectionReason}
                  onChange={(e) => setRejectionReason(e.target.value)}
                  className="w-full text-right p-3 border rounded-xl text-xs focus:outline-none"
                  placeholder="اكتب سبب رفض أو إعادة التقرير للمراجعة..."
                />
                <div className="flex justify-end gap-2 text-xs">
                  <button onClick={() => changeReportStatus('Returned for Correction')} className="bg-red-500 text-white font-bold px-4 py-2 rounded-lg">
                    تأكيد الإعادة
                  </button>
                  <button onClick={() => setShowRejectForm(false)} className="bg-slate-100 text-slate-600 px-4 py-2 rounded-lg">
                    إلغاء
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Ledger Dashboard Metrics Summary */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="bg-white border rounded-2xl p-5 shadow-premium text-right">
            <span className="text-xs text-slate-400 block mb-1">إجمالي المقبوضات (10 شيكل / مشترك)</span>
            <span className="text-xl font-black text-slate-800 font-cairo">{totalCollected} شيكل</span>
            <span className="text-[10px] text-emerald-600 block mt-1">({totalPaidMembers} مسددين من أصل {reportData.members?.length || 0})</span>
          </div>
          <div className="bg-white border rounded-2xl p-5 shadow-premium text-right">
            <span className="text-xs text-slate-400 block mb-1">إجمالي المصاريف والمستحقات</span>
            <span className="text-xl font-black text-red-600 font-cairo">{totalExpenses} شيكل</span>
          </div>
          <div className="bg-white border rounded-2xl p-5 shadow-premium text-right">
            <span className="text-xs text-slate-400 block mb-1">الرصيد الصافي للتقرير</span>
            <span className={`text-xl font-black font-cairo ${netBalance >= 0 ? 'text-emerald-600' : 'text-red-600'}`}>{netBalance} شيكل</span>
          </div>
        </div>

        {/* Members Payment Ledger Section */}
        <div className="bg-white border rounded-2xl p-5 shadow-premium">
          <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">سجل أسماء الجباية والتحصيل ({period})</h3>
          
          <div className="overflow-x-auto w-full">
            <table className="w-full border-collapse text-right text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-500 font-bold border-b border-slate-100">
                  <th className="p-3 text-center w-12">الرقم</th>
                  <th className="p-3">الاسم كامل</th>
                  <th className="p-3 text-center w-32">{period}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {reportData.members?.map((member: any, idx: number) => (
                  <tr key={member.id} className="hover:bg-slate-50">
                    <td className="p-3 text-center font-bold text-slate-400">{idx + 1}</td>
                    <td className="p-3 font-semibold text-slate-800">{member.name}</td>
                    <td className="p-3 text-center">
                      <span className={`inline-block w-4 h-4 rounded-full ${member.payments?.[period] ? 'bg-emerald-500' : 'bg-slate-200'}`}></span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Expenses List */}
        <div className="bg-white border rounded-2xl p-5 shadow-premium">
          <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">جدول المصاريف والصرفيات</h3>
          {(!reportData.expenses || reportData.expenses.length === 0) ? (
            <p className="text-xs text-slate-400 py-6 text-center">لم تصرف أي مستحقات في هذا التقرير.</p>
          ) : (
            <div className="divide-y divide-slate-50">
              {reportData.expenses.map((exp: any) => (
                <div key={exp.id} className="py-2.5 flex justify-between items-center text-xs">
                  <span className="font-bold text-slate-800">{exp.amount} شيكل</span>
                  <span className="text-slate-600 font-semibold">{exp.description}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Comments Section */}
        <div className="bg-white border rounded-2xl p-5 shadow-premium">
          <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 flex items-center gap-2 font-cairo">
            <MessageSquare className="h-4 w-4" />
            <span>الملاحظات وتدفق المراجعة</span>
          </h3>

          <div className="space-y-4">
            <div className="space-y-3">
              {comments.map((comm) => (
                <div key={comm.id} className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-right">
                  <div className="flex justify-between items-center text-xs mb-1.5">
                    <span className="text-slate-400">
                      {new Date(comm.created_at).toLocaleString('ar-EG')}
                    </span>
                    <span className="font-bold text-slate-700">
                      {comm.profiles?.full_name} ({comm.profiles?.role})
                    </span>
                  </div>
                  <p className="text-xs font-semibold text-slate-600 leading-relaxed">{comm.comment}</p>
                </div>
              ))}
            </div>

            <form onSubmit={submitComment} className="flex gap-2">
              <input
                type="text"
                value={newComment}
                onChange={(e) => setNewComment(e.target.value)}
                placeholder="اكتب ملاحظة أو توجيه مالي..."
                className="flex-1 text-right px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-emerald-500"
              />
              <button type="submit" disabled={addingComment} className="bg-slate-900 text-white font-bold px-5 py-2.5 rounded-xl text-sm hover:bg-slate-850 transition-all font-cairo">
                إضافة
              </button>
            </form>
          </div>
        </div>

        {/* Interactive E-Signature Modal */}
        {showSignatureModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/45">
            <div className="w-full max-w-md bg-white p-6 rounded-2xl shadow-premium">
              <h3 className="text-sm font-bold text-slate-800 mb-2 font-cairo">التوقيع الإلكتروني الرسمي</h3>
              <p className="text-[10px] text-slate-500 mb-4 font-cairo">يرجى كتابة اسمك الصريح أو الإمضاء في الحقل التالي لاعتماد وتثبيت سند الصندوق.</p>
              
              <input 
                type="text"
                value={signatureData}
                onChange={(e) => setSignatureData(e.target.value)}
                placeholder="أدخل اسمك الكريم كاملاً للمصادقة..."
                className="w-full text-right px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none mb-4"
              />
              
              <div className="flex justify-end gap-2 text-xs">
                <button onClick={saveElectronicSignature} className="bg-emerald-500 text-white font-bold px-4 py-2 rounded-xl">
                  توقيع واعتماد
                </button>
                <button onClick={() => setShowSignatureModal(false)} className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl">
                  إلغاء
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ==================================================
          A4 PRINTABLE VIEW (HIDDEN ON SCREEN, SHOWN ON PRINT)
          ================================================== */}
      {report && (
        <div className="hidden print:block text-right select-text p-6" dir="rtl">
          {/* Header block with Logo */}
          <div className="flex justify-between items-center border-b-2 border-black pb-4 mb-6">
            <div className="text-xs">
              <p className="font-extrabold text-sm">عشيرة آل اطفيحة</p>
              <p className="font-semibold text-[10px] mt-0.5 text-slate-500">اللجنة المالية والتحصيل</p>
              <p className="mt-2 text-[10px]">الكود: {report.report_number}</p>
            </div>
            {/* Safe Arabic Placeholder Logo for prints */}
            <div className="h-16 w-16 border-2 border-black rounded-full flex items-center justify-center font-bold text-lg bg-slate-50">
              ص.ا
            </div>
          </div>

          <h2 className="text-lg font-black text-center mb-6 font-cairo">{report.report_title} ({period})</h2>

          {/* KPI Mini Grid */}
          <div className="grid grid-cols-3 gap-2 border border-black p-3 mb-6 rounded-md">
            <div>
              <span className="text-[10px] text-slate-500 block font-semibold">إجمالي الجباية</span>
              <span className="font-extrabold text-sm">{totalCollected} شيكل</span>
            </div>
            <div>
              <span className="text-[10px] text-slate-500 block font-semibold">المصاريف والصرفيات</span>
              <span className="font-extrabold text-sm">{totalExpenses} شيكل</span>
            </div>
            <div>
              <span className="text-[10px] text-slate-500 block font-semibold">صافي رصيد الصندوق</span>
              <span className="font-extrabold text-sm">{netBalance} شيكل</span>
            </div>
          </div>

          {/* Members Table */}
          <table className="w-full border-collapse border border-black mb-6 text-xs">
            <thead>
              <tr className="bg-slate-100">
                <th className="border border-black p-2 text-center w-12">م</th>
                <th className="border border-black p-2">اسم المشترك كامل</th>
                <th className="border border-black p-2 text-center w-24">الحالة</th>
              </tr>
            </thead>
            <tbody>
              {reportData.members?.map((member: any, idx: number) => (
                <tr key={member.id}>
                  <td className="border border-black p-2 text-center">{idx + 1}</td>
                  <td className="border border-black p-2 font-bold">{member.name}</td>
                  <td className="border border-black p-2 text-center font-bold">
                    {member.payments?.[period] ? 'مسدد (✓)' : 'غير مسدد (✕)'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Signatures Blocks */}
          <div className="border border-black p-4 mt-12 rounded-md">
            <h4 className="text-xs font-bold mb-4">مصادقة وإعتمادات اللجنة المالية</h4>
            <div className="grid grid-cols-3 text-center text-[10px]">
              <div>
                <p className="font-bold">المحاسب (معد التقرير)</p>
                <div className="border-b border-black w-2/3 mx-auto my-4 mt-6"></div>
              </div>
              <div>
                <p className="font-bold">أمين الصندوق (المدقق)</p>
                <div className="border-b border-black w-2/3 mx-auto my-4 mt-6"></div>
              </div>
              <div>
                <p className="font-bold">عميد اللجنة (المعتمد)</p>
                <div className="border-b border-black w-2/3 mx-auto my-4 mt-6"></div>
              </div>
            </div>
          </div>
        </div>
      )}
    </MainWrapper>
  )
}
