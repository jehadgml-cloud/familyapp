'use client'

import { useState, useEffect, useRef } from 'react'
import { createClient } from '@/lib/supabase/client'
import { useRouter, useParams } from 'next/navigation'
import MainWrapper from '@/components/layout/MainWrapper'
import { 
  ArrowRight, Save, Trash2, Plus, 
  Upload, FileText, CheckCircle, AlertTriangle 
} from 'lucide-react'
import Link from 'next/link'

interface Member {
  id: number
  name: string
  payments: Record<string, boolean>
}

interface Expense {
  id: string
  description: string
  amount: number
}

interface ReportData {
  members: Member[]
  expenses: Expense[]
  month_state: Record<string, boolean>
}

export default function EditReportPage() {
  const params = useParams()
  const router = useRouter()
  const reportId = params.id as string
  const supabase = createClient()

  // State
  const [report, setReport] = useState<any>(null)
  const [reportData, setReportData] = useState<ReportData>({ members: [], expenses: [], month_state: {} })
  const [newMemberName, setNewMemberName] = useState('')
  const [expenseDesc, setExpenseDesc] = useState('')
  const [expenseAmount, setExpenseAmount] = useState('')
  const [attachments, setAttachments] = useState<any[]>([])
  
  // Save Feedback
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [lastSaved, setLastSaved] = useState<string>('')
  const [loading, setLoading] = useState(true)
  const [isDirty, setIsDirty] = useState(false)
  const [uploading, setUploading] = useState(false)

  // References to keep state synced without causing re-saves
  const reportDataRef = useRef<ReportData>(reportData)
  reportDataRef.current = reportData

  // First Load
  useEffect(() => {
    async function loadReportDetails() {
      setLoading(true)
      try {
        const { data, error } = await supabase
          .from('reports')
          .select('*')
          .eq('id', reportId)
          .single()

        if (error) throw error

        if (data) {
          if (data.status !== 'Draft' && data.status !== 'Returned for Correction') {
            alert('هذا التقرير مغلق ومعد للمراجعة فقط ولا يمكن تعديله.')
            router.push(`/reports/${reportId}`)
            return
          }
          setReport(data)
          setReportData(data.report_data as ReportData)
        }

        // Load attachments
        const { data: attachs } = await supabase
          .from('report_attachments')
          .select('*')
          .eq('report_id', reportId)
        setAttachments(attachs || [])

      } catch (err) {
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    loadReportDetails()
  }, [reportId, supabase, router])

  // Save report data
  const saveChanges = async (isAuto = false) => {
    setSaveStatus('saving')
    try {
      const { error } = await supabase
        .from('reports')
        .update({
          report_data: reportDataRef.current,
        })
        .eq('id', reportId)

      if (error) throw error

      setSaveStatus('saved')
      setIsDirty(false)
      const now = new Date()
      setLastSaved(now.toLocaleTimeString('ar-EG'))
    } catch (err) {
      console.error(err)
      setSaveStatus('error')
    }
  }

  // 20-seconds Auto Save loop
  useEffect(() => {
    const interval = setInterval(() => {
      if (isDirty) {
        saveChanges(true)
      }
    }, 20000)

    return () => clearInterval(interval)
  }, [isDirty])

  // Warn before browser exit if dirty
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (isDirty) {
        e.preventDefault()
        e.returnValue = 'هناك تغييرات لم يتم حفظها بعد، هل أنت متأكد من الخروج؟'
        return e.returnValue
      }
    }
    window.addEventListener('beforeunload', handleBeforeUnload)
    return () => window.removeEventListener('beforeunload', handleBeforeUnload)
  }, [isDirty])

  // Ledger updates
  const handlePaymentToggle = (memberId: number) => {
    setIsDirty(true)
    const period = report.reporting_period
    const updatedMembers = reportData.members.map((member) => {
      if (member.id === memberId) {
        const currentPay = !!member.payments?.[period]
        return {
          ...member,
          payments: {
            ...member.payments,
            [period]: !currentPay,
          },
        }
      }
      return member
    })
    setReportData({ ...reportData, members: updatedMembers })
    setSaveStatus('idle')
  }

  // Add Member to this ledger
  const addMember = () => {
    if (!newMemberName.trim()) return
    setIsDirty(true)
    const nextId = Math.max(...reportData.members.map(m => m.id), 0) + 1
    const newM: Member = {
      id: nextId,
      name: newMemberName.trim(),
      payments: {}
    }
    setReportData({
      ...reportData,
      members: [...reportData.members, newM]
    })
    setNewMemberName('')
  }

  // Delete Member from this ledger
  const deleteMember = (memberId: number) => {
    if (!confirm('هل أنت متأكد من حذف هذا الاسم؟')) return
    setIsDirty(true)
    setReportData({
      ...reportData,
      members: reportData.members.filter(m => m.id !== memberId)
    })
  }

  // Add general expense
  const addExpense = () => {
    if (!expenseDesc.trim() || !expenseAmount) return
    setIsDirty(true)
    const newExp: Expense = {
      id: Math.random().toString(36).substr(2, 9),
      description: expenseDesc.trim(),
      amount: parseFloat(expenseAmount)
    }
    setReportData({
      ...reportData,
      expenses: [...(reportData.expenses || []), newExp]
    })
    setExpenseDesc('')
    setExpenseAmount('')
  }

  // Delete general expense
  const deleteExpense = (expId: string) => {
    setIsDirty(true)
    setReportData({
      ...reportData,
      expenses: (reportData.expenses || []).filter(e => e.id !== expId)
    })
  }

  // Handle files upload
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files?.length) return
    const file = e.target.files[0]
    setUploading(true)

    try {
      const fileExt = file.name.split('.').pop()
      const fileName = `${Math.random().toString(36).substring(7)}-${Date.now()}.${fileExt}`
      const filePath = `${reportId}/${fileName}`

      // Upload file to Supabase private Storage
      const { error: uploadError } = await supabase.storage
        .from('report-attachments')
        .upload(filePath, file)

      if (uploadError) throw uploadError

      // Record in database
      const { data, error: dbError } = await supabase
        .from('report_attachments')
        .insert({
          report_id: reportId,
          file_name: file.name,
          file_path: filePath,
          file_type: file.type,
          file_size: file.size,
          uploaded_by: (await supabase.auth.getUser()).data.user?.id,
        })
        .select()
        .single()

      if (dbError) throw dbError

      setAttachments([...attachments, data])

    } catch (err: any) {
      alert(`خطأ في رفع الملف: ${err.message}`)
    } finally {
      setUploading(false)
    }
  }

  // Submit to Reviewing workflow
  const submitToReview = async () => {
    if (isDirty) {
      await saveChanges()
    }

    if (!confirm('هل أنت متأكد من رفع هذا التقرير للتدقيق؟ لن تتمكن من تعديله مجدداً.')) return

    try {
      const { error } = await supabase
        .from('reports')
        .update({
          status: 'Submitted',
          submitted_at: new Date().toISOString()
        })
        .eq('id', reportId)

      if (error) throw error

      router.push(`/reports/${reportId}`)
    } catch (err: any) {
      alert(`فشل إرسال التقرير: ${err.message}`)
    }
  }

  // Calculate sums
  const period = report?.reporting_period || ''
  const totalPaidMembers = reportData.members.filter(m => !!m.payments?.[period]).length
  const totalCollected = totalPaidMembers * 10
  const totalExpenses = (reportData.expenses || []).reduce((acc: number, curr: Expense) => acc + curr.amount, 0)
  const netBalance = totalCollected - totalExpenses

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
      <div className="flex flex-col gap-6 text-right">
        {/* Editor navbar heading */}
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b pb-4 border-slate-100">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">مرفوع من كود {report.report_number}</span>
              {isDirty && <span className="text-[10px] text-amber-600 bg-amber-50 px-2 py-0.5 rounded font-bold font-cairo">تغييرات غير محفوظة</span>}
            </div>
            <h1 className="text-xl font-bold text-slate-800 font-cairo">تحرير الكشف: {report.report_title}</h1>
          </div>

          <div className="flex items-center gap-2 self-end sm:self-auto text-xs">
            {saveStatus === 'saving' && <span className="text-slate-500 animate-pulse font-cairo">جاري الحفظ تلقائياً...</span>}
            {saveStatus === 'saved' && <span className="text-emerald-600 font-cairo">تم الحفظ بنجاح ({lastSaved})</span>}
            {saveStatus === 'error' && <span className="text-red-500 font-cairo font-bold">⚠️ خطأ أثناء الحفظ التلقائي</span>}

            <button onClick={() => saveChanges()} className="bg-slate-100 hover:bg-slate-100 text-slate-700 font-semibold px-4 py-2 rounded-xl border border-slate-200 transition-all font-cairo">
              💾 حفظ
            </button>
            <button onClick={submitToReview} className="bg-emerald-500 hover:bg-emerald-600 text-white font-bold px-4 py-2 rounded-xl transition-all font-cairo">
              🚀 رفع للتدقيق
            </button>
          </div>
        </div>

        {/* Ledger Dashboard Metrics Summary */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="bg-white border rounded-2xl p-5 shadow-premium text-right">
            <span className="text-xs text-slate-400 block mb-1">إجمالي المقبوضات (10 شيكل / مشترك)</span>
            <span className="text-xl font-black text-slate-800 font-cairo">{totalCollected} شيكل</span>
            <span className="text-[10px] text-emerald-600 block mt-1">({totalPaidMembers} مسددين من أصل {reportData.members.length})</span>
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
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between border-b pb-3 mb-4 gap-3 text-right">
            <h3 className="text-sm font-bold text-slate-800 font-cairo">سجل أسماء الجباية والتحصيل ({period})</h3>
            
            {/* Inline add name controls */}
            <div className="flex gap-2">
              <input
                type="text"
                value={newMemberName}
                onChange={(e) => setNewMemberName(e.target.value)}
                placeholder="اسم مشارك جديد..."
                className="px-3 py-1.5 border border-slate-200 rounded-lg text-right text-xs placeholder-slate-400 focus:outline-none"
              />
              <button onClick={addMember} className="bg-emerald-50 text-emerald-600 hover:bg-emerald-100 p-2 rounded-lg border border-emerald-100">
                <Plus className="h-4 w-4" />
              </button>
            </div>
          </div>

          <div className="overflow-x-auto w-full">
            <table className="w-full border-collapse text-right text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-500 font-bold border-b border-slate-100">
                  <th className="p-3 text-center w-12">الرقم</th>
                  <th className="p-3">الاسم كامل</th>
                  <th className="p-3 text-center w-32">{period}</th>
                  <th className="p-3 text-center w-20">إجراء</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {reportData.members.map((member, idx) => (
                  <tr key={member.id} className="hover:bg-slate-50">
                    <td className="p-3 text-center font-bold text-slate-400">{idx + 1}</td>
                    <td className="p-3 font-semibold text-slate-800">{member.name}</td>
                    <td className="p-3 text-center">
                      <input
                        type="checkbox"
                        checked={!!member.payments?.[period]}
                        onChange={() => handlePaymentToggle(member.id)}
                        className="h-5 w-5 accent-emerald-500 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer"
                      />
                    </td>
                    <td className="p-3 text-center">
                      <button onClick={() => deleteMember(member.id)} className="text-red-500 hover:text-red-700 hover:bg-red-50 p-1.5 rounded-lg border border-transparent">
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Expenditures Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* General Expenses list */}
          <div className="bg-white border rounded-2xl p-5 shadow-premium">
            <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">جدول المصاريف والصرفيات</h3>
            
            {/* Inline add cost tool */}
            <div className="grid grid-cols-3 gap-2 mb-4">
              <input
                type="text"
                value={expenseDesc}
                onChange={(e) => setExpenseDesc(e.target.value)}
                placeholder="بيان الصرف..."
                className="col-span-2 px-3 py-2 border rounded-lg text-right text-xs placeholder-slate-400 focus:outline-none"
              />
              <div className="flex gap-1.5">
                <input
                  type="number"
                  value={expenseAmount}
                  onChange={(e) => setExpenseAmount(e.target.value)}
                  placeholder="المبلغ..."
                  className="w-full px-3 py-2 border rounded-lg text-right text-xs placeholder-slate-400 focus:outline-none"
                />
                <button onClick={addExpense} className="bg-slate-900 text-white hover:bg-slate-800 p-2 rounded-lg">
                  <Plus className="h-4 w-4" />
                </button>
              </div>
            </div>

            {/* Expenses List */}
            {(!reportData.expenses || reportData.expenses.length === 0) ? (
              <p className="text-xs text-slate-400 py-6 text-center">لم تصرف أي مستحقات في هذا التقرير.</p>
            ) : (
              <div className="divide-y divide-slate-50">
                {reportData.expenses.map((exp) => (
                  <div key={exp.id} className="py-2.5 flex justify-between items-center text-xs">
                    <button onClick={() => deleteExpense(exp.id)} className="text-red-500 hover:text-red-700 p-1.5 rounded-lg">
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                    <div className="flex items-center gap-6">
                      <span className="font-bold text-slate-800">{exp.amount} شيكل</span>
                      <span className="text-slate-600 font-semibold">{exp.description}</span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Files Attachment uploads */}
          <div className="bg-white border rounded-2xl p-5 shadow-premium">
            <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">المستندات والملفات المرفقة</h3>
            
            <div className="space-y-4">
              <div className="flex items-center justify-center border-2 border-dashed border-slate-200 rounded-xl p-5 bg-slate-50 transition-all hover:bg-slate-100">
                <label className="flex flex-col items-center gap-1.5 cursor-pointer text-slate-500 hover:text-slate-800">
                  <Upload className="h-6 w-6 text-slate-400" />
                  <span className="text-xs font-bold font-cairo">{uploading ? 'جاري رفع الملف...' : 'اضغط لرفع مستند داعم (PDF, JPG, PNG)'}</span>
                  <input
                    type="file"
                    accept=".pdf,image/*"
                    className="hidden"
                    onChange={handleFileUpload}
                    disabled={uploading}
                  />
                </label>
              </div>

              {/* Attachments List */}
              {attachments.length === 0 ? (
                <p className="text-xs text-slate-400 py-4 text-center">لا توجد ملفات مرفقة حالياً.</p>
              ) : (
                <div className="divide-y divide-slate-50">
                  {attachments.map((file) => (
                    <div key={file.id} className="py-2.5 flex justify-between items-center text-xs">
                      <span className="text-[10px] text-slate-400 font-semibold">
                        {(file.file_size / (1024 * 1024)).toFixed(2)} MB
                      </span>
                      <div className="flex items-center gap-2">
                        <span className="text-slate-700 font-bold max-w-[200px] truncate">{file.file_name}</span>
                        <FileText className="h-4 w-4 text-slate-400" />
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>

      </div>
    </MainWrapper>
  )
}
