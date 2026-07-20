'use client'

import { useEffect, useState } from 'react'
import { createClient } from '@/lib/supabase/client'
import MainWrapper from '@/components/layout/MainWrapper'
import { ShieldCheck, UserPlus, Table, Trash2, KeyRound } from 'lucide-react'

export default function SettingsPage() {
  const [profiles, setProfiles] = useState<any[]>([])
  const [auditLogs, setAuditLogs] = useState<any[]>([])
  const [activeTab, setActiveTab] = useState<'users' | 'audit'>('users')
  const [currentUser, setCurrentUser] = useState<any>(null)
  const [loading, setLoading] = useState(true)
  
  // User creation states
  const [fullName, setFullName] = useState('')
  const [email, setEmail] = useState('')
  const [role, setRole] = useState('Viewer')
  const [phone, setPhone] = useState('')
  const [errorMsg, setErrorMsg] = useState('')
  const [successMsg, setSuccessMsg] = useState('')

  const supabase = createClient()

  useEffect(() => {
    async function loadConfig() {
      setLoading(true)
      try {
        const { data: { user } } = await supabase.auth.getUser()
        if (user) {
          const { data: currentProfile } = await supabase
            .from('profiles')
            .select('*')
            .eq('id', user.id)
            .single()
          setCurrentUser(currentProfile)
          
          if (currentProfile?.role === 'Super Administrator' || currentProfile?.role === 'Administrator') {
            // Load profiles
            const { data: profs } = await supabase
              .from('profiles')
              .select('*')
              .order('created_at', { ascending: false })
            setProfiles(profs || [])

            // Load audit logs
            const { data: logs } = await supabase
              .from('audit_logs')
              .select('*')
              .order('created_at', { ascending: false })
              .limit(20)
            setAuditLogs(logs || [])
          }
        }
      } catch (err) {
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    loadConfig()
  }, [supabase])

  // Change user role trigger
  const handleRoleChange = async (profileId: string, newRole: string) => {
    try {
      const { error } = await supabase
        .from('profiles')
        .update({ role: newRole })
        .eq('id', profileId)

      if (error) throw error

      setProfiles(profiles.map(p => p.id === profileId ? { ...p, role: newRole } : p))
      alert('تم تحديث صلاحية المستخدم بنجاح!')
    } catch (err: any) {
      alert(`خطأ في تحديث الصلاحية: ${err.message}`)
    }
  }

  // Toggle user state trigger
  const handleActiveToggle = async (profileId: string, currentStatus: boolean) => {
    try {
      const { error } = await supabase
        .from('profiles')
        .update({ is_active: !currentStatus })
        .eq('id', profileId)

      if (error) throw error

      setProfiles(profiles.map(p => p.id === profileId ? { ...p, is_active: !currentStatus } : p))
      alert('تم تحديث حالة قفل الحساب المالي بنجاح!')
    } catch (err: any) {
      alert(`خطأ في تحديث قفل الحساب: ${err.message}`)
    }
  }

  // Create profile stub
  const handleRegisterUser = async (e: React.FormEvent) => {
    e.preventDefault()
    setErrorMsg('')
    setSuccessMsg('')

    if (!fullName.trim() || !email.trim()) return

    try {
      // NOTE: Creating login credentials programmatically requires Supabase Admin API
      // Here we pre-create the profile stub in database so they can bind on login!
      const placeholderId = crypto.randomUUID()
      const { error } = await supabase
        .from('profiles')
        .insert({
          id: placeholderId,
          full_name: fullName.trim(),
          email: email.trim().toLowerCase(),
          phone: phone.trim(),
          role: role,
          is_active: true
        })

      if (error) throw error

      setSuccessMsg('تم إشراك المندوب/المشرف في سجلات الصندوق بنجاح!')
      setProfiles([{ id: placeholderId, full_name: fullName.trim(), email: email.trim().toLowerCase(), role, is_active: true, created_at: new Date() }, ...profiles])
      setFullName('')
      setEmail('')
      setPhone('')
    } catch (err: any) {
      setErrorMsg(err.message || 'فشلت عملية تهيئة ملف المستخدم.')
    }
  }

  if (loading) {
    return (
      <MainWrapper>
        <div className="flex justify-center items-center py-40">
          <div className="h-10 w-10 animate-spin rounded-full border-4 border-slate-200 border-t-emerald-500"></div>
        </div>
      </MainWrapper>
    )
  }

  // Role gate checking
  const isSuper = currentUser?.role === 'Super Administrator' || currentUser?.role === 'Administrator'
  if (!isSuper) {
    return (
      <MainWrapper>
        <div className="bg-red-50 text-red-700 p-6 rounded-2xl border border-red-100 text-center font-bold font-cairo">
          🔒 عذراً، صفحات الإعدادات والدخول الأمني تتطلب صلاحيات مشرف عام.
        </div>
      </MainWrapper>
    )
  }

  return (
    <MainWrapper>
      <div className="flex flex-col gap-6 text-right">
        {/* Header toolbar */}
        <div>
          <h1 className="text-xl font-bold text-slate-800 font-cairo">إدارة الأعضاء والتحكم الأمني</h1>
          <p className="text-xs text-slate-500 mt-1 font-cairo">بوابة صلاحيات العشيرة، السجلات التقنية والتحكم بالأجهزة</p>
        </div>

        {/* Tab triggers */}
        <div className="flex border-b border-slate-200">
          <button
            onClick={() => setActiveTab('users')}
            className={`px-5 py-3 text-sm font-bold border-b-2 transition-all ${
              activeTab === 'users' ? 'border-emerald-500 text-emerald-600' : 'border-transparent text-slate-500'
            }`}
          >
            المندوبين والمشرفين
          </button>
          <button
            onClick={() => setActiveTab('audit')}
            className={`px-5 py-3 text-sm font-bold border-b-2 transition-all ${
              activeTab === 'audit' ? 'border-emerald-500 text-emerald-600' : 'border-transparent text-slate-500'
            }`}
          >
            سجلات المراجعة التدقيقية (Audit Logs)
          </button>
        </div>

        {activeTab === 'users' ? (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Create new user panel */}
            <div className="lg:col-span-1 bg-white p-5 rounded-2xl border border-slate-100 shadow-premium">
              <div className="flex items-center gap-2 border-b pb-3 mb-4">
                <UserPlus className="h-4.5 w-4.5 text-emerald-600" />
                <h3 className="text-sm font-bold text-slate-800 font-cairo">تهيئة حساب جديد</h3>
              </div>

              {errorMsg && <div className="p-3 text-xs text-red-650 bg-red-50 border border-red-100 rounded-lg mb-3">{errorMsg}</div>}
              {successMsg && <div className="p-3 text-xs text-emerald-650 bg-emerald-50 border border-emerald-100 rounded-lg mb-3">{successMsg}</div>}

              <form onSubmit={handleRegisterUser} className="space-y-4 text-xs font-semibold">
                <div>
                  <label className="block text-slate-600 mb-1">الاسم الكريم كامل</label>
                  <input
                    type="text"
                    required
                    value={fullName}
                    onChange={(e) => setFullName(e.target.value)}
                    className="w-full text-right p-2.5 border rounded-lg focus:outline-none"
                    placeholder="مثال: يوسف ماجد اطفيحة"
                  />
                </div>

                <div>
                  <label className="block text-slate-600 mb-1">البريد الإلكتروني</label>
                  <input
                    type="email"
                    required
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    className="w-full text-right p-2.5 border rounded-lg focus:outline-none"
                    placeholder="mail@family.com"
                  />
                </div>

                <div>
                  <label className="block text-slate-600 mb-1">رقم الهاتف</label>
                  <input
                    type="text"
                    value={phone}
                    onChange={(e) => setPhone(e.target.value)}
                    className="w-full text-right p-2.5 border rounded-lg focus:outline-none"
                    placeholder="059xxxxxxx"
                  />
                </div>

                <div>
                  <label className="block text-slate-600 mb-1">الصلاحية والمسؤولية</label>
                  <select
                    value={role}
                    onChange={(e) => setRole(e.target.value)}
                    className="w-full text-right p-2.5 border rounded-lg focus:outline-none"
                  >
                    <option value="Viewer">مشاهد (قراءة فقط)</option>
                    <option value="Data Entry User">مدخل بيانات (معد تقارير)</option>
                    <option value="Reviewer">مدقق (مراجع حسابات)</option>
                    <option value="Approver">معتمد (لجنة الصندوق)</option>
                    <option value="Administrator">أمين الصندوق (إداري)</option>
                  </select>
                </div>

                <button type="submit" className="w-full bg-emerald-500 text-white font-bold p-3 rounded-lg text-xs hover:bg-emerald-600">
                  💾 حفظ وإشراك
                </button>
              </form>
            </div>

            {/* Users listing */}
            <div className="lg:col-span-2 bg-white rounded-2xl border border-slate-100 p-5 shadow-premium">
              <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">سجل الأعضاء والمندوبين</h3>
              
              <div className="overflow-x-auto w-full">
                <table className="w-full border-collapse text-right text-xs">
                  <thead>
                    <tr className="bg-slate-50 text-slate-500 font-bold border-b">
                      <th className="p-3">الاسم</th>
                      <th className="p-3">البريد الإلكتروني</th>
                      <th className="p-3">الصلاحية</th>
                      <th className="p-3 text-center">أمن الحساب</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {profiles.map((prof) => (
                      <tr key={prof.id} className="hover:bg-slate-50">
                        <td className="p-3 font-bold text-slate-800">{prof.full_name}</td>
                        <td className="p-3 text-slate-500">{prof.email}</td>
                        <td className="p-3">
                          <select
                            disabled={prof.role === 'Super Administrator'}
                            value={prof.role}
                            onChange={(e) => handleRoleChange(prof.id, e.target.value)}
                            className="bg-transparent border border-slate-200 rounded px-1.5 py-1"
                          >
                            <option value="Viewer">Viewer</option>
                            <option value="Data Entry User">Data Entry User</option>
                            <option value="Reviewer">Reviewer</option>
                            <option value="Approver">Approver</option>
                            <option value="Administrator">Administrator</option>
                            <option value="Super Administrator">Super Administrator</option>
                          </select>
                        </td>
                        <td className="p-3 text-center">
                          <button
                            disabled={prof.role === 'Super Administrator'}
                            onClick={() => handleActiveToggle(prof.id, prof.is_active)}
                            className={`px-3 py-1.5 rounded-lg font-bold border text-[10px] ${
                              prof.is_active 
                                ? 'bg-emerald-50 text-emerald-700 border-emerald-100' 
                                : 'bg-red-50 text-red-700 border-red-100'
                            }`}
                          >
                            {prof.is_active ? 'نشط' : 'مقفل'}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          /* Audit logs grid */
          <div className="bg-white rounded-2xl border border-slate-100 p-5 shadow-premium">
            <h3 className="text-sm font-bold text-slate-800 border-b pb-3 mb-4 font-cairo">مراقبة العمليات والتعديلات</h3>
            
            <div className="overflow-x-auto w-full">
              <table className="w-full border-collapse text-right text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 font-bold border-b">
                    <th className="p-3">التوقيت الأصلي</th>
                    <th className="p-3">الحدث / العملية</th>
                    <th className="p-3">اسم الكيان</th>
                    <th className="p-3">محدد الصف</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {auditLogs.map((log) => (
                    <tr key={log.id} className="hover:bg-slate-50 font-semibold text-slate-700">
                      <td className="p-3 text-slate-400">
                        {new Date(log.created_at).toLocaleString('ar-EG')}
                      </td>
                      <td className="p-3">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          log.action === 'INSERT' ? 'bg-emerald-50 text-emerald-700' :
                          log.action === 'UPDATE' ? 'bg-amber-50 text-amber-700' : 'bg-red-50 text-red-700'
                        }`}>
                          {log.action}
                        </span>
                      </td>
                      <td className="p-3">{log.table_name}</td>
                      <td className="p-3 text-slate-400 font-mono text-[10px]">{log.record_id}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </MainWrapper>
  )
}
