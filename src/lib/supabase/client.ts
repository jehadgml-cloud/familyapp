import { createBrowserClient } from '@supabase/ssr'

// Check if credentials are placeholder defaults
const isMockMode = 
  !process.env.NEXT_PUBLIC_SUPABASE_URL ||
  process.env.NEXT_PUBLIC_SUPABASE_URL.includes('your-project-id') ||
  !process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY.includes('your-anon-public-key-here')

// Direct static fallback database
const MOCK_ADMINS = [
  { id: 'usr-1', email: 'jehadgml@gmail.com', password: 'ABC12345', name: 'جهاد زكري إسماعيل اطفيحة', role: 'Super Administrator' },
  { id: 'usr-2', email: 'ashraf.atfihah@gmail.com', password: 'ABC12345', name: 'أشرف يوسف محمود اطفيحة', role: 'Administrator' },
  { id: 'usr-3', email: 'ibrahim.atfihah@gmail.com', password: 'ABC12345', name: 'إبراهيم محمد إبراهيم اطفيحة', role: 'Reviewer' }
]

function getStorageItem(key: string, defaultVal: string) {
  if (typeof window === 'undefined') return JSON.parse(defaultVal)
  const item = localStorage.getItem(key)
  return item ? JSON.parse(item) : JSON.parse(defaultVal)
}

function setStorageItem(key: string, val: any) {
  if (typeof window !== 'undefined') {
    localStorage.setItem(key, JSON.stringify(val))
  }
}

// In-Memory/LocalStorage Database Emulation client wrapper
class MockSupabaseClient {
  auth = {
    getUser: async () => {
      if (typeof window === 'undefined') return { data: { user: null }, error: null }
      const session = localStorage.getItem('cems_mock_session')
      if (!session) return { data: { user: null }, error: null }
      const userObj = MOCK_ADMINS.find(u => u.email === session) || MOCK_ADMINS[0]
      return {
        data: {
          user: {
            id: userObj.id,
            email: userObj.email,
            raw_user_meta_data: { full_name: userObj.name, role: userObj.role }
          }
        },
        error: null
      }
    },
    getSession: async () => {
      const { data: { user } } = await this.auth.getUser()
      if (!user) return { data: { session: null }, error: null }
      return { data: { session: { user } }, error: null }
    },
    signInWithPassword: async ({ email, password }: any) => {
      const matched = MOCK_ADMINS.find(u => u.email.toLowerCase() === email.toLowerCase() && u.password === password)
      if (!matched) {
        return { data: null, error: { message: 'خطأ في البريد الإلكتروني أو كلمة المرور' } }
      }
      localStorage.setItem('cems_mock_session', matched.email)
      return {
        data: {
          user: {
            id: matched.id,
            email: matched.email,
            raw_user_meta_data: { full_name: matched.name, role: matched.role }
          }
        },
        error: null
      }
    },
    signOut: async () => {
      localStorage.removeItem('cems_mock_session')
      return { error: null }
    }
  }

  from(table: string) {
    return {
      select: (columns: string = '*') => {
        let data: any[] = []
        if (table === 'reports') {
          data = getStorageItem('cems_mock_reports', '[]')
        } else if (table === 'profiles') {
          data = MOCK_ADMINS.map(u => ({ id: u.id, full_name: u.name, email: u.email, role: u.role, is_active: true }))
        } else if (table === 'report_attachments') {
          data = getStorageItem('cems_mock_attachments', '[]')
        } else if (table === 'report_comments') {
          data = getStorageItem('cems_mock_comments', '[]')
        }

        const buildQuery = (filteredData: any[]) => {
          const apiChain = {
            eq: (col: string, val: any) => {
              const res = filteredData.filter(item => item[col] === val)
              return buildQuery(res)
            },
            order: (col: string, options?: any) => {
              const res = [...filteredData].sort((a, b) => b[col] > a[col] ? 1 : -1)
              return buildQuery(res)
            },
            limit: (num: number) => {
              return buildQuery(filteredData.slice(0, num))
            },
            single: async () => {
              if (filteredData.length === 0) {
                return { data: null, error: { message: 'Element not found' } }
              }
              return { data: filteredData[0], error: null }
            },
            then: (resolve: any) => {
              resolve({ data: filteredData, error: null })
            }
          }
          return apiChain
        }

        return buildQuery(data)
      },
      insert: (payload: any) => {
        return {
          select: () => {
            return {
              single: async () => {
                const row = { ...payload, id: payload.id || crypto.randomUUID(), created_at: new Date().toISOString() }
                if (table === 'reports') {
                  const currentList = getStorageItem('cems_mock_reports', '[]')
                  setStorageItem('cems_mock_reports', [row, ...currentList])
                } else if (table === 'report_attachments') {
                  const currentList = getStorageItem('cems_mock_attachments', '[]')
                  setStorageItem('cems_mock_attachments', [row, ...currentList])
                } else if (table === 'report_comments') {
                  const currentList = getStorageItem('cems_mock_comments', '[]')
                  setStorageItem('cems_mock_comments', [row, ...currentList])
                }
                return { data: row, error: null }
              }
            }
          }
        }
      },
      update: (payload: any) => {
        return {
          eq: (colName: string, colValue: any) => {
            if (table === 'reports') {
              const currentList = getStorageItem('cems_mock_reports', '[]')
              const updated = currentList.map((item: any) => {
                if (item[colName] === colValue) {
                  return { ...item, ...payload, updated_at: new Date().toISOString() }
                }
                return item
              })
              setStorageItem('cems_mock_reports', updated)
            }
            return {
              then: (resolve: any) => resolve({ data: payload, error: null })
            }
          }
        }
      }
    }
  }

  storage = {
    from: (bucket: string) => ({
      upload: async (filePath: string, file: any) => {
        return { data: { path: filePath }, error: null }
      }
    })
  }
}

export function createClient() {
  if (isMockMode) {
    return new MockSupabaseClient() as any
  }
  return createBrowserClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
  )
}
