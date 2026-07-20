import { createServerClient, type CookieOptions } from '@supabase/ssr'
import { cookies } from 'next/headers'

// Check if credentials are placeholder defaults
const isMockMode = 
  !process.env.NEXT_PUBLIC_SUPABASE_URL ||
  process.env.NEXT_PUBLIC_SUPABASE_URL.includes('your-project-id') ||
  !process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY.includes('your-anon-public-key-here')

const MOCK_ADMINS = [
  { id: 'usr-1', email: 'jehadgml@gmail.com', name: 'جهاد زكري إسماعيل اطفيحة', role: 'Super Administrator' }
]

class MockSupabaseServerClient {
  auth = {
    getUser: async () => {
      // Simulate authenticated super admin in server environment as fallback
      const userObj = MOCK_ADMINS[0]
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
      return { data: { session: { user } }, error: null }
    }
  }
  from(table: string) {
    return {
      select: () => ({
        eq: () => ({
          single: async () => ({ data: null, error: null }),
          then: (resolve: any) => resolve({ data: [], error: null })
        }),
        then: (resolve: any) => resolve({ data: [], error: null })
      })
    }
  }
}

export function createClient() {
  if (isMockMode) {
    return new MockSupabaseServerClient() as any
  }

  const cookieStore = cookies()

  return createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        get(name: string) {
          return cookieStore.get(name)?.value
        },
        set(name: string, value: string, options: CookieOptions) {
          try {
            cookieStore.set({ name, value, ...options })
          } catch (error) {
            // Safe to ignore in Server Components
          }
        },
        remove(name: string, options: CookieOptions) {
          try {
            cookieStore.set({ name, value: '', ...options })
          } catch (error) {
            // Safe to ignore in Server Components
          }
        },
      },
    }
  )
}
