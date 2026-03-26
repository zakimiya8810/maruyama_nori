import { create } from 'zustand'
import { persist } from 'zustand/middleware'
import type { User, UserRole } from '../types'
import { supabase } from '../lib/supabase'

interface AuthState {
  user: User | null
  isAuthenticated: boolean
  isLoading: boolean
  error: string | null
  login: (id: string, password: string) => Promise<boolean>
  logout: () => void
  refreshSession: () => Promise<void>
  getUserRole: () => UserRole
}

export const useAuthStore = create<AuthState>()(
  persist(
    (set, get) => ({
      user: null,
      isAuthenticated: false,
      isLoading: false,
      error: null,

      login: async (id: string, password: string) => {
        set({ isLoading: true, error: null })

        try {
          const { data: employee, error } = await supabase
            .from('employees')
            .select(`
              id,
              employee_code,
              login_id,
              name,
              email,
              role,
              department_id,
              division,
              admin_permission,
              departments(name)
            `)
            .eq('login_id', id)
            .eq('password_hash', password)
            .eq('is_retired', false)
            .maybeSingle()

          if (error) throw error
          if (!employee) {
            set({ isLoading: false, error: 'IDまたはパスワードが正しくありません。' })
            return false
          }

          const dept = employee.departments as unknown as { name: string } | null
          const user: User = {
            id: employee.id,
            employeeCode: employee.employee_code || '',
            loginId: employee.login_id,
            name: employee.name,
            email: employee.email || '',
            role: employee.role || '',
            departmentId: employee.department_id || '',
            departmentName: dept?.name || '',
            division: employee.division || '',
            adminPermission: employee.admin_permission || '',
          }

          set({ user, isAuthenticated: true, isLoading: false })
          return true
        } catch (err) {
          const message = err instanceof Error ? err.message : 'ログインに失敗しました'
          set({ isLoading: false, error: message })
          return false
        }
      },

      logout: () => {
        set({ user: null, isAuthenticated: false, error: null })
      },

      refreshSession: async () => {
        const { user } = get()
        if (!user) return

        try {
          const { data: employee, error } = await supabase
            .from('employees')
            .select(`
              id,
              employee_code,
              login_id,
              name,
              email,
              role,
              department_id,
              division,
              admin_permission,
              departments(name)
            `)
            .eq('id', user.id)
            .eq('is_retired', false)
            .maybeSingle()

          if (error || !employee) {
            set({ user: null, isAuthenticated: false })
            return
          }

          const dept = employee.departments as unknown as { name: string } | null
          const updatedUser: User = {
            id: employee.id,
            employeeCode: employee.employee_code || '',
            loginId: employee.login_id,
            name: employee.name,
            email: employee.email || '',
            role: employee.role || '',
            departmentId: employee.department_id || '',
            departmentName: dept?.name || '',
            division: employee.division || '',
            adminPermission: employee.admin_permission || '',
          }

          set({ user: updatedUser })
        } catch {
          set({ user: null, isAuthenticated: false })
        }
      },

      getUserRole: (): UserRole => {
        const { user } = get()
        if (!user) return 'general'

        const role = user.role.toLowerCase()
        if (role.includes('決裁者') || role.includes('社長')) return 'ceo'
        if (role.includes('常務')) return 'director'
        if (role.includes('管理部門') || role.includes('管理')) return 'admin'
        if (role.includes('上長')) return 'supervisor'
        return 'general'
      },
    }),
    {
      name: 'auth-storage',
      partialize: (state) => ({
        user: state.user,
        isAuthenticated: state.isAuthenticated,
      }),
    }
  )
)
