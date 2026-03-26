import { create } from 'zustand'
import type { Department, Employee, Rank, BusinessType, PaymentTerm } from '../types'
import { supabase } from '../lib/supabase'

interface MasterDataState {
  departments: Department[]
  employees: Employee[]
  ranks: Rank[]
  businessTypes: BusinessType[]
  paymentTerms: PaymentTerm[]
  isLoading: boolean
  error: string | null
  lastFetched: number | null
  fetchAllMasterData: () => Promise<void>
  clearCache: () => void
}

const CACHE_DURATION = 10 * 60 * 1000

export const useMasterDataStore = create<MasterDataState>((set, get) => ({
  departments: [],
  employees: [],
  ranks: [],
  businessTypes: [],
  paymentTerms: [],
  isLoading: false,
  error: null,
  lastFetched: null,

  fetchAllMasterData: async () => {
    const { lastFetched, isLoading } = get()
    const now = Date.now()

    if (isLoading) return
    if (lastFetched && now - lastFetched < CACHE_DURATION) return

    set({ isLoading: true, error: null })

    try {
      const [
        { data: deptsData, error: deptsError },
        { data: empsData, error: empsError },
        { data: ranksData, error: ranksError },
        { data: btData, error: btError },
        { data: ptData, error: ptError },
      ] = await Promise.all([
        supabase.from('departments').select('*').order('display_order'),
        supabase.from('employees').select('*, departments(name)').eq('is_retired', false).order('display_order'),
        supabase.from('ranks').select('*').order('display_order'),
        supabase.from('business_types').select('*').order('code'),
        supabase.from('payment_terms').select('*').order('display_order'),
      ])

      if (deptsError) throw deptsError
      if (empsError) throw empsError
      if (ranksError) throw ranksError
      if (btError) throw btError
      if (ptError) throw ptError

      const departments: Department[] = (deptsData || []).map(d => ({
        id: d.id,
        code: d.code,
        name: d.name,
        division: d.division || '',
        displayOrder: d.display_order || 999,
      }))

      const employees: Employee[] = (empsData || []).map(e => ({
        id: e.id,
        employeeCode: e.employee_code || '',
        loginId: e.login_id,
        name: e.name,
        email: e.email || '',
        role: e.role || '',
        departmentId: e.department_id || '',
        departmentName: (e.departments as { name: string } | null)?.name || '',
        division: e.division || '',
        adminPermission: e.admin_permission || '',
        isRetired: e.is_retired || false,
        displayOrder: e.display_order || 999,
      }))

      const ranks: Rank[] = (ranksData || []).map(r => ({
        id: r.id,
        code: r.code || '',
        name: r.name,
        visibility: r.visibility || '表示',
        displayOrder: r.display_order || 999,
      }))

      const businessTypes: BusinessType[] = (btData || []).map(bt => ({
        id: bt.id,
        code: bt.code,
        name: bt.name,
      }))

      const paymentTerms: PaymentTerm[] = (ptData || []).map(pt => ({
        id: pt.id,
        pattern: pt.pattern,
        displayOrder: pt.display_order || 999,
      }))

      set({
        departments,
        employees,
        ranks,
        businessTypes,
        paymentTerms,
        isLoading: false,
        lastFetched: now,
      })
    } catch (err) {
      const message = err instanceof Error ? err.message : 'マスタデータの取得に失敗しました'
      set({ isLoading: false, error: message })
    }
  },

  clearCache: () => {
    set({ lastFetched: null })
  },
}))
