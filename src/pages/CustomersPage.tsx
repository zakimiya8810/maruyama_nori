import { useState, useEffect, useMemo } from 'react'
import { Link } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { useMasterDataStore } from '../stores/masterDataStore'
import { useAuthStore } from '../stores/authStore'
import { Loader, Button } from '../components/common'
import type { Customer } from '../types'

export default function CustomersPage() {
  const [customers, setCustomers] = useState<Customer[]>([])
  const [isLoading, setIsLoading] = useState(true)
  const [searchTerm, setSearchTerm] = useState('')
  const [showHidden, setShowHidden] = useState(false)
  const [selectedDivision, setSelectedDivision] = useState('')
  const [selectedDepartment, setSelectedDepartment] = useState('')
  const [selectedHandler, setSelectedHandler] = useState('')
  const [selectedRank, setSelectedRank] = useState('')
  const [selectedBusinessType, setSelectedBusinessType] = useState('')
  const [filterAlertOnly, setFilterAlertOnly] = useState(false)
  const [filterJutaikinOnly, setFilterJutaikinOnly] = useState(false)

  const { departments, employees, ranks, businessTypes } = useMasterDataStore()
  const { user } = useAuthStore()

  useEffect(() => {
    fetchCustomers()
  }, [])

  const fetchCustomers = async () => {
    setIsLoading(true)
    try {
      const { data, error } = await supabase
        .from('customers')
        .select(`
          *,
          business_types(name),
          ranks(name),
          employees:handler_employee_id(name, department_id, division, departments(name))
        `)
        .order('customer_code')

      if (error) throw error

      const mapped: Customer[] = (data || []).map((c) => {
        const handler = c.employees as {
          name: string
          department_id: string
          division: string
          departments: { name: string } | null
        } | null

        return {
          id: c.id,
          customerCode: c.customer_code,
          name: c.name,
          shortName: c.short_name || '',
          kanaName: c.kana_name || '',
          postalCode: c.postal_code || '',
          address1: c.address_1 || '',
          address2: c.address_2 || '',
          tel: c.tel || '',
          fax: c.fax || '',
          businessTypeId: c.business_type_id || '',
          businessTypeName: (c.business_types as { name: string } | null)?.name || '',
          rankId: c.rank_id || '',
          rankName: (c.ranks as { name: string } | null)?.name || '',
          handlerEmployeeId: c.handler_employee_id || '',
          handlerName: handler?.name || '',
          departmentName: handler?.departments?.name || '',
          division: handler?.division || '',
          billingCustomerId: c.billing_customer_id || '',
          paymentTermId: c.payment_term_id || '',
          paymentTermName: '',
          taxMethod: c.tax_method || '',
          creditLimit: c.credit_limit || 0,
          shippingLocation: c.shipping_location || '',
          deliveryMethod: c.delivery_method || '',
          groupCode: c.group_code || '',
          representativeName: c.representative_name || '',
          contactPerson: c.contact_person || '',
          noriUsage: c.nori_usage || 0,
          teaUsage: c.tea_usage || 0,
          naturalFoodHandling: c.natural_food_handling || '',
          shippingFee: c.shipping_fee || '',
          registrationDate: c.registration_date || '',
          tradeStartDate: c.trade_start_date || '',
          deleteFlag: c.delete_flag || 0,
          billingPostalCode: c.billing_postal_code || '',
          billingName: c.billing_name || '',
          billingAddress1: c.billing_address_1 || '',
          billingAddress2: c.billing_address_2 || '',
          billingTel: c.billing_tel || '',
          billingFax: c.billing_fax || '',
          needsAlert: false,
          hasJutaikin: false,
          isHidden:
            c.name?.includes('O_') ||
            (c.ranks as { name: string } | null)?.name === 'OLD' ||
            (c.delete_flag && c.delete_flag >= 1),
        }
      })

      setCustomers(mapped)
    } catch (err) {
      console.error('Failed to fetch customers:', err)
    } finally {
      setIsLoading(false)
    }
  }

  const divisions = useMemo(
    () => [...new Set(employees.map((e) => e.division).filter(Boolean))],
    [employees]
  )

  const filteredDepartments = useMemo(() => {
    if (!selectedDivision) return departments
    const deptIds = employees
      .filter((e) => e.division === selectedDivision)
      .map((e) => e.departmentId)
    return departments.filter((d) => deptIds.includes(d.id))
  }, [selectedDivision, employees, departments])

  const filteredEmployees = useMemo(() => {
    let filtered = employees
    if (selectedDivision) {
      filtered = filtered.filter((e) => e.division === selectedDivision)
    }
    if (selectedDepartment) {
      filtered = filtered.filter((e) => e.departmentId === selectedDepartment)
    }
    return filtered
  }, [selectedDivision, selectedDepartment, employees])

  const filteredCustomers = useMemo(() => {
    return customers.filter((c) => {
      if (!showHidden && c.isHidden) return false

      if (searchTerm) {
        const term = searchTerm.toLowerCase()
        const matchesSearch =
          c.customerCode.toLowerCase().includes(term) ||
          c.name.toLowerCase().includes(term) ||
          c.kanaName.toLowerCase().includes(term) ||
          c.address1.toLowerCase().includes(term)
        if (!matchesSearch) return false
      }

      if (selectedDivision && c.division !== selectedDivision) return false
      if (selectedHandler && c.handlerEmployeeId !== selectedHandler) return false
      if (selectedRank && c.rankId !== selectedRank) return false
      if (selectedBusinessType && c.businessTypeId !== selectedBusinessType) return false
      if (filterAlertOnly && !c.needsAlert) return false
      if (filterJutaikinOnly && !c.hasJutaikin) return false

      return true
    })
  }, [
    customers,
    showHidden,
    searchTerm,
    selectedDivision,
    selectedHandler,
    selectedRank,
    selectedBusinessType,
    filterAlertOnly,
    filterJutaikinOnly,
  ])

  const clearFilters = () => {
    setSearchTerm('')
    setSelectedDivision('')
    setSelectedDepartment('')
    setSelectedHandler('')
    setSelectedRank('')
    setSelectedBusinessType('')
    setFilterAlertOnly(false)
    setFilterJutaikinOnly(false)
  }

  const filterByMyCustomers = () => {
    if (user?.id) {
      setSelectedHandler(user.id)
    }
  }

  if (isLoading) {
    return <Loader message="顧客データを読み込み中..." />
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6 flex-wrap gap-4">
        <h2 className="text-2xl font-bold text-primary-900">顧客一覧</h2>
        <div className="flex gap-3">
          <Link to="/customers/new">
            <Button>
              <i className="fas fa-plus mr-2"></i>
              新規顧客登録申請
            </Button>
          </Link>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-gray-200 p-4 mb-6">
        <div className="flex flex-wrap gap-4 items-center">
          <input
            type="text"
            placeholder="顧客名・コード・住所で検索..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="flex-1 min-w-[200px] px-4 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
          />

          <select
            value={selectedDivision}
            onChange={(e) => {
              setSelectedDivision(e.target.value)
              setSelectedDepartment('')
              setSelectedHandler('')
            }}
            className="px-3 py-2 border border-gray-300 rounded"
          >
            <option value="">大区分</option>
            {divisions.map((div) => (
              <option key={div} value={div}>
                {div}
              </option>
            ))}
          </select>

          <select
            value={selectedDepartment}
            onChange={(e) => {
              setSelectedDepartment(e.target.value)
              setSelectedHandler('')
            }}
            className="px-3 py-2 border border-gray-300 rounded"
          >
            <option value="">部門</option>
            {filteredDepartments.map((dept) => (
              <option key={dept.id} value={dept.id}>
                {dept.name}
              </option>
            ))}
          </select>

          <select
            value={selectedHandler}
            onChange={(e) => setSelectedHandler(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded"
          >
            <option value="">担当者</option>
            {filteredEmployees.map((emp) => (
              <option key={emp.id} value={emp.id}>
                {emp.name}
              </option>
            ))}
          </select>

          <select
            value={selectedRank}
            onChange={(e) => setSelectedRank(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded"
          >
            <option value="">ランク</option>
            {ranks
              .filter((r) => r.visibility !== '非表示')
              .map((rank) => (
                <option key={rank.id} value={rank.id}>
                  {rank.name}
                </option>
              ))}
          </select>

          <select
            value={selectedBusinessType}
            onChange={(e) => setSelectedBusinessType(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded"
          >
            <option value="">業態</option>
            {businessTypes.map((bt) => (
              <option key={bt.id} value={bt.id}>
                {bt.name}
              </option>
            ))}
          </select>
        </div>

        <div className="flex flex-wrap gap-4 items-center mt-4">
          <label className="flex items-center gap-2 cursor-pointer">
            <input
              type="checkbox"
              checked={showHidden}
              onChange={(e) => setShowHidden(e.target.checked)}
              className="rounded"
            />
            <span className="text-sm">旧顧客を表示</span>
          </label>

          <label className="flex items-center gap-2 cursor-pointer">
            <input
              type="checkbox"
              checked={filterAlertOnly}
              onChange={(e) => setFilterAlertOnly(e.target.checked)}
              className="rounded"
            />
            <span className="text-sm">訪問遅れのみ</span>
          </label>

          <label className="flex items-center gap-2 cursor-pointer">
            <input
              type="checkbox"
              checked={filterJutaikinOnly}
              onChange={(e) => setFilterJutaikinOnly(e.target.checked)}
              className="rounded"
            />
            <span className="text-sm">渋滞金あり</span>
          </label>

          <Button variant="secondary" size="sm" onClick={filterByMyCustomers}>
            <i className="fas fa-user mr-1"></i>
            自分の担当
          </Button>

          <Button variant="secondary" size="sm" onClick={clearFilters}>
            <i className="fas fa-times mr-1"></i>
            フィルタクリア
          </Button>

          <span className="ml-auto text-sm text-gray-500">
            {filteredCustomers.length.toLocaleString()} 件表示
          </span>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto max-h-[calc(100vh-350px)]">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 bg-gray-50 z-10">
              <tr>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  得意先コード
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  得意先名称
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  ランク
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  部門
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  担当者
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  業態
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  住所
                </th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b border-gray-200">
                  ステータス
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredCustomers.map((customer) => (
                <tr
                  key={customer.id}
                  className="hover:bg-blue-50 cursor-pointer transition-colors"
                  onClick={() => (window.location.href = `/customers/${customer.id}`)}
                >
                  <td className="px-4 py-3 border-b border-gray-200 whitespace-nowrap">
                    {customer.customerCode}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200">
                    <span className={customer.isHidden ? 'text-gray-400' : ''}>
                      {customer.name}
                    </span>
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200 whitespace-nowrap">
                    {customer.rankName}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200 whitespace-nowrap">
                    {customer.departmentName}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200 whitespace-nowrap">
                    {customer.handlerName}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200 whitespace-nowrap">
                    {customer.businessTypeName}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200 truncate max-w-[200px]">
                    {customer.address1}
                  </td>
                  <td className="px-4 py-3 border-b border-gray-200">
                    <div className="flex gap-1">
                      {customer.needsAlert && (
                        <span className="px-2 py-0.5 bg-yellow-100 text-yellow-800 text-xs rounded">
                          訪問遅れ
                        </span>
                      )}
                      {customer.hasJutaikin && (
                        <span className="px-2 py-0.5 bg-red-100 text-red-800 text-xs rounded">
                          渋滞金
                        </span>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
