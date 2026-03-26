import { useState, useEffect, useMemo } from 'react'
import { supabase } from '../lib/supabase'
import { useMasterDataStore } from '../stores/masterDataStore'
import { useAuthStore } from '../stores/authStore'
import { Loader, Button, Modal } from '../components/common'
import type { Meeting } from '../types'

type TabType = 'all' | 'upcoming' | 'overdue' | 'completed'

export default function MeetingsPage() {
  const [meetings, setMeetings] = useState<Meeting[]>([])
  const [isLoading, setIsLoading] = useState(true)
  const [activeTab, setActiveTab] = useState<TabType>('all')
  const [searchTerm, setSearchTerm] = useState('')
  const [selectedDivision, setSelectedDivision] = useState('')
  const [selectedDepartment, setSelectedDepartment] = useState('')
  const [selectedHandler, setSelectedHandler] = useState('')
  const [showNewModal, setShowNewModal] = useState(false)
  const [newMeetingData, setNewMeetingData] = useState({
    customerId: '',
    customerName: '',
    scheduleDate: '',
    purpose: '',
    planNotes: '',
    appointmentType: '',
  })

  const { departments, employees } = useMasterDataStore()
  const { user } = useAuthStore()

  useEffect(() => {
    fetchMeetings()
  }, [])

  const fetchMeetings = async () => {
    setIsLoading(true)
    try {
      const { data, error } = await supabase
        .from('meetings')
        .select(`
          *,
          customers(customer_code, name, ranks(name)),
          employees:handler_id(name, division, departments(name))
        `)
        .order('schedule_date', { ascending: false })

      if (error) throw error

      const mapped: Meeting[] = (data || []).map((m) => {
        const handler = m.employees as {
          name: string
          division: string
          departments: { name: string } | null
        } | null
        const customer = m.customers as {
          customer_code: string
          name: string
          ranks: { name: string } | null
        } | null

        return {
          id: m.id,
          customerId: m.customer_id,
          customerCode: customer?.customer_code || '',
          customerName: customer?.name || '',
          rankName: customer?.ranks?.name || '',
          handlerId: m.handler_id || '',
          handlerName: handler?.name || '',
          departmentName: handler?.departments?.name || '',
          division: handler?.division || '',
          scheduleDate: m.schedule_date,
          actualDate: m.actual_date,
          appointmentType: m.appointment_type || '',
          purpose: m.purpose || '',
          planNotes: m.plan_notes || '',
          result: m.result || '',
          resultNotes: m.result_notes || '',
          delayReason: m.delay_reason || '',
          status: m.status || 'upcoming',
          createdAt: m.created_at,
        }
      })

      setMeetings(mapped)
    } catch (err) {
      console.error('Failed to fetch meetings:', err)
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

  const filteredMeetings = useMemo(() => {
    return meetings.filter((m) => {
      if (activeTab === 'upcoming' && m.status !== 'upcoming') return false
      if (activeTab === 'overdue' && m.status !== 'overdue') return false
      if (activeTab === 'completed' && !['completed', 'delayed'].includes(m.status)) return false

      if (searchTerm) {
        const term = searchTerm.toLowerCase()
        const matchesSearch =
          m.customerName.toLowerCase().includes(term) ||
          m.customerCode.toLowerCase().includes(term) ||
          m.handlerName.toLowerCase().includes(term)
        if (!matchesSearch) return false
      }

      if (selectedDivision && m.division !== selectedDivision) return false
      if (selectedHandler && m.handlerId !== selectedHandler) return false

      return true
    })
  }, [meetings, activeTab, searchTerm, selectedDivision, selectedHandler])

  const handleCreateMeeting = async () => {
    if (!newMeetingData.customerId || !newMeetingData.scheduleDate) {
      alert('顧客と予定日は必須です')
      return
    }

    try {
      const { error } = await supabase.from('meetings').insert({
        customer_id: newMeetingData.customerId,
        handler_id: user?.id,
        schedule_date: newMeetingData.scheduleDate,
        purpose: newMeetingData.purpose,
        plan_notes: newMeetingData.planNotes,
        appointment_type: newMeetingData.appointmentType,
        status: 'upcoming',
      })

      if (error) throw error

      setShowNewModal(false)
      setNewMeetingData({
        customerId: '',
        customerName: '',
        scheduleDate: '',
        purpose: '',
        planNotes: '',
        appointmentType: '',
      })
      fetchMeetings()
    } catch (err) {
      console.error('Failed to create meeting:', err)
      alert('商談の登録に失敗しました')
    }
  }

  const statusStyles: Record<string, string> = {
    upcoming: 'bg-blue-100 text-blue-800',
    completed: 'bg-green-100 text-green-800',
    delayed: 'bg-yellow-100 text-yellow-800',
    overdue: 'bg-red-100 text-red-800',
  }

  const statusLabels: Record<string, string> = {
    upcoming: '予定',
    completed: '完了',
    delayed: '遅延完了',
    overdue: '対応遅れ',
  }

  if (isLoading) {
    return <Loader message="商談データを読み込み中..." />
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6 flex-wrap gap-4">
        <h2 className="text-2xl font-bold text-primary-900">商談管理</h2>
        <Button onClick={() => setShowNewModal(true)}>
          <i className="fas fa-plus mr-2"></i>
          商談予定登録
        </Button>
      </div>

      <div className="flex gap-0 mb-5 border-b-2 border-gray-200">
        {[
          { key: 'all', label: 'すべて', count: meetings.length },
          {
            key: 'upcoming',
            label: '予定',
            count: meetings.filter((m) => m.status === 'upcoming').length,
          },
          {
            key: 'overdue',
            label: '対応遅れ',
            count: meetings.filter((m) => m.status === 'overdue').length,
          },
          {
            key: 'completed',
            label: '完了',
            count: meetings.filter((m) => ['completed', 'delayed'].includes(m.status)).length,
          },
        ].map((tab) => (
          <button
            key={tab.key}
            onClick={() => setActiveTab(tab.key as TabType)}
            className={`px-6 py-3 text-sm font-medium border-b-[3px] transition-colors -mb-[2px] ${
              activeTab === tab.key
                ? 'text-primary-900 border-primary-900'
                : 'text-gray-500 border-transparent hover:text-primary-900 hover:bg-primary-50/50'
            }`}
          >
            {tab.label} ({tab.count})
          </button>
        ))}
      </div>

      <div className="bg-white rounded-lg border border-gray-200 p-4 mb-6">
        <div className="flex flex-wrap gap-4 items-center">
          <input
            type="text"
            placeholder="顧客名・コード・担当者で検索..."
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

          <Button
            variant="secondary"
            size="sm"
            onClick={() => user?.id && setSelectedHandler(user.id)}
          >
            <i className="fas fa-user mr-1"></i>
            自分の商談
          </Button>

          <span className="ml-auto text-sm text-gray-500">
            {filteredMeetings.length.toLocaleString()} 件
          </span>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto max-h-[calc(100vh-350px)]">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 bg-gray-50 z-10">
              <tr>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">予定日</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">実施日</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">顧客名</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">ランク</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">担当者</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">部門</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">目的</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">結果</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">ステータス</th>
              </tr>
            </thead>
            <tbody>
              {filteredMeetings.map((meeting) => (
                <tr
                  key={meeting.id}
                  className="hover:bg-blue-50 cursor-pointer transition-colors"
                  onClick={() => (window.location.href = `/meetings/${meeting.id}`)}
                >
                  <td className="px-4 py-3 border-b whitespace-nowrap">{meeting.scheduleDate}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">
                    {meeting.actualDate || '-'}
                  </td>
                  <td className="px-4 py-3 border-b">{meeting.customerName}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{meeting.rankName}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{meeting.handlerName}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{meeting.departmentName}</td>
                  <td className="px-4 py-3 border-b truncate max-w-[150px]">{meeting.purpose}</td>
                  <td className="px-4 py-3 border-b truncate max-w-[100px]">{meeting.result}</td>
                  <td className="px-4 py-3 border-b">
                    <span
                      className={`px-2 py-1 text-xs rounded ${statusStyles[meeting.status] || 'bg-gray-100'}`}
                    >
                      {statusLabels[meeting.status] || meeting.status}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <Modal
        isOpen={showNewModal}
        onClose={() => setShowNewModal(false)}
        title="商談予定登録"
        footer={
          <>
            <Button variant="secondary" onClick={() => setShowNewModal(false)}>
              キャンセル
            </Button>
            <Button onClick={handleCreateMeeting}>登録</Button>
          </>
        }
      >
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium mb-1">顧客検索</label>
            <CustomerSearch
              onSelect={(customer) =>
                setNewMeetingData({
                  ...newMeetingData,
                  customerId: customer.id,
                  customerName: customer.name,
                })
              }
            />
            {newMeetingData.customerName && (
              <p className="mt-1 text-sm text-primary-600">選択: {newMeetingData.customerName}</p>
            )}
          </div>

          <div>
            <label className="block text-sm font-medium mb-1">予定日</label>
            <input
              type="date"
              value={newMeetingData.scheduleDate}
              onChange={(e) =>
                setNewMeetingData({ ...newMeetingData, scheduleDate: e.target.value })
              }
              className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
            />
          </div>

          <div>
            <label className="block text-sm font-medium mb-1">アポイント</label>
            <select
              value={newMeetingData.appointmentType}
              onChange={(e) =>
                setNewMeetingData({ ...newMeetingData, appointmentType: e.target.value })
              }
              className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
            >
              <option value="">選択してください</option>
              <option value="アポあり">アポあり</option>
              <option value="アポなし">アポなし</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium mb-1">商談目的</label>
            <input
              type="text"
              value={newMeetingData.purpose}
              onChange={(e) => setNewMeetingData({ ...newMeetingData, purpose: e.target.value })}
              className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
              placeholder="商談の目的を入力"
            />
          </div>

          <div>
            <label className="block text-sm font-medium mb-1">予定備考</label>
            <textarea
              value={newMeetingData.planNotes}
              onChange={(e) => setNewMeetingData({ ...newMeetingData, planNotes: e.target.value })}
              className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500 min-h-[80px]"
              placeholder="事前メモなど"
            />
          </div>
        </div>
      </Modal>
    </div>
  )
}

function CustomerSearch({
  onSelect,
}: {
  onSelect: (customer: { id: string; name: string }) => void
}) {
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<{ id: string; customer_code: string; name: string }[]>([])
  const [isSearching, setIsSearching] = useState(false)

  useEffect(() => {
    if (query.length < 2) {
      setResults([])
      return
    }

    const timer = setTimeout(async () => {
      setIsSearching(true)
      try {
        const { data, error } = await supabase
          .from('customers')
          .select('id, customer_code, name')
          .or(`name.ilike.%${query}%,customer_code.ilike.%${query}%`)
          .eq('delete_flag', 0)
          .limit(10)

        if (error) throw error
        setResults(data || [])
      } catch (err) {
        console.error('Search error:', err)
      } finally {
        setIsSearching(false)
      }
    }, 300)

    return () => clearTimeout(timer)
  }, [query])

  return (
    <div className="relative">
      <input
        type="text"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder="顧客名またはコードで検索..."
        className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
      />
      {isSearching && (
        <div className="absolute right-3 top-2.5">
          <div className="w-5 h-5 border-2 border-gray-300 border-t-primary-500 rounded-full animate-spin"></div>
        </div>
      )}
      {results.length > 0 && (
        <div className="absolute z-10 w-full mt-1 bg-white border border-gray-300 rounded shadow-lg max-h-48 overflow-y-auto">
          {results.map((customer) => (
            <button
              key={customer.id}
              onClick={() => {
                onSelect({ id: customer.id, name: customer.name })
                setQuery('')
                setResults([])
              }}
              className="w-full px-3 py-2 text-left hover:bg-gray-100 text-sm"
            >
              <span className="text-gray-500">{customer.customer_code}</span>
              <span className="ml-2">{customer.name}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
