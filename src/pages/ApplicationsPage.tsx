import { useState, useEffect, useMemo } from 'react'
import { Link } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { useAuthStore } from '../stores/authStore'
import { Loader, Button } from '../components/common'
import type { Application, ApprovalStage } from '../types'

type TabType = 'my' | 'pending' | 'all'

export default function ApplicationsPage() {
  const [applications, setApplications] = useState<Application[]>([])
  const [isLoading, setIsLoading] = useState(true)
  const [activeTab, setActiveTab] = useState<TabType>('my')
  const [searchTerm, setSearchTerm] = useState('')

  const { user, getUserRole } = useAuthStore()
  const userRole = getUserRole()

  useEffect(() => {
    fetchApplications()
  }, [])

  const fetchApplications = async () => {
    setIsLoading(true)
    try {
      const { data, error } = await supabase
        .from('applications')
        .select(`
          *,
          customers(name),
          employees:applicant_id(name, departments(name))
        `)
        .order('application_date', { ascending: false })

      if (error) throw error

      const mapped: Application[] = (data || []).map((app) => {
        const applicant = app.employees as {
          name: string
          departments: { name: string } | null
        } | null

        return {
          id: app.id,
          applicationId: app.application_id,
          groupId: app.group_id || app.application_id,
          applicationType: app.application_type,
          targetMaster: app.target_master || '',
          customerId: app.customer_id || '',
          customerCode: app.customer_code || '',
          customerName: app.customer_name || (app.customers as { name: string } | null)?.name || '',
          applicantId: app.applicant_id || '',
          applicantName: app.applicant_name || applicant?.name || '',
          applicantDepartment: applicant?.departments?.name || '',
          applicationDate: app.application_date,
          status: app.status,
          stage: app.stage,
          effectiveDate: app.effective_date || '',
          reflectionStatus: app.reflection_status || '',
          rejectReason: app.reject_reason || '',
          requiredFields: app.required_fields || [],
          resubmitCount: app.resubmit_count || 0,
          originalApplicationId: app.original_application_id || '',
        }
      })

      setApplications(mapped)
    } catch (err) {
      console.error('Failed to fetch applications:', err)
    } finally {
      setIsLoading(false)
    }
  }

  const filteredApplications = useMemo(() => {
    return applications.filter((app) => {
      if (activeTab === 'my' && app.applicantId !== user?.id) return false

      if (activeTab === 'pending') {
        if (!canApprove(app.stage as ApprovalStage, userRole)) return false
        if (['決裁完了', '却下'].includes(app.stage)) return false
      }

      if (searchTerm) {
        const term = searchTerm.toLowerCase()
        const matchesSearch =
          app.applicationId.toLowerCase().includes(term) ||
          app.customerName.toLowerCase().includes(term) ||
          app.applicantName.toLowerCase().includes(term)
        if (!matchesSearch) return false
      }

      return true
    })
  }, [applications, activeTab, user?.id, userRole, searchTerm])

  const stageStyles: Record<string, string> = {
    申請中: 'bg-blue-100 text-blue-800',
    上長承認済: 'bg-cyan-100 text-cyan-800',
    管理承認済: 'bg-teal-100 text-teal-800',
    常務承認済: 'bg-indigo-100 text-indigo-800',
    決裁完了: 'bg-green-100 text-green-800',
    却下: 'bg-red-100 text-red-800',
  }

  if (isLoading) {
    return <Loader message="申請データを読み込み中..." />
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6 flex-wrap gap-4">
        <h2 className="text-2xl font-bold text-primary-900">申請管理</h2>
      </div>

      <div className="flex gap-0 mb-5 border-b-2 border-gray-200">
        {[
          { key: 'my', label: '自分の申請', count: applications.filter((a) => a.applicantId === user?.id).length },
          {
            key: 'pending',
            label: '承認待ち',
            count: applications.filter(
              (a) => canApprove(a.stage as ApprovalStage, userRole) && !['決裁完了', '却下'].includes(a.stage)
            ).length,
          },
          { key: 'all', label: 'すべての申請', count: applications.length },
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
            placeholder="申請ID・顧客名・申請者で検索..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="flex-1 min-w-[200px] px-4 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500"
          />
          <span className="ml-auto text-sm text-gray-500">
            {filteredApplications.length.toLocaleString()} 件
          </span>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto max-h-[calc(100vh-350px)]">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 bg-gray-50 z-10">
              <tr>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">申請ID</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">申請日時</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">申請種別</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">対象顧客</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">申請者</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">部門</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">承認段階</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">操作</th>
              </tr>
            </thead>
            <tbody>
              {filteredApplications.map((app) => (
                <tr
                  key={app.id}
                  className="hover:bg-blue-50 transition-colors"
                >
                  <td className="px-4 py-3 border-b whitespace-nowrap font-mono text-sm">
                    {app.applicationId}
                    {app.resubmitCount > 0 && (
                      <span className="ml-1 text-xs text-gray-500">
                        (再{app.resubmitCount})
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3 border-b whitespace-nowrap text-sm">
                    {formatDate(app.applicationDate)}
                  </td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{app.applicationType}</td>
                  <td className="px-4 py-3 border-b">{app.customerName}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{app.applicantName}</td>
                  <td className="px-4 py-3 border-b whitespace-nowrap">{app.applicantDepartment}</td>
                  <td className="px-4 py-3 border-b">
                    <span className={`px-2 py-1 text-xs rounded ${stageStyles[app.stage] || 'bg-gray-100'}`}>
                      {app.stage}
                    </span>
                  </td>
                  <td className="px-4 py-3 border-b">
                    <Link to={`/applications/${app.id}`}>
                      <Button size="sm">詳細</Button>
                    </Link>
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

function canApprove(stage: ApprovalStage, userRole: string): boolean {
  switch (stage) {
    case '申請中':
      return userRole === 'supervisor' || userRole === 'admin' || userRole === 'director' || userRole === 'ceo'
    case '上長承認済':
      return userRole === 'admin' || userRole === 'director' || userRole === 'ceo'
    case '管理承認済':
      return userRole === 'director' || userRole === 'ceo'
    case '常務承認済':
      return userRole === 'ceo'
    default:
      return false
  }
}

function formatDate(dateStr: string): string {
  if (!dateStr) return '-'
  try {
    const date = new Date(dateStr)
    return date.toLocaleString('ja-JP', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
    })
  } catch {
    return dateStr
  }
}
