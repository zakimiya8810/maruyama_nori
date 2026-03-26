import { useState, useEffect } from 'react'
import { Link } from 'react-router-dom'
import { useAuthStore } from '../stores/authStore'
import { supabase } from '../lib/supabase'
import { Loader } from '../components/common'

interface DashboardStats {
  totalCustomers: number
  upcomingMeetings: number
  pendingApplications: number
  overdueVisits: number
}

export default function DashboardPage() {
  const { user } = useAuthStore()
  const [stats, setStats] = useState<DashboardStats | null>(null)
  const [isLoading, setIsLoading] = useState(true)
  const [recentMeetings, setRecentMeetings] = useState<
    { id: string; customerName: string; scheduleDate: string; status: string }[]
  >([])

  useEffect(() => {
    fetchDashboardData()
  }, [])

  const fetchDashboardData = async () => {
    setIsLoading(true)
    try {
      const [
        { count: customerCount },
        { count: meetingCount },
        { count: applicationCount },
        { data: meetingsData },
      ] = await Promise.all([
        supabase.from('customers').select('*', { count: 'exact', head: true }).eq('delete_flag', 0),
        supabase
          .from('meetings')
          .select('*', { count: 'exact', head: true })
          .gte('schedule_date', new Date().toISOString().split('T')[0])
          .is('actual_date', null),
        supabase
          .from('applications')
          .select('*', { count: 'exact', head: true })
          .in('stage', ['申請中', '上長承認済', '管理承認済', '常務承認済']),
        supabase
          .from('meetings')
          .select('id, schedule_date, status, customers(name)')
          .order('schedule_date', { ascending: false })
          .limit(5),
      ])

      setStats({
        totalCustomers: customerCount || 0,
        upcomingMeetings: meetingCount || 0,
        pendingApplications: applicationCount || 0,
        overdueVisits: 0,
      })

      setRecentMeetings(
        (meetingsData || []).map((m) => {
          const customer = m.customers as unknown as { name: string } | null
          return {
            id: m.id,
            customerName: customer?.name || '不明',
            scheduleDate: m.schedule_date,
            status: m.status || 'upcoming',
          }
        })
      )
    } catch (err) {
      console.error('Dashboard fetch error:', err)
    } finally {
      setIsLoading(false)
    }
  }

  if (isLoading) {
    return <Loader message="データを読み込み中..." />
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-2xl font-bold text-primary-900">ダッシュボード</h2>
        <span className="text-gray-500">
          {new Date().toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            weekday: 'long',
          })}
        </span>
      </div>

      <div className="bg-primary-50 border-l-4 border-primary-900 p-4 rounded mb-6">
        <p className="text-primary-900">
          <i className="fas fa-user mr-2"></i>
          ようこそ、<strong>{user?.name}</strong>さん
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <StatCard
          icon="fa-building"
          label="登録顧客数"
          value={stats?.totalCustomers || 0}
          color="primary"
          link="/customers"
        />
        <StatCard
          icon="fa-calendar-check"
          label="今後の商談予定"
          value={stats?.upcomingMeetings || 0}
          color="blue"
          link="/meetings"
        />
        <StatCard
          icon="fa-file-signature"
          label="承認待ち申請"
          value={stats?.pendingApplications || 0}
          color="orange"
          link="/applications"
        />
        <StatCard
          icon="fa-exclamation-triangle"
          label="訪問遅延顧客"
          value={stats?.overdueVisits || 0}
          color="red"
          link="/customers?filter=overdue"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-lg border border-gray-200 p-6">
          <h3 className="text-lg font-medium text-primary-900 mb-4 pb-2 border-b-2 border-primary-900">
            <i className="fas fa-clock mr-2"></i>
            最近の商談予定
          </h3>
          {recentMeetings.length === 0 ? (
            <p className="text-gray-500 text-center py-8">商談予定がありません</p>
          ) : (
            <div className="space-y-3">
              {recentMeetings.map((meeting) => (
                <Link
                  key={meeting.id}
                  to={`/meetings/${meeting.id}`}
                  className="flex justify-between items-center p-3 bg-gray-50 rounded hover:bg-gray-100 transition-colors"
                >
                  <span className="font-medium">{meeting.customerName}</span>
                  <span className="text-sm text-gray-500">{meeting.scheduleDate}</span>
                </Link>
              ))}
            </div>
          )}
          <Link
            to="/meetings"
            className="block text-center text-primary-600 hover:text-primary-800 mt-4"
          >
            すべて見る <i className="fas fa-arrow-right ml-1"></i>
          </Link>
        </div>

        <div className="bg-white rounded-lg border border-gray-200 p-6">
          <h3 className="text-lg font-medium text-primary-900 mb-4 pb-2 border-b-2 border-primary-900">
            <i className="fas fa-tasks mr-2"></i>
            クイックアクション
          </h3>
          <div className="grid grid-cols-2 gap-3">
            <Link
              to="/customers/new"
              className="flex items-center gap-2 p-4 bg-primary-50 rounded hover:bg-primary-100 transition-colors text-primary-900"
            >
              <i className="fas fa-plus-circle"></i>
              <span>新規顧客登録</span>
            </Link>
            <Link
              to="/meetings/new"
              className="flex items-center gap-2 p-4 bg-blue-50 rounded hover:bg-blue-100 transition-colors text-blue-700"
            >
              <i className="fas fa-calendar-plus"></i>
              <span>商談予定登録</span>
            </Link>
            <Link
              to="/applications"
              className="flex items-center gap-2 p-4 bg-orange-50 rounded hover:bg-orange-100 transition-colors text-orange-700"
            >
              <i className="fas fa-file-alt"></i>
              <span>申請一覧</span>
            </Link>
            <Link
              to="/manual"
              className="flex items-center gap-2 p-4 bg-gray-50 rounded hover:bg-gray-100 transition-colors text-gray-700"
            >
              <i className="fas fa-book"></i>
              <span>マニュアル</span>
            </Link>
          </div>
        </div>
      </div>
    </div>
  )
}

interface StatCardProps {
  icon: string
  label: string
  value: number
  color: 'primary' | 'blue' | 'orange' | 'red'
  link: string
}

function StatCard({ icon, label, value, color, link }: StatCardProps) {
  const colorClasses = {
    primary: 'bg-primary-50 text-primary-900 border-primary-200',
    blue: 'bg-blue-50 text-blue-700 border-blue-200',
    orange: 'bg-orange-50 text-orange-700 border-orange-200',
    red: 'bg-red-50 text-red-700 border-red-200',
  }

  return (
    <Link
      to={link}
      className={`${colorClasses[color]} border rounded-lg p-5 hover:shadow-md transition-shadow`}
    >
      <div className="flex items-center gap-4">
        <i className={`fas ${icon} text-3xl`}></i>
        <div>
          <p className="text-sm opacity-80">{label}</p>
          <p className="text-3xl font-bold">{value.toLocaleString()}</p>
        </div>
      </div>
    </Link>
  )
}
