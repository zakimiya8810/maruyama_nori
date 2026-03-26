import { useState, useEffect } from 'react'
import { Link, useLocation, Outlet, useNavigate } from 'react-router-dom'
import { useAuthStore } from '../stores/authStore'
import { useMasterDataStore } from '../stores/masterDataStore'

const navItems = [
  { path: '/dashboard', icon: 'fas fa-chart-line', label: 'ダッシュボード' },
  { path: '/customers', icon: 'fas fa-building', label: '顧客一覧' },
  { path: '/meetings', icon: 'fas fa-calendar-check', label: '商談管理' },
  { path: '/applications', icon: 'fas fa-file-signature', label: '申請管理' },
  { path: '/manual', icon: 'fas fa-book', label: '操作マニュアル' },
]

export default function Layout() {
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false)
  const location = useLocation()
  const navigate = useNavigate()
  const { user, logout } = useAuthStore()
  const { fetchAllMasterData } = useMasterDataStore()

  useEffect(() => {
    fetchAllMasterData()
  }, [fetchAllMasterData])

  const handleLogout = () => {
    logout()
    navigate('/login')
  }

  return (
    <div className="flex h-screen">
      <aside
        className={`bg-white border-r border-gray-200 flex flex-col pt-5 transition-all duration-300 ${
          isSidebarCollapsed ? 'w-[70px]' : 'w-[260px]'
        }`}
      >
        <div className="px-5 pb-5 border-b border-gray-200 flex items-center justify-between">
          {!isSidebarCollapsed && (
            <h1 className="text-xl font-bold text-primary-900 whitespace-nowrap">顧客カルテ</h1>
          )}
          <button
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            className="p-2 hover:bg-gray-100 rounded text-primary-900"
          >
            <i className={`fas ${isSidebarCollapsed ? 'fa-angle-right' : 'fa-angle-left'}`}></i>
          </button>
        </div>

        <nav className="flex-1 mt-5">
          {navItems.map((item) => (
            <Link
              key={item.path}
              to={item.path}
              className={`flex items-center py-4 px-5 text-gray-700 font-medium border-l-[3px] border-transparent transition-all hover:bg-gray-100 ${
                location.pathname.startsWith(item.path)
                  ? 'bg-primary-50 text-primary-900 border-l-primary-900'
                  : ''
              } ${isSidebarCollapsed ? 'justify-center' : ''}`}
            >
              <i className={`${item.icon} w-5 text-center ${isSidebarCollapsed ? 'mr-0' : 'mr-4'}`}></i>
              {!isSidebarCollapsed && <span>{item.label}</span>}
            </Link>
          ))}
        </nav>

        <div className="p-5 border-t border-gray-200">
          {!isSidebarCollapsed && user && (
            <div className="flex flex-col gap-2 text-sm">
              <span className="font-medium">{user.name}</span>
              <span className="text-gray-500">{user.departmentName}</span>
              <button
                onClick={handleLogout}
                className="mt-2 text-left text-gray-600 hover:text-primary-900"
              >
                <i className="fas fa-sign-out-alt mr-2"></i>
                ログアウト
              </button>
            </div>
          )}
          {isSidebarCollapsed && (
            <button
              onClick={handleLogout}
              className="w-full text-center text-gray-600 hover:text-primary-900"
              title="ログアウト"
            >
              <i className="fas fa-sign-out-alt"></i>
            </button>
          )}
        </div>
      </aside>

      <main className="flex-1 flex flex-col h-screen overflow-y-auto min-w-0">
        <header className="bg-white border-b border-gray-200 px-6 py-3 flex justify-end items-center min-h-[57px]">
          <div className="flex items-center gap-4 text-sm">
            <span className="text-gray-600">{user?.name}</span>
            <span className="bg-gray-100 px-3 py-1.5 rounded">{user?.role || '一般'}</span>
          </div>
        </header>
        <div className="flex-1 overflow-y-auto">
          <Outlet />
        </div>
      </main>
    </div>
  )
}
