import { useState, useEffect } from 'react'
import { Link, useLocation, Outlet, useNavigate } from 'react-router-dom'
import { useAuthStore } from '../stores/authStore'
import { useMasterDataStore } from '../stores/masterDataStore'

const navItems = [
  { path: '/customers', icon: 'fa-solid fa-users', label: '顧客一覧' },
  { path: '/applications', icon: 'fa-solid fa-file-alt', label: '申請一覧' },
  { path: '/meetings', icon: 'fa-solid fa-list-alt', label: '商談登録・履歴' },
]

const masterDataItems = [
  { path: '/master/departments', label: '部署' },
  { path: '/master/users', label: 'ユーザー' },
  { path: '/master/meeting-types', label: '商談種別' },
  { path: '/master/application-types', label: '申請種別' },
]

export default function Layout() {
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false)
  const [isMasterExpanded, setIsMasterExpanded] = useState(false)
  const location = useLocation()
  const navigate = useNavigate()
  const { user, logout } = useAuthStore()
  const { fetchAllMasterData } = useMasterDataStore()

  useEffect(() => {
    fetchAllMasterData()
  }, [fetchAllMasterData])

  useEffect(() => {
    if (location.pathname.startsWith('/master')) {
      setIsMasterExpanded(true)
    }
  }, [location.pathname])

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
            <i className={`fas ${isSidebarCollapsed ? 'fa-bars' : 'fa-bars'}`}></i>
          </button>
        </div>

        <nav className="flex-1 mt-5 overflow-y-auto">
          {navItems.map((item) => (
            <Link
              key={item.path}
              to={item.path}
              className={`flex items-center py-3 px-5 text-gray-700 font-medium transition-all hover:bg-gray-50 ${
                location.pathname.startsWith(item.path)
                  ? 'bg-blue-50 text-blue-700'
                  : ''
              } ${isSidebarCollapsed ? 'justify-center' : ''}`}
            >
              <i className={`${item.icon} w-5 text-center ${isSidebarCollapsed ? 'mr-0' : 'mr-3'}`}></i>
              {!isSidebarCollapsed && <span>{item.label}</span>}
            </Link>
          ))}

          {!isSidebarCollapsed && (
            <div>
              <button
                onClick={() => setIsMasterExpanded(!isMasterExpanded)}
                className={`w-full flex items-center justify-between py-3 px-5 text-gray-700 font-medium transition-all hover:bg-gray-50 ${
                  location.pathname.startsWith('/master') ? 'bg-blue-50 text-blue-700' : ''
                }`}
              >
                <span className="flex items-center">
                  <i className="fa-solid fa-database w-5 text-center mr-3"></i>
                  <span>マスタ・データ管理</span>
                </span>
                <i className={`fas ${isMasterExpanded ? 'fa-chevron-down' : 'fa-chevron-right'} text-xs`}></i>
              </button>
              {isMasterExpanded && (
                <div className="bg-gray-50">
                  {masterDataItems.map((item) => (
                    <Link
                      key={item.path}
                      to={item.path}
                      className={`flex items-center py-2.5 pl-12 pr-5 text-sm text-gray-600 transition-all hover:bg-gray-100 ${
                        location.pathname === item.path ? 'text-blue-700 font-medium' : ''
                      }`}
                    >
                      {item.label}
                    </Link>
                  ))}
                </div>
              )}
            </div>
          )}

          {isSidebarCollapsed && (
            <Link
              to="/master/departments"
              className={`flex items-center justify-center py-3 px-5 text-gray-700 font-medium transition-all hover:bg-gray-50 ${
                location.pathname.startsWith('/master') ? 'bg-blue-50 text-blue-700' : ''
              }`}
              title="マスタ・データ管理"
            >
              <i className="fa-solid fa-database w-5 text-center"></i>
            </Link>
          )}

          <Link
            to="/manual"
            className={`flex items-center py-3 px-5 text-gray-700 font-medium transition-all hover:bg-gray-50 ${
              location.pathname === '/manual' ? 'bg-blue-50 text-blue-700' : ''
            } ${isSidebarCollapsed ? 'justify-center' : ''}`}
          >
            <i className={`fa-solid fa-book w-5 text-center ${isSidebarCollapsed ? 'mr-0' : 'mr-3'}`}></i>
            {!isSidebarCollapsed && <span>操作マニュアル</span>}
          </Link>
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
        <div className="flex-1 overflow-y-auto">
          <Outlet />
        </div>
      </main>
    </div>
  )
}
