import { useState, FormEvent } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuthStore } from '../stores/authStore'
import { Button, Loader } from '../components/common'

export default function LoginPage() {
  const [id, setId] = useState('')
  const [password, setPassword] = useState('')
  const { login, isLoading, error } = useAuthStore()
  const navigate = useNavigate()

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault()
    const success = await login(id, password)
    if (success) {
      navigate('/dashboard')
    }
  }

  if (isLoading) {
    return <Loader message="ログイン中..." />
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-100">
      <div className="bg-white rounded-lg shadow-lg w-full max-w-md p-8">
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-primary-900">顧客カルテ</h1>
          <p className="text-gray-500 mt-2">Customer Karte System</p>
        </div>

        <form onSubmit={handleSubmit}>
          {error && (
            <div className="mb-4 p-3 bg-red-50 border border-red-200 text-red-700 rounded">
              {error}
            </div>
          )}

          <div className="mb-5">
            <label className="block text-sm font-medium text-gray-700 mb-1">
              <i className="fas fa-user mr-2"></i>
              ユーザーID
            </label>
            <input
              type="text"
              value={id}
              onChange={(e) => setId(e.target.value)}
              className="w-full px-4 py-3 border border-gray-300 rounded focus:outline-none focus:border-primary-500 focus:ring-1 focus:ring-primary-500"
              placeholder="IDを入力"
              required
            />
          </div>

          <div className="mb-6">
            <label className="block text-sm font-medium text-gray-700 mb-1">
              <i className="fas fa-lock mr-2"></i>
              パスワード
            </label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full px-4 py-3 border border-gray-300 rounded focus:outline-none focus:border-primary-500 focus:ring-1 focus:ring-primary-500"
              placeholder="パスワードを入力"
              required
            />
          </div>

          <Button type="submit" className="w-full py-3">
            <i className="fas fa-sign-in-alt mr-2"></i>
            ログイン
          </Button>
        </form>

        <div className="mt-6 text-center text-sm text-gray-500">
          <p>株式会社丸山海苔店 顧客管理システム</p>
        </div>
      </div>
    </div>
  )
}
