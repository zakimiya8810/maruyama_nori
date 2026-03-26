import { useState, useEffect } from 'react'
import { supabase } from '../../lib/supabase'
import { Button, Modal } from '../../components/common'

interface User {
  id: string
  name: string
  email: string
  department_id: string | null
  role: string
  department?: { name: string }
}

interface Department {
  id: string
  name: string
}

export default function UsersPage() {
  const [users, setUsers] = useState<User[]>([])
  const [departments, setDepartments] = useState<Department[]>([])
  const [loading, setLoading] = useState(true)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [editingUser, setEditingUser] = useState<User | null>(null)
  const [formData, setFormData] = useState({
    name: '',
    email: '',
    department_id: '',
    role: '一般',
    password: '',
  })

  useEffect(() => {
    fetchUsers()
    fetchDepartments()
  }, [])

  async function fetchUsers() {
    setLoading(true)
    const { data, error } = await supabase
      .from('users')
      .select('*, department:departments(name)')
      .order('name')

    if (!error && data) {
      setUsers(data)
    }
    setLoading(false)
  }

  async function fetchDepartments() {
    const { data, error } = await supabase
      .from('departments')
      .select('*')
      .order('name')

    if (!error && data) {
      setDepartments(data)
    }
  }

  function openAddModal() {
    setEditingUser(null)
    setFormData({ name: '', email: '', department_id: '', role: '一般', password: '' })
    setIsModalOpen(true)
  }

  function openEditModal(user: User) {
    setEditingUser(user)
    setFormData({
      name: user.name,
      email: user.email,
      department_id: user.department_id || '',
      role: user.role,
      password: '',
    })
    setIsModalOpen(true)
  }

  async function handleSubmit() {
    if (!formData.name.trim() || !formData.email.trim()) return

    if (editingUser) {
      const updateData: Record<string, string | null> = {
        name: formData.name.trim(),
        email: formData.email.trim(),
        department_id: formData.department_id || null,
        role: formData.role,
      }
      if (formData.password) {
        updateData.password_hash = formData.password
      }

      const { error } = await supabase
        .from('users')
        .update(updateData)
        .eq('id', editingUser.id)

      if (!error) {
        setIsModalOpen(false)
        fetchUsers()
      }
    } else {
      const { error } = await supabase
        .from('users')
        .insert({
          name: formData.name.trim(),
          email: formData.email.trim(),
          department_id: formData.department_id || null,
          role: formData.role,
          password_hash: formData.password || 'password',
        })

      if (!error) {
        setIsModalOpen(false)
        fetchUsers()
      }
    }
  }

  async function handleDelete(id: string) {
    if (!confirm('このユーザーを削除してもよろしいですか？')) return

    const { error } = await supabase
      .from('users')
      .delete()
      .eq('id', id)

    if (!error) {
      fetchUsers()
    }
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6">
        <h1 className="text-2xl font-bold">ユーザーマスタ</h1>
        <Button onClick={openAddModal}>+ ユーザー追加</Button>
      </div>

      <div className="bg-white rounded-lg shadow">
        {loading ? (
          <div className="p-4 text-center text-gray-500">読み込み中...</div>
        ) : (
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="text-left px-4 py-3 font-medium">名前</th>
                <th className="text-left px-4 py-3 font-medium">メール</th>
                <th className="text-left px-4 py-3 font-medium">部署</th>
                <th className="text-left px-4 py-3 font-medium">権限</th>
                <th className="text-right px-4 py-3 font-medium w-32">操作</th>
              </tr>
            </thead>
            <tbody>
              {users.map((user) => (
                <tr key={user.id} className="border-t">
                  <td className="px-4 py-3">{user.name}</td>
                  <td className="px-4 py-3">{user.email}</td>
                  <td className="px-4 py-3">{user.department?.name || '-'}</td>
                  <td className="px-4 py-3">{user.role}</td>
                  <td className="px-4 py-3 text-right">
                    <div className="flex justify-end gap-2">
                      <button
                        onClick={() => openEditModal(user)}
                        className="text-blue-600 hover:underline text-sm"
                      >
                        編集
                      </button>
                      <button
                        onClick={() => handleDelete(user.id)}
                        className="text-red-600 hover:underline text-sm"
                      >
                        削除
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {users.length === 0 && (
                <tr>
                  <td colSpan={5} className="px-4 py-8 text-center text-gray-500">
                    ユーザーが登録されていません
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>

      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={editingUser ? 'ユーザー編集' : 'ユーザー追加'}>
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium mb-1">名前 *</label>
            <input
              type="text"
              value={formData.name}
              onChange={(e) => setFormData({ ...formData, name: e.target.value })}
              className="w-full border rounded px-3 py-2"
            />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">メール *</label>
            <input
              type="email"
              value={formData.email}
              onChange={(e) => setFormData({ ...formData, email: e.target.value })}
              className="w-full border rounded px-3 py-2"
            />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">部署</label>
            <select
              value={formData.department_id}
              onChange={(e) => setFormData({ ...formData, department_id: e.target.value })}
              className="w-full border rounded px-3 py-2"
            >
              <option value="">選択してください</option>
              {departments.map((dept) => (
                <option key={dept.id} value={dept.id}>{dept.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">権限</label>
            <select
              value={formData.role}
              onChange={(e) => setFormData({ ...formData, role: e.target.value })}
              className="w-full border rounded px-3 py-2"
            >
              <option value="一般">一般</option>
              <option value="上長">上長</option>
              <option value="管理者">管理者</option>
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">
              パスワード {editingUser ? '(変更する場合のみ)' : '*'}
            </label>
            <input
              type="password"
              value={formData.password}
              onChange={(e) => setFormData({ ...formData, password: e.target.value })}
              className="w-full border rounded px-3 py-2"
            />
          </div>
          <div className="flex justify-end gap-2 pt-4">
            <Button variant="secondary" onClick={() => setIsModalOpen(false)}>キャンセル</Button>
            <Button onClick={handleSubmit}>{editingUser ? '更新' : '追加'}</Button>
          </div>
        </div>
      </Modal>
    </div>
  )
}
