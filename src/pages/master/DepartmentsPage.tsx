import { useState, useEffect } from 'react'
import { supabase } from '../../lib/supabase'
import { Button } from '../../components/common'

interface Department {
  id: string
  name: string
  created_at: string
}

export default function DepartmentsPage() {
  const [departments, setDepartments] = useState<Department[]>([])
  const [loading, setLoading] = useState(true)
  const [newDepartmentName, setNewDepartmentName] = useState('')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editingName, setEditingName] = useState('')

  useEffect(() => {
    fetchDepartments()
  }, [])

  async function fetchDepartments() {
    setLoading(true)
    const { data, error } = await supabase
      .from('departments')
      .select('*')
      .order('name')

    if (!error && data) {
      setDepartments(data)
    }
    setLoading(false)
  }

  async function handleAdd() {
    if (!newDepartmentName.trim()) return

    const { error } = await supabase
      .from('departments')
      .insert({ name: newDepartmentName.trim() })

    if (!error) {
      setNewDepartmentName('')
      fetchDepartments()
    }
  }

  async function handleUpdate(id: string) {
    if (!editingName.trim()) return

    const { error } = await supabase
      .from('departments')
      .update({ name: editingName.trim() })
      .eq('id', id)

    if (!error) {
      setEditingId(null)
      setEditingName('')
      fetchDepartments()
    }
  }

  async function handleDelete(id: string) {
    if (!confirm('この部署を削除してもよろしいですか？')) return

    const { error } = await supabase
      .from('departments')
      .delete()
      .eq('id', id)

    if (!error) {
      fetchDepartments()
    }
  }

  return (
    <div className="p-6">
      <h1 className="text-2xl font-bold mb-6">部署マスタ</h1>

      <div className="bg-white rounded-lg shadow">
        <div className="p-4 border-b flex gap-2">
          <input
            type="text"
            value={newDepartmentName}
            onChange={(e) => setNewDepartmentName(e.target.value)}
            placeholder="新しい部署名"
            className="flex-1 border rounded px-3 py-2"
            onKeyDown={(e) => e.key === 'Enter' && handleAdd()}
          />
          <Button onClick={handleAdd}>追加</Button>
        </div>

        {loading ? (
          <div className="p-4 text-center text-gray-500">読み込み中...</div>
        ) : (
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="text-left px-4 py-3 font-medium">部署名</th>
                <th className="text-right px-4 py-3 font-medium w-32">操作</th>
              </tr>
            </thead>
            <tbody>
              {departments.map((dept) => (
                <tr key={dept.id} className="border-t">
                  <td className="px-4 py-3">
                    {editingId === dept.id ? (
                      <input
                        type="text"
                        value={editingName}
                        onChange={(e) => setEditingName(e.target.value)}
                        className="border rounded px-2 py-1 w-full"
                        onKeyDown={(e) => e.key === 'Enter' && handleUpdate(dept.id)}
                        autoFocus
                      />
                    ) : (
                      dept.name
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {editingId === dept.id ? (
                      <div className="flex justify-end gap-2">
                        <button
                          onClick={() => handleUpdate(dept.id)}
                          className="text-blue-600 hover:underline text-sm"
                        >
                          保存
                        </button>
                        <button
                          onClick={() => setEditingId(null)}
                          className="text-gray-600 hover:underline text-sm"
                        >
                          キャンセル
                        </button>
                      </div>
                    ) : (
                      <div className="flex justify-end gap-2">
                        <button
                          onClick={() => {
                            setEditingId(dept.id)
                            setEditingName(dept.name)
                          }}
                          className="text-blue-600 hover:underline text-sm"
                        >
                          編集
                        </button>
                        <button
                          onClick={() => handleDelete(dept.id)}
                          className="text-red-600 hover:underline text-sm"
                        >
                          削除
                        </button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
              {departments.length === 0 && (
                <tr>
                  <td colSpan={2} className="px-4 py-8 text-center text-gray-500">
                    部署が登録されていません
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  )
}
