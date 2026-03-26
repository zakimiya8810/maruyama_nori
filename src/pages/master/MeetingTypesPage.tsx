import { useState, useEffect } from 'react'
import { supabase } from '../../lib/supabase'
import { Button } from '../../components/common'

interface MeetingType {
  id: string
  name: string
  created_at: string
}

export default function MeetingTypesPage() {
  const [meetingTypes, setMeetingTypes] = useState<MeetingType[]>([])
  const [loading, setLoading] = useState(true)
  const [newTypeName, setNewTypeName] = useState('')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editingName, setEditingName] = useState('')

  useEffect(() => {
    fetchMeetingTypes()
  }, [])

  async function fetchMeetingTypes() {
    setLoading(true)
    const { data, error } = await supabase
      .from('meeting_types')
      .select('*')
      .order('name')

    if (!error && data) {
      setMeetingTypes(data)
    }
    setLoading(false)
  }

  async function handleAdd() {
    if (!newTypeName.trim()) return

    const { error } = await supabase
      .from('meeting_types')
      .insert({ name: newTypeName.trim() })

    if (!error) {
      setNewTypeName('')
      fetchMeetingTypes()
    }
  }

  async function handleUpdate(id: string) {
    if (!editingName.trim()) return

    const { error } = await supabase
      .from('meeting_types')
      .update({ name: editingName.trim() })
      .eq('id', id)

    if (!error) {
      setEditingId(null)
      setEditingName('')
      fetchMeetingTypes()
    }
  }

  async function handleDelete(id: string) {
    if (!confirm('この商談種別を削除してもよろしいですか？')) return

    const { error } = await supabase
      .from('meeting_types')
      .delete()
      .eq('id', id)

    if (!error) {
      fetchMeetingTypes()
    }
  }

  return (
    <div className="p-6">
      <h1 className="text-2xl font-bold mb-6">商談種別マスタ</h1>

      <div className="bg-white rounded-lg shadow">
        <div className="p-4 border-b flex gap-2">
          <input
            type="text"
            value={newTypeName}
            onChange={(e) => setNewTypeName(e.target.value)}
            placeholder="新しい商談種別名"
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
                <th className="text-left px-4 py-3 font-medium">商談種別名</th>
                <th className="text-right px-4 py-3 font-medium w-32">操作</th>
              </tr>
            </thead>
            <tbody>
              {meetingTypes.map((type) => (
                <tr key={type.id} className="border-t">
                  <td className="px-4 py-3">
                    {editingId === type.id ? (
                      <input
                        type="text"
                        value={editingName}
                        onChange={(e) => setEditingName(e.target.value)}
                        className="border rounded px-2 py-1 w-full"
                        onKeyDown={(e) => e.key === 'Enter' && handleUpdate(type.id)}
                        autoFocus
                      />
                    ) : (
                      type.name
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {editingId === type.id ? (
                      <div className="flex justify-end gap-2">
                        <button
                          onClick={() => handleUpdate(type.id)}
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
                            setEditingId(type.id)
                            setEditingName(type.name)
                          }}
                          className="text-blue-600 hover:underline text-sm"
                        >
                          編集
                        </button>
                        <button
                          onClick={() => handleDelete(type.id)}
                          className="text-red-600 hover:underline text-sm"
                        >
                          削除
                        </button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
              {meetingTypes.length === 0 && (
                <tr>
                  <td colSpan={2} className="px-4 py-8 text-center text-gray-500">
                    商談種別が登録されていません
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
