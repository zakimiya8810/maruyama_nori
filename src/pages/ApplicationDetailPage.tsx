import { useState, useEffect } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { useAuthStore } from '../stores/authStore'
import { Loader, Button, Modal } from '../components/common'
import type { ApplicationDetail, ApprovalStage } from '../types'

export default function ApplicationDetailPage() {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()
  const [detail, setDetail] = useState<ApplicationDetail | null>(null)
  const [isLoading, setIsLoading] = useState(true)
  const [isProcessing, setIsProcessing] = useState(false)
  const [showRejectModal, setShowRejectModal] = useState(false)
  const [rejectReason, setRejectReason] = useState('')
  const [requiredFields, setRequiredFields] = useState<string[]>([])

  const { user, getUserRole } = useAuthStore()
  const userRole = getUserRole()

  useEffect(() => {
    if (id) fetchApplicationDetail(id)
  }, [id])

  const fetchApplicationDetail = async (appId: string) => {
    setIsLoading(true)
    try {
      const { data: appData, error: appError } = await supabase
        .from('applications')
        .select(`
          *,
          customers(name),
          employees:applicant_id(name, email, departments(name))
        `)
        .eq('id', appId)
        .maybeSingle()

      if (appError) throw appError
      if (!appData) {
        navigate('/applications')
        return
      }

      const applicant = appData.employees as {
        name: string
        email: string
        departments: { name: string } | null
      } | null

      const [{ data: customerData }, { data: priceData }] = await Promise.all([
        supabase.from('application_customer_data').select('*').eq('application_id', appId),
        supabase.from('application_price_data').select('*').eq('application_id', appId),
      ])

      const detail: ApplicationDetail = {
        id: appData.id,
        applicationId: appData.application_id,
        groupId: appData.group_id || appData.application_id,
        applicationType: appData.application_type,
        targetMaster: appData.target_master || '',
        customerId: appData.customer_id || '',
        customerCode: appData.customer_code || '',
        customerName: appData.customer_name || (appData.customers as { name: string } | null)?.name || '',
        applicantId: appData.applicant_id || '',
        applicantName: appData.applicant_name || applicant?.name || '',
        applicantDepartment: applicant?.departments?.name || '',
        applicationDate: appData.application_date,
        status: appData.status,
        stage: appData.stage,
        effectiveDate: appData.effective_date || '',
        reflectionStatus: appData.reflection_status || '',
        rejectReason: appData.reject_reason || '',
        requiredFields: appData.required_fields || [],
        resubmitCount: appData.resubmit_count || 0,
        originalApplicationId: appData.original_application_id || '',
        detailData: (customerData || []).map((d) => ({
          field: d.field_name,
          before: d.before_value || '',
          after: d.after_value || '',
        })),
        prices: (priceData || []).map((p) => ({
          productCode: p.product_code,
          productName: p.product_name || '',
          registrationType: p.registration_type || '',
          priceBefore: p.price_before,
          priceAfter: p.price_after,
        })),
        supervisorApprover: appData.supervisor_approver,
        supervisorApprovedAt: appData.supervisor_approved_at,
        adminApprover: appData.admin_approver,
        adminApprovedAt: appData.admin_approved_at,
        directorApprover: appData.director_approver,
        directorApprovedAt: appData.director_approved_at,
        ceoApprover: appData.ceo_approver,
        ceoApprovedAt: appData.ceo_approved_at,
      }

      setDetail(detail)
    } catch (err) {
      console.error('Failed to fetch application detail:', err)
    } finally {
      setIsLoading(false)
    }
  }

  const handleApprove = async () => {
    if (!detail || !user) return
    setIsProcessing(true)

    try {
      const nextStage = getNextStage(detail.stage as ApprovalStage)
      const updateData: Record<string, unknown> = {
        stage: nextStage,
        status: nextStage === '決裁完了' ? '承認済' : '承認中',
      }

      switch (detail.stage) {
        case '申請中':
          updateData.supervisor_approver = user.name
          updateData.supervisor_approved_at = new Date().toISOString()
          break
        case '上長承認済':
          updateData.admin_approver = user.name
          updateData.admin_approved_at = new Date().toISOString()
          break
        case '管理承認済':
          updateData.director_approver = user.name
          updateData.director_approved_at = new Date().toISOString()
          break
        case '常務承認済':
          updateData.ceo_approver = user.name
          updateData.ceo_approved_at = new Date().toISOString()
          break
      }

      const { error } = await supabase.from('applications').update(updateData).eq('id', detail.id)

      if (error) throw error

      alert('承認しました')
      fetchApplicationDetail(detail.id)
    } catch (err) {
      console.error('Approve error:', err)
      alert('承認処理に失敗しました')
    } finally {
      setIsProcessing(false)
    }
  }

  const handleReject = async () => {
    if (!detail || !user || !rejectReason.trim()) {
      alert('却下理由を入力してください')
      return
    }
    setIsProcessing(true)

    try {
      const { error } = await supabase
        .from('applications')
        .update({
          stage: '却下',
          status: '却下',
          reject_reason: rejectReason,
          required_fields: requiredFields,
        })
        .eq('id', detail.id)

      if (error) throw error

      setShowRejectModal(false)
      alert('却下しました')
      fetchApplicationDetail(detail.id)
    } catch (err) {
      console.error('Reject error:', err)
      alert('却下処理に失敗しました')
    } finally {
      setIsProcessing(false)
    }
  }

  const canApproveThis = (): boolean => {
    if (!detail) return false
    const stage = detail.stage as ApprovalStage

    switch (stage) {
      case '申請中':
        return ['supervisor', 'admin', 'director', 'ceo'].includes(userRole)
      case '上長承認済':
        return ['admin', 'director', 'ceo'].includes(userRole)
      case '管理承認済':
        return ['director', 'ceo'].includes(userRole)
      case '常務承認済':
        return userRole === 'ceo'
      default:
        return false
    }
  }

  const canResubmit = (): boolean => {
    if (!detail || !user) return false
    return detail.stage === '却下' && detail.applicantId === user.id
  }

  if (isLoading) {
    return <Loader message="申請詳細を読み込み中..." />
  }

  if (!detail) {
    return (
      <div className="p-6 text-center">
        <p className="text-gray-500">申請が見つかりません</p>
      </div>
    )
  }

  const stageStyles: Record<string, string> = {
    申請中: 'bg-blue-100 text-blue-800',
    上長承認済: 'bg-cyan-100 text-cyan-800',
    管理承認済: 'bg-teal-100 text-teal-800',
    常務承認済: 'bg-indigo-100 text-indigo-800',
    決裁完了: 'bg-green-100 text-green-800',
    却下: 'bg-red-100 text-red-800',
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-start mb-6">
        <div>
          <button onClick={() => navigate(-1)} className="text-gray-500 hover:text-gray-700 mb-2">
            <i className="fas fa-arrow-left mr-2"></i>
            戻る
          </button>
          <h2 className="text-2xl font-bold text-primary-900">
            申請詳細
            <span className="ml-3 text-base font-mono text-gray-500">
              {detail.applicationId}
            </span>
          </h2>
        </div>
        <span className={`px-4 py-2 text-sm rounded-lg font-medium ${stageStyles[detail.stage] || 'bg-gray-100'}`}>
          {detail.stage}
        </span>
      </div>

      <section className="bg-white rounded-lg border border-gray-200 p-6 mb-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          申請基本情報
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <InfoItem label="申請ID" value={detail.applicationId} />
          <InfoItem label="申請種別" value={detail.applicationType} />
          <InfoItem label="対象マスタ" value={detail.targetMaster || '-'} />
          <InfoItem label="対象顧客" value={detail.customerName} />
          <InfoItem label="得意先コード" value={detail.customerCode} />
          <InfoItem label="申請者" value={detail.applicantName} />
          <InfoItem label="申請者部門" value={detail.applicantDepartment} />
          <InfoItem label="申請日時" value={formatDate(detail.applicationDate)} />
          <InfoItem label="登録有効日" value={detail.effectiveDate || '-'} />
        </div>
      </section>

      <section className="bg-white rounded-lg border border-gray-200 p-6 mb-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          承認フロー
        </h3>
        <ApprovalFlow detail={detail} />
      </section>

      {detail.detailData && detail.detailData.length > 0 && (
        <section className="bg-white rounded-lg border border-gray-200 p-6 mb-6">
          <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
            顧客情報変更内容
          </h3>
          <div className="overflow-x-auto">
            <table className="w-full border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">項目名</th>
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">修正前</th>
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">修正後</th>
                </tr>
              </thead>
              <tbody>
                {detail.detailData.map((item, idx) => (
                  <tr key={idx} className="hover:bg-gray-50">
                    <td className="px-4 py-3 border-b font-medium">{item.field}</td>
                    <td className="px-4 py-3 border-b text-gray-500">{item.before || '-'}</td>
                    <td className="px-4 py-3 border-b text-primary-700">{item.after || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {detail.prices && detail.prices.length > 0 && (
        <section className="bg-white rounded-lg border border-gray-200 p-6 mb-6">
          <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
            単価変更内容
          </h3>
          <div className="overflow-x-auto">
            <table className="w-full border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">登録区分</th>
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">商品コード</th>
                  <th className="px-4 py-3 text-left text-sm font-medium border-b">商品名</th>
                  <th className="px-4 py-3 text-right text-sm font-medium border-b">変更前単価</th>
                  <th className="px-4 py-3 text-right text-sm font-medium border-b">変更後単価</th>
                </tr>
              </thead>
              <tbody>
                {detail.prices.map((price, idx) => (
                  <tr key={idx} className="hover:bg-gray-50">
                    <td className="px-4 py-3 border-b">
                      <span
                        className={`px-2 py-1 text-xs rounded ${
                          price.registrationType === '追加'
                            ? 'bg-green-100 text-green-800'
                            : price.registrationType === '削除'
                            ? 'bg-red-100 text-red-800'
                            : 'bg-yellow-100 text-yellow-800'
                        }`}
                      >
                        {price.registrationType}
                      </span>
                    </td>
                    <td className="px-4 py-3 border-b font-mono">{price.productCode}</td>
                    <td className="px-4 py-3 border-b">{price.productName}</td>
                    <td className="px-4 py-3 border-b text-right text-gray-500">
                      {price.priceBefore != null ? `${price.priceBefore.toLocaleString()}円` : '-'}
                    </td>
                    <td className="px-4 py-3 border-b text-right text-primary-700">
                      {price.priceAfter != null ? `${price.priceAfter.toLocaleString()}円` : '-'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {detail.rejectReason && (
        <section className="bg-red-50 rounded-lg border border-red-200 p-6 mb-6">
          <h3 className="text-lg font-medium text-red-800 mb-4">
            <i className="fas fa-times-circle mr-2"></i>
            却下理由
          </h3>
          <p className="text-red-700 whitespace-pre-wrap">{detail.rejectReason}</p>
          {detail.requiredFields && detail.requiredFields.length > 0 && (
            <div className="mt-4">
              <p className="font-medium text-red-800 mb-2">修正が必要な項目:</p>
              <ul className="list-disc list-inside text-red-700">
                {detail.requiredFields.map((field, idx) => (
                  <li key={idx}>{field}</li>
                ))}
              </ul>
            </div>
          )}
        </section>
      )}

      {(canApproveThis() || canResubmit()) && (
        <div className="flex justify-center gap-6 mt-8">
          {canApproveThis() && (
            <>
              <Button
                size="lg"
                variant="success"
                onClick={handleApprove}
                disabled={isProcessing}
                className="shadow-lg hover:shadow-xl transition-shadow"
              >
                <i className="fas fa-check mr-2"></i>
                承認
              </Button>
              <Button
                size="lg"
                variant="danger"
                onClick={() => setShowRejectModal(true)}
                disabled={isProcessing}
                className="shadow-lg hover:shadow-xl transition-shadow"
              >
                <i className="fas fa-times mr-2"></i>
                却下
              </Button>
            </>
          )}
          {canResubmit() && (
            <Button
              size="lg"
              variant="info"
              onClick={() => alert('再申請機能は開発中です')}
              disabled={isProcessing}
              className="shadow-lg hover:shadow-xl transition-shadow"
            >
              <i className="fas fa-redo mr-2"></i>
              再申請
            </Button>
          )}
        </div>
      )}

      <Modal
        isOpen={showRejectModal}
        onClose={() => setShowRejectModal(false)}
        title="申請を却下"
        footer={
          <>
            <Button variant="secondary" onClick={() => setShowRejectModal(false)}>
              キャンセル
            </Button>
            <Button variant="danger" onClick={handleReject} disabled={isProcessing}>
              却下する
            </Button>
          </>
        }
      >
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium mb-1">却下理由</label>
            <textarea
              value={rejectReason}
              onChange={(e) => setRejectReason(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded focus:outline-none focus:border-primary-500 min-h-[120px]"
              placeholder="却下理由を入力してください"
            />
          </div>
          <div>
            <label className="block text-sm font-medium mb-2">修正が必要な項目</label>
            <div className="space-y-2">
              {['得意先名称', '住所', '連絡先', '取引条件', '単価', 'その他'].map((field) => (
                <label key={field} className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={requiredFields.includes(field)}
                    onChange={(e) => {
                      if (e.target.checked) {
                        setRequiredFields([...requiredFields, field])
                      } else {
                        setRequiredFields(requiredFields.filter((f) => f !== field))
                      }
                    }}
                    className="rounded"
                  />
                  <span>{field}</span>
                </label>
              ))}
            </div>
          </div>
        </div>
      </Modal>
    </div>
  )
}

function ApprovalFlow({ detail }: { detail: ApplicationDetail }) {
  const stages = [
    { key: '申請中', label: '申請', approver: null, approvedAt: null },
    { key: '上長承認済', label: '上長承認', approver: detail.supervisorApprover, approvedAt: detail.supervisorApprovedAt },
    { key: '管理承認済', label: '管理承認', approver: detail.adminApprover, approvedAt: detail.adminApprovedAt },
    { key: '常務承認済', label: '常務承認', approver: detail.directorApprover, approvedAt: detail.directorApprovedAt },
    { key: '決裁完了', label: '決裁', approver: detail.ceoApprover, approvedAt: detail.ceoApprovedAt },
  ]

  const currentIndex = stages.findIndex((s) => s.key === detail.stage)
  const isRejected = detail.stage === '却下'

  return (
    <div className="flex items-center justify-between">
      {stages.map((stage, idx) => {
        const isPast = idx < currentIndex || detail.stage === '決裁完了'
        const isCurrent = idx === currentIndex && !isRejected

        return (
          <div key={stage.key} className="flex items-center flex-1">
            <div className="flex flex-col items-center">
              <div
                className={`w-10 h-10 rounded-full flex items-center justify-center text-sm font-medium ${
                  isPast
                    ? 'bg-green-500 text-white'
                    : isCurrent
                    ? 'bg-primary-500 text-white animate-pulse'
                    : isRejected && idx === currentIndex
                    ? 'bg-red-500 text-white'
                    : 'bg-gray-200 text-gray-500'
                }`}
              >
                {isPast ? (
                  <i className="fas fa-check"></i>
                ) : isRejected && idx === currentIndex ? (
                  <i className="fas fa-times"></i>
                ) : (
                  idx + 1
                )}
              </div>
              <span className="text-xs mt-2 text-center">{stage.label}</span>
              {stage.approver && (
                <span className="text-xs text-gray-500 mt-1">{stage.approver}</span>
              )}
              {stage.approvedAt && (
                <span className="text-xs text-gray-400">{formatDate(stage.approvedAt)}</span>
              )}
            </div>
            {idx < stages.length - 1 && (
              <div
                className={`flex-1 h-1 mx-2 ${
                  isPast ? 'bg-green-500' : 'bg-gray-200'
                }`}
              />
            )}
          </div>
        )
      })}
    </div>
  )
}

function InfoItem({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <span className="text-xs text-gray-500 font-medium">{label}</span>
      <div className="mt-0.5">{value || '-'}</div>
    </div>
  )
}

function getNextStage(current: ApprovalStage): ApprovalStage {
  const flow: Record<ApprovalStage, ApprovalStage> = {
    '申請中': '上長承認済',
    '上長承認済': '管理承認済',
    '管理承認済': '常務承認済',
    '常務承認済': '決裁完了',
    '決裁完了': '決裁完了',
    '却下': '却下',
  }
  return flow[current]
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
