import { useState, useEffect } from 'react'
import { useParams, useNavigate, Link } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { Loader, Button, Modal } from '../components/common'
import type { Customer, CustomerPrice, Meeting, Jutaikin, CustomerSales } from '../types'

interface CustomerDetails {
  basicInfo: Customer
  prices: CustomerPrice[]
  meetings: Meeting[]
  jutaikin: Jutaikin[]
  sales: CustomerSales[]
}

type TabType = 'basic' | 'prices' | 'meetings' | 'sales' | 'jutaikin'

export default function CustomerDetailPage() {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()
  const [details, setDetails] = useState<CustomerDetails | null>(null)
  const [isLoading, setIsLoading] = useState(true)
  const [activeTab, setActiveTab] = useState<TabType>('basic')
  const [showNewMeetingModal, setShowNewMeetingModal] = useState(false)

  useEffect(() => {
    if (id) fetchCustomerDetails(id)
  }, [id])

  const fetchCustomerDetails = async (customerId: string) => {
    setIsLoading(true)
    try {
      const { data: customerData, error: customerError } = await supabase
        .from('customers')
        .select(`
          *,
          business_types(name),
          ranks(name),
          payment_terms(pattern),
          employees:handler_employee_id(name, department_id, division, departments(name))
        `)
        .eq('id', customerId)
        .maybeSingle()

      if (customerError) throw customerError
      if (!customerData) {
        navigate('/customers')
        return
      }

      const handler = customerData.employees as {
        name: string
        division: string
        departments: { name: string } | null
      } | null

      const basicInfo: Customer = {
        id: customerData.id,
        customerCode: customerData.customer_code,
        name: customerData.name,
        shortName: customerData.short_name || '',
        kanaName: customerData.kana_name || '',
        postalCode: customerData.postal_code || '',
        address1: customerData.address_1 || '',
        address2: customerData.address_2 || '',
        tel: customerData.tel || '',
        fax: customerData.fax || '',
        businessTypeId: customerData.business_type_id || '',
        businessTypeName: (customerData.business_types as { name: string } | null)?.name || '',
        rankId: customerData.rank_id || '',
        rankName: (customerData.ranks as { name: string } | null)?.name || '',
        handlerEmployeeId: customerData.handler_employee_id || '',
        handlerName: handler?.name || '',
        departmentName: handler?.departments?.name || '',
        division: handler?.division || '',
        billingCustomerId: customerData.billing_customer_id || '',
        paymentTermId: customerData.payment_term_id || '',
        paymentTermName: (customerData.payment_terms as { pattern: string } | null)?.pattern || '',
        taxMethod: customerData.tax_method || '',
        creditLimit: customerData.credit_limit || 0,
        shippingLocation: customerData.shipping_location || '',
        deliveryMethod: customerData.delivery_method || '',
        groupCode: customerData.group_code || '',
        representativeName: customerData.representative_name || '',
        contactPerson: customerData.contact_person || '',
        noriUsage: customerData.nori_usage || 0,
        teaUsage: customerData.tea_usage || 0,
        naturalFoodHandling: customerData.natural_food_handling || '',
        shippingFee: customerData.shipping_fee || '',
        registrationDate: customerData.registration_date || '',
        tradeStartDate: customerData.trade_start_date || '',
        deleteFlag: customerData.delete_flag || 0,
        billingPostalCode: customerData.billing_postal_code || '',
        billingName: customerData.billing_name || '',
        billingAddress1: customerData.billing_address_1 || '',
        billingAddress2: customerData.billing_address_2 || '',
        billingTel: customerData.billing_tel || '',
        billingFax: customerData.billing_fax || '',
        needsAlert: false,
        hasJutaikin: false,
        isHidden: false,
      }

      const [
        { data: pricesData },
        { data: meetingsData },
        { data: jutaikinData },
        { data: salesData },
      ] = await Promise.all([
        supabase
          .from('customer_prices')
          .select('*, products(product_code, name)')
          .eq('customer_id', customerId)
          .order('products(product_code)'),
        supabase
          .from('meetings')
          .select('*, employees:handler_id(name, departments(name))')
          .eq('customer_id', customerId)
          .order('schedule_date', { ascending: false }),
        supabase.from('jutaikin').select('*').eq('customer_id', customerId),
        supabase
          .from('customer_sales')
          .select('*')
          .eq('customer_id', customerId)
          .order('year', { ascending: false })
          .order('month', { ascending: false }),
      ])

      const prices: CustomerPrice[] = (pricesData || []).map((p) => ({
        id: p.id,
        customerId: p.customer_id,
        productId: p.product_id,
        productCode: (p.products as { product_code: string } | null)?.product_code || '',
        productName: (p.products as { name: string } | null)?.name || '',
        price: p.price || 0,
        effectiveDate: p.effective_date || '',
      }))

      const meetings: Meeting[] = (meetingsData || []).map((m) => {
        const mHandler = m.employees as {
          name: string
          departments: { name: string } | null
        } | null
        return {
          id: m.id,
          customerId: m.customer_id,
          customerCode: basicInfo.customerCode,
          customerName: basicInfo.name,
          rankName: basicInfo.rankName,
          handlerId: m.handler_id || '',
          handlerName: mHandler?.name || '',
          departmentName: mHandler?.departments?.name || '',
          division: '',
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

      const jutaikin: Jutaikin[] = (jutaikinData || []).map((j) => ({
        id: j.id,
        customerId: j.customer_id,
        amount: j.amount || 0,
        baseYearMonth: j.base_year_month || '',
      }))

      const sales: CustomerSales[] = (salesData || []).map((s) => ({
        id: s.id,
        customerId: s.customer_id,
        year: s.year,
        month: s.month,
        amount: s.amount || 0,
      }))

      setDetails({ basicInfo, prices, meetings, jutaikin, sales })
    } catch (err) {
      console.error('Failed to fetch customer details:', err)
    } finally {
      setIsLoading(false)
    }
  }

  if (isLoading) {
    return <Loader message="顧客情報を読み込み中..." />
  }

  if (!details) {
    return (
      <div className="p-6 text-center">
        <p className="text-gray-500">顧客が見つかりません</p>
        <Link to="/customers" className="text-primary-600 hover:underline">
          顧客一覧に戻る
        </Link>
      </div>
    )
  }

  const { basicInfo, prices, meetings, jutaikin, sales } = details

  return (
    <div className="p-6">
      <div className="flex justify-between items-start mb-6 flex-wrap gap-4">
        <div>
          <div className="flex items-center gap-3 mb-2">
            <button onClick={() => navigate(-1)} className="text-gray-500 hover:text-gray-700">
              <i className="fas fa-arrow-left mr-2"></i>
              戻る
            </button>
          </div>
          <h2 className="text-2xl font-bold text-primary-900">
            {basicInfo.name}
            <span className="ml-3 text-base font-normal text-gray-500">
              ({basicInfo.customerCode})
            </span>
          </h2>
          <div className="flex items-center gap-3 mt-1 text-sm text-gray-600">
            <span>
              <i className="fas fa-tag mr-1"></i>
              {basicInfo.rankName || '未設定'}
            </span>
            <span>
              <i className="fas fa-user mr-1"></i>
              {basicInfo.handlerName || '未設定'}
            </span>
            <span>
              <i className="fas fa-building mr-1"></i>
              {basicInfo.departmentName || '未設定'}
            </span>
          </div>
        </div>
        <div className="flex gap-3">
          <Button onClick={() => setShowNewMeetingModal(true)}>
            <i className="fas fa-calendar-plus mr-2"></i>
            商談登録
          </Button>
          <Link to={`/customers/${id}/edit`}>
            <Button variant="secondary">
              <i className="fas fa-edit mr-2"></i>
              顧客情報修正申請
            </Button>
          </Link>
        </div>
      </div>

      <div className="flex gap-0 mb-5 border-b-2 border-gray-200">
        {[
          { key: 'basic', label: '基本情報', icon: 'fa-info-circle' },
          { key: 'prices', label: `単価情報 (${prices.length})`, icon: 'fa-yen-sign' },
          { key: 'meetings', label: `商談履歴 (${meetings.length})`, icon: 'fa-handshake' },
          { key: 'sales', label: '売上推移', icon: 'fa-chart-bar' },
          { key: 'jutaikin', label: `渋滞金 (${jutaikin.length})`, icon: 'fa-exclamation-triangle' },
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
            <i className={`fas ${tab.icon} mr-2`}></i>
            {tab.label}
          </button>
        ))}
      </div>

      {activeTab === 'basic' && <BasicInfoTab info={basicInfo} />}
      {activeTab === 'prices' && <PricesTab prices={prices} customerId={basicInfo.id} />}
      {activeTab === 'meetings' && <MeetingsTab meetings={meetings} />}
      {activeTab === 'sales' && <SalesTab sales={sales} />}
      {activeTab === 'jutaikin' && <JutaikinTab jutaikin={jutaikin} />}

      <Modal
        isOpen={showNewMeetingModal}
        onClose={() => setShowNewMeetingModal(false)}
        title="商談予定登録"
      >
        <p className="text-gray-500">商談予定の登録機能は商談管理ページから行えます。</p>
        <div className="mt-4 flex gap-3 justify-end">
          <Button variant="secondary" onClick={() => setShowNewMeetingModal(false)}>
            閉じる
          </Button>
          <Link to={`/meetings/new?customerId=${id}`}>
            <Button>商談登録へ</Button>
          </Link>
        </div>
      </Modal>
    </div>
  )
}

function BasicInfoTab({ info }: { info: Customer }) {
  return (
    <div className="space-y-6">
      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          基本情報
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <InfoItem label="得意先コード" value={info.customerCode} />
          <InfoItem label="得意先名称" value={info.name} />
          <InfoItem label="略名称" value={info.shortName} />
          <InfoItem label="カナ名称" value={info.kanaName} />
          <InfoItem label="郵便番号" value={info.postalCode} />
          <InfoItem label="住所1" value={info.address1} />
          <InfoItem label="住所2" value={info.address2} />
          <InfoItem label="TEL" value={info.tel} />
          <InfoItem label="FAX" value={info.fax} />
          <InfoItem label="業態" value={info.businessTypeName} />
          <InfoItem label="ランク" value={info.rankName} />
          <InfoItem label="担当者" value={info.handlerName} />
        </div>
      </section>

      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          取引条件
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <InfoItem label="締日パターン" value={info.paymentTermName} />
          <InfoItem label="課税方式" value={info.taxMethod} />
          <InfoItem label="与信限度額" value={`${info.creditLimit.toLocaleString()}円`} />
          <InfoItem label="出荷場所" value={info.shippingLocation} />
          <InfoItem label="配送方法" value={info.deliveryMethod} />
          <InfoItem label="得意先グループコード" value={info.groupCode} />
          <InfoItem label="配送料" value={info.shippingFee} />
        </div>
      </section>

      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          担当者・代表者情報
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <InfoItem label="代表者氏名" value={info.representativeName} />
          <InfoItem label="得意先担当者" value={info.contactPerson} />
        </div>
      </section>

      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
          商品取扱
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <InfoItem label="海苔使用本数" value={info.noriUsage.toString()} />
          <InfoItem label="お茶使用量" value={info.teaUsage.toString()} />
          <InfoItem label="自然食品取扱" value={info.naturalFoodHandling} />
        </div>
      </section>

      {info.billingName && (
        <section className="bg-white rounded-lg border border-gray-200 p-6">
          <h3 className="text-lg font-medium text-primary-900 border-b-2 border-primary-900 pb-2 mb-6">
            請求先情報
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            <InfoItem label="請求先名称" value={info.billingName} />
            <InfoItem label="請求先郵便番号" value={info.billingPostalCode} />
            <InfoItem label="請求先住所1" value={info.billingAddress1} />
            <InfoItem label="請求先住所2" value={info.billingAddress2} />
            <InfoItem label="請求先TEL" value={info.billingTel} />
            <InfoItem label="請求先FAX" value={info.billingFax} />
          </div>
        </section>
      )}
    </div>
  )
}

function PricesTab({ prices, customerId }: { prices: CustomerPrice[]; customerId: string }) {
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6">
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-lg font-medium text-primary-900">単価情報</h3>
        <Link to={`/customers/${customerId}/prices/edit`}>
          <Button size="sm">
            <i className="fas fa-edit mr-2"></i>
            単価修正申請
          </Button>
        </Link>
      </div>
      {prices.length === 0 ? (
        <p className="text-gray-500 text-center py-8">単価情報がありません</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr className="bg-gray-50">
                <th className="px-4 py-3 text-left text-sm font-medium border-b">商品コード</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">商品名</th>
                <th className="px-4 py-3 text-right text-sm font-medium border-b">単価</th>
                <th className="px-4 py-3 text-left text-sm font-medium border-b">有効日</th>
              </tr>
            </thead>
            <tbody>
              {prices.map((price) => (
                <tr key={price.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 border-b">{price.productCode}</td>
                  <td className="px-4 py-3 border-b">{price.productName}</td>
                  <td className="px-4 py-3 border-b text-right">
                    {price.price.toLocaleString()}円
                  </td>
                  <td className="px-4 py-3 border-b">{price.effectiveDate || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function MeetingsTab({ meetings }: { meetings: Meeting[] }) {
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

  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6">
      <h3 className="text-lg font-medium text-primary-900 mb-4">商談履歴</h3>
      {meetings.length === 0 ? (
        <p className="text-gray-500 text-center py-8">商談履歴がありません</p>
      ) : (
        <div className="space-y-4">
          {meetings.map((meeting) => (
            <div
              key={meeting.id}
              className="border border-gray-200 rounded-lg p-4 hover:shadow-md transition-shadow"
            >
              <div className="flex justify-between items-start mb-2">
                <div>
                  <span className="font-medium">{meeting.scheduleDate}</span>
                  {meeting.actualDate && (
                    <span className="text-gray-500 ml-2">(実施: {meeting.actualDate})</span>
                  )}
                </div>
                <span
                  className={`px-2 py-1 text-xs rounded ${statusStyles[meeting.status] || 'bg-gray-100'}`}
                >
                  {statusLabels[meeting.status] || meeting.status}
                </span>
              </div>
              <div className="text-sm text-gray-600 mb-2">
                <span className="mr-4">
                  <i className="fas fa-user mr-1"></i>
                  {meeting.handlerName}
                </span>
                <span>
                  <i className="fas fa-bullseye mr-1"></i>
                  {meeting.purpose || '目的未設定'}
                </span>
              </div>
              {meeting.result && (
                <div className="text-sm">
                  <strong>結果:</strong> {meeting.result}
                </div>
              )}
              {meeting.resultNotes && (
                <div className="text-sm text-gray-600 mt-1">{meeting.resultNotes}</div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

function SalesTab({ sales }: { sales: CustomerSales[] }) {
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6">
      <h3 className="text-lg font-medium text-primary-900 mb-4">売上推移</h3>
      {sales.length === 0 ? (
        <p className="text-gray-500 text-center py-8">売上データがありません</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr className="bg-gray-50">
                <th className="px-4 py-3 text-left text-sm font-medium border-b">年月</th>
                <th className="px-4 py-3 text-right text-sm font-medium border-b">売上金額</th>
              </tr>
            </thead>
            <tbody>
              {sales.map((s) => (
                <tr key={s.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 border-b">
                    {s.year}年{s.month}月
                  </td>
                  <td className="px-4 py-3 border-b text-right">{s.amount.toLocaleString()}円</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function JutaikinTab({ jutaikin }: { jutaikin: Jutaikin[] }) {
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6">
      <h3 className="text-lg font-medium text-primary-900 mb-4">渋滞金情報</h3>
      {jutaikin.length === 0 ? (
        <p className="text-gray-500 text-center py-8">渋滞金データがありません</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead>
              <tr className="bg-gray-50">
                <th className="px-4 py-3 text-left text-sm font-medium border-b">基準年月</th>
                <th className="px-4 py-3 text-right text-sm font-medium border-b">金額</th>
              </tr>
            </thead>
            <tbody>
              {jutaikin.map((j) => (
                <tr key={j.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 border-b">{j.baseYearMonth}</td>
                  <td className="px-4 py-3 border-b text-right text-red-600 font-medium">
                    {j.amount.toLocaleString()}円
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
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
