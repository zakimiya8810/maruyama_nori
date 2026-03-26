import { useState } from 'react'

type Mode = 'user' | 'admin' | 'dev'

export default function ManualPage() {
  const [mode, setMode] = useState<Mode>('user')

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-primary-500 text-white py-6 px-10 sticky top-0 z-50 shadow-md">
        <h1 className="text-2xl font-bold mb-1">顧客カルテ 操作マニュアル</h1>
        <p className="text-sm opacity-90">株式会社丸山海苔店 - 顧客管理・申請承認システム</p>
      </div>

      <div className="bg-white border-b border-gray-200 py-5 px-10 flex items-center gap-3">
        <label className="font-semibold mr-4">閲覧モード:</label>
        {[
          { key: 'user', label: 'エンドユーザー向け', icon: 'fa-user' },
          { key: 'admin', label: '管理者向け', icon: 'fa-cog' },
          { key: 'dev', label: '開発者向け', icon: 'fa-code' },
        ].map((m) => (
          <button
            key={m.key}
            onClick={() => setMode(m.key as Mode)}
            className={`px-5 py-2.5 rounded-lg border-2 font-semibold transition-all ${
              mode === m.key
                ? 'bg-primary-500 text-white border-primary-500'
                : 'bg-white text-gray-700 border-gray-300 hover:border-primary-300 hover:shadow'
            }`}
          >
            <i className={`fas ${m.icon} mr-2`}></i>
            {m.label}
          </button>
        ))}
      </div>

      <div className="max-w-5xl mx-auto p-10">
        <Section title="システム概要">
          <p className="mb-4">
            顧客カルテは、顧客情報の管理、商談履歴の記録、申請・承認ワークフローを
            一元管理するためのシステムです。
          </p>
          <InfoBox type="info">
            <strong>主な機能</strong>
            <ul className="list-disc list-inside mt-2 space-y-1">
              <li>顧客一覧・詳細表示</li>
              <li>商談予定・履歴管理</li>
              <li>顧客情報修正申請</li>
              <li>単価修正申請</li>
              <li>5段階承認ワークフロー</li>
            </ul>
          </InfoBox>
        </Section>

        <Section title="ログイン方法">
          <StepGuide
            steps={[
              { title: 'ログイン画面を開く', content: 'ブラウザでシステムのURLにアクセスします。' },
              {
                title: 'ID・パスワードを入力',
                content: '社員マスタに登録されているIDとパスワードを入力します。',
              },
              {
                title: 'ログインボタンをクリック',
                content: '認証が成功するとダッシュボードが表示されます。',
              },
            ]}
          />
          <InfoBox type="warning">
            <strong>注意</strong>: パスワードを忘れた場合は、管理者にお問い合わせください。
          </InfoBox>
        </Section>

        <Section title="顧客一覧の使い方">
          <p className="mb-4">
            サイドメニューの「顧客一覧」をクリックすると、登録されている顧客の一覧が表示されます。
          </p>
          <h4 className="text-lg font-semibold mt-6 mb-3">フィルタ機能</h4>
          <ul className="list-disc list-inside space-y-2">
            <li><strong>検索ボックス</strong>: 顧客名、コード、住所で絞り込み</li>
            <li><strong>大区分</strong>: 営業エリアで絞り込み</li>
            <li><strong>部門</strong>: 部門単位で絞り込み</li>
            <li><strong>担当者</strong>: 担当者で絞り込み</li>
            <li><strong>ランク</strong>: 顧客ランクで絞り込み</li>
            <li><strong>業態</strong>: 業態で絞り込み</li>
          </ul>
          <InfoBox type="success">
            <strong>便利な機能</strong>: 「自分の担当」ボタンで自分が担当する顧客のみ表示できます。
          </InfoBox>
        </Section>

        <Section title="顧客詳細画面">
          <p className="mb-4">顧客一覧から顧客をクリックすると、詳細画面が表示されます。</p>
          <h4 className="text-lg font-semibold mt-6 mb-3">タブの説明</h4>
          <Table
            headers={['タブ', '内容']}
            rows={[
              ['基本情報', '顧客の基本情報（住所、連絡先、取引条件など）'],
              ['単価情報', 'その顧客専用の商品単価一覧'],
              ['商談履歴', '過去の商談記録と今後の予定'],
              ['売上推移', '月別の売上データ'],
              ['渋滞金', '未回収金の情報（ある場合）'],
            ]}
          />
        </Section>

        <Section title="商談管理">
          <p className="mb-4">
            商談の予定登録、実績入力を行います。
          </p>
          <h4 className="text-lg font-semibold mt-6 mb-3">商談ステータス</h4>
          <div className="flex gap-3 flex-wrap">
            <StatusBadge color="blue">予定</StatusBadge>
            <StatusBadge color="green">完了</StatusBadge>
            <StatusBadge color="yellow">遅延完了</StatusBadge>
            <StatusBadge color="red">対応遅れ</StatusBadge>
          </div>
        </Section>

        <Section title="申請・承認フロー">
          <p className="mb-4">
            顧客情報の新規登録・修正、単価の変更は申請・承認フローを通して行います。
          </p>
          <FlowChart
            steps={[
              '申請中（申請者が申請を作成）',
              '上長承認（同じ大区分の上長が承認）',
              '管理承認（管理部門が承認）',
              '常務承認（常務が承認）',
              '決裁完了（社長が最終承認）',
            ]}
          />
          <InfoBox type="info">
            <strong>ポイント</strong>: 修正申請は常務承認時点でマスタに先行反映されます。
            社長却下時には自動でロールバックされます。
          </InfoBox>
        </Section>

        {mode !== 'user' && (
          <Section title="管理者向け機能">
            <p className="mb-4">
              管理権限を持つユーザーは、以下の追加機能を利用できます。
            </p>
            <ul className="list-disc list-inside space-y-2">
              <li>全ての申請を閲覧・承認</li>
              <li>得意先コード・グループコードの採番</li>
              <li>マスタデータの管理</li>
            </ul>
          </Section>
        )}

        {mode === 'dev' && (
          <Section title="開発者向け情報">
            <h4 className="text-lg font-semibold mt-6 mb-3">技術スタック</h4>
            <ul className="list-disc list-inside space-y-2">
              <li><strong>フロントエンド</strong>: React + TypeScript + Vite</li>
              <li><strong>スタイリング</strong>: Tailwind CSS</li>
              <li><strong>状態管理</strong>: Zustand</li>
              <li><strong>データベース</strong>: Supabase (PostgreSQL)</li>
              <li><strong>認証</strong>: カスタム認証 + Supabase RLS</li>
            </ul>
            <h4 className="text-lg font-semibold mt-6 mb-3">データベーステーブル</h4>
            <CodeBlock>{`
-- 主要テーブル
departments      -- 部門マスタ
employees        -- 社員マスタ
customers        -- 得意先マスタ
products         -- 商品マスタ
customer_prices  -- 単価マスタ
meetings         -- 商談管理
applications     -- 申請管理
            `.trim()}</CodeBlock>
          </Section>
        )}

        <Section title="お問い合わせ">
          <p className="mb-4">
            システムに関するお問い合わせは、管理部門までご連絡ください。
          </p>
        </Section>
      </div>
    </div>
  )
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="bg-white rounded-xl shadow-sm p-10 mb-8">
      <h2 className="text-2xl font-bold text-primary-500 border-b-[3px] border-primary-500 pb-3 mb-6">
        {title}
      </h2>
      {children}
    </section>
  )
}

function InfoBox({ type, children }: { type: 'info' | 'warning' | 'success' | 'danger'; children: React.ReactNode }) {
  const styles = {
    info: 'bg-blue-50 border-blue-500',
    warning: 'bg-orange-50 border-orange-500',
    success: 'bg-green-50 border-green-500',
    danger: 'bg-red-50 border-red-500',
  }

  return (
    <div className={`p-5 rounded-lg border-l-4 my-5 ${styles[type]}`}>
      {children}
    </div>
  )
}

function StepGuide({ steps }: { steps: { title: string; content: string }[] }) {
  return (
    <div className="my-6">
      {steps.map((step, idx) => (
        <div
          key={idx}
          className="relative pl-20 py-5 mb-5 bg-gray-50 rounded-lg border-l-[3px] border-primary-500"
        >
          <div className="absolute left-5 top-1/2 -translate-y-1/2 w-10 h-10 bg-primary-500 text-white rounded-full flex items-center justify-center font-bold">
            {idx + 1}
          </div>
          <h4 className="text-primary-500 font-semibold mb-1">{step.title}</h4>
          <p className="text-gray-600">{step.content}</p>
        </div>
      ))}
    </div>
  )
}

function Table({ headers, rows }: { headers: string[]; rows: string[][] }) {
  return (
    <div className="overflow-x-auto my-5">
      <table className="w-full border border-gray-200 rounded-lg overflow-hidden">
        <thead>
          <tr className="bg-primary-500 text-white">
            {headers.map((h, i) => (
              <th key={i} className="px-4 py-3 text-left font-semibold text-sm">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri} className="hover:bg-gray-50">
              {row.map((cell, ci) => (
                <td key={ci} className="px-4 py-3 border-b border-gray-200">
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function StatusBadge({ color, children }: { color: 'blue' | 'green' | 'yellow' | 'red'; children: React.ReactNode }) {
  const styles = {
    blue: 'bg-blue-100 text-blue-800',
    green: 'bg-green-100 text-green-800',
    yellow: 'bg-yellow-100 text-yellow-800',
    red: 'bg-red-100 text-red-800',
  }

  return (
    <span className={`px-3 py-1.5 rounded text-sm font-medium ${styles[color]}`}>
      {children}
    </span>
  )
}

function FlowChart({ steps }: { steps: string[] }) {
  return (
    <div className="my-6 p-5 bg-gray-50 rounded-lg border border-gray-200">
      {steps.map((step, idx) => (
        <div key={idx}>
          <div className="py-4 px-5 bg-white rounded-lg border-l-4 border-primary-500">
            <span className="font-semibold text-primary-500 mr-2">{idx + 1}.</span>
            {step}
          </div>
          {idx < steps.length - 1 && (
            <div className="text-center text-2xl text-primary-500 my-2">
              <i className="fas fa-arrow-down"></i>
            </div>
          )}
        </div>
      ))}
    </div>
  )
}

function CodeBlock({ children }: { children: React.ReactNode }) {
  return (
    <pre className="bg-gray-800 text-gray-100 rounded-lg p-5 overflow-x-auto text-sm my-5">
      <code>{children}</code>
    </pre>
  )
}
