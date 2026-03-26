interface LoaderProps {
  message?: string
}

export default function Loader({ message = '読み込み中...' }: LoaderProps) {
  return (
    <div className="fixed inset-0 bg-white/80 flex flex-col justify-center items-center z-[9999]">
      <div className="w-12 h-12 border-4 border-gray-200 border-t-primary-900 rounded-full animate-spin"></div>
      <p className="mt-4 font-medium text-gray-700">{message}</p>
    </div>
  )
}
