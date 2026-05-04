'use client'

export default function DashboardError({
  error,
  reset,
}: {
  error: Error & { digest?: string }
  reset: () => void
}) {
  return (
    <div className="flex flex-col items-center justify-center min-h-64 gap-4 text-center">
      <div className="text-4xl">⚠️</div>
      <h2 className="text-lg font-semibold text-gray-800">Algo deu errado</h2>
      <p className="text-sm text-gray-500 max-w-sm">{error.message || 'Ocorreu um erro inesperado.'}</p>
      <button
        onClick={reset}
        className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
      >
        Tentar novamente
      </button>
    </div>
  )
}
