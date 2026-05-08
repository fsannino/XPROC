'use client'

export default function DashboardError({
  error,
  reset,
}: {
  error: Error & { digest?: string }
  reset: () => void
}) {
  return (
    <div className="flex flex-col items-center justify-center min-h-64 gap-4 text-center bg-white border border-[#E2E8F0] border-t-4 border-t-[#E05040] rounded-lg p-10">
      <p className="text-[10px] font-bold tracking-[0.2em] uppercase text-[#E05040]">
        Falha inesperada
      </p>
      <h2 className="font-display text-2xl text-navy">Algo deu errado</h2>
      <p className="text-sm text-gray-medium max-w-sm">
        {error.message || 'Ocorreu um erro inesperado ao processar sua solicitação.'}
      </p>
      <button
        onClick={reset}
        className="bg-navy hover:bg-teal text-white px-5 py-2.5 rounded-md text-sm font-semibold transition-all hover:-translate-y-0.5 hover:shadow-md"
      >
        Tentar novamente
      </button>
    </div>
  )
}
