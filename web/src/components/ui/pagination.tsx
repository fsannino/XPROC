import Link from 'next/link'

interface PaginationProps {
  page: number
  total: number
  perPage: number
  basePath: string
  busca?: string
}

export function Pagination({ page, total, perPage, basePath, busca }: PaginationProps) {
  const totalPages = Math.ceil(total / perPage)
  if (totalPages <= 1) return null

  function href(p: number) {
    const params = new URLSearchParams()
    if (busca) params.set('busca', busca)
    params.set('pagina', String(p))
    return `${basePath}?${params.toString()}`
  }

  return (
    <div className="flex items-center justify-between px-4 py-3 border-t border-gray-100 bg-white">
      <span className="text-sm text-gray-500">
        {(page - 1) * perPage + 1}–{Math.min(page * perPage, total)} de {total}
      </span>
      <div className="flex gap-1">
        {page > 1 && (
          <Link href={href(page - 1)} className="px-3 py-1.5 rounded-lg border border-gray-200 text-sm text-gray-600 hover:bg-gray-50">
            ← Anterior
          </Link>
        )}
        {page < totalPages && (
          <Link href={href(page + 1)} className="px-3 py-1.5 rounded-lg border border-gray-200 text-sm text-gray-600 hover:bg-gray-50">
            Próxima →
          </Link>
        )}
      </div>
    </div>
  )
}
