'use client'

export function ExportButton({ tipo }: { tipo: 'processos' | 'transacoes' | 'usuarios' }) {
  return (
    <a
      href={`/api/export?tipo=${tipo}`}
      className="border border-gray-300 text-gray-600 px-3 py-2 rounded-lg text-sm font-medium hover:bg-gray-50 transition-colors"
    >
      ↓ CSV
    </a>
  )
}
