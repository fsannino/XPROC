import { prisma } from '@/lib/prisma'
import Link from 'next/link'
import { excluirTransacao } from '@/actions/transacoes'

export default async function TransacoesPage() {
  const transacoes = await prisma.transacao.findMany({
    orderBy: { id: 'asc' },
    include: { _count: { select: { megaProcessos: true } } },
    take: 100,
  })

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Transações</h1>
        <Link
          href="/dashboard/transacoes/nova"
          className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
        >
          + Nova Transação
        </Link>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-100">
            <tr>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Código</th>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Descrição</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Mega-Processos</th>
              <th className="text-right px-4 py-3 font-medium text-gray-600">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {transacoes.length === 0 && (
              <tr>
                <td colSpan={4} className="text-center py-8 text-gray-400">
                  Nenhuma transação cadastrada.
                </td>
              </tr>
            )}
            {transacoes.map((t) => (
              <tr key={t.id} className="hover:bg-gray-50">
                <td className="px-4 py-3 font-mono text-gray-700 font-medium">{t.id}</td>
                <td className="px-4 py-3 text-gray-800">{t.descricao || '—'}</td>
                <td className="px-4 py-3 text-center text-gray-600">{t._count.megaProcessos}</td>
                <td className="px-4 py-3 text-right">
                  <form action={excluirTransacao.bind(null, t.id)} className="inline">
                    <button type="submit" className="text-red-600 hover:text-red-800 font-medium text-xs">
                      Excluir
                    </button>
                  </form>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
