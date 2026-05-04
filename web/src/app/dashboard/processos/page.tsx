import { prisma } from '@/lib/prisma'
import Link from 'next/link'
import { excluirMegaProcesso } from '@/actions/processos'
import { DeleteButton } from '@/components/ui/delete-button'
import { SearchInput } from '@/components/ui/search'

export default async function ProcessosPage({ searchParams }: { searchParams: Promise<{ busca?: string }> }) {
  const { busca } = await searchParams
  const megaProcessos = await prisma.megaProcesso.findMany({
    where: busca ? { descricao: { contains: busca, mode: 'insensitive' } } : undefined,
    orderBy: { id: 'asc' },
    include: {
      _count: { select: { processos: true, acessos: true } },
    },
  })

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Processos</h1>
        <div className="flex items-center gap-3">
          <SearchInput placeholder="Buscar mega-processo..." />
          <Link
            href="/dashboard/processos/novo"
            className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 transition-colors"
          >
            + Novo Mega-Processo
          </Link>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-100">
            <tr>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Cód.</th>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Mega-Processo</th>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Abrev.</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Processos</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Acessos</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Status</th>
              <th className="text-right px-4 py-3 font-medium text-gray-600">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {megaProcessos.length === 0 && (
              <tr>
                <td colSpan={7} className="text-center py-8 text-gray-400">
                  Nenhum mega-processo cadastrado.
                </td>
              </tr>
            )}
            {megaProcessos.map((mp) => (
              <tr key={mp.id} className="hover:bg-gray-50 transition-colors">
                <td className="px-4 py-3 text-gray-500">{mp.id}</td>
                <td className="px-4 py-3 font-medium text-gray-900">{mp.descricao}</td>
                <td className="px-4 py-3 text-gray-500">{mp.abreviacao || '—'}</td>
                <td className="px-4 py-3 text-center text-gray-700">{mp._count.processos}</td>
                <td className="px-4 py-3 text-center text-gray-700">{mp._count.acessos}</td>
                <td className="px-4 py-3 text-center">
                  {mp.bloqueado ? (
                    <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700">Bloqueado</span>
                  ) : (
                    <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700">Ativo</span>
                  )}
                </td>
                <td className="px-4 py-3 text-right space-x-2">
                  <Link
                    href={`/dashboard/processos/${mp.id}`}
                    className="text-blue-600 hover:text-blue-800 font-medium text-sm"
                  >
                    Ver
                  </Link>
                  <Link
                    href={`/dashboard/processos/${mp.id}/editar`}
                    className="text-amber-600 hover:text-amber-800 font-medium text-sm"
                  >
                    Editar
                  </Link>
                  <DeleteButton action={excluirMegaProcesso.bind(null, mp.id)} confirmText={`Excluir "${mp.descricao}"?`} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
