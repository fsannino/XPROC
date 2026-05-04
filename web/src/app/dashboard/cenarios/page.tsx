export const metadata = { title: 'Cenários' }

import { prisma } from '@/lib/prisma'
import { excluirCenario } from '@/actions/cenarios'
import NovoCenarioForm from './form'
import { DeleteButton } from '@/components/ui/delete-button'

export default async function CenariosPage() {
  const cenarios = await prisma.cenario.findMany({
    orderBy: { id: 'asc' },
    include: { _count: { select: { transacoes: true } } },
  })

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Cenários</h1>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <NovoCenarioForm />

        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-100">
              <tr>
                <th className="text-left px-4 py-3 font-medium text-gray-600">Descrição</th>
                <th className="text-left px-4 py-3 font-medium text-gray-600">Situação</th>
                <th className="text-center px-4 py-3 font-medium text-gray-600">Transações</th>
                <th className="text-right px-4 py-3 font-medium text-gray-600">Ações</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {cenarios.length === 0 && (
                <tr>
                  <td colSpan={4} className="text-center py-8 text-gray-400">
                    Nenhum cenário cadastrado.
                  </td>
                </tr>
              )}
              {cenarios.map((c) => (
                <tr key={c.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">{c.descricao}</td>
                  <td className="px-4 py-3 text-gray-500">{c.situacao || '—'}</td>
                  <td className="px-4 py-3 text-center text-gray-600">{c._count.transacoes}</td>
                  <td className="px-4 py-3 text-right">
                    <DeleteButton action={excluirCenario.bind(null, c.id)} confirmText={`Excluir cenário "${c.descricao}"?`} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
