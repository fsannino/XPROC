export const metadata = { title: 'Empresas' }

import { prisma } from '@/lib/prisma'
import { excluirEmpresa } from '@/actions/transacoes'
import NovaEmpresaForm from './form'
import { DeleteButton } from '@/components/ui/delete-button'

export default async function EmpresasPage() {
  const empresas = await prisma.empresaUnidade.findMany({
    orderBy: { nome: 'asc' },
    include: { _count: { select: { subProcessos: true } } },
  })

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Empresas / Unidades</h1>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <NovaEmpresaForm />

        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-100">
              <tr>
                <th className="text-left px-4 py-3 font-medium text-gray-600">Nome</th>
                <th className="text-center px-4 py-3 font-medium text-gray-600">Sub-proc.</th>
                <th className="text-right px-4 py-3 font-medium text-gray-600">Ações</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {empresas.length === 0 && (
                <tr>
                  <td colSpan={3} className="text-center py-8 text-gray-400">
                    Nenhuma empresa cadastrada.
                  </td>
                </tr>
              )}
              {empresas.map((e) => (
                <tr key={e.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-900">{e.nome}</td>
                  <td className="px-4 py-3 text-center text-gray-600">{e._count.subProcessos}</td>
                  <td className="px-4 py-3 text-right">
                    <DeleteButton action={excluirEmpresa.bind(null, e.id)} confirmText={`Excluir empresa "${e.nome}"?`} />
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
