import { prisma } from '@/lib/prisma'
import { excluirEmpresa, adicionarEmpresa } from '@/actions/transacoes'

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
        {/* Formulário de criação */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
          <h2 className="text-lg font-semibold text-gray-800 mb-4">Nova Empresa</h2>
          <form action={adicionarEmpresa} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Nome <span className="text-red-500">*</span>
              </label>
              <input
                name="nome"
                required
                maxLength={150}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>
            <button
              type="submit"
              className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
            >
              Criar Empresa
            </button>
          </form>
        </div>

        {/* Lista */}
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
                    <form action={excluirEmpresa.bind(null, e.id)} className="inline">
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
    </div>
  )
}
