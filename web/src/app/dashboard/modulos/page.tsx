import { prisma } from '@/lib/prisma'
import { excluirModulo, excluirSubModulo, criarModulo, criarSubModulo } from '@/actions/modulos'
import { DeleteButton } from '@/components/ui/delete-button'
import ModulosForm from './form'

export const metadata = { title: 'Módulos — Collab:Flow' }

export default async function ModulosPage() {
  const [modulos, subModulosSemModulo] = await Promise.all([
    prisma.modulo.findMany({
      orderBy: { id: 'asc' },
      include: { subModulos: { orderBy: { id: 'asc' } } },
    }),
    prisma.subModulo.findMany({ where: { moduloId: null }, orderBy: { id: 'asc' } }),
  ])

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Módulos</h1>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <ModulosForm modulos={modulos} criarModulo={criarModulo} criarSubModulo={criarSubModulo} />

        <div className="lg:col-span-2 space-y-4">
          {modulos.length === 0 && (
            <div className="bg-white rounded-xl border border-gray-100 p-8 text-center text-gray-400">
              Nenhum módulo cadastrado.
            </div>
          )}

          {modulos.map((m) => (
            <div key={m.id} className="bg-white rounded-xl border border-gray-100 overflow-hidden">
              <div className="flex items-center justify-between px-4 py-3 bg-gray-50 border-b border-gray-100">
                <span className="font-semibold text-gray-800">{m.descricao}</span>
                <DeleteButton action={excluirModulo.bind(null, m.id)} confirmText={`Excluir módulo "${m.descricao}"?`} />
              </div>
              <div className="divide-y divide-gray-50">
                {m.subModulos.length === 0 ? (
                  <p className="px-4 py-3 text-sm text-gray-400">Sem sub-módulos.</p>
                ) : (
                  m.subModulos.map((sm) => (
                    <div key={sm.id} className="flex items-center justify-between px-6 py-2.5">
                      <div className="flex items-center gap-2">
                        <span className="text-sm text-gray-700">{sm.descricao}</span>
                        {sm.abreviacao && <span className="text-xs text-gray-400 font-mono">{sm.abreviacao}</span>}
                      </div>
                      <DeleteButton action={excluirSubModulo.bind(null, sm.id)} confirmText={`Excluir sub-módulo "${sm.descricao}"?`} />
                    </div>
                  ))
                )}
              </div>
            </div>
          ))}

          {subModulosSemModulo.length > 0 && (
            <div className="bg-white rounded-xl border border-gray-100 overflow-hidden">
              <div className="px-4 py-3 bg-gray-50 border-b border-gray-100">
                <span className="font-semibold text-gray-600 text-sm">Sem módulo vinculado</span>
              </div>
              <div className="divide-y divide-gray-50">
                {subModulosSemModulo.map((sm) => (
                  <div key={sm.id} className="flex items-center justify-between px-4 py-2.5">
                    <span className="text-sm text-gray-700">{sm.descricao}</span>
                    <DeleteButton action={excluirSubModulo.bind(null, sm.id)} confirmText={`Excluir "${sm.descricao}"?`} />
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
