import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import Link from 'next/link'
import { excluirProcesso, excluirSubProcesso } from '@/actions/processos'

export default async function MegaProcessoDetalhe({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params
  const megaProcesso = await prisma.megaProcesso.findUnique({
    where: { id: Number(id) },
    include: {
      processos: {
        orderBy: { sequencia: 'asc' },
        include: {
          subProcessos: { orderBy: { sequencia: 'asc' } },
        },
      },
    },
  })

  if (!megaProcesso) notFound()

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <Link href="/dashboard/processos" className="text-sm text-blue-600 hover:underline">
            ← Mega-Processos
          </Link>
          <h1 className="text-2xl font-bold text-gray-900 mt-1">{megaProcesso.descricao}</h1>
          {megaProcesso.abreviacao && (
            <span className="text-sm text-gray-500">Abrev: {megaProcesso.abreviacao}</span>
          )}
        </div>
        <div className="flex gap-2">
          <Link
            href={`/dashboard/processos/${megaProcesso.id}/editar`}
            className="border border-gray-300 text-gray-700 px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-50"
          >
            Editar
          </Link>
          <Link
            href={`/dashboard/processos/${megaProcesso.id}/sub-processos`}
            className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
          >
            + Processo
          </Link>
        </div>
      </div>

      {megaProcesso.descricaoLonga && (
        <p className="text-gray-600 mb-6 bg-white rounded-xl border border-gray-100 p-4 text-sm">
          {megaProcesso.descricaoLonga}
        </p>
      )}

      <div className="space-y-4">
        {megaProcesso.processos.length === 0 && (
          <div className="bg-white rounded-xl border border-gray-100 p-8 text-center text-gray-400">
            Nenhum processo cadastrado para este mega-processo.
          </div>
        )}
        {megaProcesso.processos.map((proc) => (
          <div key={proc.id} className="bg-white rounded-xl border border-gray-100 overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-gray-50 bg-gray-50">
              <div className="flex items-center gap-2">
                <span className="text-xs font-mono text-gray-400">{proc.sequencia ?? '—'}</span>
                <span className="font-semibold text-gray-800">{proc.descricao}</span>
              </div>
              <form action={excluirProcesso.bind(null, proc.id)} className="inline">
                <button type="submit" className="text-xs text-red-500 hover:text-red-700">
                  Excluir
                </button>
              </form>
            </div>

            <div className="divide-y divide-gray-50">
              {proc.subProcessos.length === 0 && (
                <p className="px-4 py-3 text-sm text-gray-400">Sem sub-processos.</p>
              )}
              {proc.subProcessos.map((sub) => (
                <div key={sub.id} className="flex items-center justify-between px-6 py-2.5">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-mono text-gray-400">{sub.sequencia ?? '—'}</span>
                    <span className="text-sm text-gray-700">{sub.descricao}</span>
                  </div>
                  <form action={excluirSubProcesso.bind(null, sub.id)} className="inline">
                    <button type="submit" className="text-xs text-red-500 hover:text-red-700">
                      Excluir
                    </button>
                  </form>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}
