export const metadata = { title: 'Mapa de Processos' }

import { prisma } from '@/lib/prisma'
import Link from 'next/link'
import { StatusBadge } from '@/components/ui/status-badge'

export default async function MapaPage() {
  const megaProcessos = await prisma.megaProcesso.findMany({
    orderBy: { id: 'asc' },
    include: {
      responsavel: { select: { nome: true } },
      processos: {
        orderBy: { sequencia: 'asc' },
        include: {
          subProcessos: { orderBy: { sequencia: 'asc' }, select: { id: true, descricao: true, sequencia: true } },
          _count: { select: { subProcessos: true } },
        },
      },
      _count: { select: { riscos: true, comentarios: true } },
    },
  })

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Mapa de Processos</h1>
          <p className="text-sm text-gray-500 mt-0.5">Visão hierárquica completa — MegaProcesso → Processo → Sub-Processo</p>
        </div>
        <Link
          href="/dashboard/processos"
          className="border border-gray-300 text-gray-700 px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-50"
        >
          Lista
        </Link>
      </div>

      {megaProcessos.length === 0 && (
        <div className="bg-white rounded-xl border border-gray-100 p-12 text-center text-gray-400">
          Nenhum mega-processo cadastrado.
        </div>
      )}

      <div className="space-y-4">
        {megaProcessos.map((mp) => (
          <details key={mp.id} className="group bg-white rounded-xl border border-gray-200 overflow-hidden">
            <summary className="flex items-center justify-between px-5 py-4 cursor-pointer hover:bg-gray-50 list-none">
              <div className="flex items-center gap-3">
                <span className="text-xs font-mono text-gray-400 w-8">{mp.id}</span>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-gray-900">{mp.descricao}</span>
                    {mp.abreviacao && (
                      <span className="text-xs bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded font-mono">{mp.abreviacao}</span>
                    )}
                    <StatusBadge status={mp.status} />
                  </div>
                  <div className="flex items-center gap-3 mt-0.5 text-xs text-gray-500">
                    {mp.responsavel && <span>Owner: {mp.responsavel.nome}</span>}
                    <span>{mp.processos.length} processos</span>
                    {mp._count.riscos > 0 && <span className="text-red-500">{mp._count.riscos} riscos</span>}
                    {mp._count.comentarios > 0 && <span>{mp._count.comentarios} comentários</span>}
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <Link
                  href={`/dashboard/processos/${mp.id}`}
                  onClick={(e) => e.stopPropagation()}
                  className="text-blue-600 hover:text-blue-800 text-sm font-medium"
                >
                  Ver detalhes →
                </Link>
                <span className="text-gray-400 group-open:rotate-180 transition-transform">▾</span>
              </div>
            </summary>

            <div className="border-t border-gray-100 divide-y divide-gray-50">
              {mp.processos.length === 0 && (
                <p className="px-5 py-3 text-sm text-gray-400">Nenhum processo.</p>
              )}
              {mp.processos.map((proc) => (
                <details key={proc.id} className="group/proc">
                  <summary className="flex items-center gap-3 px-5 py-3 bg-gray-50 cursor-pointer hover:bg-gray-100 list-none">
                    <span className="text-xs font-mono bg-white border border-gray-200 text-gray-500 px-1.5 py-0.5 rounded w-6 text-center">
                      {proc.sequencia ?? '—'}
                    </span>
                    <span className="text-sm font-medium text-gray-800 flex-1">{proc.descricao}</span>
                    {(proc.tempoMedioCiclo || proc.custoEstimado || proc.volumeMensal) && (
                      <span className="text-xs text-gray-400 hidden sm:block">
                        {proc.tempoMedioCiclo && `⏱ ${proc.tempoMedioCiclo}d`}
                        {proc.custoEstimado && ` · R$ ${proc.custoEstimado.toLocaleString('pt-BR')}`}
                        {proc.volumeMensal && ` · ${proc.volumeMensal}/mês`}
                      </span>
                    )}
                    <span className="text-xs text-gray-400">{proc._count.subProcessos} etapas</span>
                    <span className="text-gray-400 group-open/proc:rotate-180 transition-transform">▾</span>
                  </summary>

                  {proc.subProcessos.length > 0 && (
                    <div className="px-8 py-2 space-y-1">
                      {proc.subProcessos.map((sub) => (
                        <div key={sub.id} className="flex items-center gap-2 py-1.5 text-sm text-gray-700">
                          <span className="text-xs font-mono text-gray-400 w-5 text-right">{sub.sequencia ?? '—'}</span>
                          <span className="w-1.5 h-1.5 rounded-full bg-gray-300 shrink-0"></span>
                          {sub.descricao}
                        </div>
                      ))}
                    </div>
                  )}
                </details>
              ))}
            </div>
          </details>
        ))}
      </div>
    </div>
  )
}
