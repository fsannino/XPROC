import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import Link from 'next/link'
import { excluirProcesso, excluirSubProcesso } from '@/actions/processos'
import { excluirComentario, adicionarComentarioForm } from '@/actions/lifecycle'
import { proximosStatus } from '@/lib/lifecycle-utils'
import { excluirRisco, adicionarRiscoForm } from '@/actions/riscos'
import { DeleteButton } from '@/components/ui/delete-button'
import { StatusBadge } from '@/components/ui/status-badge'
import { StatusSelector } from '@/components/ui/status-selector'
import { getSession } from '@/lib/session'

export default async function MegaProcessoDetalhe({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params
  const session = await getSession()

  const megaProcesso = await prisma.megaProcesso.findUnique({
    where: { id: Number(id) },
    include: {
      responsavel: { select: { codigo: true, nome: true } },
      processos: {
        orderBy: { sequencia: 'asc' },
        include: {
          subProcessos: {
            orderBy: { sequencia: 'asc' },
            include: {
              atividades: {
                orderBy: { sequencia: 'asc' },
                include: { transacao: { select: { id: true, descricao: true } } },
              },
              empresas: { include: { empresa: { select: { nome: true } } } },
            },
          },
        },
      },
      comentarios: {
        where: { parentId: null },
        orderBy: { criadoEm: 'desc' },
        take: 20,
        include: {
          usuario: { select: { codigo: true, nome: true } },
          respostas: {
            orderBy: { criadoEm: 'asc' },
            include: { usuario: { select: { codigo: true, nome: true } } },
          },
        },
      },
      riscos: { orderBy: { criadoEm: 'desc' } },
      versoes: {
        orderBy: { versao: 'desc' },
        take: 10,
        include: { criadoPor: { select: { codigo: true, nome: true } } },
      },
    },
  })

  if (!megaProcesso) notFound()

  const proximos = proximosStatus(megaProcesso.status)

  const labelMap: Record<string, string> = {
    A: 'Alta', M: 'Média', B: 'Baixa',
  }
  const riskColor: Record<string, string> = {
    A: 'text-red-600', M: 'text-yellow-600', B: 'text-green-600',
  }
  return (
    <div>
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <div>
          <Link href="/dashboard/processos" className="text-sm text-blue-600 hover:underline">
            ← Mega-Processos
          </Link>
          <h1 className="text-2xl font-bold text-gray-900 mt-1">{megaProcesso.descricao}</h1>
          <div className="flex items-center gap-3 mt-2 flex-wrap">
            {megaProcesso.abreviacao && (
              <span className="text-sm text-gray-500">Abrev: {megaProcesso.abreviacao}</span>
            )}
            {megaProcesso.bloqueado ? (
              <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700">Bloqueado</span>
            ) : (
              <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700">Ativo</span>
            )}
            <StatusSelector
              megaProcessoId={megaProcesso.id}
              statusAtual={megaProcesso.status}
              proximosStatus={proximos}
            />
            {megaProcesso.responsavel && (
              <span className="text-xs text-gray-500">
                Owner: <span className="font-medium text-gray-700">{megaProcesso.responsavel.nome}</span>
              </span>
            )}
          </div>
        </div>
        <div className="flex gap-2">
          <Link
            href={`/api/bpmn/${megaProcesso.id}`}
            target="_blank"
            className="border border-gray-300 text-gray-700 px-3 py-2 rounded-lg text-sm font-medium hover:bg-gray-50"
          >
            BPMN XML
          </Link>
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

      {/* Process hierarchy */}
      <div className="space-y-4 mb-8">
        {megaProcesso.processos.length === 0 && (
          <div className="bg-white rounded-xl border border-gray-100 p-8 text-center text-gray-400">
            Nenhum processo cadastrado para este mega-processo.
          </div>
        )}

        {megaProcesso.processos.map((proc) => (
          <div key={proc.id} className="bg-white rounded-xl border border-gray-100 overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 bg-gray-50 border-b border-gray-100">
              <div className="flex items-center gap-2">
                <span className="text-xs font-mono bg-gray-200 text-gray-600 px-1.5 py-0.5 rounded">{proc.sequencia ?? '—'}</span>
                <span className="font-semibold text-gray-800">{proc.descricao}</span>
                {(proc.tempoMedioCiclo || proc.custoEstimado || proc.volumeMensal) && (
                  <span className="text-xs text-gray-400 ml-2">
                    {proc.tempoMedioCiclo && `⏱ ${proc.tempoMedioCiclo}d`}
                    {proc.custoEstimado && ` · R$ ${proc.custoEstimado.toLocaleString('pt-BR')}`}
                    {proc.volumeMensal && ` · ${proc.volumeMensal}/mês`}
                  </span>
                )}
              </div>
              <DeleteButton action={excluirProcesso.bind(null, proc.id)} confirmText={`Excluir processo "${proc.descricao}"?`} />
            </div>

            <div className="divide-y divide-gray-50">
              {proc.subProcessos.length === 0 && (
                <p className="px-4 py-3 text-sm text-gray-400">Sem sub-processos.</p>
              )}

              {proc.subProcessos.map((sub) => (
                <div key={sub.id}>
                  <div className="flex items-center justify-between px-6 py-2.5 bg-white">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-mono text-gray-400">{sub.sequencia ?? '—'}</span>
                      <span className="text-sm font-medium text-gray-700">{sub.descricao}</span>
                      {sub.empresas.length > 0 && (
                        <span className="text-xs text-gray-400">
                          ({sub.empresas.map((se) => se.empresa.nome).join(', ')})
                        </span>
                      )}
                    </div>
                    <DeleteButton action={excluirSubProcesso.bind(null, sub.id)} confirmText={`Excluir sub-processo "${sub.descricao}"?`} />
                  </div>

                  {sub.atividades.length > 0 && (
                    <div className="ml-10 mb-2 space-y-1">
                      {sub.atividades.map((at) => (
                        <div key={at.id} className="flex items-center gap-2 px-3 py-1 bg-blue-50 rounded text-xs text-gray-600">
                          <span className="font-mono text-gray-400">{at.sequencia ?? '—'}</span>
                          <span>{at.descricao || '—'}</span>
                          {at.transacao && (
                            <span className="ml-auto font-mono text-blue-600">{at.transacao.id}</span>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      {/* Risks */}
      <div className="mb-8">
        <h2 className="text-lg font-semibold text-gray-800 mb-3">Riscos</h2>
        <div className="bg-white rounded-xl border border-gray-100 overflow-hidden mb-3">
          {megaProcesso.riscos.length === 0 ? (
            <p className="p-4 text-sm text-gray-400">Nenhum risco cadastrado.</p>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Descrição</th>
                  <th className="text-center px-3 py-2 font-medium text-gray-600">Prob.</th>
                  <th className="text-center px-3 py-2 font-medium text-gray-600">Impacto</th>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Controle</th>
                  <th className="w-10"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {megaProcesso.riscos.map((r) => (
                  <tr key={r.id}>
                    <td className="px-4 py-2 text-gray-800">{r.descricao}</td>
                    <td className={`px-3 py-2 text-center font-semibold ${riskColor[r.probabilidade]}`}>{labelMap[r.probabilidade]}</td>
                    <td className={`px-3 py-2 text-center font-semibold ${riskColor[r.impacto]}`}>{labelMap[r.impacto]}</td>
                    <td className="px-4 py-2 text-gray-500 text-xs">{r.controle || '—'}</td>
                    <td className="px-2 py-2">
                      <DeleteButton action={excluirRisco.bind(null, r.id)} confirmText="Excluir risco?" />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
        <form action={adicionarRiscoForm} className="bg-white rounded-xl border border-gray-100 p-4 space-y-3">
          <input type="hidden" name="megaProcessoId" value={megaProcesso.id} />
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div className="md:col-span-3">
              <input
                name="descricao"
                placeholder="Descrição do risco *"
                required
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>
            <select name="probabilidade" className="rounded-lg border border-gray-300 px-3 py-2 text-sm">
              <option value="B">Probabilidade: Baixa</option>
              <option value="M" selected>Probabilidade: Média</option>
              <option value="A">Probabilidade: Alta</option>
            </select>
            <select name="impacto" className="rounded-lg border border-gray-300 px-3 py-2 text-sm">
              <option value="B">Impacto: Baixo</option>
              <option value="M" selected>Impacto: Médio</option>
              <option value="A">Impacto: Alto</option>
            </select>
            <input
              name="controle"
              placeholder="Controle mitigador (opcional)"
              className="rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <button type="submit" className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800">
            + Adicionar Risco
          </button>
        </form>
      </div>

      {/* Comments */}
      <div className="mb-8">
        <h2 className="text-lg font-semibold text-gray-800 mb-3">Comentários</h2>
        <div className="space-y-3 mb-3">
          {megaProcesso.comentarios.length === 0 && (
            <p className="text-sm text-gray-400">Nenhum comentário ainda.</p>
          )}
          {megaProcesso.comentarios.map((c) => (
            <div key={c.id} className="bg-white rounded-xl border border-gray-100 p-4">
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-semibold text-gray-700">{c.usuario.nome}</span>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-gray-400">{new Date(c.criadoEm).toLocaleDateString('pt-BR')}</span>
                  {(session?.userId === c.usuarioId || session?.categoria === 'A') && (
                    <DeleteButton action={excluirComentario.bind(null, c.id)} confirmText="Excluir comentário?" />
                  )}
                </div>
              </div>
              <p className="text-sm text-gray-800">{c.texto}</p>
              {c.respostas.map((r) => (
                <div key={r.id} className="mt-2 ml-4 pl-3 border-l-2 border-gray-200">
                  <span className="text-xs font-semibold text-gray-600">{r.usuario.nome}</span>
                  <p className="text-xs text-gray-700 mt-0.5">{r.texto}</p>
                </div>
              ))}
            </div>
          ))}
        </div>
        <form action={adicionarComentarioForm} className="bg-white rounded-xl border border-gray-100 p-4 flex gap-2">
          <input type="hidden" name="megaProcessoId" value={megaProcesso.id} />
          <input
            name="texto"
            placeholder="Adicionar comentário..."
            required
            className="flex-1 rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <button type="submit" className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 shrink-0">
            Comentar
          </button>
        </form>
      </div>

      {/* Version history */}
      {megaProcesso.versoes.length > 0 && (
        <div>
          <h2 className="text-lg font-semibold text-gray-800 mb-3">Histórico de Status</h2>
          <div className="bg-white rounded-xl border border-gray-100 overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Versão</th>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">De</th>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Para</th>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Por</th>
                  <th className="text-left px-4 py-2 font-medium text-gray-600">Data</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {megaProcesso.versoes.map((v) => (
                  <tr key={v.id}>
                    <td className="px-4 py-2 text-gray-500 font-mono">v{v.versao}</td>
                    <td className="px-4 py-2"><StatusBadge status={v.statusAnterior} /></td>
                    <td className="px-4 py-2"><StatusBadge status={v.statusNovo} /></td>
                    <td className="px-4 py-2 text-gray-700">{v.criadoPor.nome}</td>
                    <td className="px-4 py-2 text-gray-500">{new Date(v.criadoEm).toLocaleDateString('pt-BR')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
