import { notFound } from 'next/navigation'
import Link from 'next/link'
import { prisma } from '@/lib/prisma'
import CenarioEditarForm from './form'

export const metadata = { title: 'Editar Cenário' }

export default async function CenarioEditarPage({
  params,
}: {
  params: Promise<{ id: string }>
}) {
  const { id } = await params
  const cenarioId = Number(id)
  if (Number.isNaN(cenarioId)) notFound()

  const [cenario, processos, atividades, transacoes] = await Promise.all([
    prisma.cenario.findUnique({
      where: { id: cenarioId },
      include: {
        processos: { select: { processoId: true } },
        atividades: { select: { atividadeId: true } },
        transacoes: { select: { transacaoId: true } },
      },
    }),
    prisma.processo.findMany({
      orderBy: [{ megaProcessoId: 'asc' }, { sequencia: 'asc' }],
      select: {
        id: true,
        descricao: true,
        sequencia: true,
        megaProcesso: { select: { descricao: true, abreviacao: true } },
      },
    }),
    prisma.atividade.findMany({
      orderBy: { id: 'asc' },
      select: {
        id: true,
        descricao: true,
        subProcesso: {
          select: {
            descricao: true,
            processo: {
              select: {
                descricao: true,
                megaProcesso: { select: { abreviacao: true } },
              },
            },
          },
        },
      },
    }),
    prisma.transacao.findMany({
      orderBy: { id: 'asc' },
      select: { id: true, descricao: true },
    }),
  ])

  if (!cenario) notFound()

  return (
    <div className="max-w-5xl">
      <div className="mb-6">
        <p className="section-tag">Editar Cenário #{cenario.id}</p>
        <h1 className="section-title">{cenario.descricao}</h1>
        <div className="gold-bar w-24 rounded-full mt-3" />
      </div>

      <CenarioEditarForm
        cenario={{
          id: cenario.id,
          descricao: cenario.descricao,
          situacao: cenario.situacao,
        }}
        processos={processos.map((p) => ({
          value: String(p.id),
          label: p.descricao,
          hint: `${p.megaProcesso.abreviacao ?? p.megaProcesso.descricao} · seq ${p.sequencia ?? '-'}`,
        }))}
        atividades={atividades.map((a) => ({
          value: String(a.id),
          label: a.descricao,
          hint: `${a.subProcesso.processo.megaProcesso.abreviacao ?? ''} › ${a.subProcesso.processo.descricao} › ${a.subProcesso.descricao}`,
        }))}
        transacoes={transacoes.map((t) => ({
          value: t.id,
          label: t.descricao || t.id,
          hint: t.id,
        }))}
        initialProcessoIds={cenario.processos.map((r) => String(r.processoId))}
        initialAtividadeIds={cenario.atividades.map((r) => String(r.atividadeId))}
        initialTransacaoIds={cenario.transacoes.map((r) => r.transacaoId)}
      />

      <div className="mt-8 flex gap-4">
        <Link
          href={`/dashboard/cenarios/${cenario.id}`}
          className="text-sm text-teal hover:text-navy font-semibold"
        >
          ← Visualizar cenário
        </Link>
        <Link href="/dashboard/cenarios" className="text-sm text-teal hover:text-navy font-semibold">
          Lista de cenários
        </Link>
      </div>
    </div>
  )
}
