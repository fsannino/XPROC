import { notFound } from 'next/navigation'
import Link from 'next/link'
import { prisma } from '@/lib/prisma'

export const metadata = { title: 'Cenário' }

export default async function CenarioViewPage({
  params,
}: {
  params: Promise<{ id: string }>
}) {
  const { id } = await params
  const cenarioId = Number(id)
  if (Number.isNaN(cenarioId)) notFound()

  const cenario = await prisma.cenario.findUnique({
    where: { id: cenarioId },
    include: {
      transacoes: {
        include: { transacao: { select: { id: true, descricao: true } } },
      },
      processos: {
        include: {
          processo: {
            include: {
              megaProcesso: { select: { id: true, descricao: true, abreviacao: true } },
            },
          },
        },
      },
      atividades: {
        include: {
          atividade: {
            include: {
              subProcesso: {
                include: {
                  processo: {
                    include: {
                      megaProcesso: { select: { id: true, descricao: true, abreviacao: true } },
                    },
                  },
                },
              },
            },
          },
        },
      },
    },
  })

  if (!cenario) notFound()

  return (
    <div className="max-w-5xl">
      <div className="mb-6 flex items-start justify-between gap-4">
        <div>
          <p className="section-tag">Cenário #{cenario.id}</p>
          <h1 className="section-title">{cenario.descricao}</h1>
          {cenario.situacao && (
            <p className="text-sm text-gray-medium">{cenario.situacao}</p>
          )}
          <div className="gold-bar w-24 rounded-full mt-3" />
        </div>
        <Link
          href={`/dashboard/cenarios/${cenario.id}/editar`}
          className="shrink-0 px-4 py-2 rounded-md text-sm font-semibold bg-navy hover:bg-teal text-white transition-all hover:-translate-y-0.5 hover:shadow-md"
        >
          Editar
        </Link>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <Stat label="Processos" value={cenario.processos.length} />
        <Stat label="Atividades" value={cenario.atividades.length} />
        <Stat label="Transações" value={cenario.transacoes.length} />
      </div>

      <Section title="Processos relacionados">
        {cenario.processos.length === 0 ? (
          <Empty>Nenhum processo relacionado.</Empty>
        ) : (
          <ul className="divide-y divide-[#F5F6F8]">
            {cenario.processos.map((cp) => (
              <li key={cp.processoId} className="flex items-center gap-3 py-2.5">
                <span className="text-[10px] font-mono text-gray-medium uppercase tracking-wider w-12">
                  {cp.processo.megaProcesso.abreviacao || `MP${cp.processo.megaProcesso.id}`}
                </span>
                <span className="text-sm text-slate flex-1">{cp.processo.descricao}</span>
                <span className="text-[10px] text-gray-medium">
                  seq {cp.processo.sequencia ?? '-'}
                </span>
              </li>
            ))}
          </ul>
        )}
      </Section>

      <Section title="Atividades relacionadas">
        {cenario.atividades.length === 0 ? (
          <Empty>Nenhuma atividade relacionada.</Empty>
        ) : (
          <ul className="divide-y divide-[#F5F6F8]">
            {cenario.atividades.map((ca) => (
              <li key={ca.atividadeId} className="py-2.5">
                <p className="text-sm text-slate">{ca.atividade.descricao}</p>
                <p className="text-[10px] text-gray-medium font-mono mt-0.5">
                  {ca.atividade.subProcesso.processo.megaProcesso.abreviacao ?? ''}
                  {' › '}
                  {ca.atividade.subProcesso.processo.descricao}
                  {' › '}
                  {ca.atividade.subProcesso.descricao}
                </p>
              </li>
            ))}
          </ul>
        )}
      </Section>

      <Section title="Transações relacionadas">
        {cenario.transacoes.length === 0 ? (
          <Empty>Nenhuma transação relacionada.</Empty>
        ) : (
          <ul className="divide-y divide-[#F5F6F8]">
            {cenario.transacoes.map((ct) => (
              <li key={ct.transacaoId} className="flex items-center gap-3 py-2.5">
                <span className="text-xs font-mono text-navy bg-[#F5F6F8] rounded px-2 py-0.5">
                  {ct.transacao.id}
                </span>
                <span className="text-sm text-slate flex-1">
                  {ct.transacao.descricao || <span className="text-gray-medium italic">sem nome</span>}
                </span>
              </li>
            ))}
          </ul>
        )}
      </Section>

      <div className="mt-8">
        <Link href="/dashboard/cenarios" className="text-sm text-teal hover:text-navy font-semibold">
          ← Voltar para a lista
        </Link>
      </div>
    </div>
  )
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="bg-white rounded-lg border border-[#E2E8F0] border-t-4 border-t-teal px-4 py-3">
      <p className="text-[10px] font-bold tracking-[0.18em] uppercase text-teal">{label}</p>
      <p className="font-display text-3xl text-navy mt-1">{value}</p>
    </div>
  )
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-lg border border-[#E2E8F0] p-6 mb-4">
      <h2 className="font-display text-xl text-navy mb-3">{title}</h2>
      {children}
    </div>
  )
}

function Empty({ children }: { children: React.ReactNode }) {
  return <p className="text-sm text-gray-medium italic py-4">{children}</p>
}
