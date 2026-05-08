import { prisma } from '@/lib/prisma'
import Link from 'next/link'

export const metadata = { title: 'Painel — XPROC' }

async function getStats() {
  const [megaProcessos, processos, subProcessos, transacoes, usuarios, empresas] =
    await Promise.all([
      prisma.megaProcesso.count(),
      prisma.processo.count(),
      prisma.subProcesso.count(),
      prisma.transacao.count(),
      prisma.usuario.count({ where: { ativo: true } }),
      prisma.empresaUnidade.count(),
    ])
  return { megaProcessos, processos, subProcessos, transacoes, usuarios, empresas }
}

const cards = [
  {
    label: 'Mega-Processos',
    key: 'megaProcessos',
    href: '/dashboard/processos',
    accent: 'navy',
    description: 'Visão estratégica dos processos.',
  },
  {
    label: 'Processos',
    key: 'processos',
    href: '/dashboard/processos',
    accent: 'teal',
    description: 'Mapa operacional ponta-a-ponta.',
  },
  {
    label: 'Sub-Processos',
    key: 'subProcessos',
    href: '/dashboard/processos',
    accent: 'teal',
    description: 'Atividades detalhadas em cada fluxo.',
  },
  {
    label: 'Transações',
    key: 'transacoes',
    href: '/dashboard/transacoes',
    accent: 'gold',
    description: 'Movimentações registradas no sistema.',
  },
  {
    label: 'Usuários Ativos',
    key: 'usuarios',
    href: '/dashboard/usuarios',
    accent: 'navy',
    description: 'Pessoas com acesso vigente.',
  },
  {
    label: 'Empresas',
    key: 'empresas',
    href: '/dashboard/empresas',
    accent: 'gold',
    description: 'Unidades organizacionais cadastradas.',
  },
] as const

const accentMap = {
  navy: { tile: 'bg-navy/8 text-navy', border: 'border-t-navy', hover: 'group-hover:text-teal' },
  teal: { tile: 'bg-teal/8 text-teal', border: 'border-t-teal', hover: 'group-hover:text-navy' },
  gold: { tile: 'bg-gold/10 text-gold', border: 'border-t-gold', hover: 'group-hover:text-teal' },
} as const

export default async function DashboardPage({
  searchParams,
}: {
  searchParams: Promise<{ acesso?: string }>
}) {
  const { acesso } = await searchParams
  const stats = await getStats()

  return (
    <div>
      <div className="mb-8">
        <p className="section-tag">Visão Geral</p>
        <h1 className="section-title">Painel de Processos</h1>
        <p className="section-subtitle">
          Acompanhe os indicadores-chave da operação. Cada bloco abre o
          módulo correspondente para gerenciar dados, fluxos e acessos.
        </p>
        <div className="gold-bar w-24 rounded-full" />
      </div>

      {acesso === 'negado' && (
        <div className="mb-6 bg-[rgba(224,80,64,0.06)] border-l-4 border-[#E05040] text-[#9A2E1F] text-sm px-4 py-3 rounded-md">
          <strong className="font-semibold">Acesso negado.</strong> Esta seção
          requer permissão de administrador.
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
        {cards.map((card) => {
          const accent = accentMap[card.accent]
          const value = stats[card.key]
          return (
            <Link
              key={card.key}
              href={card.href}
              className={`group block bg-white rounded-lg border border-[#E2E8F0] ${accent.border} border-t-4 p-6 hover:shadow-lg hover:-translate-y-1 transition-all`}
            >
              <div
                className={`w-11 h-11 rounded-lg ${accent.tile} flex items-center justify-center mb-5 font-mono font-bold text-sm`}
              >
                {String(value).padStart(2, '0').slice(0, 3)}
              </div>
              <p className="font-display text-3xl text-navy mb-1">
                {value.toLocaleString('pt-BR')}
              </p>
              <p className={`text-sm font-semibold text-navy ${accent.hover} transition-colors`}>
                {card.label}
              </p>
              <p className="text-xs text-gray-medium mt-1">{card.description}</p>
            </Link>
          )
        })}
      </div>
    </div>
  )
}
