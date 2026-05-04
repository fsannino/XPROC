import { prisma } from '@/lib/prisma'
import Link from 'next/link'

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
  { label: 'Mega-Processos', key: 'megaProcessos', href: '/dashboard/processos', color: 'bg-blue-500' },
  { label: 'Processos', key: 'processos', href: '/dashboard/processos', color: 'bg-indigo-500' },
  { label: 'Sub-Processos', key: 'subProcessos', href: '/dashboard/processos', color: 'bg-violet-500' },
  { label: 'Transações', key: 'transacoes', href: '/dashboard/transacoes', color: 'bg-emerald-500' },
  { label: 'Usuários Ativos', key: 'usuarios', href: '/dashboard/usuarios', color: 'bg-amber-500' },
  { label: 'Empresas', key: 'empresas', href: '/dashboard/empresas', color: 'bg-rose-500' },
] as const

export default async function DashboardPage() {
  const stats = await getStats()

  return (
    <div>
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Painel Geral</h1>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        {cards.map((card) => (
          <Link
            key={card.key}
            href={card.href}
            className="block bg-white rounded-xl shadow-sm border border-gray-100 p-6 hover:shadow-md transition-shadow"
          >
            <div className={`inline-flex items-center justify-center w-10 h-10 rounded-lg ${card.color} text-white text-lg mb-4`}>
              {stats[card.key]}
            </div>
            <p className="text-3xl font-bold text-gray-900">{stats[card.key].toLocaleString('pt-BR')}</p>
            <p className="text-sm text-gray-500 mt-1">{card.label}</p>
          </Link>
        ))}
      </div>
    </div>
  )
}
