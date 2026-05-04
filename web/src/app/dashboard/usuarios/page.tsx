export const metadata = { title: 'Usuários' }

import { prisma } from '@/lib/prisma'
import Link from 'next/link'
import { alternarStatusUsuario } from '@/actions/usuarios'
import { SearchInput } from '@/components/ui/search'
import { Pagination } from '@/components/ui/pagination'
import { ExportButton } from '@/components/ui/export-button'

const PER_PAGE = 25

export default async function UsuariosPage({ searchParams }: { searchParams: Promise<{ busca?: string; pagina?: string }> }) {
  const { busca, pagina } = await searchParams
  const page = Math.max(1, Number(pagina) || 1)
  const where = busca
    ? { OR: [{ nome: { contains: busca, mode: 'insensitive' as const } }, { codigo: { contains: busca, mode: 'insensitive' as const } }] }
    : undefined

  const [total, usuarios] = await Promise.all([
    prisma.usuario.count({ where }),
    prisma.usuario.findMany({
      where,
      orderBy: { nome: 'asc' },
      include: { _count: { select: { acessos: true } } },
      skip: (page - 1) * PER_PAGE,
      take: PER_PAGE,
    }),
  ])

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Usuários</h1>
        <div className="flex items-center gap-3">
          <SearchInput placeholder="Buscar usuário..." />
          <ExportButton tipo="usuarios" />
          <Link
            href="/dashboard/usuarios/novo"
            className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
          >
            + Novo Usuário
          </Link>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-100">
            <tr>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Código</th>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Nome</th>
              <th className="text-left px-4 py-3 font-medium text-gray-600">Email</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Cat.</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Acessos</th>
              <th className="text-center px-4 py-3 font-medium text-gray-600">Status</th>
              <th className="text-right px-4 py-3 font-medium text-gray-600">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {usuarios.length === 0 && (
              <tr>
                <td colSpan={7} className="text-center py-8 text-gray-400">
                  Nenhum usuário cadastrado.
                </td>
              </tr>
            )}
            {usuarios.map((u) => (
              <tr key={u.id} className="hover:bg-gray-50">
                <td className="px-4 py-3 font-mono text-gray-700">{u.codigo}</td>
                <td className="px-4 py-3 font-medium text-gray-900">{u.nome}</td>
                <td className="px-4 py-3 text-gray-500">{u.email || '—'}</td>
                <td className="px-4 py-3 text-center text-gray-500">{u.categoria || '—'}</td>
                <td className="px-4 py-3 text-center text-gray-700">{u._count.acessos}</td>
                <td className="px-4 py-3 text-center">
                  {u.ativo ? (
                    <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700">Ativo</span>
                  ) : (
                    <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium bg-gray-100 text-gray-500">Inativo</span>
                  )}
                </td>
                <td className="px-4 py-3 text-right space-x-2">
                  <Link href={`/dashboard/usuarios/${u.id}`} className="text-blue-600 hover:text-blue-800 font-medium text-sm">
                    Acessos
                  </Link>
                  <form action={alternarStatusUsuario.bind(null, u.id, !u.ativo)} className="inline">
                    <button type="submit" className={`font-medium text-sm ${u.ativo ? 'text-amber-600 hover:text-amber-800' : 'text-green-600 hover:text-green-800'}`}>
                      {u.ativo ? 'Desativar' : 'Ativar'}
                    </button>
                  </form>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <Pagination page={page} total={total} perPage={PER_PAGE} basePath="/dashboard/usuarios" busca={busca} />
      </div>
    </div>
  )
}
