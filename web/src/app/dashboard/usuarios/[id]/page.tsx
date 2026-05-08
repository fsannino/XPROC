import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import Link from 'next/link'
import { concederAcesso, revogarAcesso } from '@/actions/usuarios'

export async function generateMetadata({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params
  const u = await prisma.usuario.findUnique({ where: { id }, select: { nome: true } })
  return { title: u ? `${u.nome} — Collab:Flow` : 'Usuário — Collab:Flow' }
}

export default async function UsuarioAcessosPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params

  const [usuario, megaProcessos, historico] = await Promise.all([
    prisma.usuario.findUnique({
      where: { id },
      include: { acessos: { include: { megaProcesso: true } } },
    }),
    prisma.megaProcesso.findMany({ orderBy: { id: 'asc' } }),
    prisma.usuarioHistorico.findMany({
      where: { usuarioId: id },
      orderBy: { criadoEm: 'desc' },
      take: 50,
    }),
  ])

  if (!usuario) notFound()

  const acessoIds = new Set(usuario.acessos.map((a) => a.megaProcessoId))

  const opLabel: Record<string, string> = {
    LOGIN: '✓ Login',
    LOGIN_FALHA: '✕ Falha de login',
    LOGOUT: '↩ Logout',
  }

  return (
    <div className="max-w-2xl space-y-6">
      <div>
        <Link href="/dashboard/usuarios" className="text-sm text-blue-600 hover:underline">
          ← Usuários
        </Link>
        <h1 className="text-2xl font-bold text-gray-900 mt-2 mb-1">{usuario.nome}</h1>
        <p className="text-sm text-gray-500">Código: {usuario.codigo}</p>
      </div>

      {/* Controle de acesso */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 bg-gray-50">
          <h2 className="font-semibold text-gray-800">Controle de Acesso — Mega-Processos</h2>
        </div>
        <div className="divide-y divide-gray-50">
          {megaProcessos.map((mp) => {
            const temAcesso = acessoIds.has(mp.id)
            return (
              <div key={mp.id} className="flex items-center justify-between px-4 py-3">
                <div>
                  <span className="text-sm font-medium text-gray-800">{mp.descricao}</span>
                  {mp.abreviacao && (
                    <span className="ml-2 text-xs text-gray-400">{mp.abreviacao}</span>
                  )}
                </div>
                <form
                  action={
                    temAcesso
                      ? revogarAcesso.bind(null, usuario.id, mp.id)
                      : concederAcesso.bind(null, usuario.id, mp.id)
                  }
                >
                  <button
                    type="submit"
                    className={`text-sm font-medium px-3 py-1 rounded-lg transition-colors ${
                      temAcesso
                        ? 'bg-red-50 text-red-600 hover:bg-red-100'
                        : 'bg-green-50 text-green-600 hover:bg-green-100'
                    }`}
                  >
                    {temAcesso ? 'Revogar' : 'Conceder'}
                  </button>
                </form>
              </div>
            )
          })}
        </div>
      </div>

      {/* Histórico de atividades */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 bg-gray-50">
          <h2 className="font-semibold text-gray-800">Histórico de Atividades</h2>
        </div>
        {historico.length === 0 ? (
          <p className="px-4 py-6 text-sm text-gray-400 text-center">Nenhum registro.</p>
        ) : (
          <div className="divide-y divide-gray-50">
            {historico.map((h) => (
              <div key={h.id} className="flex items-center justify-between px-4 py-2.5">
                <div className="flex items-center gap-2">
                  <span className={`text-xs font-medium ${h.operacao === 'LOGIN_FALHA' ? 'text-red-600' : h.operacao === 'LOGOUT' ? 'text-gray-500' : 'text-green-600'}`}>
                    {opLabel[h.operacao] ?? h.operacao}
                  </span>
                  {h.descricao && <span className="text-xs text-gray-400">{h.descricao}</span>}
                </div>
                <span className="text-xs text-gray-400">
                  {new Date(h.criadoEm).toLocaleString('pt-BR')}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
