export const metadata = { title: 'Catálogo' }

import { prisma } from '@/lib/prisma'
import { excluirProduto } from '@/actions/produtos'
import { excluirInsumo } from '@/actions/insumos'
import { excluirSistema } from '@/actions/sistemas'
import { DeleteButton } from '@/components/ui/delete-button'
import {
  PRODUTO_TIPO_LABEL,
  INSUMO_TIPO_LABEL,
  type ProdutoTipo,
  type InsumoTipo,
} from '@/lib/definitions'
import CatalogoForms from './forms'

export default async function CatalogoPage() {
  const [produtos, insumos, sistemas] = await Promise.all([
    prisma.produto.findMany({
      orderBy: { descricao: 'asc' },
      include: { _count: { select: { processos: true } } },
    }),
    prisma.insumo.findMany({
      orderBy: { descricao: 'asc' },
      include: { _count: { select: { processos: true, atividades: true } } },
    }),
    prisma.sistema.findMany({
      orderBy: { nome: 'asc' },
      include: { _count: { select: { processos: true } } },
    }),
  ])

  return (
    <div>
      <div className="mb-8">
        <p className="section-tag">Mapa</p>
        <h1 className="section-title">Catálogo</h1>
        <p className="text-sm text-gray-medium">
          Produtos, insumos (entradas/saídas) e sistemas externos disponíveis para vincular aos processos do mapa.
        </p>
        <div className="gold-bar w-24 rounded-full mt-3" />
      </div>

      <CatalogoForms />

      <div className="grid grid-cols-1 xl:grid-cols-3 gap-6 mt-8">
        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Produtos ({produtos.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {produtos.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhum produto cadastrado.</li>
            )}
            {produtos.map((p) => (
              <li key={p.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{p.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{p.descricao}</p>
                  <p className="text-xs text-gray-medium">
                    {PRODUTO_TIPO_LABEL[p.tipo as ProdutoTipo]} · {p._count.processos} proc
                  </p>
                </div>
                <DeleteButton
                  action={excluirProduto.bind(null, p.id)}
                  confirmText={`Excluir produto "${p.descricao}"?`}
                />
              </li>
            ))}
          </ul>
        </section>

        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Insumos ({insumos.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {insumos.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhum insumo cadastrado.</li>
            )}
            {insumos.map((i) => (
              <li key={i.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{i.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{i.descricao}</p>
                  <p className="text-xs text-gray-medium">
                    {INSUMO_TIPO_LABEL[i.tipo as InsumoTipo]} · {i._count.processos} proc · {i._count.atividades} ativ
                  </p>
                </div>
                <DeleteButton
                  action={excluirInsumo.bind(null, i.id)}
                  confirmText={`Excluir insumo "${i.descricao}"?`}
                />
              </li>
            ))}
          </ul>
        </section>

        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Sistemas ({sistemas.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {sistemas.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhum sistema cadastrado.</li>
            )}
            {sistemas.map((s) => (
              <li key={s.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{s.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{s.nome}</p>
                  <p className="text-xs text-gray-medium">
                    {s.tipo} · {s._count.processos} proc
                  </p>
                </div>
                <DeleteButton
                  action={excluirSistema.bind(null, s.id)}
                  confirmText={`Excluir sistema "${s.nome}"?`}
                />
              </li>
            ))}
          </ul>
        </section>
      </div>
    </div>
  )
}
