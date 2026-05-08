export const metadata = { title: 'Catálogo' }

import { prisma } from '@/lib/prisma'
import {
  type ProdutoTipo,
  type InsumoTipo,
  type SistemaTipo,
} from '@/lib/definitions'
import CatalogoForms from './forms'
import { ProdutoItem, InsumoItem, SistemaItem } from './items'

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
              <li key={p.id} className="px-5 py-3">
                <ProdutoItem
                  p={{
                    id: p.id,
                    codigo: p.codigo,
                    descricao: p.descricao,
                    tipo: p.tipo as ProdutoTipo,
                    processosCount: p._count.processos,
                  }}
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
              <li key={i.id} className="px-5 py-3">
                <InsumoItem
                  i={{
                    id: i.id,
                    codigo: i.codigo,
                    descricao: i.descricao,
                    tipo: i.tipo as InsumoTipo,
                    processosCount: i._count.processos,
                    atividadesCount: i._count.atividades,
                  }}
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
              <li key={s.id} className="px-5 py-3">
                <SistemaItem
                  s={{
                    id: s.id,
                    codigo: s.codigo,
                    nome: s.nome,
                    tipo: s.tipo as SistemaTipo,
                    processosCount: s._count.processos,
                  }}
                />
              </li>
            ))}
          </ul>
        </section>
      </div>
    </div>
  )
}
