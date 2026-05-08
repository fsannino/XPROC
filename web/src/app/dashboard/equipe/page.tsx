export const metadata = { title: 'Equipe' }

import { prisma } from '@/lib/prisma'
import { excluirArea, excluirFuncao, excluirPessoa } from '@/actions/equipe'
import { DeleteButton } from '@/components/ui/delete-button'
import EquipeForms from './forms'

export default async function EquipePage() {
  const [areas, funcoes, pessoas] = await Promise.all([
    prisma.area.findMany({
      orderBy: { codigo: 'asc' },
      include: {
        _count: { select: { funcoes: true, pessoas: true } },
        parent: { select: { codigo: true, descricao: true } },
      },
    }),
    prisma.funcao.findMany({
      orderBy: { codigo: 'asc' },
      include: {
        _count: { select: { pessoas: true } },
        area: { select: { codigo: true, descricao: true } },
      },
    }),
    prisma.pessoa.findMany({
      orderBy: { nome: 'asc' },
      include: {
        area: { select: { codigo: true, descricao: true } },
        funcao: { select: { codigo: true, descricao: true } },
        _count: { select: { raciAtribuicoes: true } },
      },
    }),
  ])

  return (
    <div>
      <div className="mb-8">
        <p className="section-tag">RACI</p>
        <h1 className="section-title">Equipe</h1>
        <p className="text-sm text-gray-medium">
          Áreas, funções e pessoas usadas na matriz RACI dos processos.
        </p>
        <div className="gold-bar w-24 rounded-full mt-3" />
      </div>

      <EquipeForms
        areas={areas.map((a) => ({ id: a.id, codigo: a.codigo, descricao: a.descricao }))}
        funcoes={funcoes.map((f) => ({ id: f.id, codigo: f.codigo, descricao: f.descricao }))}
      />

      <div className="grid grid-cols-1 xl:grid-cols-3 gap-6 mt-8">
        {/* Áreas */}
        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Áreas ({areas.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {areas.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhuma área cadastrada.</li>
            )}
            {areas.map((a) => (
              <li key={a.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{a.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{a.descricao}</p>
                  <p className="text-xs text-gray-medium">
                    {a.parent ? `↳ ${a.parent.codigo} ${a.parent.descricao}` : '— raiz —'} ·{' '}
                    {a._count.funcoes} func · {a._count.pessoas} pess
                  </p>
                </div>
                <DeleteButton
                  action={excluirArea.bind(null, a.id)}
                  confirmText={`Excluir área "${a.descricao}"?`}
                />
              </li>
            ))}
          </ul>
        </section>

        {/* Funções */}
        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Funções ({funcoes.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {funcoes.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhuma função cadastrada.</li>
            )}
            {funcoes.map((f) => (
              <li key={f.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{f.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{f.descricao}</p>
                  <p className="text-xs text-gray-medium">
                    {f.area ? `${f.area.codigo} ${f.area.descricao}` : 'sem área'} ·{' '}
                    {f._count.pessoas} pess
                  </p>
                </div>
                <DeleteButton
                  action={excluirFuncao.bind(null, f.id)}
                  confirmText={`Excluir função "${f.descricao}"?`}
                />
              </li>
            ))}
          </ul>
        </section>

        {/* Pessoas */}
        <section className="bg-white rounded-lg border border-[#E2E8F0]">
          <header className="px-5 py-4 border-b border-[#E2E8F0]">
            <h2 className="font-display text-lg text-navy">Pessoas ({pessoas.length})</h2>
          </header>
          <ul className="divide-y divide-[#E2E8F0]">
            {pessoas.length === 0 && (
              <li className="px-5 py-6 text-sm text-gray-medium italic">Nenhuma pessoa cadastrada.</li>
            )}
            {pessoas.map((p) => (
              <li key={p.id} className="px-5 py-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-mono text-xs text-teal">{p.codigo}</p>
                  <p className="text-sm font-medium text-navy truncate">{p.nome}</p>
                  <p className="text-xs text-gray-medium">
                    {p.funcao?.descricao ?? 'sem função'}
                    {p.area ? ` · ${p.area.descricao}` : ''}
                    {p._count.raciAtribuicoes > 0 ? ` · ${p._count.raciAtribuicoes} RACI` : ''}
                  </p>
                </div>
                <DeleteButton
                  action={excluirPessoa.bind(null, p.id)}
                  confirmText={`Excluir "${p.nome}"?`}
                />
              </li>
            ))}
          </ul>
        </section>
      </div>
    </div>
  )
}
