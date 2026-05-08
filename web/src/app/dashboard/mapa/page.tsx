import { getMapa } from '@/actions/mapa'
import { listarPessoasParaRaci } from '@/actions/raci'
import { listarProdutos } from '@/actions/produtos'
import { listarInsumos } from '@/actions/insumos'
import { listarSistemas } from '@/actions/sistemas'
import { listarTodasDependenciasDeProcesso, listarTodasDependenciasDeAtividade } from '@/actions/dependencias'
import { prisma } from '@/lib/prisma'
import MapaCanvas from './MapaCanvas'

export const metadata = { title: 'Mapa Visual' }

export default async function MapaPage() {
  const [initial, transacoes, pessoas, produtos, insumos, sistemas, depsProcesso, depsAtividade] = await Promise.all([
    getMapa(),
    prisma.transacao.findMany({
      orderBy: { id: 'asc' },
      select: { id: true, descricao: true },
    }),
    listarPessoasParaRaci(),
    listarProdutos(),
    listarInsumos(),
    listarSistemas(),
    listarTodasDependenciasDeProcesso(),
    listarTodasDependenciasDeAtividade(),
  ])

  return (
    <div className="flex flex-col h-[calc(100vh-7rem)]">
      <div className="mb-4">
        <p className="section-tag">Cadeia de Valor</p>
        <h1 className="section-title">Mapa Visual</h1>
        <p className="text-sm text-gray-medium">
          Cadeia de Valor → Macroprocesso → Processo → Macroatividade → Atividade.
          Arraste para reposicionar, clique para editar, use “+” para adicionar filhos.
        </p>
        <div className="gold-bar w-24 rounded-full mt-3" />
      </div>

      <div className="flex-1 rounded-lg border border-[#E2E8F0] bg-white overflow-hidden">
        <MapaCanvas
          initialNodes={initial.nodes}
          initialEdges={initial.edges}
          transacoes={transacoes.map((t) => ({
            value: t.id,
            label: t.descricao || t.id,
            hint: t.id,
          }))}
          pessoas={pessoas}
          produtos={produtos}
          insumos={insumos}
          sistemas={sistemas}
          dependenciasProcesso={depsProcesso}
          dependenciasAtividade={depsAtividade}
        />
      </div>
    </div>
  )
}
