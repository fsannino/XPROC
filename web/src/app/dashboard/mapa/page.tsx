import { getMapa } from '@/actions/mapa'
import MapaCanvas from './MapaCanvas'

export const metadata = { title: 'Mapa Visual' }

export default async function MapaPage() {
  const initial = await getMapa()

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
        <MapaCanvas initialNodes={initial.nodes} initialEdges={initial.edges} />
      </div>
    </div>
  )
}
