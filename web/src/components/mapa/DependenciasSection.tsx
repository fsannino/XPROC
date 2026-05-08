'use client'

import { useState } from 'react'
import { DEPENDENCIA_TIPOS, DEPENDENCIA_TIPO_LABEL, type DependenciaTipo } from '@/lib/definitions'

export type NoOption = { id: number; descricao: string }

export type DependenciaItem = {
  id: string
  origemId: number
  origemDescricao: string
  destinoId: number
  destinoDescricao: string
  tipo: DependenciaTipo
  descricao: string | null
}

type Props = {
  proprioId: number
  outrosNos: NoOption[] | null
  saidas: DependenciaItem[] | null
  entradas: DependenciaItem[] | null
  onAdd: (input: { destinoId: number; tipo: DependenciaTipo; descricao?: string }) => Promise<boolean>
  onRemove: (id: string) => Promise<boolean>
}

export default function DependenciasSection({ proprioId, outrosNos, saidas, entradas, onAdd, onRemove }: Props) {
  const [destinoId, setDestinoId] = useState<string>('')
  const [tipo, setTipo] = useState<DependenciaTipo>('PRECEDE')
  const [descricao, setDescricao] = useState('')
  const [pending, setPending] = useState(false)
  const [erro, setErro] = useState<string | null>(null)

  const carregando = saidas == null || entradas == null

  async function adicionar() {
    if (!destinoId) return
    const id = Number(destinoId)
    if (!Number.isFinite(id) || id === proprioId) return
    setPending(true)
    setErro(null)
    const ok = await onAdd({ destinoId: id, tipo, descricao: descricao || undefined })
    setPending(false)
    if (ok) {
      setDestinoId('')
      setDescricao('')
    } else {
      setErro('Não foi possível adicionar (duplicada?).')
    }
  }

  async function remover(id: string) {
    setPending(true)
    await onRemove(id)
    setPending(false)
  }

  function renderLista(titulo: string, itens: DependenciaItem[] | null, lado: 'saida' | 'entrada') {
    if (itens == null) return null
    if (itens.length === 0) {
      return (
        <div>
          <p className="text-[10px] font-bold tracking-[0.14em] uppercase text-navy/70 mb-1.5">{titulo}</p>
          <p className="text-xs text-gray-medium italic">Nenhuma.</p>
        </div>
      )
    }
    return (
      <div>
        <p className="text-[10px] font-bold tracking-[0.14em] uppercase text-navy/70 mb-1.5">{titulo}</p>
        <ul className="space-y-1.5">
          {itens.map((d) => (
            <li key={d.id} className="flex items-center gap-2 text-xs">
              <span className="flex-1 min-w-0">
                <span className="block font-medium text-navy truncate">
                  {lado === 'saida' ? d.destinoDescricao : d.origemDescricao}
                </span>
                <span className="block text-[10px] text-gray-medium truncate">
                  {DEPENDENCIA_TIPO_LABEL[d.tipo]}
                  {d.descricao ? ` · ${d.descricao}` : ''}
                </span>
              </span>
              <button
                type="button"
                onClick={() => remover(d.id)}
                disabled={pending}
                aria-label="Remover"
                className="w-6 h-6 rounded text-gray-medium hover:text-[#9A2E1F] hover:bg-[rgba(224,80,64,0.08)] transition-colors disabled:opacity-50"
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      </div>
    )
  }

  const candidatosDestino = (outrosNos ?? []).filter((n) => n.id !== proprioId)

  return (
    <fieldset className="border border-[#E2E8F0] rounded-md p-4 space-y-3">
      <legend className="px-2 text-[10px] font-bold tracking-[0.18em] uppercase text-teal">
        Dependências
      </legend>

      {carregando && <p className="text-xs text-gray-medium italic py-1">Carregando dependências...</p>}

      {!carregando && (
        <div className="space-y-3">
          {renderLista('Saídas (este → outro)', saidas, 'saida')}
          {renderLista('Entradas (outro → este)', entradas, 'entrada')}

          <div className="border-t border-[#E2E8F0] pt-3 space-y-2">
            <p className="text-[10px] font-bold tracking-[0.14em] uppercase text-navy/70">Adicionar dependência</p>
            <div className="grid grid-cols-2 gap-2">
              <select
                value={destinoId}
                onChange={(e) => setDestinoId(e.target.value)}
                className="rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-2 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
              >
                <option value="">Destino…</option>
                {candidatosDestino.map((n) => (
                  <option key={n.id} value={n.id}>
                    {n.descricao}
                  </option>
                ))}
              </select>
              <select
                value={tipo}
                onChange={(e) => setTipo(e.target.value as DependenciaTipo)}
                className="rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-2 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
              >
                {DEPENDENCIA_TIPOS.map((t) => (
                  <option key={t} value={t}>
                    {DEPENDENCIA_TIPO_LABEL[t]}
                  </option>
                ))}
              </select>
            </div>
            <input
              type="text"
              value={descricao}
              onChange={(e) => setDescricao(e.target.value)}
              maxLength={300}
              placeholder="Descrição (opcional)"
              className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-2 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
            />
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={adicionar}
                disabled={pending || !destinoId}
                className="px-3 py-1.5 rounded-md text-xs font-semibold bg-navy hover:bg-teal text-white transition-all disabled:opacity-50"
              >
                Adicionar
              </button>
              {erro && <span className="text-[11px] text-[#9A2E1F]">{erro}</span>}
            </div>
          </div>
        </div>
      )}
    </fieldset>
  )
}
