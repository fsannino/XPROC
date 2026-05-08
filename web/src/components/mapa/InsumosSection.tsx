'use client'

import { useMemo } from 'react'
import { INSUMO_TIPO_LABEL, type InsumoTipo, type InsumoDirecao } from '@/lib/definitions'

export type InsumoOption = {
  id: number
  codigo: string
  descricao: string
  tipo: InsumoTipo
}

export type InsumoVinculo = { insumoId: number; direcao: InsumoDirecao }

type Props = {
  insumos: InsumoOption[] | null
  vinculos: InsumoVinculo[] | null
  onChange: (next: InsumoVinculo[]) => void
}

const DIRECAO_LABEL: Record<InsumoDirecao, string> = {
  INPUT: 'Entrada',
  OUTPUT: 'Saída',
}

export default function InsumosSection({ insumos, vinculos, onChange }: Props) {
  const carregando = vinculos == null
  const lista = vinculos ?? []

  const index = useMemo(() => {
    const m = new Map<number, InsumoOption>()
    for (const i of insumos ?? []) m.set(i.id, i)
    return m
  }, [insumos])

  const inputs = lista.filter((v) => v.direcao === 'INPUT')
  const outputs = lista.filter((v) => v.direcao === 'OUTPUT')

  function adicionar(insumoId: number, direcao: InsumoDirecao) {
    if (lista.some((v) => v.insumoId === insumoId && v.direcao === direcao)) return
    onChange([...lista, { insumoId, direcao }])
  }

  function remover(insumoId: number, direcao: InsumoDirecao) {
    onChange(lista.filter((v) => !(v.insumoId === insumoId && v.direcao === direcao)))
  }

  function renderLista(itens: InsumoVinculo[], direcao: InsumoDirecao) {
    return (
      <div>
        <p className="text-[10px] font-bold tracking-[0.14em] uppercase text-navy/70 mb-1.5">
          {DIRECAO_LABEL[direcao]}s
        </p>
        {itens.length === 0 && (
          <p className="text-xs text-gray-medium italic mb-2">Nenhum.</p>
        )}
        {itens.length > 0 && (
          <ul className="space-y-1.5 mb-2">
            {itens.map((v) => {
              const i = index.get(v.insumoId)
              return (
                <li key={`${v.insumoId}-${v.direcao}`} className="flex items-center gap-2 text-xs">
                  <span className="flex-1 min-w-0">
                    <span className="block font-medium text-navy truncate">
                      {i ? i.descricao : `#${v.insumoId}`}
                    </span>
                    <span className="block text-[10px] text-gray-medium truncate">
                      {i ? `${INSUMO_TIPO_LABEL[i.tipo]} · ${i.codigo}` : ''}
                    </span>
                  </span>
                  <button
                    type="button"
                    onClick={() => remover(v.insumoId, v.direcao)}
                    aria-label="Remover"
                    className="w-6 h-6 rounded text-gray-medium hover:text-[#9A2E1F] hover:bg-[rgba(224,80,64,0.08)] transition-colors"
                  >
                    ×
                  </button>
                </li>
              )
            })}
          </ul>
        )}
        <select
          value=""
          onChange={(e) => {
            const id = Number(e.target.value)
            if (id) adicionar(id, direcao)
            e.target.value = ''
          }}
          className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
        >
          <option value="">+ {DIRECAO_LABEL[direcao].toLowerCase()}…</option>
          {(insumos ?? [])
            .filter((i) => !itens.some((v) => v.insumoId === i.id))
            .map((i) => (
              <option key={i.id} value={i.id}>
                {i.descricao} — {INSUMO_TIPO_LABEL[i.tipo]}
              </option>
            ))}
        </select>
      </div>
    )
  }

  return (
    <fieldset className="border border-[#E2E8F0] rounded-md p-4 space-y-3">
      <legend className="px-2 text-[10px] font-bold tracking-[0.18em] uppercase text-teal">
        Inputs / Outputs
      </legend>

      {carregando && <p className="text-xs text-gray-medium italic py-1">Carregando insumos...</p>}

      {!carregando && (
        <div className="space-y-3">
          {renderLista(inputs, 'INPUT')}
          {renderLista(outputs, 'OUTPUT')}
        </div>
      )}

      {!carregando && (insumos == null || insumos.length === 0) && (
        <p className="text-xs text-gray-medium italic">
          Cadastre insumos em <a href="/dashboard/catalogo" className="text-teal hover:underline">/catálogo</a>.
        </p>
      )}
    </fieldset>
  )
}
