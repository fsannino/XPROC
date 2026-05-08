'use client'

import { useMemo } from 'react'
import { PRODUTO_TIPO_LABEL, type ProdutoTipo } from '@/lib/definitions'

export type ProdutoOption = {
  id: number
  codigo: string
  descricao: string
  tipo: ProdutoTipo
}

type Props = {
  produtos: ProdutoOption[] | null
  selecionados: number[] | null
  onChange: (next: number[]) => void
}

export default function ProdutosSection({ produtos, selecionados, onChange }: Props) {
  const carregando = selecionados == null
  const lista = selecionados ?? []

  const index = useMemo(() => {
    const m = new Map<number, ProdutoOption>()
    for (const p of produtos ?? []) m.set(p.id, p)
    return m
  }, [produtos])

  function remover(id: number) {
    onChange(lista.filter((x) => x !== id))
  }

  function adicionar(id: number) {
    if (lista.includes(id)) return
    onChange([...lista, id])
  }

  const naoSelecionados = (produtos ?? []).filter((p) => !lista.includes(p.id))

  return (
    <fieldset className="border border-[#E2E8F0] rounded-md p-4 space-y-3">
      <legend className="px-2 text-[10px] font-bold tracking-[0.18em] uppercase text-teal">
        Produtos
      </legend>

      {carregando && <p className="text-xs text-gray-medium italic py-1">Carregando produtos...</p>}

      {!carregando && lista.length === 0 && (
        <p className="text-xs text-gray-medium italic">
          Nenhum produto associado. Cadastre em <a href="/dashboard/catalogo" className="text-teal hover:underline">/catálogo</a>.
        </p>
      )}

      {!carregando && lista.length > 0 && (
        <ul className="space-y-2">
          {lista.map((id) => {
            const p = index.get(id)
            return (
              <li key={id} className="flex items-center gap-2 text-xs">
                <span className="flex-1 min-w-0">
                  <span className="block font-medium text-navy truncate">{p ? p.descricao : `#${id}`}</span>
                  <span className="block text-[10px] text-gray-medium truncate">
                    {p ? `${PRODUTO_TIPO_LABEL[p.tipo]} · ${p.codigo}` : ''}
                  </span>
                </span>
                <button
                  type="button"
                  onClick={() => remover(id)}
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

      {!carregando && naoSelecionados.length > 0 && (
        <div>
          <label htmlFor="produto-add" className="sr-only">Adicionar produto</label>
          <select
            id="produto-add"
            value=""
            onChange={(e) => {
              const id = Number(e.target.value)
              if (id) adicionar(id)
              e.target.value = ''
            }}
            className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
          >
            <option value="">+ Adicionar produto…</option>
            {naoSelecionados.map((p) => (
              <option key={p.id} value={p.id}>
                {p.descricao} — {PRODUTO_TIPO_LABEL[p.tipo]}
              </option>
            ))}
          </select>
        </div>
      )}
    </fieldset>
  )
}
