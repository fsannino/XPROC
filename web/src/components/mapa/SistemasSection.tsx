'use client'

import { useMemo } from 'react'
import {
  SISTEMA_PAPEIS,
  SISTEMA_PAPEL_LABEL,
  type SistemaTipo,
  type SistemaPapel,
} from '@/lib/definitions'

export type SistemaOption = {
  id: number
  codigo: string
  nome: string
  tipo: SistemaTipo
}

export type SistemaVinculo = { sistemaId: number; papel: SistemaPapel }

type Props = {
  sistemas: SistemaOption[] | null
  vinculos: SistemaVinculo[] | null
  onChange: (next: SistemaVinculo[]) => void
}

export default function SistemasSection({ sistemas, vinculos, onChange }: Props) {
  const carregando = vinculos == null
  const lista = vinculos ?? []

  const index = useMemo(() => {
    const m = new Map<number, SistemaOption>()
    for (const s of sistemas ?? []) m.set(s.id, s)
    return m
  }, [sistemas])

  function setPapel(sistemaId: number, papel: SistemaPapel) {
    const idx = lista.findIndex((v) => v.sistemaId === sistemaId)
    if (idx >= 0) {
      const next = lista.slice()
      next[idx] = { sistemaId, papel }
      onChange(next)
    } else {
      onChange([...lista, { sistemaId, papel }])
    }
  }

  function remover(sistemaId: number) {
    onChange(lista.filter((v) => v.sistemaId !== sistemaId))
  }

  const naoSelecionados = (sistemas ?? []).filter((s) => !lista.some((v) => v.sistemaId === s.id))

  return (
    <fieldset className="border border-[#E2E8F0] rounded-md p-4 space-y-3">
      <legend className="px-2 text-[10px] font-bold tracking-[0.18em] uppercase text-teal">
        Sistemas
      </legend>

      {carregando && <p className="text-xs text-gray-medium italic py-1">Carregando sistemas...</p>}

      {!carregando && lista.length === 0 && (
        <p className="text-xs text-gray-medium italic">
          Nenhum sistema vinculado. Cadastre em <a href="/dashboard/catalogo" className="text-teal hover:underline">/catálogo</a>.
        </p>
      )}

      {!carregando && lista.length > 0 && (
        <ul className="space-y-2">
          {lista.map((v) => {
            const s = index.get(v.sistemaId)
            return (
              <li key={v.sistemaId} className="flex items-center gap-2 text-xs">
                <span className="flex-1 min-w-0">
                  <span className="block font-medium text-navy truncate">{s ? s.nome : `#${v.sistemaId}`}</span>
                  <span className="block text-[10px] text-gray-medium truncate">{s ? `${s.tipo} · ${s.codigo}` : ''}</span>
                </span>
                <select
                  value={v.papel}
                  onChange={(e) => setPapel(v.sistemaId, e.target.value as SistemaPapel)}
                  className="rounded-md border border-[#E2E8F0] bg-white px-2 py-1 text-xs font-mono text-navy focus:outline-none focus:ring-2 focus:ring-teal"
                >
                  {SISTEMA_PAPEIS.map((p) => (
                    <option key={p} value={p}>
                      {SISTEMA_PAPEL_LABEL[p]}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() => remover(v.sistemaId)}
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
          <label htmlFor="sistema-add" className="sr-only">Adicionar sistema</label>
          <select
            id="sistema-add"
            value=""
            onChange={(e) => {
              const id = Number(e.target.value)
              if (id) setPapel(id, 'CONSUMIDOR')
              e.target.value = ''
            }}
            className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
          >
            <option value="">+ Adicionar sistema…</option>
            {naoSelecionados.map((s) => (
              <option key={s.id} value={s.id}>
                {s.nome} — {s.tipo}
              </option>
            ))}
          </select>
        </div>
      )}
    </fieldset>
  )
}
