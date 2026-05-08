'use client'

import { useMemo } from 'react'
import { RACI_PAPEIS, RACI_PAPEL_LABEL, type RaciPapel } from '@/lib/definitions'

export type RaciAtribuicao = { pessoaId: number; papel: RaciPapel }

export type PessoaOption = {
  id: number
  codigo: string
  nome: string
  funcao: string | null
  area: string | null
}

type Props = {
  pessoas: PessoaOption[] | null
  atribuicoes: RaciAtribuicao[] | null
  onChange: (next: RaciAtribuicao[]) => void
}

export default function RaciSection({ pessoas, atribuicoes, onChange }: Props) {
  const carregando = atribuicoes == null
  const lista = atribuicoes ?? []

  const pessoasIndex = useMemo(() => {
    const m = new Map<number, PessoaOption>()
    for (const p of pessoas ?? []) m.set(p.id, p)
    return m
  }, [pessoas])

  function setPapel(pessoaId: number, papel: RaciPapel | '') {
    if (!papel) {
      onChange(lista.filter((a) => a.pessoaId !== pessoaId))
      return
    }
    const idx = lista.findIndex((a) => a.pessoaId === pessoaId)
    if (idx >= 0) {
      const next = lista.slice()
      next[idx] = { pessoaId, papel }
      onChange(next)
    } else {
      onChange([...lista, { pessoaId, papel }])
    }
  }

  function adicionarPessoa(pessoaId: number) {
    if (lista.some((a) => a.pessoaId === pessoaId)) return
    onChange([...lista, { pessoaId, papel: 'R' }])
  }

  const naoAtribuidas = (pessoas ?? []).filter(
    (p) => !lista.some((a) => a.pessoaId === p.id),
  )

  return (
    <fieldset className="border border-[#E2E8F0] rounded-md p-4 space-y-3">
      <legend className="px-2 text-[10px] font-bold tracking-[0.18em] uppercase text-teal">
        RACI
      </legend>

      {carregando && (
        <p className="text-xs text-gray-medium italic py-1">Carregando atribuições...</p>
      )}

      {!carregando && lista.length === 0 && (
        <p className="text-xs text-gray-medium italic">
          Nenhuma atribuição. Use o seletor abaixo para adicionar pessoas.
        </p>
      )}

      {!carregando && lista.length > 0 && (
        <ul className="space-y-2">
          {lista.map((a) => {
            const p = pessoasIndex.get(a.pessoaId)
            return (
              <li key={a.pessoaId} className="flex items-center gap-2 text-xs">
                <span className="flex-1 min-w-0">
                  <span className="block font-medium text-navy truncate">
                    {p ? p.nome : `#${a.pessoaId}`}
                  </span>
                  <span className="block text-[10px] text-gray-medium truncate">
                    {p?.funcao ?? p?.area ?? p?.codigo ?? ''}
                  </span>
                </span>
                <select
                  value={a.papel}
                  onChange={(e) => setPapel(a.pessoaId, e.target.value as RaciPapel | '')}
                  className="rounded-md border border-[#E2E8F0] bg-white px-2 py-1 text-xs font-mono text-navy focus:outline-none focus:ring-2 focus:ring-teal"
                >
                  {RACI_PAPEIS.map((papel) => (
                    <option key={papel} value={papel}>
                      {papel} · {RACI_PAPEL_LABEL[papel]}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() => setPapel(a.pessoaId, '')}
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

      {!carregando && naoAtribuidas.length > 0 && (
        <div>
          <label htmlFor="raci-add" className="sr-only">Adicionar pessoa</label>
          <select
            id="raci-add"
            value=""
            onChange={(e) => {
              const id = Number(e.target.value)
              if (id) adicionarPessoa(id)
              e.target.value = ''
            }}
            className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
          >
            <option value="">+ Adicionar pessoa…</option>
            {naoAtribuidas.map((p) => (
              <option key={p.id} value={p.id}>
                {p.nome}{p.funcao ? ` — ${p.funcao}` : ''}
              </option>
            ))}
          </select>
        </div>
      )}

      {!carregando && pessoas && pessoas.length === 0 && (
        <p className="text-xs text-gray-medium italic">
          Cadastre pessoas em <a href="/dashboard/equipe" className="text-teal hover:underline">/equipe</a> para atribuir.
        </p>
      )}
    </fieldset>
  )
}
