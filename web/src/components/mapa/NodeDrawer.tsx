'use client'

import { useEffect, useState } from 'react'
import { MAPA_LEVELS, type NodeType } from '@/lib/definitions'

export type DrawerState =
  | { mode: 'create'; tipoFilho: NodeType | null; parentId?: number }
  | { mode: 'edit'; tipo: NodeType; id: number; descricao: string; abreviacao?: string; sequencia?: number }

type Props = {
  state: DrawerState
  onClose: () => void
  onSubmit: (payload: {
    mode: 'create' | 'edit'
    tipo: NodeType
    id?: number
    parentId?: number
    descricao: string
    abreviacao?: string
    sequencia?: number
  }) => Promise<boolean>
}

export default function NodeDrawer({ state, onClose, onSubmit }: Props) {
  const tipo = state.mode === 'edit' ? state.tipo : (state.tipoFilho ?? 'cadeia')
  const meta = MAPA_LEVELS[tipo]
  const usaAbreviacao = tipo === 'cadeia' || tipo === 'macroprocesso'
  const usaSequencia = tipo === 'processo' || tipo === 'macroatividade' || tipo === 'atividade'

  const [descricao, setDescricao] = useState(state.mode === 'edit' ? state.descricao : '')
  const [abreviacao, setAbreviacao] = useState(state.mode === 'edit' ? state.abreviacao ?? '' : '')
  const [sequencia, setSequencia] = useState<string>(
    state.mode === 'edit' && state.sequencia != null ? String(state.sequencia) : '',
  )
  const [pending, setPending] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    setError(null)
  }, [descricao, abreviacao, sequencia])

  async function handle(e: React.FormEvent) {
    e.preventDefault()
    if (!tipo) return
    setPending(true)
    setError(null)
    const ok = await onSubmit({
      mode: state.mode,
      tipo,
      id: state.mode === 'edit' ? state.id : undefined,
      parentId: state.mode === 'create' ? state.parentId : undefined,
      descricao,
      abreviacao: abreviacao || undefined,
      sequencia: sequencia ? Number(sequencia) : undefined,
    })
    setPending(false)
    if (!ok) setError('Não foi possível salvar.')
  }

  return (
    <div className="fixed inset-0 z-50" role="dialog" aria-modal="true">
      <button
        type="button"
        aria-label="Fechar"
        onClick={onClose}
        className="absolute inset-0 bg-navy-dark/40"
      />
      <aside className="absolute right-0 top-0 bottom-0 w-full sm:w-[420px] bg-white shadow-2xl flex flex-col">
        <div className="px-6 py-5 border-b border-[#E2E8F0]">
          <p className="section-tag mb-1">{state.mode === 'edit' ? 'Editar' : 'Novo'}</p>
          <h2 className="font-display text-2xl text-navy">{meta.label}</h2>
          <p className="text-xs text-gray-medium mt-1">
            {state.mode === 'create' && state.parentId
              ? `Será criado dentro do nó pai #${state.parentId}.`
              : state.mode === 'create'
              ? 'Cadeia de Valor é o nó raiz da estrutura.'
              : 'Atualize os campos abaixo.'}
          </p>
        </div>

        <form onSubmit={handle} className="flex-1 overflow-auto px-6 py-5 space-y-5">
          <div>
            <label htmlFor="descricao" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
              Descrição<span className="text-gold ml-0.5">*</span>
            </label>
            <input
              id="descricao"
              type="text"
              value={descricao}
              onChange={(e) => setDescricao(e.target.value)}
              required
              autoFocus
              minLength={2}
              maxLength={200}
              className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
            />
          </div>

          {usaAbreviacao && (
            <div>
              <label htmlFor="abreviacao" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
                Abreviação
              </label>
              <input
                id="abreviacao"
                type="text"
                value={abreviacao}
                onChange={(e) => setAbreviacao(e.target.value.toUpperCase())}
                maxLength={8}
                className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate font-mono focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
                placeholder="Ex: FI"
              />
            </div>
          )}

          {usaSequencia && (
            <div>
              <label htmlFor="sequencia" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
                Sequência
              </label>
              <input
                id="sequencia"
                type="number"
                value={sequencia}
                onChange={(e) => setSequencia(e.target.value)}
                min={0}
                className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
              />
            </div>
          )}

          <div className="bg-[rgba(247,168,35,0.08)] border-l-3 border-gold rounded-md px-3.5 py-2.5 text-xs text-slate">
            <strong className="text-navy">Próximos passos:</strong> KPIs, Áreas, Pessoas/Funções,
            Inputs/Outputs, Produtos e Dependências serão adicionados em entregas futuras.
          </div>

          {error && (
            <p role="alert" className="text-sm text-[#9A2E1F] bg-[rgba(224,80,64,0.08)] border-l-3 border-[#E05040] rounded-md px-3.5 py-2.5">
              {error}
            </p>
          )}
        </form>

        <div className="px-6 py-4 border-t border-[#E2E8F0] flex gap-3 justify-end bg-[#F5F6F8]">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-md text-sm font-semibold text-navy border border-[#E2E8F0] bg-white hover:border-teal hover:text-teal transition-all"
          >
            Cancelar
          </button>
          <button
            type="submit"
            disabled={pending || descricao.trim().length < 2}
            onClick={handle}
            className="px-5 py-2 rounded-md text-sm font-semibold bg-navy hover:bg-teal text-white transition-all hover:-translate-y-0.5 hover:shadow-md disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
          >
            {pending ? 'Salvando...' : state.mode === 'edit' ? 'Salvar' : 'Criar'}
          </button>
        </div>
      </aside>
    </div>
  )
}
