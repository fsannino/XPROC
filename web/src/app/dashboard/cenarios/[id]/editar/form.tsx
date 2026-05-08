'use client'

import { useState, useTransition } from 'react'
import { useRouter } from 'next/navigation'
import { atualizarCenario } from '@/actions/cenarios'
import {
  setCenarioProcessos,
  setCenarioAtividades,
  setCenarioTransacoes,
} from '@/actions/relacoes'
import MultiSelect, { type MultiSelectOption } from '@/components/ui/multi-select'

type Props = {
  cenario: { id: number; descricao: string; situacao: string | null }
  processos: MultiSelectOption[]
  atividades: MultiSelectOption[]
  transacoes: MultiSelectOption[]
  initialProcessoIds: string[]
  initialAtividadeIds: string[]
  initialTransacaoIds: string[]
}

export default function CenarioEditarForm(props: Props) {
  const router = useRouter()
  const [pending, startTransition] = useTransition()
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState<string | null>(null)

  const [descricao, setDescricao] = useState(props.cenario.descricao)
  const [situacao, setSituacao] = useState(props.cenario.situacao ?? '')
  const [processoIds, setProcessoIds] = useState(props.initialProcessoIds)
  const [atividadeIds, setAtividadeIds] = useState(props.initialAtividadeIds)
  const [transacaoIds, setTransacaoIds] = useState(props.initialTransacaoIds)

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    setSuccess(null)
    startTransition(async () => {
      try {
        const fd = new FormData()
        fd.append('descricao', descricao)
        fd.append('situacao', situacao)
        await atualizarCenario(props.cenario.id, undefined, fd)

        const results = await Promise.all([
          setCenarioProcessos({
            cenarioId: props.cenario.id,
            processoIds: processoIds.map(Number),
          }),
          setCenarioAtividades({
            cenarioId: props.cenario.id,
            atividadeIds: atividadeIds.map(Number),
          }),
          setCenarioTransacoes({
            cenarioId: props.cenario.id,
            transacaoIds,
          }),
        ])

        const failed = results.find((r) => 'error' in r && r.error)
        if (failed && 'error' in failed) {
          setError(failed.error ?? 'Falha ao salvar relações.')
          return
        }

        setSuccess('Salvo com sucesso.')
        router.refresh()
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Erro inesperado.')
      }
    })
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-6">
      <div className="bg-white rounded-lg border border-[#E2E8F0] p-6 space-y-4">
        <div>
          <label htmlFor="descricao" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
            Descrição<span className="text-gold ml-0.5">*</span>
          </label>
          <input
            id="descricao"
            type="text"
            value={descricao}
            onChange={(e) => setDescricao(e.target.value)}
            minLength={2}
            maxLength={150}
            required
            className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
          />
        </div>
        <div>
          <label htmlFor="situacao" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
            Situação
          </label>
          <input
            id="situacao"
            type="text"
            value={situacao}
            onChange={(e) => setSituacao(e.target.value)}
            maxLength={20}
            className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
          />
        </div>
      </div>

      <Section title="Processos" hint="Processos contemplados pelo cenário">
        <MultiSelect
          options={props.processos}
          selected={processoIds}
          onChange={setProcessoIds}
          placeholder="Buscar processo..."
        />
      </Section>

      <Section title="Atividades" hint="Atividades contempladas pelo cenário">
        <MultiSelect
          options={props.atividades}
          selected={atividadeIds}
          onChange={setAtividadeIds}
          placeholder="Buscar atividade..."
        />
      </Section>

      <Section title="Transações" hint="Transações vinculadas ao cenário">
        <MultiSelect
          options={props.transacoes}
          selected={transacaoIds}
          onChange={setTransacaoIds}
          placeholder="Buscar transação..."
        />
      </Section>

      {error && (
        <p role="alert" className="text-sm text-[#9A2E1F] bg-[rgba(224,80,64,0.08)] border-l-3 border-[#E05040] rounded-md px-3.5 py-2.5">
          {error}
        </p>
      )}
      {success && (
        <p className="text-sm text-teal bg-teal/8 border-l-3 border-teal rounded-md px-3.5 py-2.5">
          {success}
        </p>
      )}

      <div className="flex justify-end gap-3">
        <button
          type="submit"
          disabled={pending || descricao.trim().length < 2}
          className="px-5 py-2.5 rounded-md text-sm font-semibold bg-navy hover:bg-teal text-white transition-all hover:-translate-y-0.5 hover:shadow-md disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
        >
          {pending ? 'Salvando...' : 'Salvar tudo'}
        </button>
      </div>
    </form>
  )
}

function Section({ title, hint, children }: { title: string; hint: string; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-lg border border-[#E2E8F0] p-6">
      <div className="mb-3">
        <h3 className="font-display text-lg text-navy">{title}</h3>
        <p className="text-xs text-gray-medium">{hint}</p>
      </div>
      {children}
    </div>
  )
}
