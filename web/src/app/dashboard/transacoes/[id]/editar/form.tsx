'use client'

import { useState, useTransition } from 'react'
import { useRouter } from 'next/navigation'
import { atualizarTransacao } from '@/actions/transacoes'
import {
  setTransacaoMacroprocessos,
  setTransacaoProcessos,
  setTransacaoAtividades,
} from '@/actions/relacoes'
import MultiSelect, { type MultiSelectOption } from '@/components/ui/multi-select'

type Props = {
  transacao: { id: string; descricao: string | null }
  macros: MultiSelectOption[]
  processos: MultiSelectOption[]
  atividades: MultiSelectOption[]
  initialMacroIds: string[]
  initialProcessoIds: string[]
  initialAtividadeIds: string[]
}

export default function TransacaoEditarForm(props: Props) {
  const router = useRouter()
  const [pending, startTransition] = useTransition()
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState<string | null>(null)

  const [nome, setNome] = useState(props.transacao.descricao ?? '')
  const [macroIds, setMacroIds] = useState(props.initialMacroIds)
  const [processoIds, setProcessoIds] = useState(props.initialProcessoIds)
  const [atividadeIds, setAtividadeIds] = useState(props.initialAtividadeIds)

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    setSuccess(null)
    startTransition(async () => {
      try {
        const fd = new FormData()
        fd.append('descricao', nome)
        await atualizarTransacao(props.transacao.id, undefined, fd)

        const results = await Promise.all([
          setTransacaoMacroprocessos({
            transacaoId: props.transacao.id,
            megaProcessoIds: macroIds.map(Number),
          }),
          setTransacaoProcessos({
            transacaoId: props.transacao.id,
            processoIds: processoIds.map(Number),
          }),
          setTransacaoAtividades({
            transacaoId: props.transacao.id,
            atividadeIds: atividadeIds.map(Number),
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
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
              Código
            </label>
            <input
              type="text"
              value={props.transacao.id}
              readOnly
              className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm font-mono text-slate cursor-not-allowed"
            />
          </div>
          <div className="md:col-span-2">
            <label htmlFor="nome" className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2">
              Nome<span className="text-gold ml-0.5">*</span>
            </label>
            <input
              id="nome"
              type="text"
              value={nome}
              onChange={(e) => setNome(e.target.value)}
              maxLength={150}
              className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
            />
          </div>
        </div>
      </div>

      <Section title="Macroprocessos" hint="Macroprocessos onde a transação aparece">
        <MultiSelect
          options={props.macros}
          selected={macroIds}
          onChange={setMacroIds}
          placeholder="Buscar macroprocesso..."
        />
      </Section>

      <Section title="Processos" hint="Processos dos quais a transação participa">
        <MultiSelect
          options={props.processos}
          selected={processoIds}
          onChange={setProcessoIds}
          placeholder="Buscar processo..."
        />
      </Section>

      <Section title="Atividades" hint="Atividades que a transação realiza">
        <MultiSelect
          options={props.atividades}
          selected={atividadeIds}
          onChange={setAtividadeIds}
          placeholder="Buscar atividade..."
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
          disabled={pending || nome.trim().length < 2}
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
