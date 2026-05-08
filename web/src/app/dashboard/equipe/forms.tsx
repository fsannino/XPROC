'use client'

import { useActionState, useEffect, useRef } from 'react'
import { criarAreaForm, criarFuncaoForm, criarPessoaForm } from '@/actions/equipe'
import { useToast } from '@/components/ui/toast'

type AreaOption = { id: number; codigo: string; descricao: string }
type FuncaoOption = { id: number; codigo: string; descricao: string }

type Props = {
  areas: AreaOption[]
  funcoes: FuncaoOption[]
}

export default function EquipeForms({ areas, funcoes }: Props) {
  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      <NovaAreaForm areas={areas} />
      <NovaFuncaoForm areas={areas} />
      <NovaPessoaForm areas={areas} funcoes={funcoes} />
    </div>
  )
}

function NovaAreaForm({ areas }: { areas: AreaOption[] }) {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarAreaForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Área criada.')
      formRef.current?.reset()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Nova Área</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="descricao" label="Descrição" required maxLength={150} />
      <FieldSelect name="parentId" label="Área pai" options={areas} placeholder="— raiz —" />
      <Submit pending={pending} label="Criar Área" />
    </form>
  )
}

function NovaFuncaoForm({ areas }: { areas: AreaOption[] }) {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarFuncaoForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Função criada.')
      formRef.current?.reset()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Nova Função</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="descricao" label="Descrição" required maxLength={150} />
      <FieldSelect name="areaId" label="Área" options={areas} placeholder="— sem área —" />
      <Submit pending={pending} label="Criar Função" />
    </form>
  )
}

function NovaPessoaForm({ areas, funcoes }: { areas: AreaOption[]; funcoes: FuncaoOption[] }) {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarPessoaForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Pessoa criada.')
      formRef.current?.reset()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Nova Pessoa</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="nome" label="Nome" required maxLength={150} />
      <FieldText name="email" label="Email" type="email" maxLength={150} />
      <div className="grid grid-cols-2 gap-3">
        <FieldSelect name="areaId" label="Área" options={areas} placeholder="—" />
        <FieldSelect name="funcaoId" label="Função" options={funcoes} placeholder="—" />
      </div>
      <Submit pending={pending} label="Criar Pessoa" />
    </form>
  )
}

// ─── Building blocks ───────────────────────────────────────────────

function FieldText(props: {
  name: string
  label: string
  required?: boolean
  maxLength?: number
  type?: string
  mono?: boolean
}) {
  return (
    <div>
      <label htmlFor={props.name} className="block text-xs font-semibold tracking-wider uppercase text-navy mb-1.5">
        {props.label}
        {props.required && <span className="text-gold ml-0.5">*</span>}
      </label>
      <input
        id={props.name}
        name={props.name}
        type={props.type ?? 'text'}
        required={props.required}
        maxLength={props.maxLength}
        className={`w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent ${
          props.mono ? 'font-mono uppercase' : ''
        }`}
      />
    </div>
  )
}

function FieldSelect(props: {
  name: string
  label: string
  options: { id: number; codigo: string; descricao: string }[]
  placeholder: string
}) {
  return (
    <div>
      <label htmlFor={props.name} className="block text-xs font-semibold tracking-wider uppercase text-navy mb-1.5">
        {props.label}
      </label>
      <select
        id={props.name}
        name={props.name}
        defaultValue=""
        className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
      >
        <option value="">{props.placeholder}</option>
        {props.options.map((o) => (
          <option key={o.id} value={o.id}>
            {o.codigo} — {o.descricao}
          </option>
        ))}
      </select>
    </div>
  )
}

function Submit({ pending, label }: { pending: boolean; label: string }) {
  return (
    <button
      type="submit"
      disabled={pending}
      className="w-full px-4 py-2 rounded-md text-sm font-semibold bg-navy hover:bg-teal text-white transition-all hover:-translate-y-0.5 hover:shadow-md disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
    >
      {pending ? 'Salvando...' : label}
    </button>
  )
}
