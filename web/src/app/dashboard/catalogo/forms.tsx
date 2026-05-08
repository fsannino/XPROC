'use client'

import { useActionState, useEffect, useRef } from 'react'
import { criarProdutoForm } from '@/actions/produtos'
import { criarInsumoForm } from '@/actions/insumos'
import { criarSistemaForm } from '@/actions/sistemas'
import {
  PRODUTO_TIPOS,
  PRODUTO_TIPO_LABEL,
  INSUMO_TIPOS,
  INSUMO_TIPO_LABEL,
  SISTEMA_TIPOS,
} from '@/lib/definitions'
import { useToast } from '@/components/ui/toast'

export default function CatalogoForms() {
  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      <NovoProdutoForm />
      <NovoInsumoForm />
      <NovoSistemaForm />
    </div>
  )
}

function NovoProdutoForm() {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarProdutoForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) { show('Produto criado.'); formRef.current?.reset() }
    else if ('error' in state && state.error) show(state.error, 'error')
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Novo Produto</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="descricao" label="Descrição" required maxLength={150} />
      <FieldSelect name="tipo" label="Tipo" defaultValue="BEM" options={PRODUTO_TIPOS.map((t) => ({ value: t, label: PRODUTO_TIPO_LABEL[t] }))} />
      <Submit pending={pending} label="Criar Produto" />
    </form>
  )
}

function NovoInsumoForm() {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarInsumoForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) { show('Insumo criado.'); formRef.current?.reset() }
    else if ('error' in state && state.error) show(state.error, 'error')
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Novo Insumo</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="descricao" label="Descrição" required maxLength={150} />
      <FieldSelect name="tipo" label="Tipo" defaultValue="DADO" options={INSUMO_TIPOS.map((t) => ({ value: t, label: INSUMO_TIPO_LABEL[t] }))} />
      <Submit pending={pending} label="Criar Insumo" />
    </form>
  )
}

function NovoSistemaForm() {
  const { show } = useToast()
  const formRef = useRef<HTMLFormElement>(null)
  const [state, action, pending] = useActionState(criarSistemaForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) { show('Sistema criado.'); formRef.current?.reset() }
    else if ('error' in state && state.error) show(state.error, 'error')
  }, [state, show])

  return (
    <form ref={formRef} action={action} className="bg-white border border-[#E2E8F0] rounded-lg p-5 space-y-3">
      <h3 className="font-display text-base text-navy">Novo Sistema</h3>
      <FieldText name="codigo" label="Código" required maxLength={20} mono />
      <FieldText name="nome" label="Nome" required maxLength={150} />
      <FieldSelect name="tipo" label="Tipo" defaultValue="OUTRO" options={SISTEMA_TIPOS.map((t) => ({ value: t, label: t }))} />
      <Submit pending={pending} label="Criar Sistema" />
    </form>
  )
}

function FieldText(props: { name: string; label: string; required?: boolean; maxLength?: number; mono?: boolean }) {
  return (
    <div>
      <label htmlFor={props.name} className="block text-xs font-semibold tracking-wider uppercase text-navy mb-1.5">
        {props.label}
        {props.required && <span className="text-gold ml-0.5">*</span>}
      </label>
      <input
        id={props.name}
        name={props.name}
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
  defaultValue: string
  options: { value: string; label: string }[]
}) {
  return (
    <div>
      <label htmlFor={props.name} className="block text-xs font-semibold tracking-wider uppercase text-navy mb-1.5">
        {props.label}
      </label>
      <select
        id={props.name}
        name={props.name}
        defaultValue={props.defaultValue}
        className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3 py-2 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent"
      >
        {props.options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
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
