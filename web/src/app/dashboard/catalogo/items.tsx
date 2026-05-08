'use client'

import { useActionState, useEffect, useState } from 'react'
import { atualizarProdutoForm, excluirProduto } from '@/actions/produtos'
import { atualizarInsumoForm, excluirInsumo } from '@/actions/insumos'
import { atualizarSistemaForm, excluirSistema } from '@/actions/sistemas'
import {
  PRODUTO_TIPOS,
  PRODUTO_TIPO_LABEL,
  INSUMO_TIPOS,
  INSUMO_TIPO_LABEL,
  SISTEMA_TIPOS,
  type ProdutoTipo,
  type InsumoTipo,
  type SistemaTipo,
} from '@/lib/definitions'
import { useToast } from '@/components/ui/toast'

// ─── Produto ────────────────────────────────────────────────────────

export type ProdutoLinha = {
  id: number
  codigo: string
  descricao: string
  tipo: ProdutoTipo
  processosCount: number
}

export function ProdutoItem({ p }: { p: ProdutoLinha }) {
  const [editando, setEditando] = useState(false)
  return editando ? (
    <ProdutoEditForm p={p} onCancel={() => setEditando(false)} onSaved={() => setEditando(false)} />
  ) : (
    <Linha
      codigo={p.codigo}
      titulo={p.descricao}
      meta={`${PRODUTO_TIPO_LABEL[p.tipo]} · ${p.processosCount} proc`}
      onEdit={() => setEditando(true)}
      onDelete={() => excluirProduto(p.id)}
      confirmText={`Excluir produto "${p.descricao}"?`}
    />
  )
}

function ProdutoEditForm({
  p,
  onCancel,
  onSaved,
}: {
  p: ProdutoLinha
  onCancel: () => void
  onSaved: () => void
}) {
  const { show } = useToast()
  const [state, action, pending] = useActionState(atualizarProdutoForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Produto atualizado.')
      onSaved()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show, onSaved])

  return (
    <form action={action} className="space-y-2">
      <input type="hidden" name="id" value={p.id} />
      <CodigoInput defaultValue={p.codigo} />
      <DescricaoInput defaultValue={p.descricao} placeholder="Descrição" />
      <TipoSelect
        name="tipo"
        defaultValue={p.tipo}
        options={PRODUTO_TIPOS.map((t) => ({ value: t, label: PRODUTO_TIPO_LABEL[t] }))}
      />
      <BotoesEdit pending={pending} onCancel={onCancel} />
    </form>
  )
}

// ─── Insumo ─────────────────────────────────────────────────────────

export type InsumoLinha = {
  id: number
  codigo: string
  descricao: string
  tipo: InsumoTipo
  processosCount: number
  atividadesCount: number
}

export function InsumoItem({ i }: { i: InsumoLinha }) {
  const [editando, setEditando] = useState(false)
  return editando ? (
    <InsumoEditForm i={i} onCancel={() => setEditando(false)} onSaved={() => setEditando(false)} />
  ) : (
    <Linha
      codigo={i.codigo}
      titulo={i.descricao}
      meta={`${INSUMO_TIPO_LABEL[i.tipo]} · ${i.processosCount} proc · ${i.atividadesCount} ativ`}
      onEdit={() => setEditando(true)}
      onDelete={() => excluirInsumo(i.id)}
      confirmText={`Excluir insumo "${i.descricao}"?`}
    />
  )
}

function InsumoEditForm({
  i,
  onCancel,
  onSaved,
}: {
  i: InsumoLinha
  onCancel: () => void
  onSaved: () => void
}) {
  const { show } = useToast()
  const [state, action, pending] = useActionState(atualizarInsumoForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Insumo atualizado.')
      onSaved()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show, onSaved])

  return (
    <form action={action} className="space-y-2">
      <input type="hidden" name="id" value={i.id} />
      <CodigoInput defaultValue={i.codigo} />
      <DescricaoInput defaultValue={i.descricao} placeholder="Descrição" />
      <TipoSelect
        name="tipo"
        defaultValue={i.tipo}
        options={INSUMO_TIPOS.map((t) => ({ value: t, label: INSUMO_TIPO_LABEL[t] }))}
      />
      <BotoesEdit pending={pending} onCancel={onCancel} />
    </form>
  )
}

// ─── Sistema ────────────────────────────────────────────────────────

export type SistemaLinha = {
  id: number
  codigo: string
  nome: string
  tipo: SistemaTipo
  processosCount: number
}

export function SistemaItem({ s }: { s: SistemaLinha }) {
  const [editando, setEditando] = useState(false)
  return editando ? (
    <SistemaEditForm s={s} onCancel={() => setEditando(false)} onSaved={() => setEditando(false)} />
  ) : (
    <Linha
      codigo={s.codigo}
      titulo={s.nome}
      meta={`${s.tipo} · ${s.processosCount} proc`}
      onEdit={() => setEditando(true)}
      onDelete={() => excluirSistema(s.id)}
      confirmText={`Excluir sistema "${s.nome}"?`}
    />
  )
}

function SistemaEditForm({
  s,
  onCancel,
  onSaved,
}: {
  s: SistemaLinha
  onCancel: () => void
  onSaved: () => void
}) {
  const { show } = useToast()
  const [state, action, pending] = useActionState(atualizarSistemaForm, undefined)

  useEffect(() => {
    if (!state) return
    if ('success' in state) {
      show('Sistema atualizado.')
      onSaved()
    } else if ('error' in state && state.error) {
      show(state.error, 'error')
    }
  }, [state, show, onSaved])

  return (
    <form action={action} className="space-y-2">
      <input type="hidden" name="id" value={s.id} />
      <CodigoInput defaultValue={s.codigo} />
      <DescricaoInput defaultValue={s.nome} placeholder="Nome" name="nome" />
      <TipoSelect name="tipo" defaultValue={s.tipo} options={SISTEMA_TIPOS.map((t) => ({ value: t, label: t }))} />
      <BotoesEdit pending={pending} onCancel={onCancel} />
    </form>
  )
}

// ─── Building blocks ────────────────────────────────────────────────

function Linha({
  codigo,
  titulo,
  meta,
  onEdit,
  onDelete,
  confirmText,
}: {
  codigo: string
  titulo: string
  meta: string
  onEdit: () => void
  onDelete: () => Promise<unknown>
  confirmText: string
}) {
  const [pending, setPending] = useState(false)
  async function handleDelete() {
    if (!confirm(confirmText)) return
    setPending(true)
    await onDelete()
    setPending(false)
  }
  return (
    <div className="flex items-start justify-between gap-3">
      <div className="min-w-0">
        <p className="font-mono text-xs text-teal">{codigo}</p>
        <p className="text-sm font-medium text-navy truncate">{titulo}</p>
        <p className="text-xs text-gray-medium">{meta}</p>
      </div>
      <div className="flex items-center gap-1 shrink-0">
        <button
          type="button"
          onClick={onEdit}
          className="px-2 py-1 rounded text-[11px] font-semibold text-teal hover:bg-teal/10 transition-colors"
        >
          Editar
        </button>
        <button
          type="button"
          onClick={handleDelete}
          disabled={pending}
          aria-label="Excluir"
          className="w-7 h-7 rounded text-gray-medium hover:text-[#9A2E1F] hover:bg-[rgba(224,80,64,0.08)] transition-colors disabled:opacity-50"
        >
          ×
        </button>
      </div>
    </div>
  )
}

function CodigoInput({ defaultValue }: { defaultValue: string }) {
  return (
    <input
      name="codigo"
      defaultValue={defaultValue}
      required
      maxLength={20}
      className="w-full rounded-md border border-[#E2E8F0] bg-white px-2 py-1.5 text-xs font-mono uppercase text-slate focus:outline-none focus:ring-2 focus:ring-teal"
    />
  )
}

function DescricaoInput({
  defaultValue,
  placeholder,
  name = 'descricao',
}: {
  defaultValue: string
  placeholder: string
  name?: string
}) {
  return (
    <input
      name={name}
      defaultValue={defaultValue}
      placeholder={placeholder}
      required
      maxLength={150}
      className="w-full rounded-md border border-[#E2E8F0] bg-white px-2 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
    />
  )
}

function TipoSelect({
  name,
  defaultValue,
  options,
}: {
  name: string
  defaultValue: string
  options: { value: string; label: string }[]
}) {
  return (
    <select
      name={name}
      defaultValue={defaultValue}
      className="w-full rounded-md border border-[#E2E8F0] bg-white px-2 py-1.5 text-xs text-slate focus:outline-none focus:ring-2 focus:ring-teal"
    >
      {options.map((o) => (
        <option key={o.value} value={o.value}>
          {o.label}
        </option>
      ))}
    </select>
  )
}

function BotoesEdit({ pending, onCancel }: { pending: boolean; onCancel: () => void }) {
  return (
    <div className="flex gap-2">
      <button
        type="submit"
        disabled={pending}
        className="px-3 py-1.5 rounded-md text-xs font-semibold bg-navy hover:bg-teal text-white transition-all disabled:opacity-50"
      >
        {pending ? 'Salvando…' : 'Salvar'}
      </button>
      <button
        type="button"
        onClick={onCancel}
        disabled={pending}
        className="px-3 py-1.5 rounded-md text-xs font-semibold text-navy border border-[#E2E8F0] bg-white hover:border-teal hover:text-teal transition-all"
      >
        Cancelar
      </button>
    </div>
  )
}
