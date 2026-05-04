'use client'

import { useActionState, useEffect } from 'react'
import { useToast } from '@/components/ui/toast'
import type { Modulo } from '@prisma/client'

type ActionFn = (_state: unknown, formData: FormData) => Promise<{ success?: boolean; errors?: Record<string, string[]> }>

export default function ModulosForm({
  modulos,
  criarModulo,
  criarSubModulo,
}: {
  modulos: Modulo[]
  criarModulo: ActionFn
  criarSubModulo: ActionFn
}) {
  const { show } = useToast()
  const [mState, mAction, mPending] = useActionState(criarModulo, undefined)
  const [sState, sAction, sPending] = useActionState(criarSubModulo, undefined)

  useEffect(() => {
    if (mState && 'success' in mState) show('Módulo criado!')
    if (mState && 'errors' in mState) show('Erro ao criar módulo.', 'error')
  }, [mState, show])

  useEffect(() => {
    if (sState && 'success' in sState) show('Sub-módulo criado!')
    if (sState && 'errors' in sState) show('Erro ao criar sub-módulo.', 'error')
  }, [sState, show])

  return (
    <div className="space-y-4">
      <form action={mAction} className="bg-white rounded-xl border border-gray-100 p-5 space-y-3">
        <h2 className="font-semibold text-gray-800">Novo Módulo</h2>
        <input
          name="descricao"
          required
          maxLength={80}
          placeholder="Descrição"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        <button type="submit" disabled={mPending}
          className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60 w-full">
          {mPending ? 'Criando...' : 'Criar Módulo'}
        </button>
      </form>

      <form action={sAction} className="bg-white rounded-xl border border-gray-100 p-5 space-y-3">
        <h2 className="font-semibold text-gray-800">Novo Sub-Módulo</h2>
        <select
          name="moduloId"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        >
          <option value="">— Sem módulo —</option>
          {modulos.map((m) => (
            <option key={m.id} value={m.id}>{m.descricao}</option>
          ))}
        </select>
        <input
          name="descricao"
          required
          maxLength={80}
          placeholder="Descrição"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        <input
          name="abreviacao"
          maxLength={10}
          placeholder="Abreviação (opcional)"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        <button type="submit" disabled={sPending}
          className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60 w-full">
          {sPending ? 'Criando...' : 'Criar Sub-Módulo'}
        </button>
      </form>
    </div>
  )
}
