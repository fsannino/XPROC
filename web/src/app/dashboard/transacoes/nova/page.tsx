'use client'

import { useActionState, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import { criarTransacao } from '@/actions/transacoes'

export default function NovaTransacaoPage() {
  const router = useRouter()
  const [state, action, pending] = useActionState(criarTransacao, undefined)

  useEffect(() => {
    if (state && 'success' in state) router.push('/dashboard/transacoes')
  }, [state, router])

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Nova Transação</h1>

      <form action={action} className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Código <span className="text-red-500">*</span>
          </label>
          <input
            name="id"
            required
            maxLength={30}
            className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="Ex: MM01"
          />
          {state?.errors?.id && (
            <p className="text-xs text-red-600 mt-1">{state.errors.id[0]}</p>
          )}
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Descrição</label>
          <input
            name="descricao"
            maxLength={150}
            className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>

        <div className="flex gap-3 pt-2">
          <button
            type="submit"
            disabled={pending}
            className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60"
          >
            {pending ? 'Salvando...' : 'Salvar'}
          </button>
          <button
            type="button"
            onClick={() => router.back()}
            className="border border-gray-300 text-gray-700 px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-50"
          >
            Cancelar
          </button>
        </div>
      </form>
    </div>
  )
}
