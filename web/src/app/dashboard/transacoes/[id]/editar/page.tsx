'use client'

import { useActionState, useEffect } from 'react'
import { useRouter, useParams } from 'next/navigation'
import { atualizarTransacao } from '@/actions/transacoes'
import { useToast } from '@/components/ui/toast'

export default function EditarTransacaoPage() {
  const { id } = useParams<{ id: string }>()
  const router = useRouter()
  const { show } = useToast()
  const action = atualizarTransacao.bind(null, id)
  const [state, formAction, pending] = useActionState(action, undefined)

  useEffect(() => {
    if (state && 'success' in state) {
      show('Transação atualizada!')
      router.push('/dashboard/transacoes')
    }
  }, [state, router, show])

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Editar Transação</h1>

      <form action={formAction} className="bg-white rounded-xl border border-gray-100 p-6 space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Código</label>
          <input
            value={id}
            readOnly
            className="w-full rounded-lg border border-gray-200 bg-gray-50 px-3 py-2 text-sm font-mono text-gray-500"
          />
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
