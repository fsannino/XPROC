'use client'

import { useActionState, useEffect } from 'react'
import { criarEmpresa } from '@/actions/transacoes'
import { useToast } from '@/components/ui/toast'

export default function NovaEmpresaForm() {
  const { show } = useToast()
  const [state, formAction, pending] = useActionState(criarEmpresa, undefined)

  useEffect(() => {
    if (state && 'success' in state) show('Empresa criada com sucesso!')
    if (state && 'errors' in state) show('Erro ao criar empresa.', 'error')
  }, [state, show])

  return (
    <form action={formAction} className="bg-white rounded-xl border border-gray-100 p-6 space-y-4">
      <h2 className="text-lg font-semibold text-gray-800">Nova Empresa</h2>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Nome <span className="text-red-500">*</span>
        </label>
        <input
          name="nome"
          required
          maxLength={150}
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        {state?.errors?.nome && (
          <p className="text-xs text-red-600 mt-1">{state.errors.nome[0]}</p>
        )}
      </div>

      <button
        type="submit"
        disabled={pending}
        className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60"
      >
        {pending ? 'Criando...' : 'Criar Empresa'}
      </button>
    </form>
  )
}
