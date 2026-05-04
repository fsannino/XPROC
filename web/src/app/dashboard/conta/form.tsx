'use client'

import { useActionState, useEffect } from 'react'
import { trocarSenha } from '@/actions/conta'
import { useToast } from '@/components/ui/toast'

export default function TrocaSenhaForm() {
  const { show } = useToast()
  const [state, formAction, pending] = useActionState(trocarSenha, undefined)

  useEffect(() => {
    if (state && 'success' in state) show('Senha alterada com sucesso!')
    if (state && 'error' in state) show(state.error!, 'error')
  }, [state, show])

  return (
    <form action={formAction} className="space-y-4">
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Senha atual <span className="text-red-500">*</span>
        </label>
        <input
          name="senhaAtual"
          type="password"
          required
          autoComplete="current-password"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Nova senha <span className="text-red-500">*</span>
        </label>
        <input
          name="novaSenha"
          type="password"
          required
          minLength={6}
          autoComplete="new-password"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Confirmar nova senha <span className="text-red-500">*</span>
        </label>
        <input
          name="confirmar"
          type="password"
          required
          autoComplete="new-password"
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>
      <button
        type="submit"
        disabled={pending}
        className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60"
      >
        {pending ? 'Salvando...' : 'Alterar Senha'}
      </button>
    </form>
  )
}
