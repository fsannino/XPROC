'use client'

import { useActionState } from 'react'
import { useRouter } from 'next/navigation'
import { criarMegaProcesso } from '@/actions/processos'
import { useEffect } from 'react'

export default function NovoMegaProcessoPage() {
  const router = useRouter()
  const [state, action, pending] = useActionState(criarMegaProcesso, undefined)

  useEffect(() => {
    if (state && 'success' in state) router.push('/dashboard/processos')
  }, [state, router])

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Novo Mega-Processo</h1>

      <form action={action} className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Descrição <span className="text-red-500">*</span>
          </label>
          <input
            name="descricao"
            required
            maxLength={80}
            className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          {state?.errors?.descricao && (
            <p className="text-xs text-red-600 mt-1">{state.errors.descricao[0]}</p>
          )}
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Abreviação <span className="text-gray-400">(até 4 chars)</span>
          </label>
          <input
            name="abreviacao"
            maxLength={4}
            className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Descrição Longa
          </label>
          <textarea
            name="descricaoLonga"
            rows={4}
            className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
          />
        </div>

        <div className="flex items-center gap-2">
          <input type="checkbox" name="bloqueado" value="true" id="bloqueado" />
          <label htmlFor="bloqueado" className="text-sm text-gray-700">Bloqueado</label>
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
