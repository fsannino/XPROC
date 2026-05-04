'use client'

import { useActionState, useEffect } from 'react'
import { useRouter, useParams } from 'next/navigation'
import { criarProcesso, criarSubProcesso } from '@/actions/processos'

export default function AdicionarProcessoPage() {
  const router = useRouter()
  const params = useParams()
  const megaProcessoId = Number(params.id)

  const [stateProcesso, actionProcesso, pendingProcesso] = useActionState(criarProcesso, undefined)
  const [stateSub, actionSub, pendingSub] = useActionState(criarSubProcesso, undefined)

  useEffect(() => {
    if (stateProcesso && 'success' in stateProcesso) router.push(`/dashboard/processos/${megaProcessoId}`)
  }, [stateProcesso, router, megaProcessoId])

  useEffect(() => {
    if (stateSub && 'success' in stateSub) router.push(`/dashboard/processos/${megaProcessoId}`)
  }, [stateSub, router, megaProcessoId])

  return (
    <div className="max-w-xl space-y-8">
      <h1 className="text-2xl font-bold text-gray-900">Adicionar Processo / Sub-Processo</h1>

      {/* Novo Processo */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h2 className="text-lg font-semibold text-gray-800 mb-4">Novo Processo</h2>
        <form action={actionProcesso} className="space-y-4">
          <input type="hidden" name="megaProcessoId" value={megaProcessoId} />
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Descrição *</label>
            <input
              name="descricao"
              required
              maxLength={150}
              className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Sequência</label>
            <input
              name="sequencia"
              type="number"
              min={1}
              className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <button
            type="submit"
            disabled={pendingProcesso}
            className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800 disabled:opacity-60"
          >
            {pendingProcesso ? 'Salvando...' : 'Criar Processo'}
          </button>
        </form>
      </div>

      <div className="flex justify-start">
        <button
          onClick={() => router.back()}
          className="border border-gray-300 text-gray-700 px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-50"
        >
          Voltar
        </button>
      </div>
    </div>
  )
}
