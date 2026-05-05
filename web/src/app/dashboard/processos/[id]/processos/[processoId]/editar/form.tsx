'use client'

import { useActionState, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import { atualizarProcesso } from '@/actions/processos'
import { useToast } from '@/components/ui/toast'
import type { Processo } from '@prisma/client'

export default function EditarProcessoForm({
  processo,
  megaProcessoId,
}: {
  processo: Processo
  megaProcessoId: number
}) {
  const router = useRouter()
  const { show } = useToast()
  const action = atualizarProcesso.bind(null, processo.id, megaProcessoId)
  const [state, formAction, pending] = useActionState(action, undefined)

  useEffect(() => {
    if (state && 'success' in state) {
      show('Processo atualizado!')
      router.push(`/dashboard/processos/${megaProcessoId}`)
    }
    if (state && 'errors' in state) show('Corrija os erros no formulário.', 'error')
  }, [state, router, megaProcessoId, show])

  return (
    <form action={formAction} className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 space-y-4">
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">
          Descrição <span className="text-red-500">*</span>
        </label>
        <input
          name="descricao"
          defaultValue={processo.descricao}
          required
          maxLength={150}
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        {state?.errors?.descricao && (
          <p className="text-xs text-red-600 mt-1">{state.errors.descricao[0]}</p>
        )}
      </div>

      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">Sequência</label>
        <input
          name="sequencia"
          type="number"
          min={1}
          defaultValue={processo.sequencia ?? ''}
          className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>

      <fieldset className="border border-gray-200 rounded-lg p-4 space-y-3">
        <legend className="text-sm font-medium text-gray-600 px-1">KPIs</legend>
        <div className="grid grid-cols-3 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Tempo médio (dias)</label>
            <input
              name="tempoMedioCiclo"
              type="number"
              min={0}
              step="0.1"
              defaultValue={processo.tempoMedioCiclo ?? ''}
              placeholder="ex: 2.5"
              className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Custo estimado (R$)</label>
            <input
              name="custoEstimado"
              type="number"
              min={0}
              step="0.01"
              defaultValue={processo.custoEstimado ?? ''}
              placeholder="ex: 1500.00"
              className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Volume mensal</label>
            <input
              name="volumeMensal"
              type="number"
              min={1}
              defaultValue={processo.volumeMensal ?? ''}
              placeholder="ex: 200"
              className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
        </div>
      </fieldset>

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
  )
}
