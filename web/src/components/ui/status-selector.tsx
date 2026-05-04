'use client'

import { useTransition } from 'react'
import { alterarStatus } from '@/actions/lifecycle'
import { StatusBadge, statusLabel } from '@/components/ui/status-badge'

interface Props {
  megaProcessoId: number
  statusAtual: string
  proximosStatus: string[]
}

export function StatusSelector({ megaProcessoId, statusAtual, proximosStatus }: Props) {
  const [pending, startTransition] = useTransition()

  if (proximosStatus.length === 0) {
    return <StatusBadge status={statusAtual} />
  }

  function handleChange(novoStatus: string) {
    startTransition(async () => {
      await alterarStatus(megaProcessoId, novoStatus)
    })
  }

  return (
    <div className="flex items-center gap-2">
      <StatusBadge status={statusAtual} />
      <select
        disabled={pending}
        defaultValue=""
        onChange={(e) => { if (e.target.value) handleChange(e.target.value) }}
        className="text-xs border border-gray-300 rounded px-2 py-0.5 bg-white text-gray-700 disabled:opacity-50"
      >
        <option value="">Alterar...</option>
        {proximosStatus.map((s) => (
          <option key={s} value={s}>{statusLabel(s)}</option>
        ))}
      </select>
      {pending && <span className="text-xs text-gray-400">Salvando...</span>}
    </div>
  )
}
