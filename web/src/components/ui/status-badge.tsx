import type { LifecycleStatus } from '@/lib/definitions'

const CONFIG: Record<string, { label: string; classes: string }> = {
  Rascunho:  { label: 'Rascunho',    classes: 'bg-gray-100 text-gray-600' },
  EmRevisao: { label: 'Em Revisão',  classes: 'bg-yellow-100 text-yellow-700' },
  Aprovado:  { label: 'Aprovado',    classes: 'bg-blue-100 text-blue-700' },
  Publicado: { label: 'Publicado',   classes: 'bg-green-100 text-green-700' },
  Arquivado: { label: 'Arquivado',   classes: 'bg-red-100 text-red-600' },
}

export function StatusBadge({ status }: { status: string }) {
  const cfg = CONFIG[status] ?? { label: status, classes: 'bg-gray-100 text-gray-600' }
  return (
    <span className={`inline-flex px-2 py-0.5 rounded-full text-xs font-medium ${cfg.classes}`}>
      {cfg.label}
    </span>
  )
}

export function statusLabel(status: string): string {
  return CONFIG[status as LifecycleStatus]?.label ?? status
}
