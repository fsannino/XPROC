import type { LifecycleStatus } from '@/lib/definitions'

const CONFIG: Record<string, { label: string; classes: string }> = {
  Rascunho: {
    label: 'Rascunho',
    classes: 'bg-[#F5F6F8] text-gray-medium border border-[#E2E8F0]',
  },
  EmRevisao: {
    label: 'Em Revisão',
    classes: 'bg-[rgba(247,168,35,0.10)] text-[#c48500]',
  },
  Aprovado: {
    label: 'Aprovado',
    classes: 'bg-[rgba(26,110,142,0.08)] text-teal',
  },
  Publicado: {
    label: 'Publicado',
    classes: 'bg-[rgba(11,61,92,0.08)] text-navy',
  },
  Arquivado: {
    label: 'Arquivado',
    classes: 'bg-[rgba(224,80,64,0.08)] text-[#9A2E1F]',
  },
}

export function StatusBadge({ status }: { status: string }) {
  const cfg =
    CONFIG[status] ?? {
      label: status,
      classes: 'bg-[#F5F6F8] text-gray-medium border border-[#E2E8F0]',
    }
  return (
    <span
      className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-[10px] font-bold tracking-wider uppercase ${cfg.classes}`}
    >
      {cfg.label}
    </span>
  )
}

export function statusLabel(status: string): string {
  return CONFIG[status as LifecycleStatus]?.label ?? status
}
