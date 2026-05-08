'use client'

import { Handle, Position, type NodeProps } from '@xyflow/react'
import { MAPA_LEVELS, type NodeType } from '@/lib/definitions'

type Data = {
  tipo: NodeType
  label: string
  abreviacao?: string | null
  sequencia?: number | null
  tempoMedioCiclo?: number | null
  custoEstimado?: number | null
  volumeMensal?: number | null
  onAddChild?: () => void
  onDelete?: () => void
}

const styleByType: Record<NodeType, { bg: string; border: string; text: string; tag: string; kpiBg: string; kpiBorder: string }> = {
  cadeia: {
    bg: 'bg-navy', border: 'border-navy-dark',
    text: 'text-white', tag: 'bg-gold/20 text-gold',
    kpiBg: 'bg-white/10', kpiBorder: 'border-white/15',
  },
  macroprocesso: {
    bg: 'bg-teal', border: 'border-[#0f5876]',
    text: 'text-white', tag: 'bg-white/15 text-white',
    kpiBg: 'bg-white/10', kpiBorder: 'border-white/15',
  },
  processo: {
    bg: 'bg-[#2A8EAE]', border: 'border-teal',
    text: 'text-white', tag: 'bg-white/15 text-white',
    kpiBg: 'bg-white/10', kpiBorder: 'border-white/15',
  },
  macroatividade: {
    bg: 'bg-gold', border: 'border-[#c48500]',
    text: 'text-navy-dark', tag: 'bg-navy-dark/15 text-navy-dark',
    kpiBg: 'bg-navy-dark/10', kpiBorder: 'border-navy-dark/15',
  },
  atividade: {
    bg: 'bg-cream', border: 'border-gold',
    text: 'text-navy', tag: 'bg-navy/8 text-navy',
    kpiBg: 'bg-navy/5', kpiBorder: 'border-navy/10',
  },
}

function formatTempo(v: number) {
  return v % 1 === 0 ? String(v) : v.toFixed(1)
}

function formatCusto(v: number) {
  if (v >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`
  if (v >= 1_000) return `${(v / 1_000).toFixed(1)}k`
  return v.toLocaleString('pt-BR')
}

function formatVolume(v: number) {
  if (v >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`
  if (v >= 1_000) return `${(v / 1_000).toFixed(1)}k`
  return String(v)
}

export default function CustomNode({ data }: NodeProps) {
  const d = data as Data
  const s = styleByType[d.tipo]
  const meta = MAPA_LEVELS[d.tipo]
  const isLeaf = d.tipo === 'atividade'
  const childLabel = !isLeaf
    ? MAPA_LEVELS[meta.parent ? Object.entries(MAPA_LEVELS).find(([, v]) => v.parent === d.tipo)?.[0] as NodeType ?? 'atividade' : 'atividade'].label
    : null

  const temKpis =
    d.tipo === 'processo' &&
    (d.tempoMedioCiclo != null || d.custoEstimado != null || d.volumeMensal != null)

  return (
    <div
      className={`group relative w-56 rounded-lg border-2 ${s.bg} ${s.border} ${s.text} shadow-md transition-all hover:shadow-xl hover:-translate-y-0.5 cursor-pointer`}
    >
      {d.tipo !== 'cadeia' && <Handle type="target" position={Position.Top} className="!bg-navy !border-white !w-2.5 !h-2.5" />}
      <div className="p-3">
        <div className="flex items-center justify-between gap-2 mb-1.5">
          <span className={`text-[9px] font-bold tracking-[0.18em] uppercase px-1.5 py-0.5 rounded ${s.tag}`}>
            {meta.label}
          </span>
          {d.abreviacao && (
            <span className={`text-[10px] font-mono font-bold ${s.text} opacity-80`}>{d.abreviacao}</span>
          )}
        </div>
        <p className={`text-sm font-semibold leading-snug ${s.text}`}>
          {d.sequencia != null && (
            <span className="opacity-60 font-mono mr-1">{String(d.sequencia).padStart(2, '0')}</span>
          )}
          {d.label}
        </p>
      </div>

      {temKpis && (
        <div className={`grid grid-cols-3 gap-px border-t ${s.kpiBorder} ${s.kpiBg}`}>
          <KpiCell
            label="Ciclo"
            value={d.tempoMedioCiclo != null ? `${formatTempo(d.tempoMedioCiclo)}d` : '—'}
            textClass={s.text}
          />
          <KpiCell
            label="Custo"
            value={d.custoEstimado != null ? `R$${formatCusto(d.custoEstimado)}` : '—'}
            textClass={s.text}
          />
          <KpiCell
            label="Volume"
            value={d.volumeMensal != null ? `${formatVolume(d.volumeMensal)}/mês` : '—'}
            textClass={s.text}
          />
        </div>
      )}

      {!isLeaf && childLabel && d.onAddChild && (
        <button
          onClick={(e) => { e.stopPropagation(); d.onAddChild?.() }}
          className="opacity-0 group-hover:opacity-100 absolute -bottom-3 left-1/2 -translate-x-1/2 bg-gold hover:bg-gold-light text-navy-dark text-[10px] font-bold px-2.5 py-1 rounded-full shadow-md transition-all"
          title={`Adicionar ${childLabel.toLowerCase()}`}
        >
          + {childLabel}
        </button>
      )}
      {d.onDelete && (
        <button
          onClick={(e) => { e.stopPropagation(); d.onDelete?.() }}
          className="opacity-0 group-hover:opacity-100 absolute -top-2 -right-2 w-5 h-5 rounded-full bg-[#E05040] hover:bg-[#9A2E1F] text-white text-[10px] font-bold shadow-md transition-all"
          title="Excluir"
        >
          ×
        </button>
      )}

      {!isLeaf && <Handle type="source" position={Position.Bottom} className="!bg-navy !border-white !w-2.5 !h-2.5" />}
    </div>
  )
}

function KpiCell({ label, value, textClass }: { label: string; value: string; textClass: string }) {
  return (
    <div className="px-1.5 py-1 text-center">
      <p className={`text-[8px] font-bold tracking-[0.1em] uppercase opacity-60 ${textClass}`}>{label}</p>
      <p className={`text-[10px] font-mono font-semibold ${textClass}`}>{value}</p>
    </div>
  )
}
