'use client'

import { Handle, Position, type NodeProps } from '@xyflow/react'
import { MAPA_LEVELS, type NodeType } from '@/lib/definitions'

type Data = {
  tipo: NodeType
  label: string
  abreviacao?: string | null
  sequencia?: number | null
  onAddChild?: () => void
  onDelete?: () => void
}

const styleByType: Record<NodeType, { bg: string; border: string; text: string; tag: string }> = {
  cadeia: {
    bg: 'bg-navy', border: 'border-navy-dark',
    text: 'text-white', tag: 'bg-gold/20 text-gold',
  },
  macroprocesso: {
    bg: 'bg-teal', border: 'border-[#0f5876]',
    text: 'text-white', tag: 'bg-white/15 text-white',
  },
  processo: {
    bg: 'bg-[#2A8EAE]', border: 'border-teal',
    text: 'text-white', tag: 'bg-white/15 text-white',
  },
  macroatividade: {
    bg: 'bg-gold', border: 'border-[#c48500]',
    text: 'text-navy-dark', tag: 'bg-navy-dark/15 text-navy-dark',
  },
  atividade: {
    bg: 'bg-cream', border: 'border-gold',
    text: 'text-navy', tag: 'bg-navy/8 text-navy',
  },
}

export default function CustomNode({ data }: NodeProps) {
  const d = data as Data
  const s = styleByType[d.tipo]
  const meta = MAPA_LEVELS[d.tipo]
  const isLeaf = d.tipo === 'atividade'
  const childLabel = !isLeaf
    ? MAPA_LEVELS[meta.parent ? Object.entries(MAPA_LEVELS).find(([, v]) => v.parent === d.tipo)?.[0] as NodeType ?? 'atividade' : 'atividade'].label
    : null

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

      {/* Hover actions */}
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
