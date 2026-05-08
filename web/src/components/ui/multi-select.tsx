'use client'

import { useMemo, useState } from 'react'

export type MultiSelectOption = {
  value: string
  label: string
  hint?: string // texto secundário em fonte menor
}

type Props = {
  options: MultiSelectOption[]
  selected: string[]
  onChange: (next: string[]) => void
  placeholder?: string
  emptyMessage?: string
  maxHeight?: number
}

export default function MultiSelect({
  options,
  selected,
  onChange,
  placeholder = 'Buscar...',
  emptyMessage = 'Nenhuma opção.',
  maxHeight = 220,
}: Props) {
  const [filter, setFilter] = useState('')

  const filtered = useMemo(() => {
    const q = filter.trim().toLowerCase()
    if (!q) return options
    return options.filter(
      (o) =>
        o.label.toLowerCase().includes(q) ||
        o.value.toLowerCase().includes(q) ||
        (o.hint?.toLowerCase().includes(q) ?? false),
    )
  }, [options, filter])

  const selectedSet = useMemo(() => new Set(selected), [selected])

  function toggle(value: string) {
    if (selectedSet.has(value)) {
      onChange(selected.filter((v) => v !== value))
    } else {
      onChange([...selected, value])
    }
  }

  function clearAll() {
    onChange([])
  }

  return (
    <div className="rounded-md border border-[#E2E8F0] bg-white overflow-hidden">
      <div className="flex items-center justify-between gap-2 px-3 py-2 border-b border-[#E2E8F0] bg-[#F5F6F8]">
        <input
          type="text"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          placeholder={placeholder}
          className="flex-1 bg-transparent text-sm text-slate placeholder-gray-medium focus:outline-none"
        />
        <span className="text-[10px] font-mono text-gray-medium">
          {selected.length}/{options.length}
        </span>
        {selected.length > 0 && (
          <button
            type="button"
            onClick={clearAll}
            className="text-[10px] font-semibold text-[#9A2E1F] hover:text-[#E05040] uppercase tracking-wider"
          >
            Limpar
          </button>
        )}
      </div>

      <div className="overflow-auto" style={{ maxHeight }}>
        {filtered.length === 0 && (
          <p className="text-center text-xs text-gray-medium py-6">{emptyMessage}</p>
        )}
        {filtered.map((opt) => {
          const checked = selectedSet.has(opt.value)
          return (
            <label
              key={opt.value}
              className={`flex items-start gap-2.5 px-3 py-2 cursor-pointer text-sm border-b border-[#F5F6F8] last:border-b-0 transition-colors ${
                checked ? 'bg-teal/5 hover:bg-teal/10' : 'hover:bg-[#F5F6F8]'
              }`}
            >
              <input
                type="checkbox"
                checked={checked}
                onChange={() => toggle(opt.value)}
                className="mt-0.5 accent-teal"
              />
              <div className="flex-1 min-w-0">
                <p className="font-medium text-slate truncate">{opt.label}</p>
                {opt.hint && (
                  <p className="text-[10px] text-gray-medium font-mono truncate">{opt.hint}</p>
                )}
              </div>
            </label>
          )
        })}
      </div>
    </div>
  )
}
