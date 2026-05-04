'use client'

import { useRouter, useSearchParams, usePathname } from 'next/navigation'
import { useTransition } from 'react'

export function SearchInput({ placeholder = 'Buscar...' }: { placeholder?: string }) {
  const router = useRouter()
  const pathname = usePathname()
  const searchParams = useSearchParams()
  const [, startTransition] = useTransition()

  function handleChange(e: React.ChangeEvent<HTMLInputElement>) {
    const params = new URLSearchParams(searchParams.toString())
    const value = e.target.value.trim()
    if (value) params.set('busca', value)
    else params.delete('busca')
    startTransition(() => router.replace(`${pathname}?${params.toString()}`))
  }

  return (
    <input
      type="search"
      defaultValue={searchParams.get('busca') ?? ''}
      onChange={handleChange}
      placeholder={placeholder}
      className="rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-56"
    />
  )
}
