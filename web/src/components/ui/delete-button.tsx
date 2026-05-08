'use client'

import { useTransition } from 'react'

interface DeleteButtonProps {
  action: () => Promise<unknown>
  label?: string
  confirmText?: string
}

export function DeleteButton({ action, label = 'Excluir', confirmText = 'Confirmar exclusão?' }: DeleteButtonProps) {
  const [pending, startTransition] = useTransition()

  function handleClick() {
    if (!confirm(confirmText)) return
    startTransition(async () => {
      await action()
    })
  }

  return (
    <button
      type="button"
      onClick={handleClick}
      disabled={pending}
      className="text-red-600 hover:text-red-800 font-medium text-xs disabled:opacity-50"
    >
      {pending ? 'Excluindo...' : label}
    </button>
  )
}
