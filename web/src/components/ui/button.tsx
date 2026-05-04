import { ButtonHTMLAttributes } from 'react'

type Variant = 'primary' | 'secondary' | 'danger' | 'ghost'

const variants: Record<Variant, string> = {
  primary: 'bg-blue-700 text-white hover:bg-blue-800 disabled:opacity-60',
  secondary: 'border border-gray-300 text-gray-700 hover:bg-gray-50',
  danger: 'text-red-600 hover:text-red-800 font-medium',
  ghost: 'text-blue-600 hover:text-blue-800 font-medium',
}

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant
  loading?: boolean
}

export function Button({ variant = 'primary', loading, children, className = '', ...props }: ButtonProps) {
  return (
    <button
      {...props}
      disabled={props.disabled || loading}
      className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors ${variants[variant]} ${className}`}
    >
      {loading ? 'Aguarde...' : children}
    </button>
  )
}
