import { ButtonHTMLAttributes } from 'react'

type Variant = 'primary' | 'secondary' | 'danger' | 'ghost' | 'gold'

const base =
  'inline-flex items-center justify-center gap-2 px-4 py-2 rounded-md text-sm font-semibold transition-all disabled:opacity-60 disabled:cursor-not-allowed focus:outline-none focus:ring-2 focus:ring-offset-1 focus:ring-teal'

const variants: Record<Variant, string> = {
  primary:
    'bg-navy hover:bg-teal text-white hover:-translate-y-0.5 hover:shadow-md disabled:hover:translate-y-0',
  secondary:
    'border border-[#E2E8F0] bg-white text-navy hover:border-teal hover:text-teal',
  gold:
    'bg-gold hover:bg-gold-light text-navy-dark hover:-translate-y-0.5 hover:shadow-md disabled:hover:translate-y-0',
  danger: 'text-[#E05040] hover:text-[#9A2E1F] font-semibold',
  ghost: 'text-teal hover:text-navy font-semibold',
}

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant
  loading?: boolean
}

export function Button({
  variant = 'primary',
  loading,
  children,
  className = '',
  ...props
}: ButtonProps) {
  return (
    <button
      {...props}
      disabled={props.disabled || loading}
      className={`${base} ${variants[variant]} ${className}`}
    >
      {loading ? 'Aguarde...' : children}
    </button>
  )
}
