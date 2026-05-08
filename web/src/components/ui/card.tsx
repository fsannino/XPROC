import { HTMLAttributes } from 'react'

type Accent = 'none' | 'navy' | 'teal' | 'gold'

interface CardProps extends HTMLAttributes<HTMLDivElement> {
  padding?: boolean
  /** Borda superior colorida no padrão clbz (pillar-top). */
  accent?: Accent
  /** Eleva no hover. */
  interactive?: boolean
}

const accentClasses: Record<Accent, string> = {
  none: '',
  navy: 'border-t-4 border-t-navy',
  teal: 'border-t-4 border-t-teal',
  gold: 'border-t-4 border-t-gold',
}

export function Card({
  padding = true,
  accent = 'none',
  interactive = false,
  className = '',
  children,
  ...props
}: CardProps) {
  return (
    <div
      {...props}
      className={`bg-white rounded-lg border border-[#E2E8F0] ${accentClasses[accent]} ${
        padding ? 'p-6' : 'overflow-hidden'
      } ${
        interactive ? 'transition-all hover:shadow-lg hover:-translate-y-0.5 cursor-pointer' : ''
      } ${className}`}
    >
      {children}
    </div>
  )
}
