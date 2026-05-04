import { HTMLAttributes } from 'react'

interface CardProps extends HTMLAttributes<HTMLDivElement> {
  padding?: boolean
}

export function Card({ padding = true, className = '', children, ...props }: CardProps) {
  return (
    <div
      {...props}
      className={`bg-white rounded-xl shadow-sm border border-gray-100 ${padding ? 'p-6' : 'overflow-hidden'} ${className}`}
    >
      {children}
    </div>
  )
}
