import { type SVGProps } from 'react'

/**
 * Ícone-logo no padrão clbz-website: 4 quadrados com cantos suaves.
 * Disposição: navy / teal / gold / gray.
 */
export function BrandMark({ size = 32, ...props }: { size?: number } & SVGProps<SVGSVGElement>) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 32 32"
      aria-hidden="true"
      {...props}
    >
      <rect x="1" y="1" width="13" height="13" rx="1.5" fill="#0B3D5C" />
      <rect x="17" y="1" width="13" height="13" rx="1.5" fill="#1A6E8E" />
      <rect x="1" y="17" width="13" height="13" rx="1.5" fill="#F7A823" />
      <rect x="17" y="17" width="13" height="13" rx="1.5" fill="#94A3B8" />
    </svg>
  )
}

export function BrandWordmark({ light = false }: { light?: boolean }) {
  return (
    <div className="flex items-center gap-2">
      <BrandMark size={28} />
      <div className="leading-none">
        <span
          className={`font-sans font-extrabold text-lg tracking-tight ${
            light ? 'text-white' : 'text-navy'
          }`}
        >
          X<span className="text-gold">·</span>PROC
        </span>
        <span
          className={`block text-[8px] font-medium tracking-[0.2em] uppercase mt-0.5 ${
            light ? 'text-white/60' : 'text-gray-medium'
          }`}
        >
          processos &amp; governança
        </span>
      </div>
    </div>
  )
}
