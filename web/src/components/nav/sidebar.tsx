'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { logout } from '@/actions/auth'
import { BrandMark } from '@/components/brand/logo'

const navItems = [
  { href: '/dashboard', label: 'Início', icon: '◆' },
  { href: '/dashboard/mapa', label: 'Mapa', icon: '◇' },
  { href: '/dashboard/processos', label: 'Processos', icon: '◈' },
  { href: '/dashboard/transacoes', label: 'Transações', icon: '◉' },
  { href: '/dashboard/cenarios', label: 'Cenários', icon: '◎' },
  { href: '/dashboard/empresas', label: 'Empresas', icon: '◐' },
  { href: '/dashboard/modulos', label: 'Módulos', icon: '◑' },
  { href: '/dashboard/usuarios', label: 'Usuários', icon: '◒' },
]

export default function Sidebar() {
  const pathname = usePathname()

  return (
    <aside className="w-64 bg-navy-dark text-white/70 flex flex-col min-h-screen relative">
      {/* Faixa gold lateral (assinatura clbz) */}
      <span className="absolute right-0 top-0 bottom-0 w-px bg-gradient-to-b from-gold/0 via-gold/40 to-gold/0" />

      <div className="px-6 py-6 border-b border-white/10">
        <Link href="/dashboard" className="flex items-center gap-2.5 group">
          <BrandMark size={28} />
          <div className="leading-none">
            <span className="font-sans font-extrabold text-lg tracking-tight text-white">
              X<span className="text-gold">·</span>PROC
            </span>
            <span className="block text-[8px] font-medium tracking-[0.2em] uppercase text-white/50 mt-1">
              Processos &amp; Governança
            </span>
          </div>
        </Link>
      </div>

      <nav className="flex-1 px-3 py-5 space-y-0.5">
        {navItems.map((item) => {
          const active =
            pathname === item.href ||
            (item.href !== '/dashboard' && pathname.startsWith(item.href))
          return (
            <Link
              key={item.href}
              href={item.href}
              className={`relative flex items-center gap-3 px-3 py-2.5 rounded-md text-sm transition-all ${
                active
                  ? 'bg-white/8 text-white font-semibold'
                  : 'text-white/60 hover:bg-white/5 hover:text-white'
              }`}
            >
              {active && (
                <span className="absolute left-0 top-1/2 -translate-y-1/2 w-0.5 h-5 bg-gold rounded-r-full" />
              )}
              <span
                className={`text-base ${
                  active ? 'text-gold' : 'text-white/40'
                }`}
              >
                {item.icon}
              </span>
              {item.label}
            </Link>
          )
        })}
      </nav>

      <div className="px-3 py-4 border-t border-white/10">
        <form action={logout}>
          <button
            type="submit"
            className="w-full flex items-center gap-3 px-3 py-2.5 rounded-md text-sm font-medium text-white/60 hover:bg-white/5 hover:text-gold transition-all"
          >
            <span className="text-base text-white/40 group-hover:text-gold">⏻</span>
            Sair
          </button>
        </form>
        <p className="text-[10px] text-white/30 text-center mt-3 tracking-wider uppercase">
          © {new Date().getFullYear()} XPROC
        </p>
      </div>
    </aside>
  )
}
