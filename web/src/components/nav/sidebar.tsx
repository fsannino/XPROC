'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { logout } from '@/actions/auth'

const navItems = [
  { href: '/dashboard', label: 'Início', icon: '🏠' },
  { href: '/dashboard/mapa', label: 'Mapa', icon: '🗺️' },
  { href: '/dashboard/processos', label: 'Processos', icon: '🔄' },
  { href: '/dashboard/transacoes', label: 'Transações', icon: '💱' },
  { href: '/dashboard/cenarios', label: 'Cenários', icon: '📋' },
  { href: '/dashboard/empresas', label: 'Empresas', icon: '🏢' },
  { href: '/dashboard/modulos', label: 'Módulos', icon: '📦' },
  { href: '/dashboard/usuarios', label: 'Usuários', icon: '👥' },
]

export default function Sidebar() {
  const pathname = usePathname()

  return (
    <aside className="w-64 bg-blue-900 text-white flex flex-col min-h-screen">
      <div className="px-6 py-5 border-b border-blue-800">
        <h1 className="text-xl font-bold">XPROC</h1>
        <p className="text-blue-300 text-xs mt-0.5">Gestão de Processos</p>
      </div>

      <nav className="flex-1 px-3 py-4 space-y-1">
        {navItems.map((item) => {
          const active = pathname === item.href || (item.href !== '/dashboard' && pathname.startsWith(item.href))
          return (
            <Link
              key={item.href}
              href={item.href}
              className={`flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                active
                  ? 'bg-blue-700 text-white'
                  : 'text-blue-200 hover:bg-blue-800 hover:text-white'
              }`}
            >
              <span>{item.icon}</span>
              {item.label}
            </Link>
          )
        })}
      </nav>

      <div className="px-3 py-4 border-t border-blue-800">
        <form action={logout}>
          <button
            type="submit"
            className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-blue-200 hover:bg-blue-800 hover:text-white transition-colors"
          >
            <span>🚪</span>
            Sair
          </button>
        </form>
      </div>
    </aside>
  )
}
