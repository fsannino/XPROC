import Sidebar from '@/components/nav/sidebar'
import { ToastProvider } from '@/components/ui/toast'
import { getSession } from '@/lib/session'
import { redirect } from 'next/navigation'
import Link from 'next/link'

// Garante que toda página dentro de /dashboard busca dados frescos do banco
// em cada navegação, sem usar Router Cache do Next 16. Sem isto, dados
// criados no Mapa podem não aparecer imediatamente em /processos etc.
export const dynamic = 'force-dynamic'

export default async function DashboardLayout({ children }: { children: React.ReactNode }) {
  const session = await getSession()
  if (!session) redirect('/login')

  return (
    <ToastProvider>
      <div className="flex min-h-screen bg-[#F5F6F8]">
        <Sidebar />
        <div className="flex-1 flex flex-col overflow-hidden">
          <header className="sticky top-0 z-30 bg-white/90 backdrop-blur-md border-b border-[#E2E8F0] px-8 py-3 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <span className="text-[10px] font-bold tracking-[0.2em] uppercase text-teal">
                Painel
              </span>
              <span className="w-1 h-1 rounded-full bg-gold" />
              <span className="text-xs text-gray-medium">
                Gestão de Processos Corporativos
              </span>
            </div>
            <Link
              href="/dashboard/conta"
              className="group flex items-center gap-3 text-sm text-gray-medium hover:text-navy transition-colors"
            >
              <span>
                Olá,{' '}
                <span className="font-semibold text-navy group-hover:text-teal transition-colors">
                  {session.nome}
                </span>
              </span>
              <span className="w-8 h-8 rounded-full bg-navy/10 group-hover:bg-teal/10 flex items-center justify-center text-navy group-hover:text-teal text-xs font-bold transition-colors">
                {(session.nome ?? '?').slice(0, 1).toUpperCase()}
              </span>
            </Link>
          </header>
          <main className="flex-1 overflow-auto p-8">{children}</main>
        </div>
      </div>
    </ToastProvider>
  )
}
