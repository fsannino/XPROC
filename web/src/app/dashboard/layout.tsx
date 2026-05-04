import Sidebar from '@/components/nav/sidebar'
import { ToastProvider } from '@/components/ui/toast'
import { getSession } from '@/lib/session'
import { redirect } from 'next/navigation'

export default async function DashboardLayout({ children }: { children: React.ReactNode }) {
  const session = await getSession()
  if (!session) redirect('/login')

  return (
    <ToastProvider>
      <div className="flex min-h-screen">
        <Sidebar />
        <div className="flex-1 flex flex-col overflow-hidden">
          <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
            <div />
            <div className="text-sm text-gray-600">
              Olá, <span className="font-semibold text-gray-900">{session.nome}</span>
            </div>
          </header>
          <main className="flex-1 overflow-auto p-6">{children}</main>
        </div>
      </div>
    </ToastProvider>
  )
}
