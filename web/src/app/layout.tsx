import type { Metadata } from 'next'
import { DM_Serif_Display, Plus_Jakarta_Sans, JetBrains_Mono } from 'next/font/google'
import './globals.css'

const jakarta = Plus_Jakarta_Sans({
  variable: '--font-jakarta',
  subsets: ['latin'],
  weight: ['300', '400', '500', '600', '700', '800'],
  display: 'swap',
})

const serif = DM_Serif_Display({
  variable: '--font-serif',
  subsets: ['latin'],
  weight: ['400'],
  display: 'swap',
})

const mono = JetBrains_Mono({
  variable: '--font-mono-jetbrains',
  subsets: ['latin'],
  weight: ['400', '500'],
  display: 'swap',
})

export const metadata: Metadata = {
  title: { default: 'Collab:Flow — Gestão de Processos Corporativos', template: '%s — Collab:Flow' },
  description: 'Sistema de Gerenciamento de Processos Corporativos.',
  themeColor: '#0B3D5C',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html
      lang="pt-BR"
      className={`${jakarta.variable} ${serif.variable} ${mono.variable} h-full`}
    >
      <body className="h-full bg-background text-slate antialiased font-sans">
        {children}
      </body>
    </html>
  )
}
