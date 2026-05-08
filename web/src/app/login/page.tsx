'use client'

import { useActionState } from 'react'
import { login } from '@/actions/auth'

export default function LoginPage() {
  const [state, action, pending] = useActionState(login, undefined)

  return (
    <main className="min-h-screen flex">
      {/* Left panel — brand */}
      <div className="hidden lg:flex lg:w-1/2 bg-slate-900 flex-col justify-between p-12 relative overflow-hidden">
        {/* Background pattern */}
        <div className="absolute inset-0 opacity-5">
          <svg width="100%" height="100%" xmlns="http://www.w3.org/2000/svg">
            <defs>
              <pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
                <path d="M 40 0 L 0 0 0 40" fill="none" stroke="white" strokeWidth="1" />
              </pattern>
            </defs>
            <rect width="100%" height="100%" fill="url(#grid)" />
          </svg>
        </div>

        {/* Logo */}
        <div className="relative z-10 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-blue-500 flex items-center justify-center">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M10 2L17 6V14L10 18L3 14V6L10 2Z" stroke="white" strokeWidth="1.5" fill="none" />
              <path d="M10 6L14 8.5V13.5L10 16L6 13.5V8.5L10 6Z" fill="white" fillOpacity="0.3" />
            </svg>
          </div>
          <span className="text-white font-semibold text-lg tracking-tight">Collab:Flow</span>
        </div>

        {/* Central content */}
        <div className="relative z-10 space-y-6">
          <div className="space-y-3">
            <div className="inline-flex items-center gap-2 bg-blue-500/20 border border-blue-500/30 rounded-full px-3 py-1">
              <div className="w-1.5 h-1.5 rounded-full bg-blue-400"></div>
              <span className="text-blue-300 text-xs font-medium tracking-wide uppercase">Process &amp; Governance</span>
            </div>
            <h2 className="text-3xl font-bold text-white leading-snug">
              Mapeie o que importa.<br />
              <span className="text-blue-400">Otimize o que entrega.</span>
            </h2>
            <p className="text-slate-400 text-sm leading-relaxed max-w-xs">
              Processos, cenários, catálogo e macroprocesso unificados em uma única plataforma.
            </p>
          </div>

          {/* Feature pills */}
          <div className="flex flex-wrap gap-2">
            {['BPM', 'Catálogo', 'Cenários', 'Macroprocesso'].map((tag) => (
              <span key={tag} className="text-xs text-slate-400 border border-slate-700 rounded-full px-3 py-1">
                {tag}
              </span>
            ))}
          </div>
        </div>

        {/* Footer */}
        <div className="relative z-10">
          <p className="text-slate-600 text-xs">© {new Date().getFullYear()} CollabZ Consultoria</p>
        </div>
      </div>

      {/* Right panel — form */}
      <div className="flex-1 flex flex-col items-center justify-center bg-white px-6 py-12">
        <div className="w-full max-w-sm space-y-8">
          {/* Mobile logo */}
          <div className="flex lg:hidden items-center gap-2 justify-center">
            <div className="w-8 h-8 rounded-lg bg-slate-900 flex items-center justify-center">
              <svg width="18" height="18" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M10 2L17 6V14L10 18L3 14V6L10 2Z" stroke="white" strokeWidth="1.5" fill="none" />
              </svg>
            </div>
            <span className="text-slate-900 font-semibold text-lg">Collab:Flow</span>
          </div>

          {/* Heading */}
          <div className="space-y-1">
            <h1 className="text-2xl font-bold text-slate-900">Bem-vindo de volta</h1>
            <p className="text-sm text-slate-500">Entre com as suas credenciais para continuar</p>
          </div>

          <form action={action} className="space-y-5">
            {state?.error && (
              <div className="flex items-start gap-2.5 rounded-lg bg-red-50 border border-red-200 px-4 py-3">
                <svg className="w-4 h-4 text-red-500 mt-0.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" />
                </svg>
                <p className="text-red-700 text-sm">{state.error}</p>
              </div>
            )}

            <div className="space-y-1.5">
              <label htmlFor="codigo" className="block text-sm font-medium text-slate-700">
                Código do Usuário
              </label>
              <input
                id="codigo"
                name="codigo"
                type="text"
                autoComplete="username"
                required
                placeholder="Ex: JOAO"
                className="w-full rounded-lg border border-slate-300 bg-white px-3.5 py-2.5 text-sm text-slate-900 placeholder:text-slate-400 shadow-sm transition focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500/20"
              />
            </div>

            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <label htmlFor="senha" className="block text-sm font-medium text-slate-700">
                  Senha
                </label>
              </div>
              <input
                id="senha"
                name="senha"
                type="password"
                autoComplete="current-password"
                required
                placeholder="••••••••"
                className="w-full rounded-lg border border-slate-300 bg-white px-3.5 py-2.5 text-sm text-slate-900 placeholder:text-slate-400 shadow-sm transition focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500/20"
              />
            </div>

            <button
              type="submit"
              disabled={pending}
              className="w-full rounded-lg bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-900/30 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {pending ? (
                <span className="flex items-center justify-center gap-2">
                  <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                  </svg>
                  Entrando…
                </span>
              ) : (
                'Entrar'
              )}
            </button>
          </form>

          <p className="text-center text-xs text-slate-400">
            CollabZ Consultoria &middot; Uso interno
          </p>
        </div>
      </div>
    </main>
  )
}
