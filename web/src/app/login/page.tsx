'use client'

import { useActionState } from 'react'
import { login } from '@/actions/auth'
import { BrandMark } from '@/components/brand/logo'

export default function LoginPage() {
  const [state, action, pending] = useActionState(login, undefined)

  return (
    <div className="min-h-screen relative flex items-center justify-center overflow-hidden bg-navy-dark">
      {/* Camadas de gradiente — assinatura clbz */}
      <div
        className="absolute inset-0"
        style={{
          background:
            'linear-gradient(135deg, #072A40 0%, #0B3D5C 55%, #1A6E8E 100%)',
        }}
      />
      <div className="absolute inset-0">
        <div className="absolute top-[-40%] right-[-20%] w-[80%] h-[180%] bg-[radial-gradient(ellipse,rgba(247,168,35,0.10)_0%,transparent_60%)]" />
        <div className="absolute bottom-[-30%] left-[-10%] w-[60%] h-[120%] bg-[radial-gradient(ellipse,rgba(26,110,142,0.30)_0%,transparent_60%)]" />
      </div>
      {/* Faixa gold→teal→gold no topo */}
      <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-gold via-teal to-gold" />

      <div className="relative z-10 w-full max-w-md px-6">
        {/* Header com brand */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-14 h-14 mb-5">
            <BrandMark size={56} />
          </div>
          <div className="inline-flex items-center gap-3 text-gold text-[10px] font-bold tracking-[0.2em] uppercase mb-3">
            <span className="w-5 h-px bg-gold" />
            Collab:Flow — Process &amp; Governance
            <span className="w-5 h-px bg-gold" />
          </div>
          <h1 className="font-display text-4xl text-white leading-[1.1]">
            Acesse sua conta.<br />
            <em className="text-gold not-italic">Seus processos esperam.</em>
          </h1>
        </div>

        {/* Card de login */}
        <div className="bg-white rounded-xl shadow-2xl p-8 border border-white/10">
          <form action={action} className="space-y-5">
            <div>
              <label
                htmlFor="codigo"
                className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2"
              >
                Código do Usuário
              </label>
              <input
                id="codigo"
                name="codigo"
                type="text"
                autoComplete="username"
                required
                className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent transition-all"
                placeholder="Ex: JOAO"
              />
            </div>

            <div>
              <label
                htmlFor="senha"
                className="block text-xs font-semibold tracking-wider uppercase text-navy mb-2"
              >
                Senha
              </label>
              <input
                id="senha"
                name="senha"
                type="password"
                autoComplete="current-password"
                required
                className="w-full rounded-md border border-[#E2E8F0] bg-[#F5F6F8] px-3.5 py-2.5 text-sm text-slate focus:outline-none focus:ring-2 focus:ring-teal focus:border-transparent transition-all"
              />
            </div>

            {state?.error && (
              <p
                role="alert"
                className="text-sm text-[#9A2E1F] bg-[rgba(224,80,64,0.08)] border-l-3 border-[#E05040] rounded-md px-3.5 py-2.5"
              >
                {state.error}
              </p>
            )}

            <button
              type="submit"
              disabled={pending}
              className="w-full bg-navy hover:bg-teal text-white rounded-md py-3 text-sm font-semibold tracking-wide transition-all hover:-translate-y-0.5 hover:shadow-lg disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
            >
              {pending ? 'Entrando...' : 'Entrar no Collab:Flow'}
            </button>
          </form>

          <p className="text-[11px] text-gray-medium text-center mt-6 leading-relaxed">
            Ao acessar, você concorda com as políticas internas de uso e
            privacidade do Collab:Flow.
          </p>
        </div>

        {/* Footer minimal */}
        <p className="text-center text-xs text-white/45 mt-6">
          © {new Date().getFullYear()} Collab:Flow — Gestão de Processos Corporativos
        </p>
      </div>
    </div>
  )
}
