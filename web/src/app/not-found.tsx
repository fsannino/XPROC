import Link from 'next/link'

export default function NotFound() {
  return (
    <div className="min-h-screen flex flex-col items-center justify-center gap-4 text-center bg-cream px-6">
      <p className="section-tag">Erro 404</p>
      <h1 className="font-display text-5xl text-navy">Página não encontrada</h1>
      <p className="text-sm text-gray-medium max-w-sm">
        O endereço que você acessou não existe ou foi movido.
      </p>
      <Link
        href="/dashboard"
        className="inline-flex items-center gap-2 bg-navy hover:bg-teal text-white px-5 py-2.5 rounded-md text-sm font-semibold transition-all hover:-translate-y-0.5 hover:shadow-md"
      >
        Voltar ao painel
      </Link>
      <div className="mt-8 h-1 w-32 bg-gradient-to-r from-gold via-teal to-gold rounded-full" />
    </div>
  )
}
