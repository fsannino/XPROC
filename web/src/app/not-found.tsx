import Link from 'next/link'

export default function NotFound() {
  return (
    <div className="flex flex-col items-center justify-center min-h-screen gap-4 text-center bg-gray-50">
      <div className="text-6xl font-bold text-gray-200">404</div>
      <h1 className="text-xl font-semibold text-gray-800">Página não encontrada</h1>
      <p className="text-sm text-gray-500">O endereço que você acessou não existe.</p>
      <Link
        href="/dashboard"
        className="bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-800"
      >
        Voltar ao início
      </Link>
    </div>
  )
}
