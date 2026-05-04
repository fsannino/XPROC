import TrocaSenhaForm from './form'

export const metadata = { title: 'Minha Conta' }

export default function ContaPage() {
  return (
    <div className="max-w-md">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Minha Conta</h1>
      <div className="bg-white rounded-xl border border-gray-100 p-6">
        <h2 className="text-base font-semibold text-gray-800 mb-4">Alterar Senha</h2>
        <TrocaSenhaForm />
      </div>
    </div>
  )
}
