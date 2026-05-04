import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import EditarMegaProcessoForm from './form'

export default async function EditarMegaProcessoPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params
  const megaProcesso = await prisma.megaProcesso.findUnique({ where: { id: Number(id) } })
  if (!megaProcesso) notFound()

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Editar Mega-Processo</h1>
      <EditarMegaProcessoForm megaProcesso={megaProcesso} />
    </div>
  )
}
