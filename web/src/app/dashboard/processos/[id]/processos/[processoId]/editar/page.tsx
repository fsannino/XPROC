import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import EditarProcessoForm from './form'

export default async function EditarProcessoPage({
  params,
}: {
  params: Promise<{ id: string; processoId: string }>
}) {
  const { id, processoId } = await params
  const processo = await prisma.processo.findUnique({ where: { id: Number(processoId) } })
  if (!processo || processo.megaProcessoId !== Number(id)) notFound()

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Editar Processo</h1>
      <EditarProcessoForm processo={processo} megaProcessoId={Number(id)} />
    </div>
  )
}
