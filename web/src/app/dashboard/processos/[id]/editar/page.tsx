import { prisma } from '@/lib/prisma'
import { notFound } from 'next/navigation'
import EditarMegaProcessoForm from './form'

export default async function EditarMegaProcessoPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = await params
  const [megaProcesso, usuarios] = await Promise.all([
    prisma.megaProcesso.findUnique({ where: { id: Number(id) } }),
    prisma.usuario.findMany({
      where: { ativo: true },
      select: { id: true, codigo: true, nome: true },
      orderBy: { nome: 'asc' },
    }),
  ])
  if (!megaProcesso) notFound()

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Editar Mega-Processo</h1>
      <EditarMegaProcessoForm megaProcesso={megaProcesso} usuarios={usuarios} />
    </div>
  )
}
