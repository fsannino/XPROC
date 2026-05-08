'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { getSession } from '@/lib/session'
import { RiscoSchema } from '@/lib/definitions'

export async function adicionarRisco(_state: unknown, formData: FormData) {
  const session = await getSession()
  if (!session) return { error: 'Não autenticado.' }

  const validated = RiscoSchema.safeParse({
    megaProcessoId: Number(formData.get('megaProcessoId')),
    descricao: formData.get('descricao'),
    probabilidade: formData.get('probabilidade') ?? 3,
    impacto: formData.get('impacto') ?? 3,
    controle: formData.get('controle') || undefined,
  })
  if (!validated.success) return { errors: validated.error.flatten().fieldErrors }

  const { megaProcessoId, descricao, probabilidade, impacto, controle } = validated.data

  await prisma.processoRisco.create({
    data: {
      megaProcessoId,
      descricao,
      probabilidade,
      impacto,
      controle: controle || null,
    },
  })

  revalidatePath(`/dashboard/processos/${megaProcessoId}`)
  return { success: true }
}

export async function excluirRisco(id: string) {
  const session = await getSession()
  if (!session) return

  const risco = await prisma.processoRisco.findUnique({ where: { id } })
  if (!risco) return

  await prisma.processoRisco.delete({ where: { id } })
  revalidatePath(`/dashboard/processos/${risco.megaProcessoId}`)
}

export async function adicionarRiscoForm(formData: FormData): Promise<void> {
  await adicionarRisco(undefined, formData)
}
