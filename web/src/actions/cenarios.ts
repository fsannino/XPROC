'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { CenarioSchema } from '@/lib/definitions'

export async function criarCenario(_state: unknown, formData: FormData) {
  const validated = CenarioSchema.safeParse({
    descricao: formData.get('descricao'),
    situacao: formData.get('situacao') || undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.cenario.create({ data: validated.data })
  revalidatePath('/dashboard/cenarios')
  return { success: true }
}

export async function atualizarCenario(id: number, _state: unknown, formData: FormData) {
  const validated = CenarioSchema.safeParse({
    descricao: formData.get('descricao'),
    situacao: formData.get('situacao') || undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.cenario.update({ where: { id }, data: validated.data })
  revalidatePath('/dashboard/cenarios')
  return { success: true }
}

export async function excluirCenario(id: number) {
  await prisma.cenario.delete({ where: { id } })
  revalidatePath('/dashboard/cenarios')
}
