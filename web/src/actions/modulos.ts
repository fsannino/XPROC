'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { z } from 'zod'

const ModuloSchema = z.object({ descricao: z.string().min(2).max(80).trim() })
const SubModuloSchema = z.object({
  moduloId: z.number().int().positive().optional(),
  descricao: z.string().min(2).max(80).trim(),
  abreviacao: z.string().max(10).optional().or(z.literal('')),
})

export async function criarModulo(_state: unknown, formData: FormData) {
  const validated = ModuloSchema.safeParse({ descricao: formData.get('descricao') })
  if (!validated.success) return { errors: validated.error.flatten().fieldErrors }
  await prisma.modulo.create({ data: validated.data })
  revalidatePath('/dashboard/modulos')
  return { success: true }
}

export async function excluirModulo(id: number) {
  await prisma.modulo.delete({ where: { id } })
  revalidatePath('/dashboard/modulos')
}

export async function criarSubModulo(_state: unknown, formData: FormData) {
  const moduloIdRaw = formData.get('moduloId')
  const validated = SubModuloSchema.safeParse({
    moduloId: moduloIdRaw ? Number(moduloIdRaw) : undefined,
    descricao: formData.get('descricao'),
    abreviacao: formData.get('abreviacao') || undefined,
  })
  if (!validated.success) return { errors: validated.error.flatten().fieldErrors }
  await prisma.subModulo.create({ data: validated.data })
  revalidatePath('/dashboard/modulos')
  return { success: true }
}

export async function excluirSubModulo(id: number) {
  await prisma.subModulo.delete({ where: { id } })
  revalidatePath('/dashboard/modulos')
}
