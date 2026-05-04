'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { MegaProcessoSchema, ProcessoSchema, SubProcessoSchema } from '@/lib/definitions'

// ─── Mega-Processo ────────────────────────────────────────────────

export async function criarMegaProcesso(_state: unknown, formData: FormData) {
  const validated = MegaProcessoSchema.safeParse({
    descricao: formData.get('descricao'),
    abreviacao: formData.get('abreviacao') || undefined,
    descricaoLonga: formData.get('descricaoLonga') || undefined,
    bloqueado: formData.get('bloqueado') === 'true',
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.megaProcesso.create({ data: validated.data })
  revalidatePath('/dashboard/processos')
  return { success: true }
}

export async function atualizarMegaProcesso(id: number, _state: unknown, formData: FormData) {
  const validated = MegaProcessoSchema.safeParse({
    descricao: formData.get('descricao'),
    abreviacao: formData.get('abreviacao') || undefined,
    descricaoLonga: formData.get('descricaoLonga') || undefined,
    bloqueado: formData.get('bloqueado') === 'true',
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.megaProcesso.update({ where: { id }, data: validated.data })
  revalidatePath('/dashboard/processos')
  return { success: true }
}

export async function excluirMegaProcesso(id: number) {
  await prisma.megaProcesso.delete({ where: { id } })
  revalidatePath('/dashboard/processos')
}

// ─── Processo ─────────────────────────────────────────────────────

export async function criarProcesso(_state: unknown, formData: FormData) {
  const validated = ProcessoSchema.safeParse({
    megaProcessoId: Number(formData.get('megaProcessoId')),
    descricao: formData.get('descricao'),
    sequencia: formData.get('sequencia') ? Number(formData.get('sequencia')) : undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.processo.create({ data: validated.data })
  revalidatePath('/dashboard/processos')
  return { success: true }
}

export async function excluirProcesso(id: number) {
  await prisma.processo.delete({ where: { id } })
  revalidatePath('/dashboard/processos')
}

// ─── Sub-Processo ─────────────────────────────────────────────────

export async function criarSubProcesso(_state: unknown, formData: FormData) {
  const validated = SubProcessoSchema.safeParse({
    processoId: Number(formData.get('processoId')),
    megaProcessoId: Number(formData.get('megaProcessoId')),
    descricao: formData.get('descricao'),
    sequencia: formData.get('sequencia') ? Number(formData.get('sequencia')) : undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.subProcesso.create({ data: validated.data })
  revalidatePath('/dashboard/processos')
  return { success: true }
}

export async function excluirSubProcesso(id: number) {
  await prisma.subProcesso.delete({ where: { id } })
  revalidatePath('/dashboard/processos')
}
