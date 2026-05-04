'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { TransacaoSchema, EmpresaSchema } from '@/lib/definitions'

export async function criarTransacao(_state: unknown, formData: FormData) {
  const validated = TransacaoSchema.safeParse({
    id: formData.get('id'),
    descricao: formData.get('descricao') || undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.transacao.create({ data: validated.data })
  revalidatePath('/dashboard/transacoes')
  return { success: true }
}

export async function atualizarTransacao(id: string, _state: unknown, formData: FormData) {
  const descricao = (formData.get('descricao') as string)?.trim() || null
  await prisma.transacao.update({ where: { id }, data: { descricao } })
  revalidatePath('/dashboard/transacoes')
  return { success: true }
}

export async function excluirTransacao(id: string) {
  await prisma.transacao.delete({ where: { id } })
  revalidatePath('/dashboard/transacoes')
}

export async function criarEmpresa(_state: unknown, formData: FormData) {
  const validated = EmpresaSchema.safeParse({ nome: formData.get('nome') })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  await prisma.empresaUnidade.create({ data: validated.data })
  revalidatePath('/dashboard/empresas')
  return { success: true }
}

export async function excluirEmpresa(id: number) {
  await prisma.empresaUnidade.delete({ where: { id } })
  revalidatePath('/dashboard/empresas')
}
