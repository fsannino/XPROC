'use server'

import { revalidatePath } from 'next/cache'
import bcrypt from 'bcryptjs'
import { prisma } from '@/lib/prisma'
import { UsuarioSchema } from '@/lib/definitions'

export async function criarUsuario(_state: unknown, formData: FormData) {
  const validated = UsuarioSchema.safeParse({
    codigo: formData.get('codigo'),
    nome: formData.get('nome'),
    email: formData.get('email') || undefined,
    senha: formData.get('senha'),
    categoria: formData.get('categoria') || undefined,
  })

  if (!validated.success) {
    return { errors: validated.error.flatten().fieldErrors }
  }

  const { senha, email, ...rest } = validated.data
  const senhaHash = await bcrypt.hash(senha, 12)

  await prisma.usuario.create({
    data: {
      ...rest,
      email: email || null,
      senha: senhaHash,
    },
  })

  revalidatePath('/dashboard/usuarios')
  return { success: true }
}

export async function atualizarUsuario(id: string, _state: unknown, formData: FormData) {
  const nova_senha = formData.get('senha') as string

  const data: Record<string, unknown> = {
    nome: formData.get('nome'),
    email: formData.get('email') || null,
    categoria: formData.get('categoria') || null,
    ativo: formData.get('ativo') === 'true',
  }

  if (nova_senha) {
    data.senha = await bcrypt.hash(nova_senha, 12)
  }

  await prisma.usuario.update({ where: { id }, data })
  revalidatePath('/dashboard/usuarios')
  return { success: true }
}

export async function alternarStatusUsuario(id: string, ativo: boolean) {
  await prisma.usuario.update({ where: { id }, data: { ativo } })
  revalidatePath('/dashboard/usuarios')
}

export async function concederAcesso(usuarioId: string, megaProcessoId: number) {
  await prisma.acesso.upsert({
    where: { usuarioId_megaProcessoId: { usuarioId, megaProcessoId } },
    update: {},
    create: { usuarioId, megaProcessoId },
  })
  revalidatePath('/dashboard/usuarios')
}

export async function revogarAcesso(usuarioId: string, megaProcessoId: number) {
  await prisma.acesso.deleteMany({ where: { usuarioId, megaProcessoId } })
  revalidatePath('/dashboard/usuarios')
}
