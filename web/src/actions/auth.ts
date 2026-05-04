'use server'

import { redirect } from 'next/navigation'
import bcrypt from 'bcryptjs'
import { prisma } from '@/lib/prisma'
import { createSession, deleteSession } from '@/lib/session'
import { LoginSchema } from '@/lib/definitions'

export async function login(
  _state: { error?: string } | undefined,
  formData: FormData
): Promise<{ error: string }> {
  const validated = LoginSchema.safeParse({
    codigo: formData.get('codigo'),
    senha: formData.get('senha'),
  })

  if (!validated.success) {
    return { error: 'Código e senha são obrigatórios.' }
  }

  const { codigo, senha } = validated.data

  const usuario = await prisma.usuario.findUnique({ where: { codigo } })

  if (!usuario || !usuario.ativo) {
    return { error: 'Usuário não encontrado ou inativo.' }
  }

  const senhaOk = await bcrypt.compare(senha, usuario.senha)
  if (!senhaOk) {
    return { error: 'Senha incorreta.' }
  }

  await createSession({
    userId: usuario.id,
    codigo: usuario.codigo,
    nome: usuario.nome,
    categoria: usuario.categoria,
  })

  redirect('/dashboard')
}

export async function logout() {
  await deleteSession()
  redirect('/login')
}
