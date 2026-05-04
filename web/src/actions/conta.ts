'use server'

import bcrypt from 'bcryptjs'
import { prisma } from '@/lib/prisma'
import { getSession } from '@/lib/session'
import { TrocaSenhaSchema } from '@/lib/definitions'

export async function trocarSenha(_state: unknown, formData: FormData) {
  const session = await getSession()
  if (!session) return { error: 'Sessão inválida.' }

  const validated = TrocaSenhaSchema.safeParse({
    senhaAtual: formData.get('senhaAtual'),
    novaSenha: formData.get('novaSenha'),
    confirmar: formData.get('confirmar'),
  })

  if (!validated.success) {
    const erros = validated.error.flatten()
    const primeira = Object.values(erros.fieldErrors).flat()[0] ?? erros.formErrors[0]
    return { error: primeira }
  }

  const usuario = await prisma.usuario.findUnique({ where: { id: session.userId } })
  if (!usuario) return { error: 'Usuário não encontrado.' }

  const senhaOk = await bcrypt.compare(validated.data.senhaAtual, usuario.senha)
  if (!senhaOk) return { error: 'Senha atual incorreta.' }

  const hash = await bcrypt.hash(validated.data.novaSenha, 12)
  await prisma.usuario.update({ where: { id: session.userId }, data: { senha: hash } })

  return { success: true }
}
