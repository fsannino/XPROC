'use server'

import { redirect } from 'next/navigation'
import bcrypt from 'bcryptjs'
import { prisma } from '@/lib/prisma'
import { createSession, deleteSession, getSession } from '@/lib/session'
import { LoginSchema } from '@/lib/definitions'

const MAX_TENTATIVAS = 5
const JANELA_MS = 5 * 60 * 1000 // 5 minutos

async function registrarHistorico(usuarioId: string, operacao: string, descricao?: string) {
  await prisma.usuarioHistorico.create({ data: { usuarioId, operacao, descricao } })
}

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

  // Rate limiting: conta falhas recentes no histórico
  const desde = new Date(Date.now() - JANELA_MS)
  const falhasRecentes = await prisma.usuarioHistorico.count({
    where: { usuarioId: usuario.id, operacao: 'LOGIN_FALHA', criadoEm: { gte: desde } },
  })

  if (falhasRecentes >= MAX_TENTATIVAS) {
    return { error: 'Muitas tentativas. Aguarde 5 minutos.' }
  }

  const senhaOk = await bcrypt.compare(senha, usuario.senha)
  if (!senhaOk) {
    await registrarHistorico(usuario.id, 'LOGIN_FALHA', `Tentativa ${falhasRecentes + 1}`)
    return { error: 'Senha incorreta.' }
  }

  await createSession({
    userId: usuario.id,
    codigo: usuario.codigo,
    nome: usuario.nome,
    categoria: usuario.categoria,
  })

  await registrarHistorico(usuario.id, 'LOGIN')

  redirect('/dashboard')
}

export async function logout() {
  const session = await getSession()
  if (session) {
    await registrarHistorico(session.userId, 'LOGOUT').catch(() => null)
  }
  await deleteSession()
  redirect('/login')
}
