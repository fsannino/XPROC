'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { getSession } from '@/lib/session'
import { AlterarStatusSchema, ComentarioSchema } from '@/lib/definitions'
import { enviarEmail, htmlMudancaStatus } from '@/lib/email'

// ─── Status ───────────────────────────────────────────────────────

export async function alterarStatus(megaProcessoId: number, novoStatus: string) {
  const session = await getSession()
  if (!session) return { error: 'Não autenticado.' }

  const validated = AlterarStatusSchema.safeParse({ megaProcessoId, status: novoStatus })
  if (!validated.success) return { error: 'Status inválido.' }

  const mp = await prisma.megaProcesso.findUnique({
    where: { id: megaProcessoId },
    include: { responsavel: { select: { email: true, nome: true } } },
  })
  if (!mp) return { error: 'Mega-processo não encontrado.' }

  const statusAnterior = mp.status

  const totalVersoes = await prisma.megaProcessoVersao.count({ where: { megaProcessoId } })

  await prisma.$transaction([
    prisma.megaProcesso.update({
      where: { id: megaProcessoId },
      data: { status: novoStatus },
    }),
    prisma.megaProcessoVersao.create({
      data: {
        megaProcessoId,
        versao: totalVersoes + 1,
        statusAnterior,
        statusNovo: novoStatus,
        criadoPorId: session.userId,
      },
    }),
    prisma.usuarioHistorico.create({
      data: {
        usuarioId: session.userId,
        operacao: 'STATUS_ALTERADO',
        descricao: `${mp.descricao}: ${statusAnterior} → ${novoStatus}`,
      },
    }),
  ])

  if (mp.responsavel?.email) {
    const appUrl = process.env.NEXT_PUBLIC_APP_URL ?? ''
    await enviarEmail({
      para: mp.responsavel.email,
      assunto: `[XPROC] Status alterado: ${mp.descricao}`,
      html: htmlMudancaStatus(
        mp.descricao,
        statusAnterior,
        novoStatus,
        session.nome,
        `${appUrl}/dashboard/processos/${megaProcessoId}`,
      ),
    })
  }

  revalidatePath(`/dashboard/processos/${megaProcessoId}`)
  revalidatePath('/dashboard/processos')
  return { success: true }
}

// ─── Comentários ──────────────────────────────────────────────────

export async function adicionarComentario(_state: unknown, formData: FormData) {
  const session = await getSession()
  if (!session) return { error: 'Não autenticado.' }

  const validated = ComentarioSchema.safeParse({
    megaProcessoId: Number(formData.get('megaProcessoId')),
    texto: formData.get('texto'),
    parentId: (formData.get('parentId') as string) || undefined,
  })
  if (!validated.success) return { errors: validated.error.flatten().fieldErrors }

  const { megaProcessoId, texto, parentId } = validated.data

  await prisma.processoComentario.create({
    data: {
      megaProcessoId,
      usuarioId: session.userId,
      texto,
      parentId: parentId ?? null,
    },
  })

  revalidatePath(`/dashboard/processos/${megaProcessoId}`)
  return { success: true }
}

export async function excluirComentario(id: string) {
  const session = await getSession()
  if (!session) return

  const comentario = await prisma.processoComentario.findUnique({ where: { id } })
  if (!comentario) return

  if (comentario.usuarioId !== session.userId && session.categoria !== 'A') return

  await prisma.processoComentario.delete({ where: { id } })
  revalidatePath(`/dashboard/processos/${comentario.megaProcessoId}`)
}



export async function adicionarComentarioForm(formData: FormData): Promise<void> {
  await adicionarComentario(undefined, formData)
}
