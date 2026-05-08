'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { z } from 'zod'

/**
 * Substitui o conjunto de relações de uma entidade-pivot.
 * Usa transação Prisma: deleteMany + createMany para sincronizar.
 */

const SetTransacaoProcessosSchema = z.object({
  transacaoId: z.string().min(1),
  processoIds: z.array(z.number().int().positive()),
})

const SetTransacaoAtividadesSchema = z.object({
  transacaoId: z.string().min(1),
  atividadeIds: z.array(z.number().int().positive()),
})

const SetTransacaoMacrosSchema = z.object({
  transacaoId: z.string().min(1),
  megaProcessoIds: z.array(z.number().int().positive()),
})

const SetCenarioProcessosSchema = z.object({
  cenarioId: z.number().int().positive(),
  processoIds: z.array(z.number().int().positive()),
})

const SetCenarioAtividadesSchema = z.object({
  cenarioId: z.number().int().positive(),
  atividadeIds: z.array(z.number().int().positive()),
})

const SetCenarioTransacoesSchema = z.object({
  cenarioId: z.number().int().positive(),
  transacaoIds: z.array(z.string().min(1)),
})

const SetProcessoTransacoesSchema = z.object({
  processoId: z.number().int().positive(),
  transacaoIds: z.array(z.string().min(1)),
})

const SetAtividadeTransacoesSchema = z.object({
  atividadeId: z.number().int().positive(),
  transacaoIds: z.array(z.string().min(1)),
})

function revalidateAll() {
  revalidatePath('/dashboard', 'layout')
}

export async function setTransacaoProcessos(input: { transacaoId: string; processoIds: number[] }) {
  const v = SetTransacaoProcessosSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { transacaoId, processoIds } = v.data
  await prisma.$transaction([
    prisma.transacaoProcesso.deleteMany({ where: { transacaoId } }),
    prisma.transacaoProcesso.createMany({
      data: processoIds.map((processoId) => ({ transacaoId, processoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setTransacaoAtividades(input: { transacaoId: string; atividadeIds: number[] }) {
  const v = SetTransacaoAtividadesSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { transacaoId, atividadeIds } = v.data
  await prisma.$transaction([
    prisma.transacaoAtividade.deleteMany({ where: { transacaoId } }),
    prisma.transacaoAtividade.createMany({
      data: atividadeIds.map((atividadeId) => ({ transacaoId, atividadeId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setTransacaoMacroprocessos(input: { transacaoId: string; megaProcessoIds: number[] }) {
  const v = SetTransacaoMacrosSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { transacaoId, megaProcessoIds } = v.data
  await prisma.$transaction([
    prisma.transacaoMega.deleteMany({ where: { transacaoId } }),
    prisma.transacaoMega.createMany({
      data: megaProcessoIds.map((megaProcessoId) => ({ transacaoId, megaProcessoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setCenarioProcessos(input: { cenarioId: number; processoIds: number[] }) {
  const v = SetCenarioProcessosSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { cenarioId, processoIds } = v.data
  await prisma.$transaction([
    prisma.cenarioProcesso.deleteMany({ where: { cenarioId } }),
    prisma.cenarioProcesso.createMany({
      data: processoIds.map((processoId) => ({ cenarioId, processoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setCenarioAtividades(input: { cenarioId: number; atividadeIds: number[] }) {
  const v = SetCenarioAtividadesSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { cenarioId, atividadeIds } = v.data
  await prisma.$transaction([
    prisma.cenarioAtividade.deleteMany({ where: { cenarioId } }),
    prisma.cenarioAtividade.createMany({
      data: atividadeIds.map((atividadeId) => ({ cenarioId, atividadeId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setCenarioTransacoes(input: { cenarioId: number; transacaoIds: string[] }) {
  const v = SetCenarioTransacoesSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { cenarioId, transacaoIds } = v.data
  await prisma.$transaction([
    prisma.cenarioTransacao.deleteMany({ where: { cenarioId } }),
    prisma.cenarioTransacao.createMany({
      data: transacaoIds.map((transacaoId) => ({ cenarioId, transacaoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

/** Edição pelo lado oposto: do nó do mapa (Processo/Atividade) para suas transações. */
export async function setProcessoTransacoes(input: { processoId: number; transacaoIds: string[] }) {
  const v = SetProcessoTransacoesSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { processoId, transacaoIds } = v.data
  await prisma.$transaction([
    prisma.transacaoProcesso.deleteMany({ where: { processoId } }),
    prisma.transacaoProcesso.createMany({
      data: transacaoIds.map((transacaoId) => ({ processoId, transacaoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function setAtividadeTransacoes(input: { atividadeId: number; transacaoIds: string[] }) {
  const v = SetAtividadeTransacoesSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { atividadeId, transacaoIds } = v.data
  await prisma.$transaction([
    prisma.transacaoAtividade.deleteMany({ where: { atividadeId } }),
    prisma.transacaoAtividade.createMany({
      data: transacaoIds.map((transacaoId) => ({ atividadeId, transacaoId })),
      skipDuplicates: true,
    }),
  ])
  revalidateAll()
  return { success: true }
}

export async function getTransacoesDeProcesso(processoId: number): Promise<string[]> {
  const rows = await prisma.transacaoProcesso.findMany({
    where: { processoId },
    select: { transacaoId: true },
    orderBy: { transacaoId: 'asc' },
  })
  return rows.map((r) => r.transacaoId)
}

export async function getTransacoesDeAtividade(atividadeId: number): Promise<string[]> {
  const rows = await prisma.transacaoAtividade.findMany({
    where: { atividadeId },
    select: { transacaoId: true },
    orderBy: { transacaoId: 'asc' },
  })
  return rows.map((r) => r.transacaoId)
}
