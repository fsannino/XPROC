'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import {
  InsumoSchema,
  SetInsumosDeProcessoSchema,
  SetInsumosDeAtividadeSchema,
  type InsumoTipo,
  type InsumoDirecao,
} from '@/lib/definitions'

function revalidateMapa() {
  revalidatePath('/dashboard/mapa')
  revalidatePath('/dashboard/catalogo')
}

export type InsumoView = {
  id: number
  codigo: string
  descricao: string
  tipo: InsumoTipo
}

export type InsumoVinculoView = InsumoView & { direcao: InsumoDirecao }

export async function listarInsumos(): Promise<InsumoView[]> {
  const rows = await prisma.insumo.findMany({ orderBy: { descricao: 'asc' } })
  return rows.map((i) => ({
    id: i.id,
    codigo: i.codigo,
    descricao: i.descricao,
    tipo: i.tipo as InsumoTipo,
  }))
}

export async function criarInsumo(input: { codigo: string; descricao: string; tipo: InsumoTipo }) {
  const v = InsumoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const i = await prisma.insumo.create({ data: v.data })
    revalidateMapa()
    return { success: true, id: i.id }
  } catch {
    return { error: 'Não foi possível criar (código já existe?).' }
  }
}

export async function atualizarInsumo(id: number, input: { codigo: string; descricao: string; tipo: InsumoTipo }) {
  const v = InsumoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    await prisma.insumo.update({ where: { id }, data: v.data })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirInsumo(id: number) {
  try {
    await prisma.insumo.delete({ where: { id } })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir (em uso?).' }
  }
}

// ─── vinculos com Processo ─────────────────────────────────────

export async function getInsumosDeProcesso(processoId: number): Promise<InsumoVinculoView[]> {
  const rows = await prisma.insumoProcesso.findMany({
    where: { processoId },
    include: { insumo: true },
    orderBy: [{ direcao: 'asc' }, { insumo: { descricao: 'asc' } }],
  })
  return rows.map((r) => ({
    id: r.insumo.id,
    codigo: r.insumo.codigo,
    descricao: r.insumo.descricao,
    tipo: r.insumo.tipo as InsumoTipo,
    direcao: r.direcao as InsumoDirecao,
  }))
}

export async function setInsumosDeProcesso(input: {
  processoId: number
  vinculos: { insumoId: number; direcao: InsumoDirecao }[]
}) {
  const v = SetInsumosDeProcessoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { processoId, vinculos } = v.data
  await prisma.$transaction([
    prisma.insumoProcesso.deleteMany({ where: { processoId } }),
    ...(vinculos.length > 0
      ? [prisma.insumoProcesso.createMany({
          data: vinculos.map(({ insumoId, direcao }) => ({ processoId, insumoId, direcao })),
          skipDuplicates: true,
        })]
      : []),
  ])
  revalidateMapa()
  return { success: true }
}

// ─── vinculos com Atividade ────────────────────────────────────

export async function getInsumosDeAtividade(atividadeId: number): Promise<InsumoVinculoView[]> {
  const rows = await prisma.insumoAtividade.findMany({
    where: { atividadeId },
    include: { insumo: true },
    orderBy: [{ direcao: 'asc' }, { insumo: { descricao: 'asc' } }],
  })
  return rows.map((r) => ({
    id: r.insumo.id,
    codigo: r.insumo.codigo,
    descricao: r.insumo.descricao,
    tipo: r.insumo.tipo as InsumoTipo,
    direcao: r.direcao as InsumoDirecao,
  }))
}

export async function setInsumosDeAtividade(input: {
  atividadeId: number
  vinculos: { insumoId: number; direcao: InsumoDirecao }[]
}) {
  const v = SetInsumosDeAtividadeSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { atividadeId, vinculos } = v.data
  await prisma.$transaction([
    prisma.insumoAtividade.deleteMany({ where: { atividadeId } }),
    ...(vinculos.length > 0
      ? [prisma.insumoAtividade.createMany({
          data: vinculos.map(({ insumoId, direcao }) => ({ atividadeId, insumoId, direcao })),
          skipDuplicates: true,
        })]
      : []),
  ])
  revalidateMapa()
  return { success: true }
}

export async function criarInsumoForm(_prev: unknown, fd: FormData) {
  return criarInsumo({
    codigo: String(fd.get('codigo') ?? ''),
    descricao: String(fd.get('descricao') ?? ''),
    tipo: (String(fd.get('tipo') ?? 'DADO') as InsumoTipo),
  })
}
