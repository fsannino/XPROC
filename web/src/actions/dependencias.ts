'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { DependenciaSchema, type DependenciaTipo } from '@/lib/definitions'

function revalidateMapa() {
  revalidatePath('/dashboard/mapa')
}

export type DependenciaProcessoView = {
  id: string
  origemId: number
  origemDescricao: string
  destinoId: number
  destinoDescricao: string
  tipo: DependenciaTipo
  descricao: string | null
}

export type DependenciaAtividadeView = {
  id: string
  origemId: number
  origemDescricao: string
  destinoId: number
  destinoDescricao: string
  tipo: DependenciaTipo
  descricao: string | null
}

// ─── Processo ↔ Processo ───────────────────────────────────────

export async function listarDependenciasDeProcesso(processoId: number): Promise<{
  saidas: DependenciaProcessoView[]
  entradas: DependenciaProcessoView[]
}> {
  const [saidas, entradas] = await Promise.all([
    prisma.dependenciaProcesso.findMany({
      where: { origemId: processoId },
      include: { destino: { select: { id: true, descricao: true } } },
      orderBy: { criadoEm: 'asc' },
    }),
    prisma.dependenciaProcesso.findMany({
      where: { destinoId: processoId },
      include: { origem: { select: { id: true, descricao: true } } },
      orderBy: { criadoEm: 'asc' },
    }),
  ])
  return {
    saidas: saidas.map((d) => ({
      id: d.id,
      origemId: processoId,
      origemDescricao: '',
      destinoId: d.destinoId,
      destinoDescricao: d.destino.descricao,
      tipo: d.tipo as DependenciaTipo,
      descricao: d.descricao,
    })),
    entradas: entradas.map((d) => ({
      id: d.id,
      origemId: d.origemId,
      origemDescricao: d.origem.descricao,
      destinoId: processoId,
      destinoDescricao: '',
      tipo: d.tipo as DependenciaTipo,
      descricao: d.descricao,
    })),
  }
}

export async function listarTodasDependenciasDeProcesso(): Promise<DependenciaProcessoView[]> {
  const rows = await prisma.dependenciaProcesso.findMany({
    include: {
      origem: { select: { descricao: true } },
      destino: { select: { descricao: true } },
    },
  })
  return rows.map((d) => ({
    id: d.id,
    origemId: d.origemId,
    origemDescricao: d.origem.descricao,
    destinoId: d.destinoId,
    destinoDescricao: d.destino.descricao,
    tipo: d.tipo as DependenciaTipo,
    descricao: d.descricao,
  }))
}

export async function criarDependenciaProcesso(input: {
  origemId: number
  destinoId: number
  tipo: DependenciaTipo
  descricao?: string
}) {
  const v = DependenciaSchema.safeParse(input)
  if (!v.success) return { error: v.error.issues[0]?.message ?? 'Dados inválidos.' }
  try {
    const d = await prisma.dependenciaProcesso.create({
      data: {
        origemId: v.data.origemId,
        destinoId: v.data.destinoId,
        tipo: v.data.tipo,
        descricao: v.data.descricao || null,
      },
    })
    revalidateMapa()
    return { success: true, id: d.id }
  } catch {
    return { error: 'Não foi possível criar (já existe essa dependência?).' }
  }
}

export async function excluirDependenciaProcesso(id: string) {
  try {
    await prisma.dependenciaProcesso.delete({ where: { id } })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir.' }
  }
}

// ─── Atividade ↔ Atividade ────────────────────────────────────

export async function listarDependenciasDeAtividade(atividadeId: number): Promise<{
  saidas: DependenciaAtividadeView[]
  entradas: DependenciaAtividadeView[]
}> {
  const [saidas, entradas] = await Promise.all([
    prisma.dependenciaAtividade.findMany({
      where: { origemId: atividadeId },
      include: { destino: { select: { id: true, descricao: true } } },
      orderBy: { criadoEm: 'asc' },
    }),
    prisma.dependenciaAtividade.findMany({
      where: { destinoId: atividadeId },
      include: { origem: { select: { id: true, descricao: true } } },
      orderBy: { criadoEm: 'asc' },
    }),
  ])
  return {
    saidas: saidas.map((d) => ({
      id: d.id,
      origemId: atividadeId,
      origemDescricao: '',
      destinoId: d.destinoId,
      destinoDescricao: d.destino.descricao,
      tipo: d.tipo as DependenciaTipo,
      descricao: d.descricao,
    })),
    entradas: entradas.map((d) => ({
      id: d.id,
      origemId: d.origemId,
      origemDescricao: d.origem.descricao,
      destinoId: atividadeId,
      destinoDescricao: '',
      tipo: d.tipo as DependenciaTipo,
      descricao: d.descricao,
    })),
  }
}

export async function listarTodasDependenciasDeAtividade(): Promise<DependenciaAtividadeView[]> {
  const rows = await prisma.dependenciaAtividade.findMany({
    include: {
      origem: { select: { descricao: true } },
      destino: { select: { descricao: true } },
    },
  })
  return rows.map((d) => ({
    id: d.id,
    origemId: d.origemId,
    origemDescricao: d.origem.descricao,
    destinoId: d.destinoId,
    destinoDescricao: d.destino.descricao,
    tipo: d.tipo as DependenciaTipo,
    descricao: d.descricao,
  }))
}

export async function criarDependenciaAtividade(input: {
  origemId: number
  destinoId: number
  tipo: DependenciaTipo
  descricao?: string
}) {
  const v = DependenciaSchema.safeParse(input)
  if (!v.success) return { error: v.error.issues[0]?.message ?? 'Dados inválidos.' }
  try {
    const d = await prisma.dependenciaAtividade.create({
      data: {
        origemId: v.data.origemId,
        destinoId: v.data.destinoId,
        tipo: v.data.tipo,
        descricao: v.data.descricao || null,
      },
    })
    revalidateMapa()
    return { success: true, id: d.id }
  } catch {
    return { error: 'Não foi possível criar (já existe essa dependência?).' }
  }
}

export async function excluirDependenciaAtividade(id: string) {
  try {
    await prisma.dependenciaAtividade.delete({ where: { id } })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir.' }
  }
}
