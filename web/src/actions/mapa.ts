'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import {
  NODE_TYPES,
  NodePositionSchema,
  NodeUpsertSchema,
  type NodeType,
} from '@/lib/definitions'

export type MapaNode = {
  tipo: NodeType
  id: number
  parentId: number | null
  descricao: string
  abreviacao?: string | null
  sequencia?: number | null
  posicaoX: number
  posicaoY: number
}

export type MapaEdge = {
  source: string // `${tipo}:${id}`
  target: string
}

/** Invalida todas as telas que mostram dados criados/editados via mapa. */
function revalidateAll() {
  revalidatePath('/dashboard', 'layout')
}

/**
 * Carrega todos os nós e arestas (pai→filho) da cadeia de valor.
 * Edges derivam das FKs (cadeiaValorId, megaProcessoId, processoId, subProcessoId).
 */
export async function getMapa(): Promise<{ nodes: MapaNode[]; edges: MapaEdge[] }> {
  const [cadeias, macros, processos, macroAtivs, atividades] = await Promise.all([
    prisma.cadeiaValor.findMany({ orderBy: { id: 'asc' } }),
    prisma.megaProcesso.findMany({ orderBy: { id: 'asc' } }),
    prisma.processo.findMany({ orderBy: { sequencia: 'asc' } }),
    prisma.subProcesso.findMany({ orderBy: { sequencia: 'asc' } }),
    prisma.atividade.findMany({ orderBy: { sequencia: 'asc' } }),
  ])

  const nodes: MapaNode[] = [
    ...cadeias.map((c): MapaNode => ({
      tipo: 'cadeia', id: c.id, parentId: null,
      descricao: c.descricao, abreviacao: c.abreviacao,
      posicaoX: c.posicaoX, posicaoY: c.posicaoY,
    })),
    ...macros.map((m): MapaNode => ({
      tipo: 'macroprocesso', id: m.id, parentId: m.cadeiaValorId,
      descricao: m.descricao, abreviacao: m.abreviacao,
      posicaoX: m.posicaoX, posicaoY: m.posicaoY,
    })),
    ...processos.map((p): MapaNode => ({
      tipo: 'processo', id: p.id, parentId: p.megaProcessoId,
      descricao: p.descricao, sequencia: p.sequencia,
      posicaoX: p.posicaoX, posicaoY: p.posicaoY,
    })),
    ...macroAtivs.map((s): MapaNode => ({
      tipo: 'macroatividade', id: s.id, parentId: s.processoId,
      descricao: s.descricao, sequencia: s.sequencia,
      posicaoX: s.posicaoX, posicaoY: s.posicaoY,
    })),
    ...atividades.map((a): MapaNode => ({
      tipo: 'atividade', id: a.id, parentId: a.subProcessoId,
      descricao: a.descricao, sequencia: a.sequencia,
      posicaoX: a.posicaoX, posicaoY: a.posicaoY,
    })),
  ]

  const parentTipo: Record<NodeType, NodeType | null> = {
    cadeia: null, macroprocesso: 'cadeia', processo: 'macroprocesso',
    macroatividade: 'processo', atividade: 'macroatividade',
  }

  const edges: MapaEdge[] = []
  for (const n of nodes) {
    if (n.parentId == null) continue
    const ptipo = parentTipo[n.tipo]
    if (!ptipo) continue
    edges.push({
      source: `${ptipo}:${n.parentId}`,
      target: `${n.tipo}:${n.id}`,
    })
  }

  return { nodes, edges }
}

/** Cria ou atualiza um nó (descrição/abreviação/sequência). */
export async function upsertNode(input: {
  tipo: NodeType
  id?: number
  parentId?: number
  descricao: string
  abreviacao?: string
  sequencia?: number
}) {
  const validated = NodeUpsertSchema.safeParse(input)
  if (!validated.success) {
    return { error: 'Dados inválidos.', issues: validated.error.flatten().fieldErrors }
  }
  const { tipo, id, parentId, descricao, abreviacao, sequencia } = validated.data

  switch (tipo) {
    case 'cadeia': {
      if (id) {
        await prisma.cadeiaValor.update({ where: { id }, data: { descricao, abreviacao: abreviacao || null } })
      } else {
        await prisma.cadeiaValor.create({ data: { descricao, abreviacao: abreviacao || null } })
      }
      break
    }
    case 'macroprocesso': {
      if (id) {
        await prisma.megaProcesso.update({
          where: { id },
          data: { descricao, abreviacao: abreviacao || null, ...(parentId ? { cadeiaValorId: parentId } : {}) },
        })
      } else {
        if (!parentId) return { error: 'Macroprocesso requer uma Cadeia de Valor pai.' }
        await prisma.megaProcesso.create({
          data: { descricao, abreviacao: abreviacao || null, cadeiaValorId: parentId },
        })
      }
      break
    }
    case 'processo': {
      if (id) {
        await prisma.processo.update({
          where: { id },
          data: { descricao, sequencia, ...(parentId ? { megaProcessoId: parentId } : {}) },
        })
      } else {
        if (!parentId) return { error: 'Processo requer um Macroprocesso pai.' }
        await prisma.processo.create({
          data: { descricao, sequencia, megaProcessoId: parentId },
        })
      }
      break
    }
    case 'macroatividade': {
      if (id) {
        await prisma.subProcesso.update({
          where: { id },
          data: { descricao, sequencia, ...(parentId ? { processoId: parentId } : {}) },
        })
      } else {
        if (!parentId) return { error: 'Macroatividade requer um Processo pai.' }
        const proc = await prisma.processo.findUnique({
          where: { id: parentId },
          select: { megaProcessoId: true },
        })
        if (!proc) return { error: 'Processo pai não encontrado.' }
        await prisma.subProcesso.create({
          data: {
            descricao,
            sequencia,
            processoId: parentId,
            megaProcessoId: proc.megaProcessoId,
          },
        })
      }
      break
    }
    case 'atividade': {
      if (id) {
        await prisma.atividade.update({
          where: { id },
          data: { descricao, sequencia, ...(parentId ? { subProcessoId: parentId } : {}) },
        })
      } else {
        if (!parentId) return { error: 'Atividade requer uma Macroatividade pai.' }
        await prisma.atividade.create({
          data: { descricao, sequencia, subProcessoId: parentId },
        })
      }
      break
    }
  }

  revalidateAll()
  return { success: true }
}

/** Atualiza posição (drag) — chamada com debounce no client. */
export async function updateNodePosition(input: { tipo: NodeType; id: number; posicaoX: number; posicaoY: number }) {
  const validated = NodePositionSchema.safeParse(input)
  if (!validated.success) return { error: 'Posição inválida.' }
  const { tipo, id, posicaoX, posicaoY } = validated.data

  const map: Record<NodeType, () => Promise<unknown>> = {
    cadeia:         () => prisma.cadeiaValor.update({ where: { id }, data: { posicaoX, posicaoY } }),
    macroprocesso:  () => prisma.megaProcesso.update({ where: { id }, data: { posicaoX, posicaoY } }),
    processo:       () => prisma.processo.update({ where: { id }, data: { posicaoX, posicaoY } }),
    macroatividade: () => prisma.subProcesso.update({ where: { id }, data: { posicaoX, posicaoY } }),
    atividade:      () => prisma.atividade.update({ where: { id }, data: { posicaoX, posicaoY } }),
  }
  await map[tipo]()
  // posição é meta-dado de UI — não precisa invalidar telas legadas.
  revalidatePath('/dashboard/mapa')
  return { success: true }
}

/** Exclui um nó (cascade do Postgres cuida dos filhos onde definido). */
export async function deleteNode(input: { tipo: NodeType; id: number }) {
  if (!NODE_TYPES.includes(input.tipo)) return { error: 'Tipo inválido.' }
  const { tipo, id } = input

  switch (tipo) {
    case 'cadeia':         await prisma.cadeiaValor.delete({ where: { id } }); break
    case 'macroprocesso':  await prisma.megaProcesso.delete({ where: { id } }); break
    case 'processo':       await prisma.processo.delete({ where: { id } }); break
    case 'macroatividade': await prisma.subProcesso.delete({ where: { id } }); break
    case 'atividade':      await prisma.atividade.delete({ where: { id } }); break
  }

  revalidateAll()
  return { success: true }
}
