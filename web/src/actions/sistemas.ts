'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import {
  SistemaSchema,
  SetSistemasDoProcessoSchema,
  type SistemaTipo,
  type SistemaPapel,
} from '@/lib/definitions'

function revalidateMapa() {
  revalidatePath('/dashboard/mapa')
  revalidatePath('/dashboard/catalogo')
}

export type SistemaView = {
  id: number
  codigo: string
  nome: string
  tipo: SistemaTipo
}

export type SistemaVinculoView = SistemaView & { papel: SistemaPapel }

export async function listarSistemas(): Promise<SistemaView[]> {
  const rows = await prisma.sistema.findMany({ orderBy: { nome: 'asc' } })
  return rows.map((s) => ({
    id: s.id,
    codigo: s.codigo,
    nome: s.nome,
    tipo: s.tipo as SistemaTipo,
  }))
}

export async function criarSistema(input: { codigo: string; nome: string; tipo: SistemaTipo }) {
  const v = SistemaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const s = await prisma.sistema.create({ data: v.data })
    revalidateMapa()
    return { success: true, id: s.id }
  } catch {
    return { error: 'Não foi possível criar (código já existe?).' }
  }
}

export async function atualizarSistema(id: number, input: { codigo: string; nome: string; tipo: SistemaTipo }) {
  const v = SistemaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    await prisma.sistema.update({ where: { id }, data: v.data })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirSistema(id: number) {
  try {
    await prisma.sistema.delete({ where: { id } })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir (em uso?).' }
  }
}

export async function getSistemasDoProcesso(processoId: number): Promise<SistemaVinculoView[]> {
  const rows = await prisma.processoSistema.findMany({
    where: { processoId },
    include: { sistema: true },
    orderBy: { sistema: { nome: 'asc' } },
  })
  return rows.map((r) => ({
    id: r.sistema.id,
    codigo: r.sistema.codigo,
    nome: r.sistema.nome,
    tipo: r.sistema.tipo as SistemaTipo,
    papel: r.papel as SistemaPapel,
  }))
}

export async function setSistemasDoProcesso(input: {
  processoId: number
  vinculos: { sistemaId: number; papel: SistemaPapel }[]
}) {
  const v = SetSistemasDoProcessoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { processoId, vinculos } = v.data
  await prisma.$transaction([
    prisma.processoSistema.deleteMany({ where: { processoId } }),
    ...(vinculos.length > 0
      ? [prisma.processoSistema.createMany({
          data: vinculos.map(({ sistemaId, papel }) => ({ processoId, sistemaId, papel })),
          skipDuplicates: true,
        })]
      : []),
  ])
  revalidateMapa()
  return { success: true }
}

export async function criarSistemaForm(_prev: unknown, fd: FormData) {
  return criarSistema({
    codigo: String(fd.get('codigo') ?? ''),
    nome: String(fd.get('nome') ?? ''),
    tipo: (String(fd.get('tipo') ?? 'OUTRO') as SistemaTipo),
  })
}

export async function atualizarSistemaForm(_prev: unknown, fd: FormData) {
  const id = Number(fd.get('id'))
  if (!Number.isFinite(id) || id <= 0) return { error: 'ID inválido.' }
  return atualizarSistema(id, {
    codigo: String(fd.get('codigo') ?? ''),
    nome: String(fd.get('nome') ?? ''),
    tipo: (String(fd.get('tipo') ?? 'OUTRO') as SistemaTipo),
  })
}
