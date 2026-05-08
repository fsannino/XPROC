'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { ProdutoSchema, SetProdutosDoProcessoSchema, type ProdutoTipo } from '@/lib/definitions'

function revalidateMapa() {
  revalidatePath('/dashboard/mapa')
  revalidatePath('/dashboard/catalogo')
}

export type ProdutoView = {
  id: number
  codigo: string
  descricao: string
  tipo: ProdutoTipo
}

export async function listarProdutos(): Promise<ProdutoView[]> {
  const rows = await prisma.produto.findMany({ orderBy: { descricao: 'asc' } })
  return rows.map((p) => ({
    id: p.id,
    codigo: p.codigo,
    descricao: p.descricao,
    tipo: p.tipo as ProdutoTipo,
  }))
}

export async function criarProduto(input: {
  codigo: string
  descricao: string
  tipo: ProdutoTipo
}) {
  const v = ProdutoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const p = await prisma.produto.create({ data: v.data })
    revalidateMapa()
    return { success: true, id: p.id }
  } catch {
    return { error: 'Não foi possível criar (código já existe?).' }
  }
}

export async function atualizarProduto(id: number, input: {
  codigo: string
  descricao: string
  tipo: ProdutoTipo
}) {
  const v = ProdutoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    await prisma.produto.update({ where: { id }, data: v.data })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirProduto(id: number) {
  try {
    await prisma.produto.delete({ where: { id } })
    revalidateMapa()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir (em uso?).' }
  }
}

export async function getProdutosDoProcesso(processoId: number): Promise<ProdutoView[]> {
  const rows = await prisma.processoProduto.findMany({
    where: { processoId },
    include: { produto: true },
    orderBy: { produto: { descricao: 'asc' } },
  })
  return rows.map((r) => ({
    id: r.produto.id,
    codigo: r.produto.codigo,
    descricao: r.produto.descricao,
    tipo: r.produto.tipo as ProdutoTipo,
  }))
}

export async function setProdutosDoProcesso(input: { processoId: number; produtoIds: number[] }) {
  const v = SetProdutosDoProcessoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  const { processoId, produtoIds } = v.data
  await prisma.$transaction([
    prisma.processoProduto.deleteMany({ where: { processoId } }),
    ...(produtoIds.length > 0
      ? [prisma.processoProduto.createMany({
          data: produtoIds.map((produtoId) => ({ processoId, produtoId })),
          skipDuplicates: true,
        })]
      : []),
  ])
  revalidateMapa()
  return { success: true }
}

export async function criarProdutoForm(_prev: unknown, fd: FormData) {
  return criarProduto({
    codigo: String(fd.get('codigo') ?? ''),
    descricao: String(fd.get('descricao') ?? ''),
    tipo: (String(fd.get('tipo') ?? 'BEM') as ProdutoTipo),
  })
}
