'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { SetRaciDoProcessoSchema, type RaciPapel } from '@/lib/definitions'

export type RaciAtribuicaoView = {
  pessoaId: number
  pessoaNome: string
  pessoaCodigo: string
  funcaoDescricao: string | null
  areaDescricao: string | null
  papel: RaciPapel
}

export async function getRaciDoProcesso(processoId: number): Promise<RaciAtribuicaoView[]> {
  const rows = await prisma.raciAtribuicao.findMany({
    where: { processoId },
    include: {
      pessoa: {
        include: {
          funcao: { select: { descricao: true } },
          area: { select: { descricao: true } },
        },
      },
    },
    orderBy: [{ papel: 'asc' }, { pessoa: { nome: 'asc' } }],
  })
  return rows.map((r) => ({
    pessoaId: r.pessoaId,
    pessoaNome: r.pessoa.nome,
    pessoaCodigo: r.pessoa.codigo,
    funcaoDescricao: r.pessoa.funcao?.descricao ?? null,
    areaDescricao: r.pessoa.area?.descricao ?? null,
    papel: r.papel as RaciPapel,
  }))
}

export async function setRaciDoProcesso(input: {
  processoId: number
  atribuicoes: { pessoaId: number; papel: RaciPapel }[]
}) {
  const v = SetRaciDoProcessoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }

  const { processoId, atribuicoes } = v.data
  // Sincroniza: deleta tudo e recria. Simples e correto p/ N pequeno.
  await prisma.$transaction([
    prisma.raciAtribuicao.deleteMany({ where: { processoId } }),
    ...(atribuicoes.length > 0
      ? [prisma.raciAtribuicao.createMany({
          data: atribuicoes.map(({ pessoaId, papel }) => ({ processoId, pessoaId, papel })),
          skipDuplicates: true,
        })]
      : []),
  ])

  revalidatePath('/dashboard/mapa')
  revalidatePath('/dashboard/processos')
  return { success: true }
}

export async function listarPessoasParaRaci() {
  const rows = await prisma.pessoa.findMany({
    where: { ativo: true },
    select: {
      id: true,
      codigo: true,
      nome: true,
      funcao: { select: { descricao: true } },
      area: { select: { descricao: true } },
    },
    orderBy: { nome: 'asc' },
  })
  return rows.map((p) => ({
    id: p.id,
    codigo: p.codigo,
    nome: p.nome,
    funcao: p.funcao?.descricao ?? null,
    area: p.area?.descricao ?? null,
  }))
}
