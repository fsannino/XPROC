'use server'

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { AreaSchema, FuncaoSchema, PessoaSchema } from '@/lib/definitions'

function revalidateEquipe() {
  revalidatePath('/dashboard/equipe')
  revalidatePath('/dashboard/mapa')
}

// ─── Áreas ───────────────────────────────────────────────────

export async function criarArea(input: {
  codigo: string
  descricao: string
  parentId?: number | null
}) {
  const v = AreaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const a = await prisma.area.create({
      data: { codigo: v.data.codigo, descricao: v.data.descricao, parentId: v.data.parentId ?? null },
    })
    revalidateEquipe()
    return { success: true, id: a.id }
  } catch {
    return { error: 'Não foi possível criar (código já existe?).' }
  }
}

export async function atualizarArea(id: number, input: {
  codigo: string
  descricao: string
  parentId?: number | null
}) {
  const v = AreaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  if (v.data.parentId === id) return { error: 'Área não pode ser pai de si mesma.' }
  try {
    await prisma.area.update({
      where: { id },
      data: { codigo: v.data.codigo, descricao: v.data.descricao, parentId: v.data.parentId ?? null },
    })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirArea(id: number) {
  try {
    await prisma.area.delete({ where: { id } })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir.' }
  }
}

// ─── Funções ──────────────────────────────────────────────────

export async function criarFuncao(input: {
  codigo: string
  descricao: string
  areaId?: number | null
}) {
  const v = FuncaoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const f = await prisma.funcao.create({
      data: { codigo: v.data.codigo, descricao: v.data.descricao, areaId: v.data.areaId ?? null },
    })
    revalidateEquipe()
    return { success: true, id: f.id }
  } catch {
    return { error: 'Não foi possível criar (código já existe?).' }
  }
}

export async function atualizarFuncao(id: number, input: {
  codigo: string
  descricao: string
  areaId?: number | null
}) {
  const v = FuncaoSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    await prisma.funcao.update({
      where: { id },
      data: { codigo: v.data.codigo, descricao: v.data.descricao, areaId: v.data.areaId ?? null },
    })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirFuncao(id: number) {
  try {
    await prisma.funcao.delete({ where: { id } })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir.' }
  }
}

// ─── Pessoas ──────────────────────────────────────────────────

export async function criarPessoa(input: {
  codigo: string
  nome: string
  email?: string
  areaId?: number | null
  funcaoId?: number | null
  usuarioId?: string | null
}) {
  const v = PessoaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    const p = await prisma.pessoa.create({
      data: {
        codigo: v.data.codigo,
        nome: v.data.nome,
        email: v.data.email || null,
        areaId: v.data.areaId ?? null,
        funcaoId: v.data.funcaoId ?? null,
        usuarioId: v.data.usuarioId || null,
      },
    })
    revalidateEquipe()
    return { success: true, id: p.id }
  } catch {
    return { error: 'Não foi possível criar (código/email/usuário já vinculado?).' }
  }
}

export async function atualizarPessoa(id: number, input: {
  codigo: string
  nome: string
  email?: string
  areaId?: number | null
  funcaoId?: number | null
  usuarioId?: string | null
}) {
  const v = PessoaSchema.safeParse(input)
  if (!v.success) return { error: 'Dados inválidos.' }
  try {
    await prisma.pessoa.update({
      where: { id },
      data: {
        codigo: v.data.codigo,
        nome: v.data.nome,
        email: v.data.email || null,
        areaId: v.data.areaId ?? null,
        funcaoId: v.data.funcaoId ?? null,
        usuarioId: v.data.usuarioId || null,
      },
    })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível atualizar.' }
  }
}

export async function excluirPessoa(id: number) {
  try {
    await prisma.pessoa.delete({ where: { id } })
    revalidateEquipe()
    return { success: true }
  } catch {
    return { error: 'Não foi possível excluir.' }
  }
}

// ─── FormData adapters (para useActionState em Server Components) ─

function toIntOrNull(v: FormDataEntryValue | null): number | null {
  if (v == null) return null
  const s = String(v).trim()
  if (!s) return null
  const n = Number(s)
  return Number.isFinite(n) && n > 0 ? n : null
}

export async function criarAreaForm(_prev: unknown, fd: FormData) {
  return criarArea({
    codigo: String(fd.get('codigo') ?? ''),
    descricao: String(fd.get('descricao') ?? ''),
    parentId: toIntOrNull(fd.get('parentId')),
  })
}

export async function criarFuncaoForm(_prev: unknown, fd: FormData) {
  return criarFuncao({
    codigo: String(fd.get('codigo') ?? ''),
    descricao: String(fd.get('descricao') ?? ''),
    areaId: toIntOrNull(fd.get('areaId')),
  })
}

export async function criarPessoaForm(_prev: unknown, fd: FormData) {
  return criarPessoa({
    codigo: String(fd.get('codigo') ?? ''),
    nome: String(fd.get('nome') ?? ''),
    email: String(fd.get('email') ?? ''),
    areaId: toIntOrNull(fd.get('areaId')),
    funcaoId: toIntOrNull(fd.get('funcaoId')),
  })
}
