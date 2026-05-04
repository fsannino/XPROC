import { NextRequest, NextResponse } from 'next/server'
import { prisma } from '@/lib/prisma'
import { getSession } from '@/lib/session'
import { cookies } from 'next/headers'
import { decrypt } from '@/lib/session'

function toCSV(rows: string[][]): string {
  return rows.map((r) => r.map((c) => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')).join('\n')
}

export async function GET(req: NextRequest) {
  const cookieStore = await cookies()
  const token = cookieStore.get('session')?.value
  const session = token ? await decrypt(token) : null
  if (!session) return NextResponse.json({ error: 'Não autorizado' }, { status: 401 })

  const tipo = req.nextUrl.searchParams.get('tipo')

  let csv = ''
  let filename = 'export.csv'

  if (tipo === 'processos') {
    const data = await prisma.megaProcesso.findMany({
      orderBy: { id: 'asc' },
      include: { _count: { select: { processos: true } } },
    })
    csv = toCSV([
      ['ID', 'Descrição', 'Abreviação', 'Processos', 'Bloqueado'],
      ...data.map((d) => [String(d.id), d.descricao, d.abreviacao ?? '', String(d._count.processos), d.bloqueado ? 'Sim' : 'Não']),
    ])
    filename = 'processos.csv'
  } else if (tipo === 'transacoes') {
    const data = await prisma.transacao.findMany({ orderBy: { id: 'asc' } })
    csv = toCSV([
      ['Código', 'Descrição'],
      ...data.map((d) => [d.id, d.descricao ?? '']),
    ])
    filename = 'transacoes.csv'
  } else if (tipo === 'usuarios') {
    const data = await prisma.usuario.findMany({ orderBy: { nome: 'asc' } })
    csv = toCSV([
      ['Código', 'Nome', 'Email', 'Categoria', 'Ativo'],
      ...data.map((d) => [d.codigo, d.nome, d.email ?? '', d.categoria ?? '', d.ativo ? 'Sim' : 'Não']),
    ])
    filename = 'usuarios.csv'
  } else {
    return NextResponse.json({ error: 'Tipo inválido' }, { status: 400 })
  }

  return new NextResponse('﻿' + csv, {
    headers: {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': `attachment; filename="${filename}"`,
    },
  })
}
