import { NextRequest, NextResponse } from 'next/server'
import { prisma } from '@/lib/prisma'

function authenticate(req: NextRequest) {
  const key = req.headers.get('x-api-key')
  return key && key === process.env.API_KEY
}

export async function GET(req: NextRequest) {
  if (!authenticate(req)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
  }

  const { searchParams } = req.nextUrl
  const status = searchParams.get('status')
  const busca = searchParams.get('busca')

  const where: Record<string, unknown> = {}
  if (status) where.status = status
  if (busca) where.descricao = { contains: busca, mode: 'insensitive' }

  const processos = await prisma.megaProcesso.findMany({
    where,
    orderBy: { id: 'asc' },
    select: {
      id: true,
      descricao: true,
      abreviacao: true,
      status: true,
      bloqueado: true,
      criadoEm: true,
      atualizadoEm: true,
      responsavel: { select: { codigo: true, nome: true, email: true } },
      _count: { select: { processos: true, riscos: true, comentarios: true } },
    },
  })

  return NextResponse.json({ data: processos, total: processos.length })
}
