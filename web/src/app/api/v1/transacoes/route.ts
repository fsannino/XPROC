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
  const busca = searchParams.get('busca')

  const where = busca
    ? { OR: [
        { id: { contains: busca, mode: 'insensitive' as const } },
        { descricao: { contains: busca, mode: 'insensitive' as const } },
      ]}
    : undefined

  const transacoes = await prisma.transacao.findMany({
    where,
    orderBy: { id: 'asc' },
    select: {
      id: true,
      descricao: true,
      criadoEm: true,
      _count: { select: { megaProcessos: true } },
    },
  })

  return NextResponse.json({ data: transacoes, total: transacoes.length })
}
