import { NextRequest, NextResponse } from 'next/server'
import { prisma } from '@/lib/prisma'

function authenticate(req: NextRequest) {
  const key = req.headers.get('x-api-key')
  return key && key === process.env.API_KEY
}

export async function GET(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  if (!authenticate(req)) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
  }

  const { id } = await params
  const mp = await prisma.megaProcesso.findUnique({
    where: { id: Number(id) },
    include: {
      responsavel: { select: { codigo: true, nome: true } },
      processos: {
        orderBy: { sequencia: 'asc' },
        include: {
          subProcessos: {
            orderBy: { sequencia: 'asc' },
            include: {
              atividades: {
                orderBy: { sequencia: 'asc' },
                include: { transacao: { select: { id: true, descricao: true } } },
              },
            },
          },
        },
      },
      riscos: { orderBy: { criadoEm: 'desc' } },
    },
  })

  if (!mp) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return NextResponse.json({ data: mp })
}
