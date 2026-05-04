import { NextRequest, NextResponse } from 'next/server'
import { prisma } from '@/lib/prisma'

export async function GET(req: NextRequest) {
  const auth = req.headers.get('authorization')
  if (auth !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Não autorizado' }, { status: 401 })
  }

  const limite = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000)
  const { count } = await prisma.usuarioHistorico.deleteMany({
    where: { criadoEm: { lt: limite } },
  })

  return NextResponse.json({ deleted: count, before: limite.toISOString() })
}
