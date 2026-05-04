import { NextRequest, NextResponse } from 'next/server'
import { prisma } from '@/lib/prisma'
import { getSession } from '@/lib/session'

export async function GET(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const session = await getSession()
  if (!session) return NextResponse.json({ error: 'Não autorizado' }, { status: 401 })

  const { id } = await params
  const mp = await prisma.megaProcesso.findUnique({
    where: { id: Number(id) },
    include: {
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
    },
  })

  if (!mp) return NextResponse.json({ error: 'Não encontrado' }, { status: 404 })

  const esc = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

  let xml = `<?xml version="1.0" encoding="UTF-8"?>
<definitions
  xmlns="http://www.omg.org/spec/BPMN/20100524/MODEL"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  targetNamespace="http://xproc.app/bpmn"
  id="def-mp-${mp.id}">
`

  for (const proc of mp.processos) {
    xml += `  <process id="proc-${proc.id}" name="${esc(proc.descricao)}" isExecutable="false">\n`
    xml += `    <laneSet id="ls-${proc.id}">\n`

    for (const sub of proc.subProcessos) {
      xml += `      <lane id="lane-${sub.id}" name="${esc(sub.descricao)}">\n`
      for (const act of sub.atividades) {
        xml += `        <flowNodeRef>task-${act.id}</flowNodeRef>\n`
      }
      xml += `      </lane>\n`
    }

    xml += `    </laneSet>\n`

    xml += `    <startEvent id="start-${proc.id}" name="Início"/>\n`

    let prevId = `start-${proc.id}`
    let taskIdx = 0
    for (const sub of proc.subProcessos) {
      for (const act of sub.atividades) {
        const taskId = `task-${act.id}`
        const label = act.descricao
          ? esc(act.descricao)
          : act.transacao
          ? esc(act.transacao.id)
          : `Atividade ${act.id}`
        xml += `    <userTask id="${taskId}" name="${label}"/>\n`
        xml += `    <sequenceFlow id="sf-${proc.id}-${taskIdx}" sourceRef="${prevId}" targetRef="${taskId}"/>\n`
        prevId = taskId
        taskIdx++
      }
    }

    xml += `    <endEvent id="end-${proc.id}" name="Fim"/>\n`
    xml += `    <sequenceFlow id="sf-${proc.id}-end" sourceRef="${prevId}" targetRef="end-${proc.id}"/>\n`
    xml += `  </process>\n`
  }

  xml += `</definitions>\n`

  return new NextResponse(xml, {
    headers: {
      'Content-Type': 'application/xml; charset=utf-8',
      'Content-Disposition': `attachment; filename="bpmn-${mp.id}.xml"`,
    },
  })
}
