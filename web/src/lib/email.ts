import 'server-only'

let resendClient: import('resend').Resend | null = null

function getResend() {
  if (!process.env.RESEND_API_KEY) return null
  if (!resendClient) {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { Resend } = require('resend')
    resendClient = new Resend(process.env.RESEND_API_KEY)
  }
  return resendClient
}

export async function enviarEmail({
  para,
  assunto,
  html,
}: {
  para: string
  assunto: string
  html: string
}) {
  const resend = getResend()
  if (!resend) return // skip silently if not configured

  const from = process.env.RESEND_FROM || 'XPROC <noreply@xproc.app>'
  await resend.emails.send({ from, to: para, subject: assunto, html })
}

export function htmlMudancaStatus(
  megaProcesso: string,
  statusAnterior: string,
  statusNovo: string,
  autor: string,
  url: string,
) {
  const labelMap: Record<string, string> = {
    Rascunho: 'Rascunho',
    EmRevisao: 'Em Revisão',
    Aprovado: 'Aprovado',
    Publicado: 'Publicado',
    Arquivado: 'Arquivado',
  }
  return `
    <p>O status do mega-processo <strong>${megaProcesso}</strong> foi alterado.</p>
    <p><strong>De:</strong> ${labelMap[statusAnterior] ?? statusAnterior}<br/>
       <strong>Para:</strong> ${labelMap[statusNovo] ?? statusNovo}<br/>
       <strong>Por:</strong> ${autor}</p>
    <p><a href="${url}">Ver no XPROC</a></p>
  `
}
