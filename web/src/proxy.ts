// Next.js 16 — `proxy` é o novo `middleware` (renomeado em v16, file convention).
// Doc: node_modules/next/dist/docs/01-app/03-api-reference/03-file-conventions/proxy.md
//
// Roda no Edge Runtime: NÃO importar Prisma aqui.
import { NextRequest, NextResponse } from 'next/server'
import { decrypt } from '@/lib/session'

const publicRoutes = ['/login']

// Rotas que exigem categoria 'A' (admin)
const adminRoutes = ['/dashboard/usuarios', '/dashboard/modulos', '/dashboard/equipe', '/dashboard/catalogo']

export async function proxy(req: NextRequest) {
  const path = req.nextUrl.pathname
  const isPublic = publicRoutes.includes(path)

  const token = req.cookies.get('session')?.value
  const session = token ? await decrypt(token) : null

  // Token presente mas inválido/expirado: limpa o cookie e manda pro login.
  // Sem isso, o cookie persiste e gera loop de redirect a cada request.
  if (token && !session && !isPublic) {
    const response = NextResponse.redirect(new URL('/login', req.nextUrl))
    response.cookies.delete('session')
    return response
  }

  if (!session && !isPublic) {
    return NextResponse.redirect(new URL('/login', req.nextUrl))
  }

  if (session && isPublic) {
    return NextResponse.redirect(new URL('/dashboard', req.nextUrl))
  }

  if (session && adminRoutes.some((r) => path.startsWith(r)) && session.categoria !== 'A') {
    return NextResponse.redirect(new URL('/dashboard?acesso=negado', req.nextUrl))
  }

  return NextResponse.next()
}

export const config = {
  matcher: ['/((?!api|_next/static|_next/image|.*\\.png$).*)'],
}
