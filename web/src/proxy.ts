import { NextRequest, NextResponse } from 'next/server'
import { decrypt } from '@/lib/session'
import { cookies } from 'next/headers'

const publicRoutes = ['/login']

export async function proxy(req: NextRequest) {
  const path = req.nextUrl.pathname
  const isPublic = publicRoutes.includes(path)

  const cookieStore = await cookies()
  const token = cookieStore.get('session')?.value
  const session = token ? await decrypt(token) : null

  if (!session && !isPublic) {
    return NextResponse.redirect(new URL('/login', req.nextUrl))
  }

  if (session && isPublic) {
    return NextResponse.redirect(new URL('/dashboard', req.nextUrl))
  }

  return NextResponse.next()
}

export const config = {
  matcher: ['/((?!api|_next/static|_next/image|.*\\.png$).*)'],
}
