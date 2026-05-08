import 'server-only'
import { SignJWT, jwtVerify } from 'jose'
import { cookies } from 'next/headers'
import { prisma } from '@/lib/prisma'

// Fail-fast: token forjável é pior que build quebrado.
// NEXTAUTH_SECRET precisa estar definido em todos os ambientes (inclusive testes).
const secret = process.env.NEXTAUTH_SECRET
if (!secret) {
  throw new Error(
    'NEXTAUTH_SECRET não está definido. Gere com: openssl rand -base64 64'
  )
}
if (secret.length < 32) {
  throw new Error(
    `NEXTAUTH_SECRET tem apenas ${secret.length} chars; mínimo 32. Gere com: openssl rand -base64 64`
  )
}

const key = new TextEncoder().encode(secret)

export type SessionPayload = {
  userId: string
  codigo: string
  nome: string
  categoria: string | null
  expiresAt: Date
}

export async function encrypt(payload: SessionPayload) {
  return new SignJWT(payload as Record<string, unknown>)
    .setProtectedHeader({ alg: 'HS256' })
    .setIssuedAt()
    .setExpirationTime('8h')
    .sign(key)
}

export async function decrypt(token: string): Promise<SessionPayload | null> {
  try {
    const { payload } = await jwtVerify(token, key, { algorithms: ['HS256'] })
    return payload as unknown as SessionPayload
  } catch {
    return null
  }
}

export async function createSession(user: Omit<SessionPayload, 'expiresAt'>) {
  const expiresAt = new Date(Date.now() + 8 * 60 * 60 * 1000)
  const session = await encrypt({ ...user, expiresAt })
  const cookieStore = await cookies()
  cookieStore.set('session', session, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    expires: expiresAt,
    sameSite: 'lax',
    path: '/',
  })
}

export async function deleteSession() {
  const cookieStore = await cookies()
  cookieStore.delete('session')
}

export async function getSession(): Promise<SessionPayload | null> {
  const cookieStore = await cookies()
  const token = cookieStore.get('session')?.value
  if (!token) return null
  const payload = await decrypt(token)
  if (!payload) return null

  const user = await prisma.usuario.findUnique({
    where: { id: payload.userId, ativo: true },
    select: { id: true },
  })
  if (!user) return null

  return payload
}
