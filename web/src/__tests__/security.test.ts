import { describe, it, expect, beforeEach, vi } from 'vitest'
import { encrypt, decrypt, type SessionPayload } from '@/lib/session'

const samplePayload: SessionPayload = {
  userId: 'u1',
  codigo: 'ADMIN',
  nome: 'Admin',
  categoria: 'A',
  expiresAt: new Date('2030-01-01T00:00:00Z'),
}

describe('session.encrypt / decrypt', () => {
  it('faz round-trip preservando os campos', async () => {
    const token = await encrypt(samplePayload)
    const decoded = await decrypt(token)
    expect(decoded).not.toBeNull()
    expect(decoded?.userId).toBe('u1')
    expect(decoded?.codigo).toBe('ADMIN')
    expect(decoded?.categoria).toBe('A')
  })

  it('retorna null para token malformado', async () => {
    const decoded = await decrypt('isto-não-é-um-jwt-válido')
    expect(decoded).toBeNull()
  })

  it('retorna null para token assinado com outra chave', async () => {
    // jose com chave diferente: simula tampering
    const { SignJWT } = await import('jose')
    const otherKey = new TextEncoder().encode('chave-diferente-de-pelo-menos-32-bytes-aaaaa')
    const forged = await new SignJWT({ userId: 'mallory', codigo: 'X', nome: 'X', categoria: 'A' })
      .setProtectedHeader({ alg: 'HS256' })
      .setIssuedAt()
      .setExpirationTime('8h')
      .sign(otherKey)
    const decoded = await decrypt(forged)
    expect(decoded).toBeNull()
  })
})

describe('session module — fail-fast no NEXTAUTH_SECRET', () => {
  beforeEach(() => {
    vi.resetModules()
  })

  it('lança erro se NEXTAUTH_SECRET ausente', async () => {
    const original = process.env.NEXTAUTH_SECRET
    delete process.env.NEXTAUTH_SECRET
    await expect(import('@/lib/session')).rejects.toThrow(/NEXTAUTH_SECRET/)
    process.env.NEXTAUTH_SECRET = original
  })

  it('lança erro se NEXTAUTH_SECRET tiver < 32 chars', async () => {
    const original = process.env.NEXTAUTH_SECRET
    process.env.NEXTAUTH_SECRET = 'curto'
    await expect(import('@/lib/session')).rejects.toThrow(/32/)
    process.env.NEXTAUTH_SECRET = original
  })
})

describe('proxy.matcher — cobertura de rotas', () => {
  // Réplica do regex em web/src/proxy.ts.
  // Garante que /dashboard é coberto e /api é excluído.
  const matcher = /^\/((?!api|_next\/static|_next\/image|.*\.png$).*)/

  it.each([
    ['/dashboard', true],
    ['/dashboard/usuarios', true],
    ['/dashboard/mapa', true],
    ['/login', true],
    ['/', true],
  ])('cobre %s', (path, expected) => {
    expect(matcher.test(path)).toBe(expected)
  })

  it.each([
    ['/api/v1/processos', false],
    ['/api/cron/cleanup', false],
    ['/api/bpmn/123', false],
    ['/_next/static/chunk.js', false],
    ['/_next/image/foo', false],
    ['/logo.png', false],
  ])('exclui %s', (path, expected) => {
    expect(matcher.test(path)).toBe(expected)
  })
})
