import { describe, it, expect } from 'vitest'
import { LoginSchema, UsuarioSchema, MegaProcessoSchema, CenarioSchema, TrocaSenhaSchema } from '@/lib/definitions'

describe('LoginSchema', () => {
  it('válido com código e senha', () => {
    const r = LoginSchema.safeParse({ codigo: 'ADMIN', senha: '123456' })
    expect(r.success).toBe(true)
  })
  it('falha sem código', () => {
    const r = LoginSchema.safeParse({ codigo: '', senha: '123456' })
    expect(r.success).toBe(false)
  })
})

describe('UsuarioSchema', () => {
  it('aceita usuário válido', () => {
    const r = UsuarioSchema.safeParse({ codigo: 'JOAO', nome: 'João Silva', senha: 'abc123' })
    expect(r.success).toBe(true)
  })
  it('rejeita senha curta', () => {
    const r = UsuarioSchema.safeParse({ codigo: 'JOAO', nome: 'João', senha: '12' })
    expect(r.success).toBe(false)
  })
  it('rejeita email inválido', () => {
    const r = UsuarioSchema.safeParse({ codigo: 'JOAO', nome: 'João', senha: 'abc123', email: 'nao-email' })
    expect(r.success).toBe(false)
  })
})

describe('MegaProcessoSchema', () => {
  it('aceita descrição válida', () => {
    const r = MegaProcessoSchema.safeParse({ descricao: 'Gestão Financeira' })
    expect(r.success).toBe(true)
  })
  it('rejeita descrição muito curta', () => {
    const r = MegaProcessoSchema.safeParse({ descricao: 'A' })
    expect(r.success).toBe(false)
  })
})

describe('CenarioSchema', () => {
  it('aceita cenário com situação', () => {
    const r = CenarioSchema.safeParse({ descricao: 'Cenário Base', situacao: 'Ativo' })
    expect(r.success).toBe(true)
  })
  it('aceita cenário sem situação', () => {
    const r = CenarioSchema.safeParse({ descricao: 'Cenário Base' })
    expect(r.success).toBe(true)
  })
})

describe('TrocaSenhaSchema', () => {
  it('válido quando senhas coincidem', () => {
    const r = TrocaSenhaSchema.safeParse({ senhaAtual: 'antiga', novaSenha: 'nova123', confirmar: 'nova123' })
    expect(r.success).toBe(true)
  })
  it('falha quando senhas não coincidem', () => {
    const r = TrocaSenhaSchema.safeParse({ senhaAtual: 'antiga', novaSenha: 'nova123', confirmar: 'outra456' })
    expect(r.success).toBe(false)
  })
  it('falha quando nova senha é muito curta', () => {
    const r = TrocaSenhaSchema.safeParse({ senhaAtual: 'antiga', novaSenha: '12', confirmar: '12' })
    expect(r.success).toBe(false)
  })
})
