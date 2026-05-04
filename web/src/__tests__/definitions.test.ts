import { describe, it, expect } from 'vitest'
import {
  LoginSchema, UsuarioSchema, MegaProcessoSchema, CenarioSchema, TrocaSenhaSchema,
  AlterarStatusSchema, RiscoSchema, ComentarioSchema,
} from '@/lib/definitions'

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

describe('AlterarStatusSchema', () => {
  it('aceita status válido', () => {
    const r = AlterarStatusSchema.safeParse({ megaProcessoId: 1, status: 'Publicado' })
    expect(r.success).toBe(true)
  })
  it('rejeita status inválido', () => {
    const r = AlterarStatusSchema.safeParse({ megaProcessoId: 1, status: 'Invalido' })
    expect(r.success).toBe(false)
  })
})

describe('RiscoSchema', () => {
  it('aceita risco válido', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco de fraude', probabilidade: 'A', impacto: 'M' })
    expect(r.success).toBe(true)
  })
  it('rejeita descrição muito curta', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'R', probabilidade: 'A', impacto: 'M' })
    expect(r.success).toBe(false)
  })
  it('rejeita probabilidade inválida', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco X', probabilidade: 'X', impacto: 'M' })
    expect(r.success).toBe(false)
  })
})

describe('ComentarioSchema', () => {
  it('aceita comentário válido', () => {
    const r = ComentarioSchema.safeParse({ megaProcessoId: 1, texto: 'Ótimo processo!' })
    expect(r.success).toBe(true)
  })
  it('rejeita texto vazio', () => {
    const r = ComentarioSchema.safeParse({ megaProcessoId: 1, texto: '' })
    expect(r.success).toBe(false)
  })
})
