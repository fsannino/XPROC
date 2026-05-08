import { describe, it, expect } from 'vitest'
import {
  LoginSchema, UsuarioSchema, MegaProcessoSchema, CenarioSchema, TrocaSenhaSchema,
  AlterarStatusSchema, RiscoSchema, ComentarioSchema,
  AreaSchema, FuncaoSchema, PessoaSchema, RaciAtribuicaoSchema, SetRaciDoProcessoSchema,
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
  it('aceita risco válido (escala 1..5)', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco de fraude', probabilidade: 4, impacto: 3 })
    expect(r.success).toBe(true)
  })
  it('coage strings numéricas vindas do form', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco de fraude', probabilidade: '5', impacto: '1' })
    expect(r.success).toBe(true)
    if (r.success) {
      expect(r.data.probabilidade).toBe(5)
      expect(r.data.impacto).toBe(1)
    }
  })
  it('rejeita descrição muito curta', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'R', probabilidade: 4, impacto: 3 })
    expect(r.success).toBe(false)
  })
  it('rejeita probabilidade fora de 1..5', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco X', probabilidade: 6, impacto: 3 })
    expect(r.success).toBe(false)
  })
  it('rejeita impacto zero', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco Y', probabilidade: 3, impacto: 0 })
    expect(r.success).toBe(false)
  })
  it('aplica defaults quando ausentes', () => {
    const r = RiscoSchema.safeParse({ megaProcessoId: 1, descricao: 'Risco padrão' })
    expect(r.success).toBe(true)
    if (r.success) {
      expect(r.data.probabilidade).toBe(3)
      expect(r.data.impacto).toBe(3)
    }
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

describe('AreaSchema', () => {
  it('aceita área válida sem parent', () => {
    const r = AreaSchema.safeParse({ codigo: 'FIN', descricao: 'Financeiro' })
    expect(r.success).toBe(true)
  })
  it('aceita área com parent', () => {
    const r = AreaSchema.safeParse({ codigo: 'FIN-CB', descricao: 'Contas a Receber', parentId: 5 })
    expect(r.success).toBe(true)
  })
  it('rejeita código vazio', () => {
    const r = AreaSchema.safeParse({ codigo: '', descricao: 'X' })
    expect(r.success).toBe(false)
  })
})

describe('FuncaoSchema', () => {
  it('aceita função vinculada a área', () => {
    const r = FuncaoSchema.safeParse({ codigo: 'GER-FIN', descricao: 'Gerente Financeiro', areaId: 1 })
    expect(r.success).toBe(true)
  })
  it('aceita função sem área', () => {
    const r = FuncaoSchema.safeParse({ codigo: 'CEO', descricao: 'Diretor Executivo' })
    expect(r.success).toBe(true)
  })
  it('rejeita descrição muito curta', () => {
    const r = FuncaoSchema.safeParse({ codigo: 'X', descricao: 'A' })
    expect(r.success).toBe(false)
  })
})

describe('PessoaSchema', () => {
  it('aceita pessoa completa', () => {
    const r = PessoaSchema.safeParse({
      codigo: 'JS', nome: 'João Silva', email: 'joao@x.com', areaId: 1, funcaoId: 2,
    })
    expect(r.success).toBe(true)
  })
  it('aceita pessoa sem email', () => {
    const r = PessoaSchema.safeParse({ codigo: 'X', nome: 'Maria' })
    expect(r.success).toBe(true)
  })
  it('rejeita email inválido', () => {
    const r = PessoaSchema.safeParse({ codigo: 'X', nome: 'Maria', email: 'nao-email' })
    expect(r.success).toBe(false)
  })
})

describe('RaciAtribuicaoSchema', () => {
  it.each([['R'], ['A'], ['C'], ['I']])('aceita papel %s', (papel) => {
    const r = RaciAtribuicaoSchema.safeParse({ pessoaId: 1, papel })
    expect(r.success).toBe(true)
  })
  it('rejeita papel inválido', () => {
    const r = RaciAtribuicaoSchema.safeParse({ pessoaId: 1, papel: 'X' })
    expect(r.success).toBe(false)
  })
})

describe('SetRaciDoProcessoSchema', () => {
  it('aceita lista vazia (limpar atribuições)', () => {
    const r = SetRaciDoProcessoSchema.safeParse({ processoId: 1, atribuicoes: [] })
    expect(r.success).toBe(true)
  })
  it('aceita múltiplas atribuições', () => {
    const r = SetRaciDoProcessoSchema.safeParse({
      processoId: 1,
      atribuicoes: [
        { pessoaId: 10, papel: 'R' },
        { pessoaId: 11, papel: 'A' },
      ],
    })
    expect(r.success).toBe(true)
  })
})
