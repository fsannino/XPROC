import { describe, it, expect } from 'vitest'
import {
  ProdutoSchema,
  SetProdutosDoProcessoSchema,
  InsumoSchema,
  SetInsumosDeProcessoSchema,
  SetInsumosDeAtividadeSchema,
  SistemaSchema,
  SetSistemasDoProcessoSchema,
  DependenciaSchema,
} from '@/lib/definitions'

describe('ProdutoSchema', () => {
  it('aceita produto válido', () => {
    const r = ProdutoSchema.safeParse({ codigo: 'PROD-1', descricao: 'Relatório anual', tipo: 'INFORMACAO' })
    expect(r.success).toBe(true)
  })
  it('aplica default BEM quando tipo ausente', () => {
    const r = ProdutoSchema.safeParse({ codigo: 'PROD-1', descricao: 'Item' })
    expect(r.success).toBe(true)
    if (r.success) expect(r.data.tipo).toBe('BEM')
  })
  it('rejeita tipo invalido', () => {
    const r = ProdutoSchema.safeParse({ codigo: 'P', descricao: 'X', tipo: 'OUTRO' })
    expect(r.success).toBe(false)
  })
})

describe('SetProdutosDoProcessoSchema', () => {
  it('aceita lista vazia (limpa associacoes)', () => {
    const r = SetProdutosDoProcessoSchema.safeParse({ processoId: 1, produtoIds: [] })
    expect(r.success).toBe(true)
  })
  it('rejeita processoId 0', () => {
    const r = SetProdutosDoProcessoSchema.safeParse({ processoId: 0, produtoIds: [1] })
    expect(r.success).toBe(false)
  })
})

describe('InsumoSchema', () => {
  it('aceita insumo válido', () => {
    const r = InsumoSchema.safeParse({ codigo: 'IN-1', descricao: 'Pedido de venda', tipo: 'DOCUMENTO' })
    expect(r.success).toBe(true)
  })
  it('rejeita tipo invalido', () => {
    const r = InsumoSchema.safeParse({ codigo: 'I', descricao: 'X', tipo: 'INVALIDO' })
    expect(r.success).toBe(false)
  })
})

describe('SetInsumosDeProcessoSchema', () => {
  it('aceita vinculos com direcao INPUT/OUTPUT', () => {
    const r = SetInsumosDeProcessoSchema.safeParse({
      processoId: 1,
      vinculos: [
        { insumoId: 1, direcao: 'INPUT' },
        { insumoId: 2, direcao: 'OUTPUT' },
      ],
    })
    expect(r.success).toBe(true)
  })
  it('rejeita direcao desconhecida', () => {
    const r = SetInsumosDeProcessoSchema.safeParse({
      processoId: 1,
      vinculos: [{ insumoId: 1, direcao: 'INOUT' }],
    })
    expect(r.success).toBe(false)
  })
})

describe('SetInsumosDeAtividadeSchema', () => {
  it('aceita lista vazia', () => {
    const r = SetInsumosDeAtividadeSchema.safeParse({ atividadeId: 1, vinculos: [] })
    expect(r.success).toBe(true)
  })
})

describe('SistemaSchema', () => {
  it('aceita sistema valido', () => {
    const r = SistemaSchema.safeParse({ codigo: 'SAP', nome: 'SAP S/4HANA', tipo: 'ERP' })
    expect(r.success).toBe(true)
  })
  it('aplica default OUTRO', () => {
    const r = SistemaSchema.safeParse({ codigo: 'X', nome: 'Outro' })
    expect(r.success).toBe(true)
    if (r.success) expect(r.data.tipo).toBe('OUTRO')
  })
})

describe('SetSistemasDoProcessoSchema', () => {
  it('aceita papeis CONSUMIDOR/PRODUTOR/AMBOS', () => {
    const r = SetSistemasDoProcessoSchema.safeParse({
      processoId: 1,
      vinculos: [
        { sistemaId: 1, papel: 'CONSUMIDOR' },
        { sistemaId: 2, papel: 'PRODUTOR' },
        { sistemaId: 3, papel: 'AMBOS' },
      ],
    })
    expect(r.success).toBe(true)
  })
  it('rejeita papel invalido', () => {
    const r = SetSistemasDoProcessoSchema.safeParse({
      processoId: 1,
      vinculos: [{ sistemaId: 1, papel: 'OBSERVADOR' }],
    })
    expect(r.success).toBe(false)
  })
})

describe('DependenciaSchema', () => {
  it('aceita aresta dirigida valida', () => {
    const r = DependenciaSchema.safeParse({ origemId: 1, destinoId: 2, tipo: 'PRECEDE' })
    expect(r.success).toBe(true)
  })
  it('rejeita origem == destino (auto-loop)', () => {
    const r = DependenciaSchema.safeParse({ origemId: 1, destinoId: 1, tipo: 'PRECEDE' })
    expect(r.success).toBe(false)
  })
  it('rejeita tipo invalido', () => {
    const r = DependenciaSchema.safeParse({ origemId: 1, destinoId: 2, tipo: 'BLOQUEIA' })
    expect(r.success).toBe(false)
  })
  it('aceita descricao opcional', () => {
    const r = DependenciaSchema.safeParse({
      origemId: 1,
      destinoId: 2,
      tipo: 'REQUER',
      descricao: 'Necessita do output do processo X',
    })
    expect(r.success).toBe(true)
  })
})
