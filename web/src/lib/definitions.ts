import { z } from 'zod'

export const LoginSchema = z.object({
  codigo: z.string().min(1, { message: 'Código é obrigatório.' }).trim(),
  senha: z.string().min(1, { message: 'Senha é obrigatória.' }),
})

export const UsuarioSchema = z.object({
  codigo: z.string().min(2).max(10).trim(),
  nome: z.string().min(2).max(80).trim(),
  email: z.string().email({ message: 'Email inválido.' }).optional().or(z.literal('')),
  senha: z.string().min(6, { message: 'Senha deve ter pelo menos 6 caracteres.' }),
  categoria: z.string().max(1).optional(),
})

export const MegaProcessoSchema = z.object({
  descricao: z.string().min(2).max(80).trim(),
  abreviacao: z.string().max(4).optional().or(z.literal('')),
  descricaoLonga: z.string().optional().or(z.literal('')),
  bloqueado: z.boolean().optional(),
  responsavelId: z.string().optional().or(z.literal('')),
})

export const ProcessoSchema = z.object({
  megaProcessoId: z.number().int().positive(),
  descricao: z.string().min(2).max(150).trim(),
  sequencia: z.number().int().optional(),
  tempoMedioCiclo: z.number().positive().optional(),
  custoEstimado: z.number().positive().optional(),
  volumeMensal: z.number().int().positive().optional(),
})

export const SubProcessoSchema = z.object({
  processoId: z.number().int().positive(),
  megaProcessoId: z.number().int().positive(),
  descricao: z.string().min(2).max(150).trim(),
  sequencia: z.number().int().optional(),
})

export const TransacaoSchema = z.object({
  id: z.string().min(1).max(30).trim(),
  descricao: z.string().max(150).optional().or(z.literal('')),
})

export const EmpresaSchema = z.object({
  nome: z.string().min(2).max(150).trim(),
})

export const CenarioSchema = z.object({
  descricao: z.string().min(2).max(150).trim(),
  situacao: z.string().max(20).optional().or(z.literal('')),
})

export const TrocaSenhaSchema = z.object({
  senhaAtual: z.string().min(1, 'Senha atual é obrigatória.'),
  novaSenha: z.string().min(6, 'Nova senha deve ter pelo menos 6 caracteres.'),
  confirmar: z.string().min(1),
}).refine((d) => d.novaSenha === d.confirmar, {
  message: 'As senhas não coincidem.',
  path: ['confirmar'],
})

export const LIFECYCLE_STATUSES = ['Rascunho', 'EmRevisao', 'Aprovado', 'Publicado', 'Arquivado'] as const
export type LifecycleStatus = typeof LIFECYCLE_STATUSES[number]

export const AlterarStatusSchema = z.object({
  megaProcessoId: z.number().int().positive(),
  status: z.enum(LIFECYCLE_STATUSES),
})

export const ComentarioSchema = z.object({
  megaProcessoId: z.number().int().positive(),
  texto: z.string().min(1).max(2000).trim(),
  parentId: z.string().optional(),
})

export const RiscoSchema = z.object({
  megaProcessoId: z.number().int().positive(),
  descricao: z.string().min(2).max(1000).trim(),
  probabilidade: z.enum(['A', 'M', 'B']).default('M'),
  impacto: z.enum(['A', 'M', 'B']).default('M'),
  controle: z.string().max(1000).optional().or(z.literal('')),
})

export const KpiSchema = z.object({
  tempoMedioCiclo: z.number().positive().optional(),
  custoEstimado: z.number().positive().optional(),
  volumeMensal: z.number().int().positive().optional(),
})

// ─── Mapa Visual ───────────────────────────────────────────────────

export const NODE_TYPES = [
  'cadeia',
  'macroprocesso',
  'processo',
  'macroatividade',
  'atividade',
] as const
export type NodeType = typeof NODE_TYPES[number]

export const MAPA_LEVELS: Record<NodeType, { label: string; color: string; parent: NodeType | null }> = {
  cadeia:         { label: 'Cadeia de Valor', color: 'navy',       parent: null },
  macroprocesso:  { label: 'Macroprocesso',   color: 'teal',       parent: 'cadeia' },
  processo:       { label: 'Processo',        color: 'teal-light', parent: 'macroprocesso' },
  macroatividade: { label: 'Macroatividade',  color: 'gold',       parent: 'processo' },
  atividade:      { label: 'Atividade',       color: 'cream',      parent: 'macroatividade' },
}

export const NodeUpsertSchema = z.object({
  tipo: z.enum(NODE_TYPES),
  id: z.number().int().positive().optional(),
  parentId: z.number().int().positive().optional(),
  descricao: z.string().min(2).max(200).trim(),
  abreviacao: z.string().max(8).optional().or(z.literal('')),
  sequencia: z.number().int().optional(),
  // KPIs (somente Processo)
  tempoMedioCiclo: z.number().positive().optional(),
  custoEstimado: z.number().positive().optional(),
  volumeMensal: z.number().int().positive().optional(),
})

export const NodePositionSchema = z.object({
  tipo: z.enum(NODE_TYPES),
  id: z.number().int().positive(),
  posicaoX: z.number(),
  posicaoY: z.number(),
})

export type FormState<T = Record<string, string[]>> =
  | { errors?: Partial<Record<keyof T, string[]>>; message?: string }
  | undefined

// ─── RACI ──────────────────────────────────────────────────────────

export const AreaSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  descricao: z.string().min(2).max(150).trim(),
  parentId: z.number().int().positive().optional().nullable(),
})

export const FuncaoSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  descricao: z.string().min(2).max(150).trim(),
  areaId: z.number().int().positive().optional().nullable(),
})

export const PessoaSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  nome: z.string().min(2).max(150).trim(),
  email: z.string().email({ message: 'Email inválido.' }).optional().or(z.literal('')),
  areaId: z.number().int().positive().optional().nullable(),
  funcaoId: z.number().int().positive().optional().nullable(),
  usuarioId: z.string().optional().nullable(),
})

export const RACI_PAPEIS = ['R', 'A', 'C', 'I'] as const
export type RaciPapel = typeof RACI_PAPEIS[number]

export const RACI_PAPEL_LABEL: Record<RaciPapel, string> = {
  R: 'Responsável',
  A: 'Aprovador',
  C: 'Consultado',
  I: 'Informado',
}

export const RaciAtribuicaoSchema = z.object({
  pessoaId: z.number().int().positive(),
  papel: z.enum(RACI_PAPEIS),
})

export const SetRaciDoProcessoSchema = z.object({
  processoId: z.number().int().positive(),
  atribuicoes: z.array(RaciAtribuicaoSchema),
})

// ─── Atributos do Mapa: Produtos / Insumos / Sistemas / Dependencias ──

export const PRODUTO_TIPOS = ['BEM', 'SERVICO', 'INFORMACAO'] as const
export type ProdutoTipo = typeof PRODUTO_TIPOS[number]

export const PRODUTO_TIPO_LABEL: Record<ProdutoTipo, string> = {
  BEM: 'Bem',
  SERVICO: 'Serviço',
  INFORMACAO: 'Informação',
}

export const ProdutoSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  descricao: z.string().min(2).max(150).trim(),
  tipo: z.enum(PRODUTO_TIPOS).default('BEM'),
})

export const SetProdutosDoProcessoSchema = z.object({
  processoId: z.number().int().positive(),
  produtoIds: z.array(z.number().int().positive()),
})

export const INSUMO_TIPOS = ['DADO', 'DOCUMENTO', 'MATERIAL', 'SERVICO'] as const
export type InsumoTipo = typeof INSUMO_TIPOS[number]

export const INSUMO_TIPO_LABEL: Record<InsumoTipo, string> = {
  DADO: 'Dado',
  DOCUMENTO: 'Documento',
  MATERIAL: 'Material',
  SERVICO: 'Serviço',
}

export const INSUMO_DIRECOES = ['INPUT', 'OUTPUT'] as const
export type InsumoDirecao = typeof INSUMO_DIRECOES[number]

export const InsumoSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  descricao: z.string().min(2).max(150).trim(),
  tipo: z.enum(INSUMO_TIPOS).default('DADO'),
})

export const InsumoVinculoSchema = z.object({
  insumoId: z.number().int().positive(),
  direcao: z.enum(INSUMO_DIRECOES),
})

export const SetInsumosDeProcessoSchema = z.object({
  processoId: z.number().int().positive(),
  vinculos: z.array(InsumoVinculoSchema),
})

export const SetInsumosDeAtividadeSchema = z.object({
  atividadeId: z.number().int().positive(),
  vinculos: z.array(InsumoVinculoSchema),
})

export const SISTEMA_TIPOS = ['ERP', 'CRM', 'BI', 'RPA', 'OUTRO'] as const
export type SistemaTipo = typeof SISTEMA_TIPOS[number]

export const SistemaSchema = z.object({
  codigo: z.string().min(1).max(20).trim(),
  nome: z.string().min(2).max(150).trim(),
  tipo: z.enum(SISTEMA_TIPOS).default('OUTRO'),
})

export const SISTEMA_PAPEIS = ['CONSUMIDOR', 'PRODUTOR', 'AMBOS'] as const
export type SistemaPapel = typeof SISTEMA_PAPEIS[number]

export const SISTEMA_PAPEL_LABEL: Record<SistemaPapel, string> = {
  CONSUMIDOR: 'Consumidor',
  PRODUTOR: 'Produtor',
  AMBOS: 'Ambos',
}

export const SistemaVinculoSchema = z.object({
  sistemaId: z.number().int().positive(),
  papel: z.enum(SISTEMA_PAPEIS),
})

export const SetSistemasDoProcessoSchema = z.object({
  processoId: z.number().int().positive(),
  vinculos: z.array(SistemaVinculoSchema),
})

export const DEPENDENCIA_TIPOS = ['PRECEDE', 'REQUER', 'COMPARTILHA'] as const
export type DependenciaTipo = typeof DEPENDENCIA_TIPOS[number]

export const DEPENDENCIA_TIPO_LABEL: Record<DependenciaTipo, string> = {
  PRECEDE: 'Precede',
  REQUER: 'Requer',
  COMPARTILHA: 'Compartilha',
}

export const DependenciaSchema = z
  .object({
    origemId: z.number().int().positive(),
    destinoId: z.number().int().positive(),
    tipo: z.enum(DEPENDENCIA_TIPOS).default('PRECEDE'),
    descricao: z.string().max(300).optional().or(z.literal('')),
  })
  .refine((d) => d.origemId !== d.destinoId, {
    message: 'Origem e destino devem ser diferentes.',
    path: ['destinoId'],
  })
