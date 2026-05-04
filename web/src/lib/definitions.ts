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
})

export const ProcessoSchema = z.object({
  megaProcessoId: z.number().int().positive(),
  descricao: z.string().min(2).max(150).trim(),
  sequencia: z.number().int().optional(),
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

export type FormState<T = Record<string, string[]>> =
  | { errors?: Partial<Record<keyof T, string[]>>; message?: string }
  | undefined
