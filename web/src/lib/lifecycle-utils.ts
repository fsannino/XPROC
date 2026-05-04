import { LIFECYCLE_STATUSES } from '@/lib/definitions'

export const STATUS_TRANSITIONS: Record<string, string[]> = {
  Rascunho: ['EmRevisao', 'Arquivado'],
  EmRevisao: ['Aprovado', 'Rascunho'],
  Aprovado: ['Publicado', 'EmRevisao'],
  Publicado: ['Arquivado', 'EmRevisao'],
  Arquivado: ['Rascunho'],
}

export function proximosStatus(atual: string): string[] {
  return STATUS_TRANSITIONS[atual] ?? LIFECYCLE_STATUSES.slice()
}
