import { PrismaClient } from '@prisma/client'
import bcrypt from 'bcryptjs'

const prisma = new PrismaClient()

async function main() {
  const senhaAdmin = await bcrypt.hash('admin123', 12)

  const admin = await prisma.usuario.upsert({
    where: { codigo: 'ADMIN' },
    update: {},
    create: {
      codigo: 'ADMIN',
      nome: 'Administrador',
      email: 'admin@xproc.local',
      senha: senhaAdmin,
      categoria: 'A',
      ativo: true,
    },
  })

  const megaProcesso = await prisma.megaProcesso.upsert({
    where: { id: 1 },
    update: {},
    create: {
      id: 1,
      descricao: 'Gestão Financeira',
      abreviacao: 'FI',
      bloqueado: false,
    },
  })

  const processo = await prisma.processo.upsert({
    where: { id: 1 },
    update: {},
    create: {
      id: 1,
      megaProcessoId: megaProcesso.id,
      descricao: 'Contas a Pagar',
      sequencia: 1,
    },
  })

  await prisma.subProcesso.upsert({
    where: { id: 1 },
    update: {},
    create: {
      id: 1,
      megaProcessoId: megaProcesso.id,
      processoId: processo.id,
      descricao: 'Lançamento de Nota Fiscal',
      sequencia: 1,
    },
  })

  await prisma.acesso.upsert({
    where: { usuarioId_megaProcessoId: { usuarioId: admin.id, megaProcessoId: megaProcesso.id } },
    update: {},
    create: { usuarioId: admin.id, megaProcessoId: megaProcesso.id },
  })

  console.log('Seed concluído. Login: ADMIN / admin123')
}

main()
  .catch(console.error)
  .finally(() => prisma.$disconnect())
