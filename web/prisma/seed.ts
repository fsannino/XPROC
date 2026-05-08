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

  const cadeia = await prisma.cadeiaValor.upsert({
    where: { id: 1 },
    update: {},
    create: {
      id: 1,
      descricao: 'Cadeia de Valor — Empresa',
      abreviacao: 'CV',
      posicaoX: 0,
      posicaoY: 0,
    },
  })

  const megaProcesso = await prisma.megaProcesso.upsert({
    where: { id: 1 },
    update: { cadeiaValorId: cadeia.id },
    create: {
      id: 1,
      descricao: 'Gestão Financeira',
      abreviacao: 'FI',
      bloqueado: false,
      cadeiaValorId: cadeia.id,
      posicaoX: 0,
      posicaoY: 200,
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
      posicaoX: 0,
      posicaoY: 400,
    },
  })

  const subProcesso = await prisma.subProcesso.upsert({
    where: { id: 1 },
    update: {},
    create: {
      id: 1,
      megaProcessoId: megaProcesso.id,
      processoId: processo.id,
      descricao: 'Lançamento de Nota Fiscal',
      sequencia: 1,
      posicaoX: 0,
      posicaoY: 600,
    },
  })

  // Atividade exemplo (folha)
  const atividadeExistente = await prisma.atividade.findFirst({
    where: { subProcessoId: subProcesso.id, descricao: 'Conferir documentação fiscal' },
  })
  if (!atividadeExistente) {
    await prisma.atividade.create({
      data: {
        subProcessoId: subProcesso.id,
        descricao: 'Conferir documentação fiscal',
        sequencia: 1,
        posicaoX: 0,
        posicaoY: 800,
      },
    })
  }

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
