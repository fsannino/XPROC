// Variáveis exigidas no module load de session.ts e prisma.ts.
// Em testes não vamos chamar Prisma; basta DATABASE_URL existir como string.
process.env.NEXTAUTH_SECRET ||= 'test-secret-pelo-menos-32-bytes-de-comprimento-1234567890'
process.env.DATABASE_URL ||= 'postgresql://test:test@localhost:5432/test'
