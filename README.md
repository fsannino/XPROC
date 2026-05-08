# Collab:Flow (ex-XPROC) — Sistema de Gerenciamento de Processos Corporativos

Produto **Collab:Flow** da suite **Collab:Engine** da collab:Z. Originalmente conhecido como **XPROC**, o nome legado segue presente no repositório, banco de dados e código por compatibilidade.

Sistema legado de BPM (Business Process Management) desenvolvido em ASP Clássico/VBScript com SQL Server, em modernização para Next.js + TypeScript + Prisma + PostgreSQL.

## Stack Legada
- **Backend:** ASP Clássico (VBScript) / IIS
- **Banco de dados:** Microsoft SQL Server (`cogest`)
- **Frontend:** HTML, CSS, JavaScript, Flash (SWF)
- **Integração:** Lotus Notes

## Stack Moderna (em desenvolvimento)
- **Framework:** Next.js 14 (App Router) + TypeScript
- **ORM:** Prisma
- **Banco de dados:** PostgreSQL
- **Auth:** NextAuth.js
- **UI:** Tailwind CSS

## Estrutura
```
/asp          # Aplicação legada ASP (~1.735 arquivos)
/web          # Nova aplicação modernizada (Next.js)
/doc          # Documentação
/css          # Estilos legados
/js           # Scripts legados
BANCOXPROC.sql     # Schema original SQL Server
BANCOXPROCINDEX.sql # Índices originais
```
