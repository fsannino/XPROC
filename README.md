# XPROC — Sistema de Gerenciamento de Processos Corporativos

Sistema legado de BPM (Business Process Management) desenvolvido em ASP Clássico/VBScript com SQL Server.

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
