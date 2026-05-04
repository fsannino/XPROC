import "dotenv/config";
import { defineConfig } from "prisma/config";

export default defineConfig({
  schema: "prisma/schema.prisma",
  migrations: {
    path: "prisma/migrations",
  },
  datasource: {
    // DATABASE_URL  → Supabase pooler (porta 6543, usado pela app em produção)
    // DIRECT_URL    → Supabase direct (porta 5432, usado pelas migrations)
    url:       process.env["DATABASE_URL"],
    directUrl: process.env["DIRECT_URL"],
  },
});
