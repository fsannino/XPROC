import "dotenv/config";
import { defineConfig } from "prisma/config";

export default defineConfig({
  schema: "prisma/schema.prisma",
  migrations: {
    path: "prisma/migrations",
  },
  datasource: {
    // DATABASE_URL → Supabase pooler (porta 6543, app em produção)
    // DIRECT_URL   → Supabase direct (porta 5432, migrations)
    url:             process.env["DATABASE_URL"],
  },
});
