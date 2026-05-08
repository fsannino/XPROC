-- Mapa Visual: Cadeia de Valor + Atividade + posicaoX/Y em todos os níveis
-- Migration aditiva: zero rename, zero break em telas existentes.

CREATE TABLE IF NOT EXISTS "cadeia_valor" (
  "id" SERIAL PRIMARY KEY,
  "descricao" TEXT NOT NULL,
  "abreviacao" TEXT,
  "posicao_x" DOUBLE PRECISION NOT NULL DEFAULT 0,
  "posicao_y" DOUBLE PRECISION NOT NULL DEFAULT 0,
  "criado_em" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "atualizado_em" TIMESTAMP(3) NOT NULL
);

ALTER TABLE "mega_processo"
  ADD COLUMN IF NOT EXISTS "cadeia_valor_id" INTEGER,
  ADD COLUMN IF NOT EXISTS "posicao_x" DOUBLE PRECISION NOT NULL DEFAULT 0,
  ADD COLUMN IF NOT EXISTS "posicao_y" DOUBLE PRECISION NOT NULL DEFAULT 0;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.table_constraints
    WHERE constraint_name = 'mega_processo_cadeia_valor_id_fkey'
  ) THEN
    ALTER TABLE "mega_processo"
      ADD CONSTRAINT "mega_processo_cadeia_valor_id_fkey"
      FOREIGN KEY ("cadeia_valor_id") REFERENCES "cadeia_valor"("id")
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS "mega_processo_cadeia_valor_id_idx"
  ON "mega_processo"("cadeia_valor_id");

ALTER TABLE "processo"
  ADD COLUMN IF NOT EXISTS "posicao_x" DOUBLE PRECISION NOT NULL DEFAULT 0,
  ADD COLUMN IF NOT EXISTS "posicao_y" DOUBLE PRECISION NOT NULL DEFAULT 0;

ALTER TABLE "sub_processo"
  ADD COLUMN IF NOT EXISTS "posicao_x" DOUBLE PRECISION NOT NULL DEFAULT 0,
  ADD COLUMN IF NOT EXISTS "posicao_y" DOUBLE PRECISION NOT NULL DEFAULT 0;

CREATE TABLE IF NOT EXISTS "atividade" (
  "id" SERIAL PRIMARY KEY,
  "sub_processo_id" INTEGER NOT NULL,
  "descricao" TEXT NOT NULL,
  "sequencia" INTEGER,
  "posicao_x" DOUBLE PRECISION NOT NULL DEFAULT 0,
  "posicao_y" DOUBLE PRECISION NOT NULL DEFAULT 0,
  "criado_em" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "atualizado_em" TIMESTAMP(3) NOT NULL
);

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.table_constraints
    WHERE constraint_name = 'atividade_sub_processo_id_fkey'
  ) THEN
    ALTER TABLE "atividade"
      ADD CONSTRAINT "atividade_sub_processo_id_fkey"
      FOREIGN KEY ("sub_processo_id") REFERENCES "sub_processo"("id")
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS "atividade_sub_processo_id_idx"
  ON "atividade"("sub_processo_id");
