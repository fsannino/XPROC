-- Relações Transação ↔ Processo / Atividade  e  Cenário ↔ Processo / Atividade.
-- Migration aditiva, idempotente, segura para rodar mais de uma vez.

-- Garantir PK em transacao e cenario (algumas instalações legadas não tinham).
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conrelid = 'transacao'::regclass AND contype = 'p') THEN
    ALTER TABLE transacao ADD PRIMARY KEY (tran_cd_transacao);
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conrelid = 'cenario'::regclass AND contype = 'p') THEN
    ALTER TABLE cenario ADD PRIMARY KEY (cena_cd_cenario);
  END IF;
END $$;

CREATE SEQUENCE IF NOT EXISTS cenario_id_seq;
ALTER TABLE cenario ALTER COLUMN cena_cd_cenario SET DEFAULT nextval('cenario_id_seq');
ALTER SEQUENCE cenario_id_seq OWNED BY cenario.cena_cd_cenario;
SELECT setval('cenario_id_seq', GREATEST((SELECT COALESCE(MAX(cena_cd_cenario), 0) FROM cenario), 1));

-- transacao ↔ processo
CREATE TABLE IF NOT EXISTS transacao_processo (
  transacao_id TEXT NOT NULL,
  processo_id  INTEGER NOT NULL,
  criado_em    TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (transacao_id, processo_id)
);
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'transacao_processo_transacao_id_fkey') THEN
    ALTER TABLE transacao_processo
      ADD CONSTRAINT transacao_processo_transacao_id_fkey
      FOREIGN KEY (transacao_id) REFERENCES transacao(tran_cd_transacao)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'transacao_processo_processo_id_fkey') THEN
    ALTER TABLE transacao_processo
      ADD CONSTRAINT transacao_processo_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS transacao_processo_processo_id_idx ON transacao_processo(processo_id);

-- transacao ↔ atividade
CREATE TABLE IF NOT EXISTS transacao_atividade (
  transacao_id TEXT NOT NULL,
  atividade_id INTEGER NOT NULL,
  criado_em    TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (transacao_id, atividade_id)
);
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'transacao_atividade_transacao_id_fkey') THEN
    ALTER TABLE transacao_atividade
      ADD CONSTRAINT transacao_atividade_transacao_id_fkey
      FOREIGN KEY (transacao_id) REFERENCES transacao(tran_cd_transacao)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'transacao_atividade_atividade_id_fkey') THEN
    ALTER TABLE transacao_atividade
      ADD CONSTRAINT transacao_atividade_atividade_id_fkey
      FOREIGN KEY (atividade_id) REFERENCES atividade(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS transacao_atividade_atividade_id_idx ON transacao_atividade(atividade_id);

-- cenario ↔ processo
CREATE TABLE IF NOT EXISTS cenario_processo (
  cenario_id  INTEGER NOT NULL,
  processo_id INTEGER NOT NULL,
  criado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (cenario_id, processo_id)
);
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'cenario_processo_cenario_id_fkey') THEN
    ALTER TABLE cenario_processo
      ADD CONSTRAINT cenario_processo_cenario_id_fkey
      FOREIGN KEY (cenario_id) REFERENCES cenario(cena_cd_cenario)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'cenario_processo_processo_id_fkey') THEN
    ALTER TABLE cenario_processo
      ADD CONSTRAINT cenario_processo_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS cenario_processo_processo_id_idx ON cenario_processo(processo_id);

-- cenario ↔ atividade
CREATE TABLE IF NOT EXISTS cenario_atividade (
  cenario_id   INTEGER NOT NULL,
  atividade_id INTEGER NOT NULL,
  criado_em    TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (cenario_id, atividade_id)
);
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'cenario_atividade_cenario_id_fkey') THEN
    ALTER TABLE cenario_atividade
      ADD CONSTRAINT cenario_atividade_cenario_id_fkey
      FOREIGN KEY (cenario_id) REFERENCES cenario(cena_cd_cenario)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'cenario_atividade_atividade_id_fkey') THEN
    ALTER TABLE cenario_atividade
      ADD CONSTRAINT cenario_atividade_atividade_id_fkey
      FOREIGN KEY (atividade_id) REFERENCES atividade(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;
CREATE INDEX IF NOT EXISTS cenario_atividade_atividade_id_idx ON cenario_atividade(atividade_id);
