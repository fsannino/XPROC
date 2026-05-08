-- RACI: Áreas, Funções, Pessoas + atribuições por Processo.
-- Migration aditiva e idempotente; segura para rodar mais de uma vez.

-- ─── area ────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS area (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  descricao       TEXT NOT NULL,
  parent_id       INTEGER,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'area_parent_id_fkey') THEN
    ALTER TABLE area
      ADD CONSTRAINT area_parent_id_fkey
      FOREIGN KEY (parent_id) REFERENCES area(id)
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS area_parent_id_idx ON area(parent_id);

-- ─── funcao ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS funcao (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  descricao       TEXT NOT NULL,
  area_id         INTEGER,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'funcao_area_id_fkey') THEN
    ALTER TABLE funcao
      ADD CONSTRAINT funcao_area_id_fkey
      FOREIGN KEY (area_id) REFERENCES area(id)
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS funcao_area_id_idx ON funcao(area_id);

-- ─── pessoa ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS pessoa (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  nome            TEXT NOT NULL,
  email           TEXT UNIQUE,
  area_id         INTEGER,
  funcao_id       INTEGER,
  usuario_id      TEXT UNIQUE,
  ativo           BOOLEAN NOT NULL DEFAULT TRUE,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'pessoa_area_id_fkey') THEN
    ALTER TABLE pessoa
      ADD CONSTRAINT pessoa_area_id_fkey
      FOREIGN KEY (area_id) REFERENCES area(id)
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'pessoa_funcao_id_fkey') THEN
    ALTER TABLE pessoa
      ADD CONSTRAINT pessoa_funcao_id_fkey
      FOREIGN KEY (funcao_id) REFERENCES funcao(id)
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'pessoa_usuario_id_fkey') THEN
    ALTER TABLE pessoa
      ADD CONSTRAINT pessoa_usuario_id_fkey
      FOREIGN KEY (usuario_id) REFERENCES usuario(id)
      ON DELETE SET NULL ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS pessoa_area_id_idx ON pessoa(area_id);
CREATE INDEX IF NOT EXISTS pessoa_funcao_id_idx ON pessoa(funcao_id);

-- ─── raci_atribuicao ──────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS raci_atribuicao (
  id              TEXT PRIMARY KEY,
  processo_id     INTEGER NOT NULL,
  pessoa_id       INTEGER NOT NULL,
  papel           TEXT NOT NULL,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'raci_atribuicao_processo_id_fkey') THEN
    ALTER TABLE raci_atribuicao
      ADD CONSTRAINT raci_atribuicao_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'raci_atribuicao_pessoa_id_fkey') THEN
    ALTER TABLE raci_atribuicao
      ADD CONSTRAINT raci_atribuicao_pessoa_id_fkey
      FOREIGN KEY (pessoa_id) REFERENCES pessoa(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint
    WHERE conname = 'raci_atribuicao_processo_id_pessoa_id_papel_key'
  ) THEN
    ALTER TABLE raci_atribuicao
      ADD CONSTRAINT raci_atribuicao_processo_id_pessoa_id_papel_key
      UNIQUE (processo_id, pessoa_id, papel);
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS raci_atribuicao_processo_id_idx ON raci_atribuicao(processo_id);
CREATE INDEX IF NOT EXISTS raci_atribuicao_pessoa_id_idx ON raci_atribuicao(pessoa_id);
