-- Atributos extras do Mapa Visual: Produtos, Insumos (I/O), Sistemas
-- externos e Dependencias (Processo<->Processo, Atividade<->Atividade).
-- Aditiva e idempotente: segura para rodar mais de uma vez.

-- ─── produto ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS produto (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  descricao       TEXT NOT NULL,
  tipo            TEXT NOT NULL DEFAULT 'BEM',
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS processo_produto (
  processo_id     INTEGER NOT NULL,
  produto_id      INTEGER NOT NULL,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (processo_id, produto_id)
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'processo_produto_processo_id_fkey') THEN
    ALTER TABLE processo_produto
      ADD CONSTRAINT processo_produto_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'processo_produto_produto_id_fkey') THEN
    ALTER TABLE processo_produto
      ADD CONSTRAINT processo_produto_produto_id_fkey
      FOREIGN KEY (produto_id) REFERENCES produto(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS processo_produto_produto_id_idx ON processo_produto(produto_id);

-- ─── insumo (Inputs / Outputs) ──────────────────────────────────
CREATE TABLE IF NOT EXISTS insumo (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  descricao       TEXT NOT NULL,
  tipo            TEXT NOT NULL DEFAULT 'DADO',
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS insumo_processo (
  insumo_id       INTEGER NOT NULL,
  processo_id     INTEGER NOT NULL,
  direcao         TEXT NOT NULL,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (insumo_id, processo_id, direcao)
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'insumo_processo_insumo_id_fkey') THEN
    ALTER TABLE insumo_processo
      ADD CONSTRAINT insumo_processo_insumo_id_fkey
      FOREIGN KEY (insumo_id) REFERENCES insumo(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'insumo_processo_processo_id_fkey') THEN
    ALTER TABLE insumo_processo
      ADD CONSTRAINT insumo_processo_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS insumo_processo_processo_direcao_idx ON insumo_processo(processo_id, direcao);

CREATE TABLE IF NOT EXISTS insumo_atividade (
  insumo_id       INTEGER NOT NULL,
  atividade_id    INTEGER NOT NULL,
  direcao         TEXT NOT NULL,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (insumo_id, atividade_id, direcao)
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'insumo_atividade_insumo_id_fkey') THEN
    ALTER TABLE insumo_atividade
      ADD CONSTRAINT insumo_atividade_insumo_id_fkey
      FOREIGN KEY (insumo_id) REFERENCES insumo(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'insumo_atividade_atividade_id_fkey') THEN
    ALTER TABLE insumo_atividade
      ADD CONSTRAINT insumo_atividade_atividade_id_fkey
      FOREIGN KEY (atividade_id) REFERENCES atividade(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS insumo_atividade_atividade_direcao_idx ON insumo_atividade(atividade_id, direcao);

-- ─── sistema externo ────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sistema (
  id              SERIAL PRIMARY KEY,
  codigo          TEXT NOT NULL UNIQUE,
  nome            TEXT NOT NULL,
  tipo            TEXT NOT NULL DEFAULT 'OUTRO',
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  atualizado_em   TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS processo_sistema (
  processo_id     INTEGER NOT NULL,
  sistema_id      INTEGER NOT NULL,
  papel           TEXT NOT NULL DEFAULT 'CONSUMIDOR',
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (processo_id, sistema_id)
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'processo_sistema_processo_id_fkey') THEN
    ALTER TABLE processo_sistema
      ADD CONSTRAINT processo_sistema_processo_id_fkey
      FOREIGN KEY (processo_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'processo_sistema_sistema_id_fkey') THEN
    ALTER TABLE processo_sistema
      ADD CONSTRAINT processo_sistema_sistema_id_fkey
      FOREIGN KEY (sistema_id) REFERENCES sistema(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS processo_sistema_sistema_id_idx ON processo_sistema(sistema_id);

-- ─── dependencias entre Processos ────────────────────────────────
CREATE TABLE IF NOT EXISTS dependencia_processo (
  id              TEXT PRIMARY KEY,
  origem_id       INTEGER NOT NULL,
  destino_id      INTEGER NOT NULL,
  tipo            TEXT NOT NULL DEFAULT 'PRECEDE',
  descricao       TEXT,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'dependencia_processo_origem_id_fkey') THEN
    ALTER TABLE dependencia_processo
      ADD CONSTRAINT dependencia_processo_origem_id_fkey
      FOREIGN KEY (origem_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'dependencia_processo_destino_id_fkey') THEN
    ALTER TABLE dependencia_processo
      ADD CONSTRAINT dependencia_processo_destino_id_fkey
      FOREIGN KEY (destino_id) REFERENCES processo(proc_cd_processo)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE UNIQUE INDEX IF NOT EXISTS dependencia_processo_unique_idx
  ON dependencia_processo(origem_id, destino_id, tipo);
CREATE INDEX IF NOT EXISTS dependencia_processo_origem_idx ON dependencia_processo(origem_id);
CREATE INDEX IF NOT EXISTS dependencia_processo_destino_idx ON dependencia_processo(destino_id);

-- ─── dependencias entre Atividades ───────────────────────────────
CREATE TABLE IF NOT EXISTS dependencia_atividade (
  id              TEXT PRIMARY KEY,
  origem_id       INTEGER NOT NULL,
  destino_id      INTEGER NOT NULL,
  tipo            TEXT NOT NULL DEFAULT 'PRECEDE',
  descricao       TEXT,
  criado_em       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'dependencia_atividade_origem_id_fkey') THEN
    ALTER TABLE dependencia_atividade
      ADD CONSTRAINT dependencia_atividade_origem_id_fkey
      FOREIGN KEY (origem_id) REFERENCES atividade(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'dependencia_atividade_destino_id_fkey') THEN
    ALTER TABLE dependencia_atividade
      ADD CONSTRAINT dependencia_atividade_destino_id_fkey
      FOREIGN KEY (destino_id) REFERENCES atividade(id)
      ON DELETE CASCADE ON UPDATE CASCADE;
  END IF;
END $$;

CREATE UNIQUE INDEX IF NOT EXISTS dependencia_atividade_unique_idx
  ON dependencia_atividade(origem_id, destino_id, tipo);
CREATE INDEX IF NOT EXISTS dependencia_atividade_origem_idx ON dependencia_atividade(origem_id);
CREATE INDEX IF NOT EXISTS dependencia_atividade_destino_idx ON dependencia_atividade(destino_id);
