-- BPM lifecycle, comments, version history and risks
-- Run this in Supabase SQL Editor

-- Status and responsible on mega_processo
ALTER TABLE mega_processo
  ADD COLUMN IF NOT EXISTS mepr_tx_status TEXT NOT NULL DEFAULT 'Publicado',
  ADD COLUMN IF NOT EXISTS responsavel_id TEXT REFERENCES usuario(id);

-- KPI fields on processo
ALTER TABLE processo
  ADD COLUMN IF NOT EXISTS proc_nr_tempo_medio DOUBLE PRECISION,
  ADD COLUMN IF NOT EXISTS proc_nr_custo DOUBLE PRECISION,
  ADD COLUMN IF NOT EXISTS proc_nr_volume INTEGER;

-- Comments on mega_processo (threaded)
CREATE TABLE IF NOT EXISTS processo_comentario (
  id TEXT PRIMARY KEY,
  mega_processo_id INTEGER NOT NULL REFERENCES mega_processo(mepr_cd_mega_processo) ON DELETE CASCADE,
  usuario_id TEXT NOT NULL REFERENCES usuario(id),
  texto TEXT NOT NULL,
  parent_id TEXT REFERENCES processo_comentario(id),
  criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Version / lifecycle audit trail on mega_processo
CREATE TABLE IF NOT EXISTS mega_processo_versao (
  id TEXT PRIMARY KEY,
  mega_processo_id INTEGER NOT NULL REFERENCES mega_processo(mepr_cd_mega_processo) ON DELETE CASCADE,
  versao_nr INTEGER NOT NULL,
  status_anterior TEXT NOT NULL,
  status_novo TEXT NOT NULL,
  criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  criado_por_id TEXT NOT NULL REFERENCES usuario(id)
);

-- Risks on mega_processo
CREATE TABLE IF NOT EXISTS processo_risco (
  id TEXT PRIMARY KEY,
  mega_processo_id INTEGER NOT NULL REFERENCES mega_processo(mepr_cd_mega_processo) ON DELETE CASCADE,
  risco_descricao TEXT NOT NULL,
  risco_probabilidade TEXT NOT NULL DEFAULT 'M',
  risco_impacto TEXT NOT NULL DEFAULT 'M',
  risco_controle TEXT,
  criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
