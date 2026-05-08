-- Issue 008: escala de risco passa de B/M/A (string) para 1..5 (int).
-- Mapeamento legado: B->2, M->3, A->4 (deixa 1 e 5 livres para extremos).
-- Default migrado de 'M' para 3.

ALTER TABLE processo_risco
  ALTER COLUMN risco_probabilidade DROP DEFAULT,
  ALTER COLUMN risco_probabilidade TYPE integer USING (
    CASE risco_probabilidade
      WHEN 'B' THEN 2
      WHEN 'M' THEN 3
      WHEN 'A' THEN 4
      ELSE 3
    END
  ),
  ALTER COLUMN risco_probabilidade SET DEFAULT 3;

ALTER TABLE processo_risco
  ALTER COLUMN risco_impacto DROP DEFAULT,
  ALTER COLUMN risco_impacto TYPE integer USING (
    CASE risco_impacto
      WHEN 'B' THEN 2
      WHEN 'M' THEN 3
      WHEN 'A' THEN 4
      ELSE 3
    END
  ),
  ALTER COLUMN risco_impacto SET DEFAULT 3;

ALTER TABLE processo_risco
  ADD CONSTRAINT processo_risco_probabilidade_range CHECK (risco_probabilidade BETWEEN 1 AND 5),
  ADD CONSTRAINT processo_risco_impacto_range CHECK (risco_impacto BETWEEN 1 AND 5);
