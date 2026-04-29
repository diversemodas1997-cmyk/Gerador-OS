-- =====================================================================
-- POLITICAS RLS PARA ACESSO MUTUO ENTRE USUARIOS
-- =====================================================================
-- Rode este script no SQL Editor do Supabase (uma vez so).
-- Garante que QUALQUER usuario autenticado consegue ler e gravar
-- na linha unica shared_data.id='main' — assim todos veem as mesmas
-- OS, cadastros, etc., independente de qual conta criou.
--
-- Como rodar:
--   1. Abra o painel do Supabase: https://supabase.com/dashboard
--   2. Selecione o projeto (ckkqrjkhorvaahyazqsr).
--   3. Menu lateral: SQL Editor → New query
--   4. Cole este script inteiro e clique em Run.
-- =====================================================================

-- shared_data (linha unica id='main') — leitura e escrita liberadas
-- pra qualquer usuario autenticado.
ALTER TABLE shared_data ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "shared_data: authenticated select" ON shared_data;
CREATE POLICY "shared_data: authenticated select"
  ON shared_data FOR SELECT
  TO authenticated
  USING (true);

DROP POLICY IF EXISTS "shared_data: authenticated insert" ON shared_data;
CREATE POLICY "shared_data: authenticated insert"
  ON shared_data FOR INSERT
  TO authenticated
  WITH CHECK (true);

DROP POLICY IF EXISTS "shared_data: authenticated update" ON shared_data;
CREATE POLICY "shared_data: authenticated update"
  ON shared_data FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

-- shared_data_backups — todos podem ler (pra restaurar) e inserir
-- (snapshot diario rodado por qualquer usuario logado).
ALTER TABLE shared_data_backups ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "backups: authenticated select" ON shared_data_backups;
CREATE POLICY "backups: authenticated select"
  ON shared_data_backups FOR SELECT
  TO authenticated
  USING (true);

DROP POLICY IF EXISTS "backups: authenticated insert" ON shared_data_backups;
CREATE POLICY "backups: authenticated insert"
  ON shared_data_backups FOR INSERT
  TO authenticated
  WITH CHECK (true);

DROP POLICY IF EXISTS "backups: authenticated delete" ON shared_data_backups;
CREATE POLICY "backups: authenticated delete"
  ON shared_data_backups FOR DELETE
  TO authenticated
  USING (true);

-- user_roles — leitura liberada pra qualquer autenticado (cada user
-- precisa ler o proprio papel pra UI funcionar). Insercao/update so
-- via service role (admin de Supabase, fora do app).
ALTER TABLE user_roles ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "user_roles: authenticated select" ON user_roles;
CREATE POLICY "user_roles: authenticated select"
  ON user_roles FOR SELECT
  TO authenticated
  USING (true);

-- =====================================================================
-- Verificacao: depois de rodar, rode esta query pra conferir que as
-- politicas estao em vigor. Deve listar 7 linhas (3 shared_data,
-- 3 backups, 1 user_roles).
-- =====================================================================
-- SELECT schemaname, tablename, policyname, roles, cmd
--   FROM pg_policies
--  WHERE tablename IN ('shared_data','shared_data_backups','user_roles')
--  ORDER BY tablename, policyname;
