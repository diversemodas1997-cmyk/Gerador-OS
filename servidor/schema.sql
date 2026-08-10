-- =====================================================================
-- ESTRUTURA DO SERVIDOR DA FABRICA
-- =====================================================================
-- Rode UMA VEZ, no SQL Editor do Supabase LOCAL (http://IP:8000 -> SQL
-- Editor). Cria as tabelas que o Gerador-OS usa, as permissoes e o
-- Realtime. Depois disto, rode a migracao dos dados
-- (node servidor/migrar-do-backup.js).
--
-- E o mesmo desenho da nuvem, num arquivo so — na nuvem ele foi surgindo
-- aos poucos, entre varios scripts.
-- =====================================================================

-- Estado inteiro do app: UMA linha, id='main'.
CREATE TABLE IF NOT EXISTS shared_data (
  id          text PRIMARY KEY,
  data        jsonb,
  updated_at  timestamptz,
  updated_by  uuid
);

-- Snapshot diario do blob (retencao de 30 dias, feita pelo proprio app).
CREATE TABLE IF NOT EXISTS shared_data_backups (
  id             bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  snapshot_date  date UNIQUE,
  data           jsonb,
  created_by     uuid,
  created_at     timestamptz DEFAULT now()
);

-- Papel de cada conta: 'admin' escreve, 'usuario' so le.
CREATE TABLE IF NOT EXISTS user_roles (
  user_id     uuid PRIMARY KEY,
  role        text NOT NULL DEFAULT 'usuario',
  created_at  timestamptz DEFAULT now()
);

-- Tabela-sinal: avisa que alguem gravou e diz QUAIS chaves mudaram, para
-- as outras maquinas baixarem so aquilo. Ver sql/supabase-sync-signal.sql.
CREATE TABLE IF NOT EXISTS sync_signal (
  id            text PRIMARY KEY,
  updated_at    timestamptz,
  device_id     text,
  key_versions  jsonb
);
ALTER TABLE sync_signal REPLICA IDENTITY FULL;

-- ---------------------------------------------------------------------
-- PERMISSOES
-- Todo mundo autenticado le e grava: os dados sao compartilhados por
-- desenho. Quem pode editar de fato e decidido pelo papel em user_roles,
-- que o app consulta no servidor antes de qualquer acao que importe.
-- ---------------------------------------------------------------------
ALTER TABLE shared_data         ENABLE ROW LEVEL SECURITY;
ALTER TABLE shared_data_backups ENABLE ROW LEVEL SECURITY;
ALTER TABLE user_roles          ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_signal         ENABLE ROW LEVEL SECURITY;

DO $$
DECLARE
  t text; c text;
BEGIN
  FOREACH t IN ARRAY ARRAY['shared_data', 'sync_signal'] LOOP
    FOREACH c IN ARRAY ARRAY['SELECT', 'INSERT', 'UPDATE'] LOOP
      EXECUTE format(
        'DROP POLICY IF EXISTS %I ON %I', t || ': authenticated ' || lower(c), t);
      EXECUTE format(
        'CREATE POLICY %I ON %I FOR %s TO authenticated %s',
        t || ': authenticated ' || lower(c), t, c,
        CASE c WHEN 'INSERT' THEN 'WITH CHECK (true)'
               WHEN 'UPDATE' THEN 'USING (true) WITH CHECK (true)'
               ELSE 'USING (true)' END);
    END LOOP;
  END LOOP;

  -- backups: ler, inserir e apagar (a retencao de 30 dias apaga)
  FOREACH c IN ARRAY ARRAY['SELECT', 'INSERT', 'DELETE'] LOOP
    EXECUTE format('DROP POLICY IF EXISTS %I ON shared_data_backups',
                   'backups: authenticated ' || lower(c));
    EXECUTE format(
      'CREATE POLICY %I ON shared_data_backups FOR %s TO authenticated %s',
      'backups: authenticated ' || lower(c), c,
      CASE c WHEN 'INSERT' THEN 'WITH CHECK (true)' ELSE 'USING (true)' END);
  END LOOP;
END $$;

-- user_roles: qualquer autenticado LE (cada um precisa saber o proprio
-- papel para a tela funcionar). Escrever, so pelo painel do servidor —
-- de proposito nao ha politica de INSERT/UPDATE aqui.
DROP POLICY IF EXISTS "user_roles: authenticated select" ON user_roles;
CREATE POLICY "user_roles: authenticated select"
  ON user_roles FOR SELECT TO authenticated USING (true);

-- ---------------------------------------------------------------------
-- REALTIME: so a tabela-sinal. NUNCA publicar shared_data — o Postgres
-- empurraria a linha inteira (o blob) para cada cliente a cada gravacao.
-- ---------------------------------------------------------------------
DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_publication WHERE pubname = 'supabase_realtime') THEN
    CREATE PUBLICATION supabase_realtime;
  END IF;
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables
     WHERE pubname = 'supabase_realtime' AND tablename = 'sync_signal'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE sync_signal;
  END IF;
END $$;

-- ---------------------------------------------------------------------
-- STORAGE: bucket publico dos desenhos tecnicos. As imagens NAO ficam no
-- blob — o blob guarda o endereco delas. Sem este bucket (e sem copiar as
-- imagens da nuvem para ca), a folha de OS abre sem desenho.
-- ---------------------------------------------------------------------
INSERT INTO storage.buckets (id, name, public)
VALUES ('desenhos', 'desenhos', true)
ON CONFLICT (id) DO UPDATE SET public = true;

DROP POLICY IF EXISTS "desenhos: leitura publica" ON storage.objects;
CREATE POLICY "desenhos: leitura publica"
  ON storage.objects FOR SELECT
  USING (bucket_id = 'desenhos');

DROP POLICY IF EXISTS "desenhos: envio autenticado" ON storage.objects;
CREATE POLICY "desenhos: envio autenticado"
  ON storage.objects FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'desenhos');

-- =====================================================================
-- Conferencia
-- =====================================================================
-- SELECT tablename FROM pg_tables
--  WHERE tablename IN ('shared_data','shared_data_backups','user_roles','sync_signal');
-- SELECT tablename FROM pg_publication_tables WHERE pubname = 'supabase_realtime';
-- SELECT id, public FROM storage.buckets WHERE id = 'desenhos';
