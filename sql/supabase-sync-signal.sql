-- =====================================================================
-- TABELA-SINAL DO REALTIME  (economia de egress)
-- =====================================================================
-- POR QUE ESTE SCRIPT EXISTE
--
-- O app assinava as mudancas da tabela `shared_data` pelo Realtime. O
-- Postgres empurra a LINHA INTEIRA da tabela assinada para cada cliente
-- conectado, a cada gravacao — e a linha de `shared_data` carrega o blob
-- de dados do sistema inteiro (~1,8 MB, truncado no teto de 1 MB do
-- Realtime). Pior: o app JOGAVA FORA esse conteudo. Ele usava o evento
-- so como aviso de "alguem gravou, va reler", porque o payload truncado
-- nao e confiavel — o estado sempre veio por REST, completo.
--
-- Ou seja: ate ~1 MB por gravacao por cliente conectado, transportado
-- para nada. Junto com o polling (que pedia o blob a cada 15s), foi o
-- que estourou a cota de egress do plano e restringiu o projeto com
-- "exceed_egress_quota".
--
-- A tabela `sync_signal` transmite o MESMO aviso em tres campos curtos.
--
-- COMO RODAR
--   1. Abra o painel do Supabase: https://supabase.com/dashboard
--   2. Selecione o projeto (ckkqrjkhorvaahyazqsr).
--   3. Menu lateral: SQL Editor -> New query
--   4. Cole este script inteiro e clique em Run.
--
-- Enquanto este script nao for rodado, o app continua funcionando: o
-- aviso instantaneo nao chega, e a sincronizacao entre maquinas passa a
-- depender do polling de 15s (que agora custa alguns bytes por ciclo).
-- =====================================================================

-- Linha unica id='main'. So o carimbo e quem gravou — nada de dados.
CREATE TABLE IF NOT EXISTS sync_signal (
  id          text PRIMARY KEY,
  updated_at  timestamptz,
  device_id   text
);

ALTER TABLE sync_signal ENABLE ROW LEVEL SECURITY;

-- Qualquer autenticado le e grava: o aviso e coletivo, como o dado.
DROP POLICY IF EXISTS "sync_signal: authenticated select" ON sync_signal;
CREATE POLICY "sync_signal: authenticated select"
  ON sync_signal FOR SELECT
  TO authenticated
  USING (true);

DROP POLICY IF EXISTS "sync_signal: authenticated insert" ON sync_signal;
CREATE POLICY "sync_signal: authenticated insert"
  ON sync_signal FOR INSERT
  TO authenticated
  WITH CHECK (true);

DROP POLICY IF EXISTS "sync_signal: authenticated update" ON sync_signal;
CREATE POLICY "sync_signal: authenticated update"
  ON sync_signal FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

-- REPLICA IDENTITY FULL: sem isso o payload de UPDATE do Realtime pode
-- vir sem as colunas que nao mudaram. O app compara `device_id` para
-- reconhecer o proprio eco — precisa dele em toda notificacao.
ALTER TABLE sync_signal REPLICA IDENTITY FULL;

-- ---------------------------------------------------------------------
-- Publication do Realtime: ENTRA sync_signal, SAI shared_data.
-- ---------------------------------------------------------------------
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication WHERE pubname = 'supabase_realtime'
  ) THEN
    CREATE PUBLICATION supabase_realtime;
  END IF;
END $$;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables
     WHERE pubname = 'supabase_realtime' AND tablename = 'sync_signal'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE sync_signal;
  END IF;
END $$;

-- Tira `shared_data` da publication — e AQUI que a economia acontece.
-- Deixar a tabela publicada mantem o Postgres transmitindo o blob a cada
-- gravacao para qualquer cliente que ainda a assine (uma aba com a versao
-- antiga do app em cache, por exemplo). Quem estiver nessa situacao passa
-- a depender do polling de 15s, que continua funcionando.
DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM pg_publication_tables
     WHERE pubname = 'supabase_realtime' AND tablename = 'shared_data'
  ) THEN
    ALTER PUBLICATION supabase_realtime DROP TABLE shared_data;
  END IF;
END $$;

-- Semente da linha, para o primeiro aviso ja ser um UPDATE normal.
INSERT INTO sync_signal (id, updated_at, device_id)
VALUES ('main', now(), 'setup')
ON CONFLICT (id) DO NOTHING;

-- =====================================================================
-- Verificacao: deve listar sync_signal e NAO listar shared_data.
-- =====================================================================
-- SELECT tablename FROM pg_publication_tables
--  WHERE pubname = 'supabase_realtime' ORDER BY tablename;
--
-- E as politicas (3 linhas):
-- SELECT policyname, cmd FROM pg_policies WHERE tablename = 'sync_signal';
