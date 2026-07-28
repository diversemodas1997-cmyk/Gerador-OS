-- =====================================================================
-- QUEM PODE MUDAR O PAPEL DE UM USUÁRIO
-- =====================================================================
-- Rode este script no SQL Editor do Supabase (uma vez só).
--
-- POR QUE ELE EXISTE
-- A tela de "Usuários e permissões" já é escondida de quem não é admin, e o
-- app já confere o papel antes de chamar. Nada disso é tranca: a tela roda no
-- navegador de quem usa, e o papel é uma variável de JavaScript — quem abre o
-- console do navegador troca o valor num segundo, ou chama a função do Supabase
-- direto, sem passar por tela nenhuma. Era assim que um usuário comum
-- conseguia se promover a admin.
--
-- A tranca só vale quando está do lado do SERVIDOR. É o que este script faz:
-- a função `set_user_role` passa a conferir, ela mesma, se quem chamou é admin.
-- A partir daqui não importa o que o navegador diga.
--
-- COMO RODAR
--   1. Abra o painel do Supabase: https://supabase.com/dashboard
--   2. Selecione o projeto.
--   3. Menu lateral: SQL Editor → New query
--   4. Cole este script inteiro e clique em Run.
-- =====================================================================

-- Um usuário, um papel. Sem esta unicidade o mesmo usuário poderia acumular
-- duas linhas (uma 'usuario' e uma 'admin') e o app leria a que viesse primeiro.
CREATE UNIQUE INDEX IF NOT EXISTS user_roles_user_id_key ON public.user_roles (user_id);

CREATE OR REPLACE FUNCTION public.set_user_role(user_email text, new_role text)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER              -- roda com privilégio do dono, ignorando RLS...
SET search_path = public      -- ...por isso o search_path é fixado: sem isto, um
AS $$                         --    schema no caminho poderia sequestrar a função.
DECLARE
  quem uuid := auth.uid();
  alvo uuid;
  n_admins int;
BEGIN
  IF quem IS NULL THEN
    RAISE EXCEPTION 'Faça login para gerenciar usuários';
  END IF;

  -- A CONFERÊNCIA QUE FALTAVA: quem chama tem que ser admin.
  IF NOT EXISTS (
    SELECT 1 FROM public.user_roles WHERE user_id = quem AND role = 'admin'
  ) THEN
    RAISE EXCEPTION 'Apenas admin pode gerenciar usuários';
  END IF;

  IF new_role NOT IN ('admin', 'usuario') THEN
    RAISE EXCEPTION 'Papel inválido: %', new_role;
  END IF;

  SELECT id INTO alvo FROM auth.users WHERE lower(email) = lower(btrim(user_email));
  IF alvo IS NULL THEN
    RAISE EXCEPTION 'Não existe usuário com o e-mail %', user_email;
  END IF;

  -- Ninguém fica sem quem administre: o último admin não se rebaixa sozinho.
  IF alvo = quem AND new_role <> 'admin' THEN
    SELECT count(*) INTO n_admins FROM public.user_roles WHERE role = 'admin';
    IF n_admins <= 1 THEN
      RAISE EXCEPTION 'Você é o único admin — promova outro antes de se rebaixar';
    END IF;
  END IF;

  INSERT INTO public.user_roles (user_id, role)
  VALUES (alvo, new_role)
  ON CONFLICT (user_id) DO UPDATE SET role = EXCLUDED.role;
END;
$$;

-- Só quem está logado pode sequer chamar. (A função ainda recusa quem não é
-- admin — este GRANT é a porta, a conferência lá dentro é a tranca.)
REVOKE ALL ON FUNCTION public.set_user_role(text, text) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.set_user_role(text, text) TO authenticated;

-- Escrita direta na tabela continua fechada: quem quiser mudar papel passa pela
-- função acima, que conferiu quem é. (Sem política de INSERT/UPDATE, o RLS nega.)
ALTER TABLE public.user_roles ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "user_roles: authenticated insert" ON public.user_roles;
DROP POLICY IF EXISTS "user_roles: authenticated update" ON public.user_roles;
DROP POLICY IF EXISTS "user_roles: authenticated delete" ON public.user_roles;

-- =====================================================================
-- O PRIMEIRO ADMIN
-- =====================================================================
-- A função exige um admin para criar outro — então o primeiro não sai por ela.
-- Ele é posto à mão, aqui no SQL Editor (que roda como dono do banco e não
-- passa pela conferência). Hoje o admin é diversemodas1997@gmail.com; se um dia
-- o projeto for recriado do zero, troque o e-mail abaixo e rode só esta parte:
--
-- INSERT INTO public.user_roles (user_id, role)
-- SELECT id, 'admin' FROM auth.users WHERE lower(email) = 'diversemodas1997@gmail.com'
-- ON CONFLICT (user_id) DO UPDATE SET role = 'admin';

-- =====================================================================
-- CONFERIR SE PEGOU
-- =====================================================================
-- Quem é admin hoje:
--   SELECT u.email, r.role FROM public.user_roles r
--     JOIN auth.users u ON u.id = r.user_id ORDER BY r.role, u.email;
--
-- Teste real: entre no app com uma conta que NÃO é admin, abra o console do
-- navegador e rode
--   await supa.rpc('set_user_role', { user_email: 'a-propria-conta@exemplo.com', new_role: 'admin' })
-- A resposta tem que vir com erro "Apenas admin pode gerenciar usuários".
