CREATE UNIQUE INDEX IF NOT EXISTS user_roles_user_id_key ON public.user_roles (user_id);

CREATE OR REPLACE FUNCTION public.set_user_role(user_email text, new_role text)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  quem uuid := auth.uid();
  alvo uuid;
  n_admins int;
BEGIN
  IF quem IS NULL THEN
    RAISE EXCEPTION 'Faca login para gerenciar usuarios';
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM public.user_roles WHERE user_id = quem AND role = 'admin'
  ) THEN
    RAISE EXCEPTION 'Apenas admin pode gerenciar usuarios';
  END IF;

  IF new_role NOT IN ('admin', 'usuario') THEN
    RAISE EXCEPTION 'Papel invalido: %', new_role;
  END IF;

  SELECT id INTO alvo FROM auth.users WHERE lower(email) = lower(btrim(user_email));
  IF alvo IS NULL THEN
    RAISE EXCEPTION 'Nao existe usuario com o e-mail %', user_email;
  END IF;

  IF alvo = quem AND new_role <> 'admin' THEN
    SELECT count(*) INTO n_admins FROM public.user_roles WHERE role = 'admin';
    IF n_admins <= 1 THEN
      RAISE EXCEPTION 'Voce e o unico admin - promova outro antes de se rebaixar';
    END IF;
  END IF;

  INSERT INTO public.user_roles (user_id, role)
  VALUES (alvo, new_role)
  ON CONFLICT (user_id) DO UPDATE SET role = EXCLUDED.role;
END;
$$;

REVOKE ALL ON FUNCTION public.set_user_role(text, text) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.set_user_role(text, text) TO authenticated;

ALTER TABLE public.user_roles ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "user_roles: authenticated insert" ON public.user_roles;
DROP POLICY IF EXISTS "user_roles: authenticated update" ON public.user_roles;
DROP POLICY IF EXISTS "user_roles: authenticated delete" ON public.user_roles;
