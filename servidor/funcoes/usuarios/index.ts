/* Contas de acesso do Gerador-OS — listar e criar, do lado do servidor.

   POR QUE ISTO EXISTE, E NÃO É FEITO DIRETO NO NAVEGADOR

   Criar conta pelo navegador exigiria ligar o auto-cadastro do Supabase. A
   chave anônima está dentro de toda página aberta na fábrica — com o
   auto-cadastro ligado, qualquer pessoa na rede criaria a própria conta e
   passaria a ler tudo: OS, custos, produção. Então o auto-cadastro fica
   DESLIGADO e a criação passa por aqui, onde a chave de serviço nunca sai do
   servidor.

   Listar também não sai do navegador: quem é usuário comum não tem linha em
   `user_roles` (é assim que o app o trata como leitura), então a lista completa
   só existe em `auth.users`, que a chave anônima não alcança.

   SEM NENHUM IMPORT, de propósito. Uma função que baixa dependência na
   primeira execução quebraria justamente no dia em que a internet cair — que é
   o dia para o qual este servidor existe. Só `fetch` contra o próprio Supabase.
*/

const URL_BASE = Deno.env.get('SUPABASE_URL') ?? 'http://kong:8000';
const CHAVE_SERVICO = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? '';
const CHAVE_ANON = Deno.env.get('SUPABASE_ANON_KEY') ?? '';

const CABECALHOS = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

function resposta(corpo: unknown, status = 200) {
  return new Response(JSON.stringify(corpo), { status, headers: CABECALHOS });
}

/* Quem está chamando? Pergunta ao próprio Supabase, em vez de conferir a
   assinatura aqui: assim vale a mesma regra que vale para o resto do app, e
   uma sessão revogada deixa de valer aqui também. */
async function quemChamou(autorizacao: string) {
  const r = await fetch(`${URL_BASE}/auth/v1/user`, {
    headers: { Authorization: autorizacao, apikey: CHAVE_ANON },
  });
  if (!r.ok) return null;
  const u = await r.json();
  return u && u.id ? u : null;
}

async function ehAdmin(userId: string) {
  const r = await fetch(
    `${URL_BASE}/rest/v1/user_roles?user_id=eq.${encodeURIComponent(userId)}&select=role`,
    { headers: { apikey: CHAVE_SERVICO, Authorization: `Bearer ${CHAVE_SERVICO}` } },
  );
  if (!r.ok) return false;
  const linhas = await r.json();
  return Array.isArray(linhas) && linhas.length > 0 && linhas[0].role === 'admin';
}

async function listarContas() {
  // A lista de contas vem da autenticação; o papel vem da nossa tabela. Só as
  // duas juntas respondem "quem tem acesso, e o que cada um pode fazer".
  const contas: any[] = [];
  for (let pagina = 1; pagina <= 20; pagina++) {
    const r = await fetch(`${URL_BASE}/auth/v1/admin/users?page=${pagina}&per_page=200`, {
      headers: { apikey: CHAVE_SERVICO, Authorization: `Bearer ${CHAVE_SERVICO}` },
    });
    if (!r.ok) return { erro: `não consegui ler as contas (${r.status})` };
    const d = await r.json();
    const lote = Array.isArray(d) ? d : (d.users || []);
    contas.push(...lote);
    if (lote.length < 200) break;
  }

  const rp = await fetch(`${URL_BASE}/rest/v1/user_roles?select=user_id,role`, {
    headers: { apikey: CHAVE_SERVICO, Authorization: `Bearer ${CHAVE_SERVICO}` },
  });
  const papeis: Record<string, string> = {};
  if (rp.ok) for (const l of await rp.json()) papeis[l.user_id] = l.role;

  return {
    usuarios: contas.map((u) => ({
      id: u.id,
      email: u.email,
      papel: papeis[u.id] || 'usuario',
      confirmada: !!u.email_confirmed_at,
      criada_em: u.created_at,
      ultimo_acesso: u.last_sign_in_at || null,
    })).sort((a, b) => String(a.email).localeCompare(String(b.email))),
  };
}

async function criarConta(email: string, senha: string, papel: string) {
  const r = await fetch(`${URL_BASE}/auth/v1/admin/users`, {
    method: 'POST',
    headers: {
      apikey: CHAVE_SERVICO,
      Authorization: `Bearer ${CHAVE_SERVICO}`,
      'Content-Type': 'application/json',
    },
    // email_confirm: a conta já nasce válida. Não há servidor de e-mail na
    // fábrica — sem isto ela ficaria esperando uma confirmação que nunca chega,
    // e a pessoa não entraria, sem ninguém entender por quê.
    body: JSON.stringify({ email, password: senha, email_confirm: true }),
  });

  const corpo = await r.json().catch(() => ({}));
  if (!r.ok) {
    const msg = corpo.msg || corpo.message || corpo.error_description || `erro ${r.status}`;
    return { erro: /already|exists|registered/i.test(String(msg))
      ? 'Já existe uma conta com esse e-mail'
      : String(msg) };
  }

  if (papel === 'admin' && corpo.id) {
    await fetch(`${URL_BASE}/rest/v1/user_roles`, {
      method: 'POST',
      headers: {
        apikey: CHAVE_SERVICO,
        Authorization: `Bearer ${CHAVE_SERVICO}`,
        'Content-Type': 'application/json',
        Prefer: 'resolution=merge-duplicates',
      },
      body: JSON.stringify({ user_id: corpo.id, role: 'admin' }),
    });
  }

  return { ok: true, id: corpo.id, email: corpo.email };
}

Deno.serve(async (req: Request) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CABECALHOS });
  if (req.method !== 'POST') return resposta({ erro: 'use POST' }, 405);

  const autorizacao = req.headers.get('Authorization') || '';
  if (!autorizacao) return resposta({ erro: 'Faça login para gerenciar contas' }, 401);

  const usuario = await quemChamou(autorizacao);
  if (!usuario) return resposta({ erro: 'Sessão inválida — entre de novo' }, 401);

  // A tranca está AQUI, no servidor. Esconder o painel no navegador não é
  // tranca nenhuma: a página roda na máquina de quem usa.
  if (!(await ehAdmin(usuario.id))) {
    return resposta({ erro: 'Apenas admin pode gerenciar contas' }, 403);
  }

  let corpo: any = {};
  try { corpo = await req.json(); } catch { /* corpo vazio = listar */ }
  const acao = corpo.acao || 'listar';

  if (acao === 'listar') {
    const r = await listarContas();
    return resposta(r, (r as any).erro ? 500 : 200);
  }

  if (acao === 'criar') {
    const email = String(corpo.email || '').trim().toLowerCase();
    const senha = String(corpo.senha || '');
    const papel = corpo.papel === 'admin' ? 'admin' : 'usuario';

    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return resposta({ erro: 'E-mail inválido' }, 400);
    if (senha.length < 6) return resposta({ erro: 'A senha precisa de pelo menos 6 caracteres' }, 400);

    const r = await criarConta(email, senha, papel);
    return resposta(r, (r as any).erro ? 400 : 200);
  }

  return resposta({ erro: `ação desconhecida: ${acao}` }, 400);
});
