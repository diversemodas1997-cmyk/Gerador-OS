/* Gera os segredos do servidor da fábrica.
   Rode uma vez, na máquina do servidor:   node servidor/gerar-chaves.js

   O Supabase auto-hospedado não tem chaves prontas: a ANON_KEY e a
   SERVICE_ROLE_KEY são JWTs ASSINADOS com o JWT_SECRET da sua instalação. Se as
   três não combinarem, tudo responde 401 e a instalação parece quebrada sem
   dizer por quê. Este script produz o conjunto inteiro já casado.

   Não precisa de internet nem de instalar nada — só do Node, que já vem com o
   módulo de criptografia usado aqui. */
const crypto = require('crypto');

const b64url = buf => Buffer.from(buf).toString('base64')
  .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function jwt(role, segredo, anos) {
  const agora = Math.floor(Date.now() / 1000);
  const cabecalho = { alg: 'HS256', typ: 'JWT' };
  const corpo = {
    iss: 'supabase',
    role,                                   // 'anon' ou 'service_role'
    iat: agora,
    exp: agora + Math.round(anos * 365 * 24 * 3600)
  };
  const base = b64url(JSON.stringify(cabecalho)) + '.' + b64url(JSON.stringify(corpo));
  const assinatura = b64url(crypto.createHmac('sha256', segredo).update(base).digest());
  return base + '.' + assinatura;
}

// Sem caracteres que atrapalhem em arquivo .env ou na linha de comando.
const senha = n => crypto.randomBytes(n * 2).toString('base64')
  .replace(/[^A-Za-z0-9]/g, '').slice(0, n);

const JWT_SECRET      = senha(48);   // mínimo 32; 48 por folga
const POSTGRES_PASSWORD = senha(32);
const SECRET_KEY_BASE = senha(64);
const VAULT_ENC_KEY   = senha(32);
const DASHBOARD_PASSWORD = senha(20);
const ANON_KEY         = jwt('anon', JWT_SECRET, 10);
const SERVICE_ROLE_KEY = jwt('service_role', JWT_SECRET, 10);

// Conferência: reassina e compara. Uma chave que não valide contra o próprio
// segredo derrubaria a instalação inteira com erro de autenticação — melhor
// descobrir aqui do que depois de tudo montado.
function confere(token, segredo) {
  const [c, p, a] = token.split('.');
  return b64url(crypto.createHmac('sha256', segredo).update(c + '.' + p).digest()) === a;
}
if (!confere(ANON_KEY, JWT_SECRET) || !confere(SERVICE_ROLE_KEY, JWT_SECRET)) {
  console.error('ERRO: as chaves não validaram contra o segredo. Não use este resultado.');
  process.exit(1);
}

console.log(`
=====================================================================
 SEGREDOS DO SERVIDOR DA FABRICA — guarde uma copia em lugar seguro
=====================================================================
 Cole o bloco abaixo no arquivo .env do Supabase (pasta supabase/docker),
 substituindo as linhas de mesmo nome que ja existirem la.

---------------------------------------------------------------------
POSTGRES_PASSWORD=${POSTGRES_PASSWORD}
JWT_SECRET=${JWT_SECRET}
ANON_KEY=${ANON_KEY}
SERVICE_ROLE_KEY=${SERVICE_ROLE_KEY}
SECRET_KEY_BASE=${SECRET_KEY_BASE}
VAULT_ENC_KEY=${VAULT_ENC_KEY}
DASHBOARD_USERNAME=admin
DASHBOARD_PASSWORD=${DASHBOARD_PASSWORD}
---------------------------------------------------------------------

 NO GERADOR-OS, em Configuracoes -> Servidor da fabrica, use:

   Endereco:  http://IP-DO-SERVIDOR:8000
   Chave:     ${ANON_KEY}

 A SERVICE_ROLE_KEY passa por cima de todas as permissoes. Ela e so
 para a migracao dos dados e para manutencao — nunca no navegador.
=====================================================================
`);
