/* Reatribui o AUTOR das observações das OS de um login antigo para outro.

   POR QUE ISTO EXISTE
   Cada observação da folha guarda o autor congelado (`n.login`, o e-mail de
   quem escreveu), não uma referência viva à conta. Renomear uma conta pelo app
   já conserta dali em diante (_reatribuirAutorNotas). Mas notas escritas por uma
   conta que foi SUBSTITUÍDA (não renomeada) — o caso do login antigo de Gmail —
   ficam órfãs: nada nunca as religa. Este script faz esse religamento, uma vez.

   Roda NO SERVIDOR da fábrica. Lê a chave de serviço do .env do Supabase local,
   lê o shared_data, troca o login nas notas e grava de volta. É a mesma conta
   read-modify-write que a restauração credenciada usa.

   USO (confere primeiro, grava depois):
     node servidor/reatribuir-notas-usuario.js --de diversemodas1997@gmail.com --para admin@diverse.local
     node servidor/reatribuir-notas-usuario.js --de ... --para ... --gravar
*/
const fs = require('fs');
const https = require('https');

const args = process.argv.slice(2);
const opt = (n) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : null; };
const DE = (opt('--de') || '').trim().toLowerCase();
const PARA = (opt('--para') || '').trim().toLowerCase();
const GRAVAR = args.includes('--gravar');
const BASE = opt('--url') || 'https://193.168.0.200';
const ENV = opt('--env') || 'C:\\supabase\\docker\\.env';

if (!DE || !PARA) { console.error('faltou --de <email> e/ou --para <email>'); process.exit(1); }
if (DE === PARA) { console.error('--de e --para são iguais, nada a fazer'); process.exit(1); }

const envTxt = fs.readFileSync(ENV, 'utf8');
const m = envTxt.match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m);
if (!m) { console.error('não achei SERVICE_ROLE_KEY em ' + ENV); process.exit(1); }
const KEY = m[1].trim();

// O certificado é o da própria fábrica; este script fala com o servidor local.
const agente = new https.Agent({ rejectUnauthorized: false });

function req(method, path, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const u = new URL(BASE + path);
    const r = https.request({
      hostname: u.hostname, port: u.port || 443, path: u.pathname + u.search, method, agent: agente,
      headers: {
        apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json',
        Prefer: method === 'PATCH' ? 'return=minimal' : '',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
      }
    }, (resp) => {
      let d = ''; resp.on('data', c => d += c);
      resp.on('end', () => resp.statusCode < 300 ? res({ status: resp.statusCode, body: d })
                                                  : rej(new Error('HTTP ' + resp.statusCode + ': ' + d.slice(0, 300))));
    });
    r.on('error', rej); if (data) r.write(data); r.end();
  });
}

function pk(v) { return typeof v === 'string' ? JSON.parse(v) : v; }

(async () => {
  const rows = JSON.parse((await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_by')).body);
  if (!rows.length) { console.error('shared_data id=main não encontrado'); process.exit(1); }
  const data = rows[0].data || {};
  const ordensEraString = typeof data.ordens === 'string';
  const ordens = pk(data.ordens) || [];

  let notas = 0, osTocadas = 0;
  ordens.forEach(o => {
    const arr = Array.isArray(o.obsNotas) ? o.obsNotas : [];
    if (!arr.length) return;
    let mudou = false;
    arr.forEach(n => { if (String(n.login || '').trim().toLowerCase() === DE) { n.login = PARA; notas++; mudou = true; } });
    if (!mudou) return;
    osTocadas++;
    // Colapsa para uma nota por login (a reatribuição pode juntar duas do mesmo
    // autor na mesma OS): fica a de data mais recente.
    const porLogin = new Map();
    arr.forEach(n => {
      const k = String(n.login || '').trim().toLowerCase();
      const t = n.editadoEm || n.em || '';
      const at = porLogin.get(k);
      if (!at || t > (at.editadoEm || at.em || '')) porLogin.set(k, n);
    });
    o.obsNotas = Array.from(porLogin.values());
  });

  console.log(`de:   ${DE}`);
  console.log(`para: ${PARA}`);
  console.log(`OS afetadas: ${osTocadas} | notas reatribuídas: ${notas}`);

  if (!GRAVAR) { console.log('\n(conferência — nada gravado. Rode de novo com --gravar para aplicar.)'); return; }
  if (!notas) { console.log('nada a gravar.'); return; }

  data.ordens = ordensEraString ? JSON.stringify(ordens) : ordens;
  data._device = 'reparo-notas-' + Date.now();   // sentinela: todos os clientes recarregam
  await req('PATCH', '/rest/v1/shared_data?id=eq.main',
            { data, updated_at: new Date().toISOString() });
  console.log('\n✅ gravado no shared_data. Os clientes conectados recarregam sozinhos.');
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
