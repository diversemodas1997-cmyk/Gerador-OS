/*
 * DA NOME A PRATELEIRA "SEM COR" DA RIBANA MALHA ALGODAO.
 *
 * Junior, 23/09/2026: "insira nome para o cadastro sem cor em ribana malha
 * algodao em estoque de tecido. O nome da cor e algodao cru".
 *
 * A linha "sem cor" sao 4 movimentos gravados com a cor "—": a entrada de
 * abertura de 27/08 e as baixas das OS 0426, 0430 e 0437 (a gola delas saiu sem
 * cor no enfesto). O saldo da linha e zero.
 *
 * Cria no cadastro a cor "Algodao Cru Ribana Malha Algodao" — no padrao das
 * outras cores de ribana, uma por tecido — e troca a cor desses movimentos.
 *
 *   node servidor/nomear-cor-ribana-cru.js             so relata
 *   node servidor/nomear-cor-ribana-cru.js --gravar    grava (backup antes, carimbo conferido)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const TECIDO = 'Ribana Malha Algodão';
const COR = 'Algodão Cru Ribana Malha Algodão';

async function conectar() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  return { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
}

(async () => {
  const cab = await conectar();
  const l = await (await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  const { data, updated_at: updatedAt } = l[0];
  const cores = JSON.parse(data.cores || '[]');
  const mov = JSON.parse(data.estoqueMov || '[]');

  const semCor = m => m.tecidoNome === TECIDO && (!String(m.corNome || '').trim() || m.corNome === '—');
  const alvo = mov.filter(semCor);
  console.log('movimentos sem cor em ' + TECIDO + ': ' + alvo.length);
  alvo.forEach(m => console.log('  ' + m.tipo + ' ' + m.kg + ' kg  ' + (m.osNumero ? 'OS ' + m.osNumero : m.obs)));

  let cor = cores.find(c => c.nome.toLowerCase() === COR.toLowerCase());
  if (!cor) {
    const base = cores.find(c => /^algod[aã]o cru$/i.test(c.nome)) || {};
    const codigos = cores.filter(c => / Ribana Malha Algodão$/.test(c.nome)).map(c => parseInt(c.codigo, 10) || 0);
    cor = { id: 'id_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
            nome: COR, hex: base.hex || '#fffaf0',
            codigo: String(Math.max(0, ...codigos) + 1).padStart(3, '0'),
            siglaSku: base.siglaSku || 'CRU', peso: 0 };
    cores.push(cor);
    console.log('cor nova no cadastro: ' + JSON.stringify(cor));
  } else console.log('cor ja existe no cadastro: ' + cor.nome);

  alvo.forEach(m => { m.corNome = cor.nome; });
  if (!GRAVAR) { console.log('\nSIMULACAO — nada foi gravado. Rode com --gravar para aplicar.'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cor-ribana-cru-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\ncopia de seguranca: ' + path.relative(RAIZ, arq));

  data.cores = JSON.stringify(cores);
  data.estoqueMov = JSON.stringify(mov);
  const r = await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
    body: JSON.stringify({ data, updated_at: new Date().toISOString() })
  });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) throw new Error('ALGUEM GRAVOU ENTRE A LEITURA E A ESCRITA — nada foi alterado. Rode de novo.');
  console.log('gravado: ' + alvo.length + ' movimentos com a cor "' + cor.nome + '".');
})().catch(e => { console.error('\nFALHOU: ' + e.message); process.exit(1); });
