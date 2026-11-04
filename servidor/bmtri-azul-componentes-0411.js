/*
 * BM.TRI-AZUL SEM OS COMPONENTES DA TRICOLOR PRETO/MOSTARDA/OFF-WHITE.
 *
 * Junior, 04/11/2026: "analise por que a etiqueta da OS662 mostra texto
 * AZUL/PRETO/BRANCO/MOSTARDA/OFF-WHITE. As cores dessa OS não possuem mostarda
 * [nem off-white]" — as cores são Azul/Branco/Preto.
 *
 * O desenho 0049 (Blusa Moletom Tricolor | Azul/Branco/Preto) nasceu em 14/10
 * como copia do 0025 (Preto/Mostarda/Off-White) e levou junto, alem dos 15
 * componentes dele, os 9 antigos de abril — "Frente/Costa/Mangas PARTE 1/2/3
 * Blusa Moletom Tricolor" — com as cores do 0025. A etiqueta da BM.TRI junta as
 * cores dos componentes cujo nome comeca por Frente/Bolso/Costa e Capuz/Forro/
 * Mangas, e os 9 casam: dai o MOSTARDA e o OFF-WHITE.
 *
 * O script tira esses 9 (pelo componenteId no desenho, pelo nome na OS) de:
 *   - desenho 0049 (id_1791988631415_576)
 *   - OS 0661 e 0662, as unicas do 0049
 * Nenhuma fase nem enfesto dessas OS aponta para eles (conferido antes).
 * Os desenhos 0024 e 0025 tambem tem os 9, mas la as cores sao as da peca;
 * ficam como estao.
 *
 *   node servidor/bmtri-azul-componentes-0411.js            (so mostra)
 *   node servidor/bmtri-azul-componentes-0411.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const DESENHO_ID = 'id_1791988631415_576';
const OS_ALVO = ['0661', '0662'];
const VELHOS = new Set([
  'id_1777036987296_167', 'id_1777036997520_57', 'id_1777037021472_880',
  'id_1777037077928_344', 'id_1777037105104_937', 'id_1777037122272_197',
  'id_1777037153097_945', 'id_1777037177272_837', 'id_1777037192632_334'
]);
const NOME_VELHO = /^(Frente|Costa|Mangas) PARTE [123] Blusa Moletom Tricolor$/;

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
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  const data = l[0].data, updatedAt = l[0].updated_at;
  // Cada chave do blob pode estar gravada como texto JSON: le e devolve igual.
  const ler = k => (typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]);
  const pôr = (k, v) => { data[k] = (typeof data[k] === 'string') ? JSON.stringify(v) : v; };

  const desenhos = ler('desenhos');
  const ordens = ler('ordens');
  const d = desenhos.find(x => x.id === DESENHO_ID);
  if (!d) throw new Error('desenho 0049 nao encontrado');
  const antesD = d.componentes.length;
  const saiD = d.componentes.filter(c => VELHOS.has(c.componenteId));
  d.componentes = d.componentes.filter(c => !VELHOS.has(c.componenteId));
  console.log(`desenho ${d.codigo} ${d.desc}: ${antesD} -> ${d.componentes.length} componentes`);
  saiD.forEach(c => console.log('   sai', c.nome));
  if (saiD.length !== 9) throw new Error('esperava 9 componentes antigos no desenho, achei ' + saiD.length);

  for (const num of OS_ALVO) {
    const o = ordens.find(x => x.os === num && x.desenhoId === DESENHO_ID);
    if (!o) throw new Error('OS ' + num + ' nao encontrada no desenho 0049');
    const antes = o.componentes.length;
    const sai = o.componentes.filter(c => NOME_VELHO.test(c.nome || ''));
    if (sai.length !== 9) throw new Error(`OS ${num}: esperava 9 componentes antigos, achei ${sai.length}`);
    o.componentes = o.componentes.filter(c => !NOME_VELHO.test(c.nome || ''));
    console.log(`OS ${num}: ${antes} -> ${o.componentes.length} componentes (saem ${[...new Set(sai.map(c => c.corNome))].join(', ')})`);
  }

  if (!GRAVAR) { console.log('\n(so mostrou; rode com --gravar para gravar)'); return; }

  const bk = path.join(RAIZ, 'backups', `shared_data-antes-bmtri-azul-componentes-${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
  fs.writeFileSync(bk, JSON.stringify(l[0]));
  console.log('backup:', bk);

  pôr('desenhos', desenhos);
  pôr('ordens', ordens);
  const r = await fetch(`${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
    method: 'PATCH',
    headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
    body: JSON.stringify({ data, updated_at: new Date().toISOString() })
  });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  try {
    await fetch(`${SUPA}/rest/v1/sync_signal?on_conflict=id`, {
      method: 'POST',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates' }),
      body: JSON.stringify({ id: 'main', updated_at: new Date().toISOString(), device_id: 'script-bmtri-azul' })
    });
  } catch (e) { /* o polling do app acha a mudanca */ }
  console.log('gravado.');
})().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
