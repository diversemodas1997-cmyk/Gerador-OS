/*
 * ZERA O DISPONIVEL DA MALHA ALGODAO OFF-WHITE DE 117 CM.
 *
 * Junior, 02/10/2026: "Corrija o disponivel off-white 117cm para zero".
 *
 * A linha 117 cm veio das 4 bobinas de algodao cru de 04/09 (72 kg), que
 * passaram a Off-White hoje (algodao-cru-para-off-white-0210.js). Com uma
 * bobina de 117 lancada na cor, a reserva/baixa das 20 OS de camiseta (grade
 * 117) deixou de cair na 120 e caiu toda na 117: -772,746 kg. Esse consumo
 * saiu, de fato, do pano de 120 cm que o Off-White sempre teve.
 *
 * O ajuste e uma TRANSFERENCIA entre larguras da mesma cor: entrada manual na
 * 117 e saida manual igual na 120. A 117 vai a zero, o total da cor nao muda,
 * nada e apagado. Sem fechados: as bobinas saem do disponivel.
 *
 * O saldo da linha e calculado pelo proprio app.js (estoquePorLargura),
 * carregado numa VM com o blob ao vivo: a mesma conta do quadro.
 *
 *   node servidor/zerar-off-white-117-0210.js            (so mostra)
 *   node servidor/zerar-off-white-117-0210.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const TECIDO = 'Malha Algodão', COR = 'Off-White Malha Algodão', LARG = 117;

async function conectar() {
  if (!CREDS) throw new Error('nao achei supa-creds.json no Google Drive');
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^﻿/, '')).key;
  const auth = await (await fetch(`${SUPA}/auth/v1/token?grant_type=password`, {
    method: 'POST', headers: { apikey: anon, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  })).json();
  if (!auth.access_token) throw new Error('login no servidor falhou');
  return { apikey: anon, Authorization: 'Bearer ' + auth.access_token };
}
async function lerBlob(cab) {
  const l = await (await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&select=data,updated_at`, { headers: cab })).json();
  if (!l || !l[0]) throw new Error('nao achei a linha main');
  return { data: l[0].data, updatedAt: l[0].updated_at };
}
async function gravarBlob(cab, data, updatedAt) {
  const r = await fetch(
    `${SUPA}/rest/v1/shared_data?id=eq.main&updated_at=eq.${encodeURIComponent(updatedAt)}`, {
      method: 'PATCH',
      headers: Object.assign({}, cab, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
      body: JSON.stringify({ data, updated_at: new Date().toISOString() })
    });
  const volta = await r.json();
  if (!r.ok) throw new Error('o servidor recusou: ' + JSON.stringify(volta).slice(0, 300));
  if (!Array.isArray(volta) || !volta.length) {
    throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  }
}

// O app.js inteiro numa VM com DOM de mentira, e o STATE vindo do blob.
function appCom(data) {
  const el = () => new Proxy({ innerHTML: '', style: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } }, dataset: {}, value: '', children: [] },
    { get: (t, k) => k in t ? t[k] : (typeof k === 'string' ? (() => el()) : undefined), set: (t, k, v) => (t[k] = v, true) });
  const doc = new Proxy({}, { get: (t, k) => k === 'querySelectorAll' ? () => [] : k === 'getElementById' ? () => el() : () => el() });
  const nada = () => {};
  const ctx = { document: doc, console: { log: nada, warn: nada, error: nada, info: nada }, localStorage: { getItem() { return null; }, setItem: nada },
    sessionStorage: { getItem() { return null; }, setItem: nada }, navigator: { userAgent: '' }, location: { origin: '', href: '', search: '', hash: '' },
    setTimeout: nada, setInterval: nada, clearTimeout: nada, clearInterval: nada, addEventListener: nada,
    fetch: async () => ({ ok: false, json: async () => ({}) }), matchMedia: () => ({ matches: false, addEventListener: nada }),
    requestAnimationFrame: nada, Intl, Date, Math, JSON };
  ctx.window = ctx; ctx.self = ctx; ctx.globalThis = ctx;
  vm.createContext(ctx);
  try { vm.runInContext(fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8'), ctx, { filename: 'app.js' }); } catch (e) { /* inicializacao de tela */ }
  vm.runInContext('(function(d){for(const k of Object.keys(d)){try{STATE[k]=JSON.parse(d[k])}catch(e){}}})', ctx)(data);
  return ctx;
}
const larguras = data => {
  const app = appCom(data);
  return vm.runInContext('estoquePorLargura()', app).get(
    vm.runInContext('_normNome', app)(TECIDO) + '||' + vm.runInContext('_normNome', app)(COR)) || new Map();
};
const linhaDe = data => larguras(data).get(LARG);
const totalDaCor = data => Array.from(larguras(data).values()).reduce((a, v) => a + v.kg, 0);
const kg3 = n => Math.round(n * 1000) / 1000;

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const l = linhaDe(data);
  if (!l) throw new Error('a linha ' + COR + ' · ' + LARG + ' cm nao existe');
  console.log(`${COR} · ${LARG} cm agora: entradas ${kg3(l.entrada)} kg, reservado ${kg3(l.reservado)}, saidas ${kg3(l.saida)}, ` +
    `DISPONIVEL ${kg3(l.kg)} kg, fechados lancados ${l.fechados}`);
  if (Math.abs(kg3(l.kg)) < 0.0005 && !l.fechados) { console.log('ja esta em zero -- nada a fazer'); return; }
  const movs = JSON.parse(data.estoqueMov || '[]');
  const kg = kg3(l.kg);
  const id = Date.now().toString(36);
  const base = { tecidoNome: TECIDO, corNome: COR, kg: Math.abs(kg), fechados: 0, abertos: 0,
    data: '2026-10-02', origem: 'manual', osId: '', osNumero: '' };
  const obs = 'Transferencia entre larguras: o consumo das OS de grade 117 cm saiu do pano de 120 cm (02/10/2026)';
  const ajuste = [
    Object.assign({ id: 'transf-117-' + id, tipo: kg < 0 ? 'entrada' : 'saida', largura: LARG }, base, { obs }),
    Object.assign({ id: 'transf-120-' + id, tipo: kg < 0 ? 'saida' : 'entrada', largura: 120 }, base, { obs })
  ];
  console.log('ajuste: ' + JSON.stringify(ajuste));
  const novo = Object.assign({}, data, { estoqueMov: JSON.stringify(movs.concat(ajuste)) });
  const depois = linhaDe(novo);
  console.log(`depois: DISPONIVEL ${kg3(depois.kg)} kg`);
  if (Math.abs(kg3(depois.kg)) >= 0.0005) throw new Error('a conferencia nao deu zero');
  const l120 = larguras(novo).get(120);
  console.log(`120 cm depois: DISPONIVEL ${kg3(l120.kg)} kg`);
  if (kg3(l120.kg) < 0) throw new Error('a 120 cm ficaria negativa -- a transferencia nao cabe');
  const t0 = kg3(totalDaCor(data)), t1 = kg3(totalDaCor(novo));
  console.log(`total da cor: ${t0} kg -> ${t1} kg`);
  if (t0 !== t1) throw new Error('o total da cor mudou');
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-zerar-off-white-117-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, novo, updatedAt);
  const l2 = linhaDe((await lerBlob(cab)).data);
  console.log(`gravado. No servidor agora: DISPONIVEL ${kg3(l2.kg)} kg.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
