/*
 * RIBANA E 60 CM: TIRA O 120 GRAVADO NOS LANCAMENTOS DE RIBANA.
 *
 * Junior, 02/10/2026: "todas as bobinas com 60cm sao ribana" e "Corrija as
 * entradas de bobinas com 120cm, pois todas sao malha algodao, nao sao
 * ribana". Escolheu: ribana e sempre 60 cm.
 *
 * O app passou a ler ribana sem largura como 60 cm (larguraPadraoDoTecido).
 * Os lancamentos manuais de ribana tinham largura 120 GRAVADA (saldo inicial,
 * acertos, transferencias -- largura-120-estoque-3009.js), entao ficariam na
 * linha de 120: passam a 60. Junto, as duas cores cruzadas das entradas de
 * 23/09 vao para a cor que o quadro ja mostra (corCanonicaPorTecido).
 * Nenhum kg muda: confere que o total de cada cor e o mesmo antes e depois.
 *
 *   node servidor/ribana-60cm-0210.js            (so mostra)
 *   node servidor/ribana-60cm-0210.js --gravar   (grava)
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

const ehRibana = t => /^\s*ribana/i.test(String(t || ''));
function quadro(data) {
  const app = appCom(data);
  const n = vm.runInContext('_normNome', app);
  const pl = vm.runInContext('estoquePorLargura()', app);
  const tot = new Map(), r120 = [];
  for (const [k, porL] of pl) {
    let t = 0; for (const [l, v] of porL) { t += v.kg; if (/^ribana/.test(k) && l === 120) r120.push(k); }
    tot.set(k, Math.round(t * 1000) / 1000);
  }
  return { app, n, tot, r120 };
}

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = JSON.parse(data.estoqueMov);
  const { app, tot: antes, r120: r120antes } = quadro(data);
  const canon = vm.runInContext('corCanonicaPorTecido', app);
  let nLarg = 0; const cores = [];
  const novos = movs.map(m => {
    if (!ehRibana(m.tecidoNome) || m.origem === 'os') return m;
    let x = m;
    if (Number(m.largura) === 120) { x = Object.assign({}, x, { largura: 60 }); nLarg++; }
    const c = canon(m.corNome, m.tecidoNome);
    if (c && c !== m.corNome) {
      cores.push(`${m.data} ${m.tecidoNome} · "${m.corNome}" -> "${c}" (${m.kg} kg, ${m.fechados || 0} bob)`);
      x = Object.assign({}, x, { corNome: c, obs: [m.obs, 'Cor era "' + m.corNome + '" -- corrigida em 02/10/2026'].filter(Boolean).join(' · ') });
    }
    return x;
  });
  console.log(`linhas Ribana 120 cm no quadro antes: ${r120antes.length}`);
  console.log(`lancamentos de ribana 120 -> 60 cm: ${nLarg}`);
  console.log('cores cruzadas corrigidas:\n  ' + (cores.join('\n  ') || '(nenhuma)'));

  const novo = Object.assign({}, data, { estoqueMov: JSON.stringify(novos) });
  const { tot: depois, r120 } = quadro(novo);
  if (r120.length) throw new Error('ainda ha linha de ribana em 120 cm: ' + r120.join(', '));
  const dif = [...new Set([...antes.keys(), ...depois.keys()])].filter(k => (antes.get(k) || 0) !== (depois.get(k) || 0));
  if (dif.length) throw new Error('o total mudou em: ' + dif.map(k => k + ' ' + antes.get(k) + ' -> ' + depois.get(k)).join('; '));
  console.log('conferencia: nenhuma linha de ribana em 120 cm; total de cada cor igual');
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-ribana-60cm-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, novo, updatedAt);
  const { r120: r2 } = quadro((await lerBlob(cab)).data);
  console.log(`gravado. Linhas de ribana em 120 cm no servidor: ${r2.length}.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
