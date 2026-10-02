/*
 * ZERA TODOS OS SALDOS NEGATIVOS DO ESTOQUE DE TECIDOS, LINHA POR LINHA
 * (cor + largura, como o quadro mostra).
 *
 * Junior, 02/10/2026: "Corrija todos os saldos negativos para zero".
 *
 * Para cada largura negativa de uma cor:
 *   1. se a MESMA cor tem sobra em outra largura, TRANSFERE dali (entrada na
 *      negativa + saida igual na que sobra) -- e o consumo das OS que caiu
 *      numa largura sem bobina, como o Off-White 117 cm; nao inventa pano;
 *   2. o que ainda faltar e buraco de verdade (compra nao lancada) e vira
 *      ENTRADA de acerto do tamanho exato, obs "Ajuste: ...".
 * Todo lancamento leva ajuste:true, e a Ultima entrada do quadro os ignora.
 * Saldo calculado pelo proprio app.js numa VM com o blob ao vivo.
 *
 *   node servidor/zerar-negativos-por-largura.js            (so mostra)
 *   node servidor/zerar-negativos-por-largura.js --gravar   (grava)
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

/* Saldo por largura de cada cor, pela conta do app (estoquePorLargura). */
function saldos(data) {
  const app = appCom(data);
  const n = vm.runInContext('_normNome', app);
  const det = vm.runInContext('calcularSaldosEstoque()', app).detalhe;
  const pl = vm.runInContext('estoquePorLargura()', app);
  return det.map(d => ({ tecidoNome: d.tecidoNome, corNome: d.corNome,
    larg: Array.from((pl.get(n(d.tecidoNome) + '||' + n(d.corNome)) || new Map()).entries())
      .map(([l, v]) => ({ l, kg: v.kg })) }));
}
const kg3 = n => Math.round(n * 1000) / 1000;
const HOJE = '2026-10-02';

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = JSON.parse(data.estoqueMov || '[]');
  const novos = [];
  let seq = 0;
  const id = () => 'zera-neg-' + Date.now().toString(36) + '-' + (seq++);
  const base = (c, tipo, l, kg, obs) => ({ id: id(), tipo, tecidoNome: c.tecidoNome, corNome: c.corNome,
    kg: kg3(kg), fechados: 0, abertos: 0, largura: l, data: HOJE, origem: 'manual', osId: '', osNumero: '', ajuste: true, obs });
  let kgEntrada = 0;
  saldos(data).forEach(c => {
    const pos = c.larg.filter(x => kg3(x.kg) > 0).sort((a, b) => b.kg - a.kg);
    c.larg.filter(x => kg3(x.kg) < 0).forEach(neg => {
      let falta = kg3(-neg.kg);
      // 1o: transferir do que a MESMA cor tem sobrando em outra largura.
      pos.forEach(p => {
        const t = kg3(Math.min(falta, p.kg));
        if (!(t > 0)) return;
        const obs = `Transferencia entre larguras: o consumo das OS caiu na ${neg.l} cm e as bobinas estao na ${p.l} cm (${HOJE})`;
        novos.push(base(c, 'entrada', neg.l, t, obs), base(c, 'saida', p.l, t, obs));
        console.log(`  transf ${String(t.toFixed(3)).padStart(8)} kg  ${p.l} -> ${neg.l} cm   ${c.tecidoNome} · ${c.corNome}`);
        p.kg -= t; falta = kg3(falta - t);
      });
      // 2o: o que sobrar e buraco de verdade -> entrada de acerto.
      if (falta > 0) {
        novos.push(base(c, 'entrada', neg.l, falta, 'Ajuste: entrada para zerar saldo negativo'));
        kgEntrada += falta;
        console.log(`  ENTRADA ${String(falta.toFixed(3)).padStart(7)} kg  ${neg.l} cm          ${c.tecidoNome} · ${c.corNome}`);
      }
    });
  });
  if (!novos.length) { console.log('Nenhum saldo negativo. Nada a fazer.'); return; }
  console.log(`\n${novos.length} lancamentos; entrada nova de pano: ${kg3(kgEntrada)} kg`);

  const novo = Object.assign({}, data, { estoqueMov: JSON.stringify(movs.concat(novos)) });
  const sobra = saldos(novo).flatMap(c => c.larg.filter(x => kg3(x.kg) < 0).map(x => c.corNome + ' ' + x.l));
  if (sobra.length) throw new Error('conferencia: ainda negativos: ' + sobra.join(', '));
  console.log('conferencia: nenhuma linha negativa depois do ajuste');
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-zerar-negativos-largura-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, novo, updatedAt);
  const resto = saldos((await lerBlob(cab)).data).flatMap(c => c.larg.filter(x => kg3(x.kg) < 0));
  console.log(`gravado. Linhas negativas no servidor agora: ${resto.length}.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
