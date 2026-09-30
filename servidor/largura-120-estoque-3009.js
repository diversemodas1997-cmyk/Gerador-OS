/*
 * LARGURA 120 CM NOS LANCAMENTOS DE TECIDO QUE NAO TEM LARGURA.
 *
 * Junior, 30/09/2026: "Corrija os nomes das bobinas do estoque de tecidos que
 * nao tem a largura no nome, inserindo 120 cm para todos que estao sem essa
 * informacao, inclusive para os que estao sem largura no texto".
 *
 * Grava largura: 120 em todo lancamento de estoque de tecido que nao e de OS
 * (entrada e saida manual, entrada por NF) e ainda nao tem largura. Os de OS
 * (reserva e baixa) nao guardam largura: o programa os le como 120 cm (ver
 * LARGURA_BOBINA_PADRAO_CM no app.js). O algodao cru (119/117/115/80) fica como
 * esta, porque ja tem largura.
 *
 *   node servidor/largura-120-estoque-3009.js            (so mostra)
 *   node servidor/largura-120-estoque-3009.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const LARGURA = 120;

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
const arr = v => { try { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; } catch (e) { return []; } };

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = arr(data.estoqueMov);
  const alvo = m => m.origem !== 'os' && !(Number(m.largura) > 0);
  const muda = movs.filter(alvo);
  const porTipo = {};
  muda.forEach(m => { const k = (m.origem || '?') + '/' + m.tipo; porTipo[k] = (porTipo[k] || 0) + 1; });
  const tecidos = [...new Set(muda.map(m => m.tecidoNome))].sort();
  console.log(`lancamentos sem largura (fora OS): ${muda.length} de ${movs.length}`);
  console.log('  por origem/tipo:', JSON.stringify(porTipo));
  console.log('  tecidos:', tecidos.join(' | '));
  console.log(`  ja com largura: ${movs.filter(m => Number(m.largura) > 0).length}`);
  if (!muda.length) { console.log('  nada a fazer'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-largura-120-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const saida = movs.map(m => alvo(m) ? Object.assign({}, m, { largura: LARGURA }) : m);
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(saida) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.estoqueMov);
  console.log(`gravado. Sem largura fora de OS agora: ${d2.filter(alvo).length}.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
