/*
 * ETIQUETAS DE TAMANHO INFANTIL NO ESTOQUE DE AVIAMENTOS.
 *
 * Junior, 30/09/2026: "Faça o cadastro de etiquetas tamanho 2,4,6,8,10,12,14,16
 * com 200 unidades cada no estoque de aviamentos".
 *
 * Oito entradas, no mesmo formato das etiquetas P ao G3 lançadas em 24/09:
 * item Etiqueta, cor Preto (AVIAMENTO_COR_ETIQUETA, a que as OS reservam),
 * Unidade Descalvado, 200 unidades, sem kg. Não repete: se a entrada deste
 * script já existe, nada é feito.
 *
 *   node servidor/etiquetas-infantis-3009.js            (so mostra)
 *   node servidor/etiquetas-infantis-3009.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const TAMANHOS = ['2', '4', '6', '8', '10', '12', '14', '16'];
const QTD = 200;
const DATA = '2026-09-30';

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
  const movs = arr(data.aviamentosMov);
  const idDe = t => 'etiqueta-infantil-' + t + '-2026-09-30';
  const faltam = TAMANHOS.filter(t => !movs.some(m => m.id === idDe(t)));
  if (!faltam.length) { console.log('  ja lancadas antes -- nada a fazer'); return; }
  const agora = new Date().toISOString();
  const novas = faltam.map(t => ({
    id: idDe(t), tipo: 'entrada', unidade: 'desc', item: 'Etiqueta', tam: t, cor: 'Preto',
    kg: 0, qtd: QTD, data: DATA, obs: 'Cadastro das etiquetas de tamanho infantil',
    por: 'script (pedido do Junior)', em: agora
  }));
  const saldo = t => movs.filter(m => m.item === 'Etiqueta' && m.tam === t && m.tipo === 'entrada')
    .reduce((a, m) => a + (Number(m.qtd) || 0), 0);
  novas.forEach(n => console.log(`  Etiqueta ${n.tam.padStart(2)} · Preto · Descalvado · +${n.qtd} un   (antes: ${saldo(n.tam)} un de entrada)`));
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-etiquetas-infantis-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { aviamentosMov: JSON.stringify(movs.concat(novas)) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.aviamentosMov);
  console.log('gravado. Entradas infantis agora: ' + TAMANHOS.map(t => t + '=' + d2.filter(m => m.id === idDe(t)).reduce((a, m) => a + m.qtd, 0)).join(' '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
