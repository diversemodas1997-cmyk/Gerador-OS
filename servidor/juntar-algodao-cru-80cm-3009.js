/*
 * ALGODÃO CRU: UMA LINHA DE ENTRADA POR LARGURA.
 *
 * Junior, 30/09/2026: "Corrija a entrada das bobinas algodao cru, para que cada
 * item com medidas diferentes estejam em uma linha diferente, ou seja, precisa
 * existir 4 linhas com suas respectivas quantidades".
 *
 * Depois de corrigir-algodao-cru-3009.js havia CINCO entradas: 119, 117, 115 cm
 * e DUAS de 80 cm (as 40 bobinas de 04/09 e a correcao de 14 de hoje). As duas
 * de 80 cm viram uma so, na data e no id da original: 54 bobinas / 702 kg.
 * O total (68 bobinas / 954 kg) nao muda.
 *
 *   node servidor/juntar-algodao-cru-80cm-3009.js            (so mostra)
 *   node servidor/juntar-algodao-cru-80cm-3009.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const ID_40 = 'id_1788529911602_899';
const ID_CORRECAO = 'cru-80cm-correcao-2026-09-30';

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
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = arr(data.estoqueMov);
  const m40 = movs.find(m => m.id === ID_40), corr = movs.find(m => m.id === ID_CORRECAO);
  if (!corr && m40 && +m40.fechados === 54) { console.log('  ja juntado antes -- nada a fazer'); return; }
  if (!m40 || !corr) throw new Error('nao achei as duas entradas de 80 cm');
  if (+m40.fechados !== 40 || +m40.kg !== 520 || +corr.fechados !== 14 || +corr.kg !== 182) {
    throw new Error('as entradas de 80 cm mudaram desde a analise: ' + JSON.stringify([m40.fechados, m40.kg, corr.fechados, corr.kg]));
  }
  const junto = Object.assign({}, m40, { fechados: 54, kg: 702, largura: 80,
    obs: 'Bobinas de 80 cm. Contagem corrigida em 30/09: 54 bobinas (eram 40 lancadas), 13 kg cada' });
  const saida = movs.filter(m => m.id !== ID_CORRECAO).map(m => m.id === ID_40 ? junto : m);

  const cru = l => l.filter(m => m.origem !== 'os' && m.tipo === 'entrada'
    && norm(m.tecidoNome) === 'malha algodao' && /cru/.test(norm(m.corNome)));
  const resumo = l => cru(l).map(m => `    ${m.data}  ${String(m.fechados).padStart(2)} bob  ${String(m.largura || '-').padStart(3)} cm  ${String(m.kg).padStart(5)} kg`).join('\n');
  const soma = (l, k) => cru(l).reduce((a, m) => a + (+m[k] || 0), 0);
  console.log(`ANTES (${cru(movs).length} linhas):\n` + resumo(movs) + `\n    total: ${soma(movs, 'fechados')} bobinas, ${soma(movs, 'kg')} kg`);
  console.log(`DEPOIS (${cru(saida).length} linhas):\n` + resumo(saida) + `\n    total: ${soma(saida, 'fechados')} bobinas, ${soma(saida, 'kg')} kg`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-juntar-cru-80cm-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(saida) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.estoqueMov);
  console.log(`gravado. Agora: ${cru(d2).length} linhas, ${soma(d2, 'fechados')} bobinas, ${soma(d2, 'kg')} kg.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
