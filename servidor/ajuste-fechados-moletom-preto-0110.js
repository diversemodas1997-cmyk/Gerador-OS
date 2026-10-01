/*
 * AJUSTE DE -9 BOBINAS FECHADAS NO MOLETOM PRETO (120 cm).
 *
 * Junior, 01/10/2026: "lance o ajuste de -9 bobinas no moletom preto".
 *
 * A coluna Fechados do Moletom Preto marcava 11 -- as 11 bobinas da entrada de
 * 17/09 --, e desde entao as OS 0547, 0548, 0589 e 0592 baixaram 164,2 kg
 * (~9 bobinas de 18 kg). Essas baixas sao de antes da regra que faz a baixa da
 * OS descontar os fechados (01/10/2026), entao nao descontaram. O ajuste e uma
 * SAIDA MANUAL de 0 kg e 9 fechados: mexe so na contagem de bobinas, nao no kg
 * (o kg das OS ja esta baixado).
 *
 *   node servidor/ajuste-fechados-moletom-preto-0110.js            (so mostra)
 *   node servidor/ajuste-fechados-moletom-preto-0110.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

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
const ID = 'ajuste_fech_moletom_preto_20261001';
const AJUSTE = {
  id: ID, tipo: 'saida', tecidoNome: 'Moletom', corNome: 'Preto Moletom',
  kg: 0, fechados: 9, abertos: 0, largura: 120, data: '2026-10-01',
  origem: 'manual', osId: '', osNumero: '',
  obs: 'Ajuste de bobinas fechadas: baixas das OS 0547, 0548, 0589 e 0592 (164,2 kg ~ 9 bobinas de 18 kg) feitas antes da baixa descontar os fechados'
};
const arr = v => { try { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; } catch (e) { return []; } };
const ehMoletomPreto = m => String(m.tecidoNome).trim().toLowerCase() === 'moletom'
  && String(m.corNome).trim().toLowerCase() === 'preto moletom';
const fechados = movs => movs.filter(ehMoletomPreto)
  .reduce((s, m) => s + (m.tipo === 'entrada' ? 1 : -1) * (parseInt(m.fechados) || 0), 0);

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const movs = arr(data.estoqueMov);
  console.log(`Moletom Preto -- fechados hoje: ${fechados(movs)}`);
  if (movs.some(m => m.id === ID)) { console.log('o ajuste ja esta lancado -- nada a fazer'); return; }
  console.log(`depois do ajuste: ${fechados(movs.concat([AJUSTE]))}`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-ajuste-fechados-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(movs.concat([AJUSTE])) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.estoqueMov);
  console.log(`gravado. Moletom Preto -- fechados agora: ${fechados(d2)}`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
