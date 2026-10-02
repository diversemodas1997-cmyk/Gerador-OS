/*
 * COR CRUZADA NA RIBANA DE 60 CM.
 *
 * Junior, 02/10/2026: "todas as bobinas com 60cm sao ribana" / "corrija".
 * As seis entradas de 60 cm ja estao em Ribana Malha Algodao; uma delas (30/09,
 * 3 bobinas / 24 kg) tinha a cor "Vermelho Ribana Moletom". Passa para
 * "Vermelho Ribana Malha Algodao" -- kg, bobinas, data e id ficam.
 *
 *   node servidor/ribana-60-cor-vermelho-0210.js            (so mostra)
 *   node servidor/ribana-60-cor-vermelho-0210.js --gravar   (grava)
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


const ID = 'id_1790793712727_644';
const DE = 'Vermelho Ribana Moletom', PARA = 'Vermelho Ribana Malha Algodão';
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  if (!JSON.parse(data.cores).some(c => c.nome === PARA)) throw new Error('cor ' + PARA + ' nao existe');
  const movs = JSON.parse(data.estoqueMov);
  const m = movs.find(x => x.id === ID);
  if (!m) throw new Error('entrada nao encontrada');
  if (m.corNome === PARA) { console.log('ja corrigida -- nada a fazer'); return; }
  if (m.corNome !== DE || m.tecidoNome !== 'Ribana Malha Algodão' || Number(m.largura) !== 60) throw new Error('a entrada mudou: ' + JSON.stringify(m));
  console.log('antes : ' + JSON.stringify(m));
  const novoM = Object.assign({}, m, { corNome: PARA, obs: [m.obs, 'Cor era "' + DE + '" -- corrigida em 02/10/2026'].filter(Boolean).join(' · ') });
  console.log('depois: ' + JSON.stringify(novoM));
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-ribana-60-vermelho-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(movs.map(x => x.id === ID ? novoM : x)) }), updatedAt);
  const m2 = JSON.parse((await lerBlob(cab)).data.estoqueMov).find(x => x.id === ID);
  console.log('gravado. No servidor: ' + m2.tecidoNome + ' · ' + m2.corNome + ' · ' + m2.largura + ' cm');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
