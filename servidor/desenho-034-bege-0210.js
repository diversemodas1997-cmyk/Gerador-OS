/*
 * DESENHO 034 "CAMISETA BASICA | BEGE" PASSA A TER COR BEGE.
 *
 * Junior, 02/10/2026: "Desenho tecnico Camiseta basica bege deve constar cor
 * bege". O 034 usava a cor "Malha Algodao cru" (apagada hoje, e por isso
 * apontado para Off-White em apagar-cor-algodao-cru-0210.js). Agora: cor
 * principal e componentes de malha -> "Bege Malha Algodao"; gola de ribana ->
 * "Bege Ribana Malha Algodao". Os componentes que ja eram Bege ficam.
 *
 *   node servidor/desenho-034-bege-0210.js            (so mostra)
 *   node servidor/desenho-034-bege-0210.js --gravar   (grava)
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


const BEGE_MALHA = 'id_1776885262694_96';        // Bege Malha Algodao
const BEGE_RIBANA = 'id_1784634012757_368';      // Bege Ribana Malha Algodao
const TEC_RIBANA = 'id_1777039139210_737';       // Ribana Malha Algodao
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const cores = JSON.parse(data.cores), desenhos = JSON.parse(data.desenhos);
  const nomeDe = id => (cores.find(c => c.id === id) || {}).nome;
  if (nomeDe(BEGE_MALHA) !== 'Bege Malha Algodão' || nomeDe(BEGE_RIBANA) !== 'Bege Ribana Malha Algodão') throw new Error('cores Bege nao conferem');
  const i = desenhos.findIndex(d => d.codigo === '034');
  if (i < 0) throw new Error('desenho 034 nao encontrado');
  const x = JSON.parse(JSON.stringify(desenhos[i]));
  if (!/bege/i.test(x.desc || '')) throw new Error('o 034 nao e o Bege: ' + x.desc);
  console.log(`desenho 034 "${x.desc}"`);
  if (x.corPrincipalId !== BEGE_MALHA) { console.log(`  cor principal: ${nomeDe(x.corPrincipalId)} -> Bege Malha Algodão`); x.corPrincipalId = BEGE_MALHA; }
  (x.componentes || []).forEach(c => {
    const alvo = c.tecidoId === TEC_RIBANA ? BEGE_RIBANA : BEGE_MALHA;
    if (c.corId !== alvo) { console.log(`  ${c.nome}: ${nomeDe(c.corId)} -> ${nomeDe(alvo)}`); c.corId = alvo; }
  });
  desenhos[i] = x;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-desenho-034-bege-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { desenhos: JSON.stringify(desenhos) }), updatedAt);
  const y = JSON.parse((await lerBlob(cab)).data.desenhos).find(d => d.codigo === '034');
  console.log(`gravado. 034: principal ${nomeDe(y.corPrincipalId)}; componentes: ` + y.componentes.map(c => c.nome + '=' + nomeDe(c.corId)).join(', '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
