/*
 * RIBANA "ALGODAO CRU" PASSA A SE CHAMAR OFF-WHITE.
 *
 * Junior, 02/10/2026: "troque o nome de ribana cor algodao cru para off-white".
 * Mesma troca feita na malha mais cedo (algodao-cru-para-off-white-0210.js).
 *
 * Renomeia a cor do cadastro (id_1790163012429_291, codigo 0013) de
 * "Algodao Cru Ribana Malha Algodao" para "Off-White Ribana Malha Algodao" e
 * leva o nome novo aos lancamentos do estoque que guardaram o antigo. Id,
 * codigo, sigla e kg ficam. O nome antigo nao aparece em mais nenhuma chave.
 *
 *   node servidor/ribana-cru-para-off-white-0210.js            (so mostra)
 *   node servidor/ribana-cru-para-off-white-0210.js --gravar   (grava)
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


const ID_COR = 'id_1790163012429_291';
const DE = 'Algodão Cru Ribana Malha Algodão', PARA = 'Off-White Ribana Malha Algodão';
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const cores = JSON.parse(data.cores), movs = JSON.parse(data.estoqueMov);
  const cor = cores.find(c => c.id === ID_COR);
  if (!cor) throw new Error('cor nao encontrada');
  if (cor.nome === PARA) { console.log('ja renomeada -- nada a fazer'); return; }
  if (cor.nome !== DE) throw new Error('a cor mudou: ' + cor.nome);
  if (cores.some(c => c.id !== ID_COR && c.nome === PARA)) throw new Error('ja existe outra cor ' + PARA);
  const alvo = movs.filter(m => m.corNome === DE);
  console.log(`cadastro: "${DE}" -> "${PARA}" (codigo ${cor.codigo})`);
  alvo.forEach(m => console.log(`  ${m.data} ${m.tipo}/${m.origem} ${m.osNumero || ''} ${m.kg} kg`));
  const novo = Object.assign({}, data, {
    cores: JSON.stringify(cores.map(c => c.id === ID_COR ? Object.assign({}, c, { nome: PARA }) : c)),
    estoqueMov: JSON.stringify(movs.map(m => m.corNome === DE ? Object.assign({}, m, { corNome: PARA }) : m))
  });
  const sobra = Object.keys(novo).filter(k => String(novo[k]).includes(DE));
  if (sobra.length) throw new Error('o nome antigo ainda aparece em: ' + sobra.join(', '));
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-ribana-cru-off-white-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, novo, updatedAt);
  const d2 = (await lerBlob(cab)).data;
  console.log(`gravado. Cadastro: ${JSON.parse(d2.cores).find(c => c.id === ID_COR).nome}; ` +
    `lancamentos com o nome novo: ${JSON.parse(d2.estoqueMov).filter(m => m.corNome === PARA).length}; ` +
    `nome antigo restante: ${Object.keys(d2).filter(k => String(d2[k]).includes(DE)).length} chave(s).`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
