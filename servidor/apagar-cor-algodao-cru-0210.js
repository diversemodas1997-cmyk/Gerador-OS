/*
 * APAGA O CADASTRO DA COR "MALHA ALGODAO CRU" (codigo 0047).
 *
 * Junior, 02/10/2026: "Apague o cadastro de algodao-cru". Algodao cru virou
 * Off-White no estoque (malha e ribana) mais cedo no mesmo dia.
 *
 * O unico registro que ainda usava a cor e o desenho 034 (cor principal e os
 * cinco componentes). Antes de apagar, ele passa para as cores Off-White do
 * mesmo tecido de cada componente: malha -> "Off-White Malha Algodao", gola de
 * ribana -> "Off-White Ribana Malha Algodao". Sem isso o desenho ficaria
 * apontando para uma cor que nao existe.
 *
 *   node servidor/apagar-cor-algodao-cru-0210.js            (so mostra)
 *   node servidor/apagar-cor-algodao-cru-0210.js --gravar   (grava)
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


const ID_CRU = 'id_1788369537414_773';
const OFF_MALHA = 'id_1776952218144_555';        // Off-White Malha Algodao
const OFF_RIBANA = 'id_1790163012429_291';       // Off-White Ribana Malha Algodao
const TEC_RIBANA = 'id_1777039139210_737';       // Ribana Malha Algodao
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const cores = JSON.parse(data.cores), desenhos = JSON.parse(data.desenhos);
  const cru = cores.find(c => c.id === ID_CRU);
  if (!cru) { console.log('cor ja apagada -- nada a fazer'); return; }
  if (cru.nome !== 'Malha Algodão cru') throw new Error('a cor mudou: ' + cru.nome);
  const nomeDe = id => (cores.find(c => c.id === id) || {}).nome;
  if (nomeDe(OFF_MALHA) !== 'Off-White Malha Algodão' || nomeDe(OFF_RIBANA) !== 'Off-White Ribana Malha Algodão') throw new Error('cores Off-White nao conferem');
  const troca = (tecidoId) => tecidoId === TEC_RIBANA ? OFF_RIBANA : OFF_MALHA;
  const novosDes = desenhos.map(d => {
    if (!JSON.stringify(d).includes(ID_CRU)) return d;
    const x = JSON.parse(JSON.stringify(d));
    if (x.corPrincipalId === ID_CRU) { x.corPrincipalId = OFF_MALHA; console.log(`desenho ${x.codigo}: cor principal -> Off-White Malha Algodao`); }
    (x.componentes || []).forEach(c => { if (c.corId === ID_CRU) { c.corId = troca(c.tecidoId); console.log(`desenho ${x.codigo}: ${c.nome} -> ${nomeDe(c.corId)}`); } });
    return x;
  });
  const novo = Object.assign({}, data, {
    cores: JSON.stringify(cores.filter(c => c.id !== ID_CRU)),
    desenhos: JSON.stringify(novosDes)
  });
  const sobra = Object.keys(novo).filter(k => String(novo[k]).includes(ID_CRU));
  if (sobra.length) throw new Error('o id da cor ainda aparece em: ' + sobra.join(', '));
  console.log(`cadastro de cores: ${cores.length} -> ${cores.length - 1} (apaga "${cru.nome}", codigo ${cru.codigo})`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-apagar-cor-cru-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, novo, updatedAt);
  const d2 = (await lerBlob(cab)).data;
  console.log(`gravado. Cores no cadastro: ${JSON.parse(d2.cores).length}; id da cor crua restante: ${Object.keys(d2).filter(k => String(d2[k]).includes(ID_CRU)).length} chave(s).`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
