/*
 * ENTRADAS DE ALGODAO CRU VIRAM OFF-WHITE.
 *
 * Junior, 02/10/2026: "altere as entradas do estoque de tecidos, transformando
 * entradas de algodao cru para entradas de Off-white".
 *
 * Sao as entradas manuais de Malha Algodao com cor "Algodao cru" (119, 117,
 * 115 e 80 cm, de 04/09) e "Malha Algodao cru" (30/09). Todas passam para a
 * cor "Off-White Malha Algodao" -- largura, bobinas, kg, data e id ficam.
 * Nenhuma OS consumiu essas cores, entao nao ha saida para mover junto.
 *
 * A ribana ("Algodao Cru Ribana Malha Algodao") NAO entra: nao existe cor
 * Off-White de Ribana Malha Algodao, e as OS que gastaram essa ribana ficariam
 * com a saida sem a entrada.
 *
 *   node servidor/algodao-cru-para-off-white-0210.js            (so mostra)
 *   node servidor/algodao-cru-para-off-white-0210.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');
const DESTINO = 'Off-White Malha Algodão';

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

const arr = s => JSON.parse(s || '[]');
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const ehCru = m => m.tipo === 'entrada' && m.origem !== 'os'
  && norm(m.tecidoNome) === 'malha algodao' && /\bcru\b/.test(norm(m.corNome));
const linha = m => `    ${m.data}  ${String(m.fechados || 0).padStart(2)} bob  ${String(m.largura || '-').padStart(3)} cm  ${String(m.kg).padStart(8)} kg   ${m.corNome}`;

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  if (!arr(data.cores).some(c => c.nome === DESTINO)) throw new Error('a cor "' + DESTINO + '" nao existe no cadastro');
  const movs = arr(data.estoqueMov);
  if (movs.some(m => m.tipo !== 'entrada' && norm(m.tecidoNome) === 'malha algodao' && /\bcru\b/.test(norm(m.corNome)))) {
    throw new Error('ha saida/ajuste de Malha Algodao cru -- analisar antes de mover so as entradas');
  }
  const alvo = movs.filter(ehCru);
  if (!alvo.length) { console.log('  Nenhuma entrada de algodao cru -- nada a fazer.'); return; }

  const kg = l => Math.round(l.reduce((a, m) => a + (+m.kg || 0), 0) * 1000) / 1000;
  const bob = l => l.reduce((a, m) => a + (+m.fechados || 0), 0);
  console.log(`ENTRADAS DE ALGODAO CRU (${alvo.length}) -> "${DESTINO}":\n` + alvo.map(linha).join('\n'));
  console.log(`    total: ${bob(alvo)} bobinas, ${kg(alvo)} kg`);

  const saida = movs.map(m => !ehCru(m) ? m : Object.assign({}, m, {
    corNome: DESTINO,
    obs: [m.obs, 'Era "' + m.corNome + '" -- passou para Off-White em 02/10/2026'].filter(Boolean).join(' · ')
  }));
  if (saida.length !== movs.length) throw new Error('contagem de movimentos mudou');
  const off = l => l.filter(m => m.tipo === 'entrada' && m.corNome === DESTINO);
  console.log(`\nEntradas Off-White Malha Algodao: ${off(movs).length} (${kg(off(movs))} kg) -> ${off(saida).length} (${kg(off(saida))} kg)`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cru-para-off-white-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { estoqueMov: JSON.stringify(saida) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const m2 = arr(d2.estoqueMov);
  console.log(`gravado. Agora: ${m2.filter(ehCru).length} entradas de cru, ${off(m2).length} entradas Off-White (${kg(off(m2))} kg); ` +
    `${arr(d2.ordens).length} OS e ${m2.length} movimentos no blob.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
