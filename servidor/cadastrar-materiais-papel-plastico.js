/*
 * CADASTRA O PAPEL E O PLASTICO EM CADASTROS > MATERIAIS.
 *
 * Junior, 28/09/2026: "Insira em Materiais na barra lateral o cadastro de
 * materiais papel e plastico". Os dois que ja estao no Estoque de materiais
 * (Papel kraft aerado e Filme plastico, os do enfesto) ganham o cadastro da
 * categoria 'material' em STATE.materiais, com o mesmo nome — e o nome que o
 * Estoque de materiais sugere a partir do cadastro. Proximo codigo livre de 3
 * digitos. Quem ja existe fica.
 *
 *   node servidor/cadastrar-materiais-papel-plastico.js            (so mostra)
 *   node servidor/cadastrar-materiais-papel-plastico.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = 'J:/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json';
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const GRAVAR = process.argv.includes('--gravar');

const normNome = s => (s || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
async function conectar() {
  const { email, password } = JSON.parse(fs.readFileSync(CREDS, 'utf8'));
  const anon = JSON.parse(fs.readFileSync(LOCAL, 'utf8').replace(/^\uFEFF/, '')).key;
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


const NOVOS = [
  { desc: 'Papel kraft aerado', tipo: 'Papel' },
  { desc: 'Filme plástico', tipo: 'Plástico' }
];
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  let mats = [];
  try { mats = typeof data.materiais === 'string' ? JSON.parse(data.materiais) : (data.materiais || []); } catch (e) { mats = []; }
  let prox = Math.max(0, ...mats.map(m => parseInt(String(m.codigo || '').replace(/\D/g, ''), 10) || 0)) + 1;
  const agora = new Date().toISOString();
  let n = 0;
  NOVOS.forEach((x, i) => {
    if (mats.some(m => normNome(m.desc) === normNome(x.desc))) { console.log('  ja existe: ' + x.desc); return; }
    const codigo = String(prox++).padStart(3, '0');
    mats.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      codigo, desc: x.desc, tipo: x.tipo, categoria: 'material' });
    n++;
    console.log('  novo: ' + codigo + ' · ' + x.desc + ' (' + x.tipo + ')');
  });
  console.log('\n' + n + ' cadastro(s) novo(s).');
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-materiais-papel-plastico-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiais: JSON.stringify(mats) }), updatedAt);
  const m2 = JSON.parse((await lerBlob(cab)).data.materiais);
  console.log('gravado. Em Cadastros > Materiais: ' + m2.filter(m => m.categoria === 'material').map(m => m.codigo + ' ' + m.desc).join(', '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
