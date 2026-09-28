/*
 * TIRA O FORNECEDOR DA ESPECIFICACAO DOS FIOS E LINHAS.
 *
 * Junior, 28/09/2026: "tire o fornecedor da especificacao dos fios e linhas"
 * (o mesmo que ja valeu para o papel e o filme: "os fornecedores podem mudar").
 * Fica so o que descreve o material.
 *
 *   node servidor/fios-linhas-sem-fornecedor.js            (so mostra)
 *   node servidor/fios-linhas-sem-fornecedor.js --gravar   (grava)
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


const DESC = {
  'fio|tp200': '100% poliéster 150 · Tex 27',
  'linha|br02c': '100% poliéster · 120 · 1828 m · 29,00 Tex'
};
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  let tipos = [];
  try { tipos = typeof data.aviamentoTipos === 'string' ? JSON.parse(data.aviamentoTipos) : (data.aviamentoTipos || []); } catch (e) { tipos = []; }
  const agora = new Date().toISOString();
  const mudou = {};
  tipos.forEach(t => {
    const nova = DESC[normNome(t.item) + '|' + normNome(t.codigo)];
    if (!nova || t.desc === nova) return;
    const k = t.item + ' ' + t.codigo;
    if (!mudou[k]) { mudou[k] = 0; console.log('  ' + k + ':\n     "' + t.desc + '"\n  -> "' + nova + '"'); }
    mudou[k]++;
    t.desc = nova; t.atualizadoPor = 'Junior'; t.atualizadoEm = agora;
  });
  const n = Object.values(mudou).reduce((a, b) => a + b, 0);
  console.log('\n' + n + ' cadastro(s) a corrigir.');
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-fios-sem-fornecedor-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos) }), updatedAt);
  const t2 = JSON.parse((await lerBlob(cab)).data.aviamentoTipos);
  const resto = t2.filter(t => /cnpj|kalina|fabricante/i.test(t.desc || '')).length;
  console.log('gravado. Cadastros ainda com fornecedor: ' + resto);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
