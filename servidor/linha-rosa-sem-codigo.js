/*
 * TIRA O CODIGO DA COR DA LINHA ROSA.
 *
 * Junior, 28/09/2026: "Retire o codigo de cor da linha rosa, pois nao se sabe
 * qual o numero desse codigo". O 000680 veio da primeira lista; a folha de
 * estoque traz a Rosa sem numero. Sai do cadastro e dos lancamentos dela.
 *
 *   node servidor/linha-rosa-sem-codigo.js            (so mostra)
 *   node servidor/linha-rosa-sem-codigo.js --gravar   (grava)
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


const OBS = 'Contagem 28/09 (folha do caderno)';
const DATA = '2026-09-28';
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const tipos = ler('aviamentoTipos'), movs = ler('aviamentosMov');
  const eRosa = x => normNome(x.item) === 'linha' && normNome(x.codigo || x.modelo) === 'br02c';
  const cads = tipos.filter(t => eRosa(t) && normNome(t.corNome) === 'rosa' && t.corCodigo);
  const ms = movs.filter(m => normNome(m.item) === 'linha' && normNome(m.modelo) === 'br02c' && normNome(m.cor) === 'rosa' && m.corCodigo);
  cads.forEach(t => { console.log('  cadastro Linha BR02C · Rosa: cod. ' + t.corCodigo + ' -> (sem)'); t.corCodigo = ''; t.atualizadoPor = 'Junior'; t.atualizadoEm = new Date().toISOString(); });
  ms.forEach(m => { console.log('  lancamento ' + m.tipo + ' ' + (m.qtd || 0) + ' un: cod. ' + m.corCodigo + ' -> (sem)'); m.corCodigo = ''; });
  if (!cads.length && !ms.length) { console.log('  nada a fazer'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-linha-rosa-sem-codigo-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos), aviamentosMov: JSON.stringify(movs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const t2 = JSON.parse(d2.aviamentoTipos).filter(t => eRosa(t) && normNome(t.corNome) === 'rosa');
  console.log('gravado. Linha Rosa: ' + t2.map(t => 'cod. ' + (t.corCodigo || '(sem)')).join(', '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
