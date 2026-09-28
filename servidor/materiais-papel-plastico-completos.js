/*
 * PAPEL E PLASTICO EM CADASTROS > MATERIAIS COM AS INFORMACOES COMPLETAS.
 *
 * Junior, 28/09/2026: "As informacoes do papel e plastico devem ser completas,
 * igual esta no estoque de materiais". O cadastro tem codigo, descricao e tipo;
 * ele passa a trazer o que o Estoque de materiais sabe de cada um, lido de la:
 *   descricao = a especificacao do estoque;
 *   tipo      = o tipo + o setor + a medida + a observacao (o rendimento).
 *
 *   node servidor/materiais-papel-plastico-completos.js            (so mostra)
 *   node servidor/materiais-papel-plastico-completos.js --gravar   (grava)
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


const PARES = [
  { cad: 'Papel kraft aerado', tipo: 'Papel' },
  { cad: 'Filme plástico', tipo: 'Plástico' }
];
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const mats = ler('materiais'), est = ler('materiaisEstCad');
  let n = 0;
  PARES.forEach(p => {
    const c = mats.find(m => m.categoria === 'material' && (normNome(m.desc) === normNome(p.cad) || normNome(m.desc).startsWith(normNome(p.cad))));
    const e = est.find(x => normNome(x.nome) === normNome(p.cad) && x.unidade !== 'sc');
    if (!c || !e) { console.log('  NAO ACHEI: ' + p.cad + (c ? '' : ' (cadastro)') + (e ? '' : ' (estoque)')); return; }
    const desc = e.desc || e.nome;
    const tipo = [p.tipo, e.setor ? 'setor ' + e.setor : '', e.medida === 'm' ? 'em metros' : '', e.obs || ''].filter(Boolean).join(' · ');
    if (c.desc === desc && c.tipo === tipo) { console.log('  ja completo: ' + p.cad); return; }
    console.log('  ' + c.codigo + ':\n     descricao "' + c.desc + '" -> "' + desc + '"\n     tipo      "' + c.tipo + '" -> "' + tipo + '"');
    c.desc = desc; c.tipo = tipo;
    n++;
  });
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-materiais-completos-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiais: JSON.stringify(mats) }), updatedAt);
  const m2 = JSON.parse((await lerBlob(cab)).data.materiais);
  m2.filter(m => m.categoria === 'material').forEach(m => console.log('gravado: ' + m.codigo + ' · ' + m.desc + ' | ' + m.tipo));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
