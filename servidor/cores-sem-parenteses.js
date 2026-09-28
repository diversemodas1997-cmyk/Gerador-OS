/*
 * TIRA O "(CODIGO)" DO NOME DAS CORES DE FIO E LINHA.
 *
 * Junior, 28/09/2026: "Corrija o nome do cadastro de fios e linhas e retire o
 * parenteses com o numero redundante de cor. Mantenha o texto cod. com a cor".
 * "Off White (002484)" volta a ser "Off White"; o codigo continua no campo
 * proprio, e a tela mostra "Off White · cod. 002484".
 *
 * Os lancamentos de fio e linha passam a guardar o CODIGO da cor (corCodigo):
 * com dois cadastros "Off White", e ele que diz de qual cone e cada entrada
 * (o app separa a linha do estoque por item + tipo + cor + codigo). O codigo
 * vem do cadastro cujo nome era igual ao da cor do lancamento, ANTES de tirar
 * o parenteses.
 *
 *   node servidor/cores-sem-parenteses.js            (so mostra)
 *   node servidor/cores-sem-parenteses.js --gravar   (grava)
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
  // 1) os lancamentos ganham o codigo da cor, pelo cadastro de mesmo nome (antes da troca)
  let nMov = 0;
  movs.forEach(m => {
    if (!m.modelo || !m.cor) return;
    const t = tipos.find(x => normNome(x.item) === normNome(m.item) && normNome(x.codigo) === normNome(m.modelo) && normNome(x.corNome) === normNome(m.cor));
    if (!t) { console.log('  AVISO: lancamento sem cadastro: ' + m.item + ' ' + m.modelo + ' ' + m.cor); return; }
    const base = String(t.corNome).replace(/\s*\(\s*[^)]*\)\s*$/, '').trim();
    const cod = t.corCodigo || '';
    if (m.cor === base && (m.corCodigo || '') === cod) return;
    m.cor = base; m.corCodigo = cod; nMov++;
  });
  // 2) os cadastros perdem o "(codigo)"
  let nCad = 0;
  tipos.forEach(t => {
    const mm = String(t.corNome || '').match(/^(.*?)\s*\(\s*([^)]*)\)\s*$/);
    if (!mm) return;
    if (normNome(mm[2]) !== normNome(t.corCodigo)) { console.log('  AVISO: parenteses diferente do codigo, fica: ' + t.corNome); return; }
    console.log('  ' + t.item + ' ' + t.codigo + ': "' + t.corNome + '" -> "' + mm[1] + '" · cod. ' + t.corCodigo);
    t.corNome = mm[1]; nCad++;
  });
  console.log('\n' + nCad + ' cadastro(s) renomeado(s); ' + nMov + ' lancamento(s) com o codigo da cor.');
  if (!nCad && !nMov) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-cores-sem-parenteses-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos), aviamentosMov: JSON.stringify(movs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const t2 = JSON.parse(d2.aviamentoTipos), m2 = JSON.parse(d2.aviamentosMov);
  console.log('gravado. Cadastros com parenteses: ' + t2.filter(t => /\(/.test(t.corNome || '')).length
    + ' · lancamentos de fio/linha sem o codigo da cor (dos que tem cadastro com codigo): '
    + m2.filter(m => m.modelo && !m.corCodigo && t2.some(t => t.codigo === m.modelo && t.corNome === m.cor && t.corCodigo)).length);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
