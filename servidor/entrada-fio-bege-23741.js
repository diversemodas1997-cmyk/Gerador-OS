/*
 * A LINHA DA MARGEM DA FOLHA DE 28/09: FIO BEGE 23741, 6 CARRETEIS.
 *
 * Junior, 28/09/2026: "Analise a linha vertical com texto na direcao vertical
 * a esquerda da imagem. Bege, 23741, 6 carreteis. Faltou esse cadastro de fio".
 * Mesma regra do entradas-fios-linhas-2809.js: o codigo 23741 nao consta, entao
 * vira cadastro novo ("Bege (23741)", ja que Bege existe com outro codigo) e a
 * entrada de 6 carreteis, Descalvado, 28/09.
 *
 *   node servidor/entrada-fio-bege-23741.js            (so mostra)
 *   node servidor/entrada-fio-bege-23741.js --gravar   (grava)
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


const OBS = 'Contagem 28/09 (folha do caderno, margem)';
const DATA = '2026-09-28';
// [item, tipo, cor, codigo da cor na folha, carreteis] -- a folha, na ordem
const p6 = c => c ? String(c).padStart(6, '0') : '';
const LINHAS = [
  ['Fio', 'TP200', 'Bege', '23741', 6]
];
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const tipos = ler('aviamentoTipos'), movs = ler('aviamentosMov');
  const agora = new Date().toISOString();
  let nCad = 0, nEnt = 0, total = { Linha: 0, Fio: 0 };
  LINHAS.forEach(([item, codigo, cor, cod, qtd], i) => {
    const doTipo = tipos.filter(x => normNome(x.item) === normNome(item) && normNome(x.codigo) === normNome(codigo));
    // Mesmo codigo; ou a mesma cor sem codigo de um lado ou do outro.
    let t = (cod && doTipo.find(x => x.corCodigo && normNome(x.corCodigo) === normNome(cod)))
      || doTipo.find(x => normNome(x.corNome) === normNome(cor) && (!x.corCodigo || !cod));
    if (!t) {
      const mesmoNome = doTipo.some(x => normNome(x.corNome) === normNome(cor));
      const irmao = doTipo[0] || {};
      t = { id: 'id_' + Date.now() + '_t' + i + String(Math.floor(Math.random() * 1000)), item, codigo,
        corNome: mesmoNome ? cor + ' (' + cod + ')' : cor, corCodigo: cod, desc: irmao.desc || '', atualizadoPor: 'Junior', atualizadoEm: agora };
      tipos.push(t); nCad++;
      console.log('  cadastro novo: ' + item + ' ' + codigo + ' · ' + t.corNome + ' · ' + (cod || '(sem codigo)'));
    }
    cor = t.corNome;
    if (!(qtd > 0)) return;
    if (movs.some(m => m.obs === OBS && normNome(m.item) === normNome(item) && normNome(m.modelo) === normNome(codigo) && normNome(m.cor) === normNome(cor))) {
      console.log('  ja lancado: ' + item + ' ' + cor); return;
    }
    movs.push({ id: 'id_' + Date.now() + '_e' + i + String(Math.floor(Math.random() * 1000)),
      tipo: 'entrada', unidade: 'desc', item, modelo: codigo, tam: '', cor, kg: 0, qtd, data: DATA, obs: OBS, por: 'Junior', em: agora });
    nEnt++; total[item] += qtd;
  });
  console.log('\n' + nCad + ' cadastro(s) novo(s); ' + nEnt + ' entrada(s): linha ' + total.Linha + ' carreteis, fio ' + total.Fio + ' carreteis.');
  if (!nCad && !nEnt) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-entrada-fio-bege-23741-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos), aviamentosMov: JSON.stringify(movs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const m2 = JSON.parse(d2.aviamentosMov).filter(m => m.obs === OBS);
  console.log('gravado. Entradas da contagem: ' + m2.length + ' (linha ' + m2.filter(m => m.item === 'Linha').reduce((a, m) => a + m.qtd, 0)
    + ', fio ' + m2.filter(m => m.item === 'Fio').reduce((a, m) => a + m.qtd, 0) + ' carreteis) · cadastros de fio/linha: ' + JSON.parse(d2.aviamentoTipos).length);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
