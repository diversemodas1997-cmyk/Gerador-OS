/*
 * COMPLETA O CODIGO DA COR DOS FIOS QUE FICARAM SEM NUMERO.
 *
 * Junior, 28/09/2026: "Analise porque alguns fios estao sem numero de codigo de
 * cor, sendo que na imagem de origem todas as quantidades estao relacionadas as
 * cores de fios com codigo de cor".
 *
 * A causa: Preto, Roxo, Branco e Vermelho de fio ja existiam no cadastro SEM
 * codigo (a primeira lista nao os tinha). O entradas-fios-linhas-2809.js casou a
 * linha da folha com esse cadastro pelo NOME e jogou fora o numero que a folha
 * trazia. Completar o codigo que falta nao e mudar o cadastro — e o que a folha
 * diz dele.
 *
 *   Fio TP200: Roxo 54, Preto 534742, Branco 10296, Vermelho 12118 (o das 12
 *   carreteis); e o Vermelho 012 da folha (sem quantidade) vira cadastro a parte.
 * Os lancamentos desses cadastros ganham o mesmo codigo.
 *
 *   node servidor/completar-codigo-fios.js            (so mostra)
 *   node servidor/completar-codigo-fios.js --gravar   (grava)
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
const CODIGOS = { 'roxo': '54', 'preto': '534742', 'branco': '10296', 'vermelho': '12118' };
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const tipos = ler('aviamentoTipos'), movs = ler('aviamentosMov');
  const agora = new Date().toISOString();
  let nCad = 0, nMov = 0;
  const fio = tipos.filter(t => normNome(t.item) === 'fio' && normNome(t.codigo) === 'tp200');
  Object.entries(CODIGOS).forEach(([cor, cod]) => {
    const t = fio.find(x => normNome(x.corNome) === cor && !x.corCodigo);
    if (!t) { console.log('  ja tem codigo (ou nao existe): ' + cor); return; }
    const ms = movs.filter(m => normNome(m.item) === 'fio' && normNome(m.modelo) === 'tp200' && normNome(m.cor) === cor && !m.corCodigo);
    ms.forEach(m => { m.corCodigo = cod; nMov++; });
    t.corCodigo = cod; t.atualizadoPor = 'Junior'; t.atualizadoEm = agora; nCad++;
    console.log('  Fio TP200 · ' + t.corNome + ' · cod. ' + cod + ' (' + ms.length + ' lancamento(s), ' + ms.reduce((a, m) => a + (m.qtd || 0), 0) + ' carreteis)');
  });
  if (!fio.some(x => normNome(x.corNome) === 'vermelho' && normNome(x.corCodigo) === '012')) {
    const irmao = fio[0] || {};
    tipos.push({ id: 'id_' + Date.now() + '_v012', item: 'Fio', codigo: 'TP200', corNome: 'Vermelho', corCodigo: '012',
      desc: irmao.desc || '', atualizadoPor: 'Junior', atualizadoEm: agora });
    nCad++;
    console.log('  cadastro novo: Fio TP200 · Vermelho · cod. 012 (sem quantidade na folha)');
  }
  console.log('\n' + nCad + ' cadastro(s); ' + nMov + ' lancamento(s).');
  if (!nCad && !nMov) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-completar-codigo-fios-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { aviamentoTipos: JSON.stringify(tipos), aviamentosMov: JSON.stringify(movs) }), updatedAt);
  const d2 = (await lerBlob(cab)).data;
  const m2 = JSON.parse(d2.aviamentosMov).filter(m => m.modelo && m.qtd > 0 && m.tipo === 'entrada');
  console.log('gravado. Entradas de fio/linha sem codigo da cor: ' + (m2.filter(m => !m.corCodigo).map(m => m.item + ' ' + m.cor).join(', ') || 'nenhuma'));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
