/*
 * REMARCA "EXPEDIÇÃO DESC X SÃO CARLOS" NAS OS ALOCADAS EM OE DE IDA QUE PERDERAM A CAIXA.
 *
 * Junior, 30/09/2026: "Analise se existem OS que foram alocadas em OE e o status
 * nao foi alterado para Em transito | IDA". Alocar numa ida marca a caixa na
 * hora (_expMarcarExpedicaoOS), mas em 10 OS ela nunca chegou ao servidor, nem
 * nas copias de backup: outra maquina, com a OS de antes da alocacao na memoria,
 * marcou o Ensaque logo depois e gravou a OS inteira por cima (o merge e por
 * registro). 0571 a 0574: alocadas 07:32 de 28/09, Ensaque 08:24-08:30.
 *
 * O script procura TODA carga de ida (desde 16/09, quando a alocacao passou a
 * marcar a caixa) cuja OS tem a etapa e esta com ela desmarcada, e marca com a
 * mesma rotina do app (pai + tarefas filhas). A HORA e a da alocacao
 * (criadaEm da carga mais antiga), que e quando o app teria marcado.
 *
 *   node servidor/remarcar-ida-oe-3009.js            (so mostra)
 *   node servidor/remarcar-ida-oe-3009.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const CREDS = ['J:', 'I:', 'G:'].map(d => d + '/Meu Drive/Backup ERP Diverse/Gerador-OS-backup-dados/supa-creds.json')
  .find(p => fs.existsSync(p));
const LOCAL = path.join(__dirname, 'tls', 'servidor-local.json');
const SUPA = 'http://localhost:8000';
const DESDE = '2026-09-16';
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

const arr = v => { try { const x = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(x) ? x : []; } catch (e) { return []; } };

// As funcoes reais do app: a marca e o status saem do mesmo codigo da tela.
const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
function recorte(de, ate) {
  const i = src.indexOf(de), j = src.indexOf(ate, i);
  if (i < 0 || j < 0) throw new Error('nao achei ' + de + ' no app.js');
  return src.slice(i, j);
}
const corta = n => recorte(n, '\n}') + '\n}';
const cortaArr = n => recorte(n, '\n];') + '\n];';
const motor = STATE => new Function('STATE', `
  ${cortaArr('const STATUS_OS')}
  ${recorte('const ETAPA_SC_NOME', 'const FASES_ESTOQUE')}
  ${cortaArr('const FASES_ESTOQUE')}
  ${corta('function osEtapaMarcada')}
  ${corta('function _marcasDoStatus')}
  ${corta('function _statusDoChecklistOS')}
  ${corta('function _ultimaMarcacaoChecklist')}
  ${corta('function _statusOS')}
  ${corta('function tarefasDaEtapa')}
  ${corta('function _tarefasDaEtapaOS')}
  ${corta('function _expMarcarExpedicaoOS')}
  return { _statusOS, _expMarcarExpedicaoOS };
`)(STATE);

(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const STATE = { ordens: arr(data.ordens), expedicaoCargas: arr(data.expedicaoCargas),
                  etapas: arr(data.etapas), tarefas: arr(data.tarefas) };
  const m = motor(STATE);

  // A alocacao mais antiga de ida de cada OS, desde que a marca automatica existe.
  const alocadaEm = new Map();
  STATE.expedicaoCargas.forEach(c => {
    if (c.perna === 'volta' || !c.osId || String(c.criadaEm || '') < DESDE) return;
    const t = Date.parse(c.criadaEm);
    if (!alocadaEm.has(c.osId) || t < alocadaEm.get(c.osId)) alocadaEm.set(c.osId, t);
  });

  let n = 0;
  const agoraReal = Date.now;
  STATE.ordens.forEach(o => {
    if (!alocadaEm.has(o.id)) return;
    const antes = m._statusOS(o);
    const t = alocadaEm.get(o.id);
    Date.now = () => t;                       // a marca leva a hora da alocacao
    const nome = m._expMarcarExpedicaoOS(o, 'ida');
    Date.now = agoraReal;
    if (!nome) return;
    n++;
    console.log(`  OS ${o.os}: marca "${nome}" em ${new Date(t - 3 * 3600e3).toISOString().slice(0, 16).replace('T', ' ')}`
      + ` · status ${antes} -> ${m._statusOS(o)}`);
  });
  if (!n) { console.log('  nada a fazer'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-remarcar-ida-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  await gravarBlob(cab, Object.assign({}, data, { ordens: JSON.stringify(STATE.ordens) }), updatedAt);
  const d2 = arr((await lerBlob(cab)).data.ordens);
  const m2 = motor({ ordens: d2, etapas: STATE.etapas, tarefas: STATE.tarefas });
  const idaNa = d2.filter(o => alocadaEm.has(o.id) && m2._statusOS(o) === 'transito-ida').map(o => o.os);
  console.log('gravado. Em transito | IDA agora: ' + idaNa.join(', '));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
