/*
 * POE MEIA BOBINA DE PAPEL KRAFT E MEIA DE FILME "EM USO".
 *
 * Junior, 28/09/2026: "Insira na coluna em uso metade de uma bobina plastico e
 * metade de bobina de papel". A meia bobina e a das 4 e meia contadas, a que
 * esta na mesa: sai do estoque e entra em uso, pelo lancamento "Posta em uso"
 * do app (saida, motivo 'uso': estoque -N, uso +N). O total nao muda.
 *   papel: meio rolo de 250 m = 125 m
 *   filme: meia bobina de ~360 m (estimativa, ver materiais-enfesto-metros.js) = 180 m
 * Rodar de novo nao repete (procura o lancamento pela observacao).
 *
 *   node servidor/meia-bobina-em-uso.js            (so mostra)
 *   node servidor/meia-bobina-em-uso.js --gravar   (grava)
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


const OBS = 'Meia bobina na mesa de corte';
const METROS = { 'papel kraft aerado': 125, 'filme plastico': 180 };
(async () => {
  const cab = await conectar();
  const { data, updatedAt } = await lerBlob(cab);
  const ler = k => { try { const v = typeof data[k] === 'string' ? JSON.parse(data[k]) : data[k]; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const cad = ler('materiaisEstCad'), movs = ler('materiaisEstMov');
  const agora = new Date().toISOString();
  const hoje = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  let n = 0;
  cad.filter(x => x.baixaOS && x.unidade !== 'sc').forEach((x, i) => {
    const q = METROS[normNome(x.nome)];
    if (!q) { console.log('  sem quantidade para: ' + x.nome); return; }
    if (movs.some(m => m.itemId === x.id && m.obs === OBS)) { console.log('  ja lancado: ' + x.nome); return; }
    const est = Number(x.emEstoque) || 0, uso = Number(x.emUso) || 0;
    if (q > est) { console.log('  estoque insuficiente: ' + x.nome); return; }
    x.emEstoque = Math.round((est - q) * 100) / 100;
    x.emUso = Math.round((uso + q) * 100) / 100;
    x.atualizadoPor = 'Junior'; x.atualizadoEm = agora;
    movs.push({ id: 'id_' + Date.now() + '_' + i + String(Math.floor(Math.random() * 1000)),
      itemId: x.id, nome: x.nome, desc: x.desc || '', unidade: 'desc', setor: x.setor || '', data: hoje,
      obs: OBS, tipo: 'saida', motivo: 'uso', qtd: q, dUso: q, dEstoque: -q, por: 'Junior', em: agora });
    n++;
    console.log('  ' + x.nome + ': em estoque ' + est + ' -> ' + x.emEstoque + ' m, em uso ' + uso + ' -> ' + x.emUso + ' m');
  });
  if (!n) return;
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-meia-bobina-em-uso-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(data), 'utf8');
  console.log('\nbackup: ' + arq);
  await gravarBlob(cab, Object.assign({}, data, { materiaisEstCad: JSON.stringify(cad), materiaisEstMov: JSON.stringify(movs) }), updatedAt);
  const c2 = JSON.parse((await lerBlob(cab)).data.materiaisEstCad);
  c2.filter(x => x.baixaOS).forEach(x => console.log('gravado: ' + x.nome + ' em uso ' + x.emUso + ' m, em estoque ' + x.emEstoque + ' m (menos 7,05 da OS 0564 na tela)'));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
