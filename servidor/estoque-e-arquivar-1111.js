/*
 * OS ANTIGAS PARA ESTOQUE, E PARA O ARQUIVO (11/11/2026).
 *
 * Junior: "Altere o status das OS anteriores às 100 OS mais recentes para
 * Estoque. Depois, arquive elas" — e: "não altere o status das OS com status
 * parado".
 *
 * 1. STATUS. Toda OS com número fora das 100 de número mais alto ganha o
 *    carimbo Estoque (statusOS/Por/Em, como o seletor de status faz, e a linha
 *    no diário de status com `c`, de carimbo). Ficam como estão:
 *      · Parado     — pedido do Junior;
 *      · Cancelado  — virar Estoque diria que a OS foi produzida;
 *      · sem número — não está na fila de numeração.
 *    O Estoque de produtos NÃO muda: a entrada de lá vem da etapa Estoque
 *    marcada no checklist, e não do status.
 *
 * 2. ARQUIVO. Escolhe com a MESMA função do programa (_osArqEscolher, recortada
 *    do app.js): fora das 100, status Estoque/Cancelado, família conjugada
 *    inteira. Grava em os_arquivadas, LÊ DE VOLTA e confere texto por texto, e
 *    só então tira do blob — numa gravação só, com o status novo das que
 *    ficam. Se o blob mudou no meio (alguém gravou), as linhas recém-gravadas
 *    são marcadas como removidas e nada muda.
 *
 *   node servidor/estoque-e-arquivar-1111.js            (só mostra)
 *   node servidor/estoque-e-arquivar-1111.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const MANTER = 100;
const QUEM = 'admin@diverse.local';
const NAO_MEXER = ['parado', 'cancelado'];

// A escolha do programa, recortada do app.js (a mesma que roda no navegador).
const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
const ini = src.indexOf('/* =========================================================\n   ARQUIVO DE OS');
const fim = src.indexOf('/* ===================== fim do ARQUIVO DE OS');
if (ini < 0 || fim < 0) throw new Error('não achei o ARQUIVO DE OS no app.js');
const ctx = { console, Map, Set, JSON, Date, Math, Number, String, parseInt, setTimeout, clearTimeout };
vm.createContext(ctx);
vm.runInContext(src.slice(ini, fim) + '\nthis._osArqEscolher = _osArqEscolher; this._osArqNumero = _osArqNumero;', ctx);

async function req(method, caminho, body, extra) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: Object.assign({ apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' }, extra || {}),
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t ? JSON.parse(t) : null;
}
const mb = s => (s.length / 1048576).toFixed(2) + ' MB';

(async () => {
  const lido = (await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const ordens = JSON.parse(data.ordens);
  const jaArq = new Set((await req('GET', '/rest/v1/os_arquivadas?select=id&removida=eq.false')).map(r => String(r.id)));
  console.log(`blob: ${ordens.length} OS (${mb(data.ordens)}) · já arquivadas: ${jaArq.size}`);

  // 1. as 100 de número mais alto (a mesma ordem da escolha do arquivo)
  const comNum = ordens.filter(o => ctx._osArqNumero(o) != null)
    .sort((a, b) => ctx._osArqNumero(b) - ctx._osArqNumero(a) || String(b.criadoEm || '').localeCompare(String(a.criadoEm || '')));
  const recentes = new Set(comNum.slice(0, MANTER).map(o => String(o.id)));
  console.log(`as ${MANTER} mais recentes: OS ${comNum[MANTER - 1].os} a ${comNum[0].os}`);

  const agora = new Date();
  const agoraIso = agora.toISOString();
  const antes = {}, mudadas = [];
  const novas = ordens.map(o => {
    if (recentes.has(String(o.id)) || ctx._osArqNumero(o) == null) return o;
    const st = String(o.statusOS || '').trim() || '(checklist)';
    if (NAO_MEXER.includes(st)) { antes[st + ' (fica)'] = (antes[st + ' (fica)'] || 0) + 1; return o; }
    antes[st] = (antes[st] || 0) + 1;
    const n = JSON.parse(JSON.stringify(o));
    n.statusOS = 'estoque';
    n.statusOSPor = QUEM;
    n.statusOSEm = agoraIso;
    delete n.statusPasso;
    const h = Array.isArray(n.statusHist) ? n.statusHist : [];
    h.push({ k: 'estoque', em: agora.getTime(), c: 'estoque' });
    n.statusHist = h;
    mudadas.push(n.os);
    return n;
  });
  console.log(`\n1. status -> Estoque: ${mudadas.length} OS. Status antes:`);
  Object.keys(antes).sort().forEach(k => console.log(`   ${k}: ${antes[k]}`));

  // 2. a escolha do arquivo, com o status novo
  const ids = ctx._osArqEscolher(novas, o => String(o.statusOS || '').trim() || 'nao-iniciado', MANTER, jaArq);
  const sai = new Set(ids);
  const ficam = novas.filter(o => !sai.has(String(o.id)));
  const ficamTxt = JSON.stringify(ficam);
  const velhasFicam = ficam.filter(o => !recentes.has(String(o.id)) && ctx._osArqNumero(o) != null);
  console.log(`\n2. arquivar: ${ids.length} OS · ficam no blob: ${ficam.length} (${mb(ficamTxt)}, era ${mb(data.ordens)})`);
  if (velhasFicam.length) console.log(`   antigas que ficam ativas: ${velhasFicam.map(o => o.os + ' ' + (o.statusOS || '')).join(', ')}`);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  if (!ids.length) { console.log('Nada a arquivar.'); return; }

  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-estoque-e-arquivar-' + agoraIso.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));

  const porId = new Map(novas.map(o => [String(o.id), o]));
  const linhas = ids.map(id => ({ id, os: String(porId.get(id).os || ''), dados: JSON.stringify(porId.get(id)), removida: false }));
  for (let i = 0; i < linhas.length; i += 40) {
    await req('POST', '/rest/v1/os_arquivadas?on_conflict=id', linhas.slice(i, i + 40), { Prefer: 'resolution=merge-duplicates,return=minimal' });
  }
  // CONFERE antes de tirar do blob.
  const conf = new Map();
  for (let i = 0; i < ids.length; i += 50) {
    const q = ids.slice(i, i + 50).map(encodeURIComponent).join(',');
    (await req('GET', `/rest/v1/os_arquivadas?select=id,dados,removida&id=in.(${q})`)).forEach(r => { if (!r.removida) conf.set(String(r.id), r.dados); });
  }
  const ruins = linhas.filter(l => conf.get(l.id) !== l.dados);
  const desfazer = async () => {
    for (let i = 0; i < ids.length; i += 50) {
      const q = ids.slice(i, i + 50).map(encodeURIComponent).join(',');
      await req('PATCH', `/rest/v1/os_arquivadas?id=in.(${q})`, { removida: true }, { Prefer: 'return=minimal' });
    }
  };
  if (ruins.length) { await desfazer(); throw new Error(`${ruins.length} OS não conferiram na tabela — desfeito, nada mudou no blob`); }
  console.log(`tabela: ${conf.size} OS gravadas e conferidas`);

  const novo = Object.assign({}, data, { ordens: ficamTxt, _device: 'estoque-e-arquivar-' + Date.now() });
  const volta = await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() });
  if (!volta || !volta.length) {
    await desfazer();
    throw new Error('ALGUÉM GRAVOU ENTRE A LEITURA E A ESCRITA — as linhas da tabela foram marcadas como removidas e nada mudou. Rode de novo.');
  }
  const d2 = (await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const o2 = JSON.parse(d2.ordens);
  const total = (await req('GET', '/rest/v1/os_arquivadas?select=id&removida=eq.false')).length;
  console.log(`gravado e relido: blob com ${o2.length} OS (${mb(d2.ordens)}) · arquivadas: ${total}`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
