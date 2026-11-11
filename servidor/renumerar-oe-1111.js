/*
 * OE 0093 <-> 0094 NA ORDEM DAS DATAS (11/11/2026).
 *
 * Junior: "Corrija a numeração da OE0093, pois a sequência correta é OE0094".
 *
 * Como estava: a 2ª expedição matinal de 10/11 foi criada depois da Matinal
 * de 11/11 e levou o número seguinte:
 *   0092  10/11 Matinal
 *   0093  11/11 Matinal                  -> passa a 0094
 *   0094  10/11 2ª expedição matinal     -> passa a 0093
 * Os dois trocam de número (a 0094 já existia: sem a troca o número repetiria).
 * O contador (meta.oeCounter) segue 94. Cada expedição é achada pelo número E
 * pela data, para o script não mexer em nada se a situação já for outra.
 *
 *   node servidor/renumerar-oe-1111.js            (só mostra)
 *   node servidor/renumerar-oe-1111.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
// [número de agora, data da expedição, número certo]
const TROCAS = [[93, '2026-11-11', 94], [94, '2026-11-10', 93]];

async function req(method, caminho, body) {
  const r = await fetch(BASE + caminho, {
    method,
    headers: { apikey: KEY, Authorization: 'Bearer ' + KEY, 'Content-Type': 'application/json', Prefer: 'return=representation' },
    body: body ? JSON.stringify(body) : undefined
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ': ' + t.slice(0, 300));
  return t;
}
const arr = s => (typeof s === 'string' ? JSON.parse(s || '[]') : (s || []));
const obj = s => (typeof s === 'string' ? JSON.parse(s || '{}') : (s || {}));

function resumo(cargas, janelas) {
  const nome = id => (janelas.find(j => j.id === id) || {}).nome || id;
  const g = new Map();
  cargas.filter(c => [92, 93, 94, 95].includes(Number(c.oe))).forEach(c => {
    const k = c.oe + ' ' + c.data + ' ' + nome(c.janelaId);
    g.set(k, (g.get(k) || 0) + 1);
  });
  return [...g.entries()].sort().map(([k, n]) => '  OE' + k.replace(/^(\d+)/, m => m.padStart(4, '0')) + ' (' + n + ' cargas)').join('\n');
}

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const cargas = arr(data.expedicaoCargas);
  const janelas = arr(data.expedicaoJanelas);
  const meta = obj(data.meta);

  console.log('antes:\n' + resumo(cargas, janelas) + '\n  contador=' + meta.oeCounter);
  // Cada número de agora tem de estar numa expedição só, e na data esperada.
  TROCAS.forEach(([de, dia]) => {
    const cs = cargas.filter(c => Number(c.oe) === de);
    const exp = new Set(cs.map(c => c.janelaId + '|' + c.data));
    if (!cs.length) throw new Error('nenhuma carga com OE ' + de + ' — nada feito');
    if (exp.size !== 1 || cs[0].data !== dia) throw new Error('a OE ' + de + ' não é mais só a de ' + dia + ' — a situação mudou, conferir antes');
  });
  if (cargas.some(c => Number(c.oe) === 95)) throw new Error('já existe OE 95 — a sequência mudou, conferir antes');

  const saida = cargas.map(c => {
    const t = TROCAS.find(([de, dia]) => Number(c.oe) === de && c.data === dia);
    return t ? Object.assign({}, c, { oe: t[2] }) : c;
  });
  const maior = saida.reduce((mx, c) => Math.max(mx, Number(c.oe) || 0), 0);
  const meta2 = Object.assign({}, meta, { oeCounter: maior });
  console.log('depois:\n' + resumo(saida, janelas) + '\n  contador=' + maior);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-renumerar-oe-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('backup: ' + path.relative(RAIZ, arq));

  const novo = Object.assign({}, data, {
    expedicaoCargas: typeof data.expedicaoCargas === 'string' ? JSON.stringify(saida) : saida,
    meta: typeof data.meta === 'string' ? JSON.stringify(meta2) : meta2,
    _device: 'renumerar-oe-' + Date.now()     // sentinela: os clientes recarregam
  });
  // Só grava se ninguém gravou desde a leitura (updated_at igual).
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUÉM GRAVOU ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');

  const d2 = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  console.log('gravado e relido:\n' + resumo(arr(d2.expedicaoCargas), arr(d2.expedicaoJanelas)) + '\n  contador=' + obj(d2.meta).oeCounter);
  console.log('\nAgora: F5 em todas as abas do programa, e reimprimir as folhas das OE 0093 e 0094.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
