/*
 * FECHA O BURACO DA OE 0086 (05/11/2026).
 *
 * Junior: "a sequência 85 e 86 são o mesmo dia, manhã e tarde. Corrija o
 * número da 87 para 86, depois corrija a 88 para 87".
 *
 * A 0086 foi emitida em 04/11 e ficou sem carga (ver _oeNumeroGarantir no
 * app.js). Aqui: cargas com oe 87 -> 86, depois 88 -> 87, e meta.oeCounter =
 * o maior número que sobrar. Confere antes que nenhuma carga tem 86.
 *
 *   node servidor/renumerar-oe-0511.js            (só mostra)
 *   node servidor/renumerar-oe-0511.js --gravar   (grava, com backup antes)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const TROCAS = [[87, 86], [88, 87]];   // nesta ordem

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

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const cargas = arr(data.expedicaoCargas);
  const meta = obj(data.meta);

  const conta = n => cargas.filter(c => Number(c.oe) === n).length;
  console.log('antes: 86=' + conta(86) + ' 87=' + conta(87) + ' 88=' + conta(88) + ' 89=' + conta(89) + ' contador=' + meta.oeCounter);
  if (conta(86)) throw new Error('já existe carga com OE 86 — nada feito');
  if (conta(89)) throw new Error('já existe OE 89 — a sequência mudou, conferir antes');

  let saida = cargas;
  TROCAS.forEach(([de, para]) => {
    saida = saida.map(c => (Number(c.oe) === de ? Object.assign({}, c, { oe: para }) : c));
  });
  const maior = saida.reduce((mx, c) => Math.max(mx, Number(c.oe) || 0), 0);
  const meta2 = Object.assign({}, meta, { oeCounter: maior });
  const c2 = n => saida.filter(c => Number(c.oe) === n).length;
  console.log('depois: 86=' + c2(86) + ' 87=' + c2(87) + ' 88=' + c2(88) + ' contador=' + maior);
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
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUÉM GRAVOU ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');

  const d2 = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const k = arr(d2.expedicaoCargas);
  const n = x => k.filter(c => Number(c.oe) === x).length;
  console.log('gravado e relido: 85=' + n(85) + ' 86=' + n(86) + ' 87=' + n(87) + ' 88=' + n(88) + ' contador=' + obj(d2.meta).oeCounter);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
