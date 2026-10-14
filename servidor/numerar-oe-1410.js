/*
 * NUMERA AS OE QUE JÁ EXISTEM.
 *
 * Junior, 14/10/2026: "cada OE deve conter um número, igual a OS".
 *
 * A OE é a expedição — uma janela numa data —, e o número mora nas cargas dela
 * (campo `oe`, o mesmo em todas as cargas da expedição; ver _oeNumeroDe no
 * app.js). Daqui em diante o programa dá o número ao alocar a primeira OS; este
 * script numera as que já existiam, em ordem de DATA, e no mesmo dia pela hora
 * da ida da janela. Grava meta.oeCounter com o último número, para a próxima
 * OE continuar dali.
 *
 * Não renumera: expedição que já tem número fica com ele, e as novas seguem do
 * maior que existir.
 *
 *   node servidor/numerar-oe-1410.js            (so mostra)
 *   node servidor/numerar-oe-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();

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
const chave = c => (c.janelaId || '') + '|' + (c.data || '');

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const cargas = arr(data.expedicaoCargas);
  const janelas = arr(data.expedicaoJanelas);
  const meta = obj(data.meta);
  const hora = id => String((janelas.find(j => j.id === id) || {}).horaIda || '99:99');
  const nomeJ = id => String((janelas.find(j => j.id === id) || {}).nome || '(janela apagada)');

  // As expedições, cada uma com o número que já tiver.
  const exps = new Map();
  cargas.forEach(c => {
    const k = chave(c);
    if (!exps.has(k)) exps.set(k, { k, janelaId: c.janelaId, data: c.data, n: 0, cargas: 0 });
    const e = exps.get(k);
    e.cargas++;
    if (Number(c.oe) > 0) e.n = Number(c.oe);
  });
  const lista = [...exps.values()].sort((a, b) =>
    String(a.data).localeCompare(String(b.data)) || hora(a.janelaId).localeCompare(hora(b.janelaId)) || String(a.janelaId).localeCompare(String(b.janelaId)));
  let ultimo = Math.max(Number(meta.oeCounter) || 0, ...lista.map(e => e.n));
  let novas = 0;
  lista.forEach(e => { if (!e.n) { e.n = ++ultimo; novas++; } });
  console.log(`${lista.length} expedições com carga, ${novas} sem número.`);
  lista.forEach(e => console.log(`  OE ${String(e.n).padStart(4, '0')}  ${e.data}  ${hora(e.janelaId)}  ${nomeJ(e.janelaId)}  (${e.cargas} carga(s))`));
  console.log(`contador: ${ultimo}`);
  if (!novas && (Number(meta.oeCounter) || 0) >= ultimo) { console.log('Nada a fazer.'); return; }
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const numero = new Map(lista.map(e => [e.k, e.n]));
  const saida = cargas.map(c => (Number(c.oe) > 0 ? c : Object.assign({}, c, { oe: numero.get(chave(c)) })));
  const meta2 = Object.assign({}, meta, { oeCounter: ultimo });

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-numerar-oe-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));

  const novo = Object.assign({}, data, {
    expedicaoCargas: typeof data.expedicaoCargas === 'string' ? JSON.stringify(saida) : saida,
    meta: typeof data.meta === 'string' ? JSON.stringify(meta2) : meta2,
    _device: 'numerar-oe-' + Date.now()     // sentinela: os clientes recarregam
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');

  const d2 = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const c2 = arr(d2.expedicaoCargas);
  console.log(`gravado. Cargas sem número: ${c2.filter(c => !(Number(c.oe) > 0)).length} de ${c2.length}; contador ${obj(d2.meta).oeCounter}.`);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
