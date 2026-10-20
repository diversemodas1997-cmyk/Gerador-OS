/*
 * OBSERVAÇÃO DA CARGA DA MANHÃ DA OS 0661 (OE 0074) SEM A MANGA.
 *
 * Junior, 20/10/2026: "Corrija a observação da manhã, retire a manga". A obs
 * dizia "Capuz + Forro de capuz, barra/punhos, manga", mas a manga foi à tarde
 * (OE 0075, aloca-0661-oe75-2010.js). Aqui só sai a palavra manga do texto.
 *
 *   node servidor/obs-0661-manha-sem-manga-2010.js            (só mostra)
 *   node servidor/obs-0661-manha-sem-manga-2010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const MATINAL = 'id_1784296716535_431';
const DATA = '2026-10-20';

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

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const o = arr(data.ordens).find(x => x.os === '0661');
  if (!o) throw new Error('OS 0661 não encontrada');
  const cargas = arr(data.expedicaoCargas);
  const manha = cargas.find(c => c.osId === o.id && c.janelaId === MATINAL && c.data === DATA && c.perna === 'ida');
  if (!manha) throw new Error('carga da manhã não encontrada');
  const antes = String(manha.obs || '');
  if (!/manga/i.test(antes)) throw new Error('a obs da manhã já não fala de manga: "' + antes + '"');
  const depois = antes
    .replace(/\s*(,|\+|\be\b)?\s*mangas?\b/gi, '')
    .replace(/\s*,\s*$/, '').replace(/\s{2,}/g, ' ').trim();
  manha.obs = depois;
  console.log('obs antes : "' + antes + '"\nobs depois: "' + depois + '"');
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-obs-0661-manha-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    expedicaoCargas: typeof data.expedicaoCargas === 'string' ? JSON.stringify(cargas) : cargas,
    _device: 'obs-0661-manha-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const c2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.expedicaoCargas)
    .find(c => c.id === manha.id);
  console.log('gravado: OE ' + c2.oe + ' · obs "' + c2.obs + '"');
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
