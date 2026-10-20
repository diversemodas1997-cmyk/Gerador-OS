/*
 * OS 0661 NA OE 0075 (Tarde de 20/10, ida) COM O QUE NÃO FOI DE MANHÃ.
 *
 * Junior, 20/10/2026: "Aloque na OE0075 a OS661 com as peças da frente e manga
 * que não foram pela manhã". A carga da manhã (OE 0074) levou Forro de capuz,
 * Barra/Punhos e Viés com os 21 pacotes e a reposição; a observação dela diz
 * "Capuz + Forro de capuz, barra/punhos, manga" — mas a manga NÃO foi (é o que
 * ele diz agora). Ela foi gravada antes de existirem os itens Mangas/Capuz no
 * checklist, então aqui ela ganha o Capuz (ordem 3,6) e `fasesComItens`.
 * A carga nova: Corpo Parte 1, 2 e 3 + Mangas (3,3), os mesmos 21 pacotes,
 * sem reposição (foi de manhã), OE 75.
 *
 *   node servidor/aloca-0661-oe75-2010.js            (só mostra)
 *   node servidor/aloca-0661-oe75-2010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const TARDE = 'id_1784631224974_386';
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
  const ordens = arr(data.ordens);
  const cargas = arr(data.expedicaoCargas);
  const o = ordens.find(x => x.os === '0661');
  if (!o) throw new Error('OS 0661 não encontrada');
  const nomes = (o.fases || []).map(f => f.nome).join(' / ');
  if (nomes !== 'Corpo Parte 1 / Corpo Parte 2 / Corpo Parte 3 / Forro de capuz / Barra/Punhos / Viés') throw new Error('fases da 0661 mudaram: ' + nomes);
  const manha = cargas.find(c => c.osId === o.id && c.janelaId === MATINAL && c.data === DATA && c.perna === 'ida');
  if (!manha || JSON.stringify(manha.fases) !== '[4,5,6]') throw new Error('carga da manhã não está como esperado');
  if (cargas.some(c => c.osId === o.id && c.janelaId === TARDE && c.data === DATA && c.perna === 'ida')) throw new Error('a 0661 já está na carga da tarde — parei');
  const oe75 = cargas.filter(c => c.janelaId === TARDE && c.data === DATA);
  if (!oe75.length || oe75.some(c => c.oe !== 75)) throw new Error('a carga da tarde de 20/10 não é a OE 75');

  manha.fases = [3.6, 4, 5, 6];
  manha.fasesComItens = true;
  const nova = {
    id: 'id_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
    janelaId: TARDE, data: DATA, perna: 'ida', osId: o.id,
    volumes: manha.pacotes.length, fases: [1, 2, 3, 3.3], fasesComItens: true,
    obs: 'Corpo partes 1, 2 e 3 (frentes) + mangas, que não foram de manhã (OE 0074).',
    oe: 75,
    pacotes: manha.pacotes.map(p => ({ tam: p.tam, tom: p.tom })), reposicao: false,
    criadaEm: new Date().toISOString(), criadaPor: 'admin@diverse.local'
  };
  cargas.push(nova);
  console.log('manhã (OE 74): fases -> ' + JSON.stringify(manha.fases) + ' (Capuz, Forro, Barra/Punhos, Viés)');
  console.log('nova (OE 75): fases ' + JSON.stringify(nova.fases) + ' (Corpo Parte 1, 2, 3, Mangas) · ' + nova.pacotes.length + ' pacotes · volumes ' + nova.volumes);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-aloca-0661-oe75-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    expedicaoCargas: typeof data.expedicaoCargas === 'string' ? JSON.stringify(cargas) : cargas,
    _device: 'aloca-0661-oe75-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const c2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.expedicaoCargas)
    .filter(c => c.osId === o.id);
  c2.forEach(c => console.log(`gravado: OE ${c.oe} ${c.data} ${c.janelaId === TARDE ? 'Tarde' : 'Matinal'} · fases ${JSON.stringify(c.fases)} · ${c.pacotes.length} pacotes${c.reposicao ? ' + reposição' : ''}`));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
