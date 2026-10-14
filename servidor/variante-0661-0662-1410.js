/*
 * VARIANTE DAS OS 0661 E 0662 = AS CORES DO DESENHO 0050.
 *
 * Junior, 14/10/2026: "Analise por que as OS661 e 662 apresentam o texto da cor
 * acima do desenho técnico de forma incompleta. Deveria ser Azul/Branco/Preto".
 *
 * O texto acima do desenho é coresDaPecaOS: as Cor 1/2/3 da variante da OS,
 * sem o tecido no nome e sem repetir. As duas OS (aprovadas de planejamentos
 * em 14/10) saíram com a variante Branco / Preto / Preto Malha Algodão — e o
 * texto ficava "Branco / Preto". O desenho 0050 (BM.TRI-AZUL, "Blusa Moletom
 * Tricolor | Azul/Branco/Preto") tem Azul / Branco / Preto Moletom: a variante
 * passa a ser essa, na ordem do desc do desenho.
 *
 * Só a variante. Tecidos, fases e componentes da OS não são tocados (os da
 * 0661 também estão sem o Azul — decisão à parte, porque mexem no material).
 *
 *   node servidor/variante-0661-0662-1410.js            (so mostra)
 *   node servidor/variante-0661-0662-1410.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const OSS = ['0661', '0662'];
const DESENHO = '0050';

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
const base = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim().split(/\s+/)[0];

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const ordens = arr(data.ordens), cores = arr(data.cores), desenhos = arr(data.desenhos);
  const d = desenhos.find(x => x.codigo === DESENHO);
  if (!d) throw new Error('desenho ' + DESENHO + ' não encontrado');
  const nome = id => (cores.find(c => c.id === id) || {}).nome || '';
  // A ordem do desc ("… | Azul/Branco/Preto"), como ordenarCoresIdsPorDesc.
  const ordem = String(d.desc || '').split('|').pop().split('/').map(base).filter(Boolean);
  const ids = [d.corPrincipalId, d.corSecundariaId, d.corTerciariaId].filter(Boolean)
    .map((id, i) => ({ id, i, pos: ordem.indexOf(base(nome(id))) }))
    .sort((a, b) => (a.pos < 0 ? 99 : a.pos) - (b.pos < 0 ? 99 : b.pos) || a.i - b.i).map(x => x.id);
  console.log(`Desenho ${d.codigo} (${d.desc}): ${ids.map(nome).join(' / ')}`);

  let mexeu = 0;
  const saida = ordens.map(o => {
    if (!OSS.includes(o.os)) return o;
    if (o.desenhoId !== d.id) throw new Error(`a OS ${o.os} não é do desenho ${DESENHO}`);
    const v0 = (o.variantes || [])[0] || {};
    console.log(`  OS ${o.os}  antes: ${[v0.cor1Nome, v0.cor2Nome, v0.cor3Nome].join(' / ')}`);
    const nova = Object.assign({}, v0, {
      num: 1, cor1: ids[0] || '', cor2: ids[1] || '', cor3: ids[2] || '',
      cor1Nome: nome(ids[0]) || '—', cor2Nome: nome(ids[1]) || '—', cor3Nome: nome(ids[2]) || '—'
    });
    console.log(`           depois: ${[nova.cor1Nome, nova.cor2Nome, nova.cor3Nome].join(' / ')}`);
    mexeu++;
    return Object.assign({}, o, { variantes: [nova].concat((o.variantes || []).slice(1)) });
  });
  if (mexeu !== OSS.length) throw new Error('não achei as duas OS');
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-variante-0661-0662-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    ordens: typeof data.ordens === 'string' ? JSON.stringify(saida) : saida,
    _device: 'variante-0661-0662-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const o2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.ordens)
    .filter(o => OSS.includes(o.os));
  o2.forEach(o => console.log(`gravado OS ${o.os}: ${[o.variantes[0].cor1Nome, o.variantes[0].cor2Nome, o.variantes[0].cor3Nome].join(' / ')}`));
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
