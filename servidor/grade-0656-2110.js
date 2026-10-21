/*
 * OS 0656 SAÍA COM 940 PEÇAS NA OE (o certo é 470).
 *
 * Junior, 21/10/2026: "a OS656 é alocada no plano de expedição com 940 peças,
 * sendo que o total é na verdade 470". A grade gravada na OS (snapshot) ainda é
 * a antiga, 2M-2G-2G1-2G2-2G3 (10 por camada), mas o cadastro da grade passou a
 * 1 por tamanho com o Corpo rendendo 2 unidades por camada. A folha (e a OE)
 * multiplica a grade GRAVADA pelo multiplicador NOVO: 10 × 47 × 2 = 940. O
 * enfesto sempre disse 470 (5 × 47 × 2). Aqui a grade da OS volta a bater com o
 * cadastro e os componentes/aviamentos são recalculados (188 -> 94 por tamanho).
 * A carga da OE 76 (5 pacotes + reposição) já está na grade nova e não muda.
 *
 *   node servidor/grade-0656-2110.js            (só mostra)
 *   node servidor/grade-0656-2110.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const K = ['p', 'm', 'g', 'gg', 'g1', 'g2', 'g3', 't2', 't4', 't6', 't8', 't10', 't12', 't14', 't16'];
const MULT = 2;

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
  const o = ordens.find(x => x.os === '0656');
  if (!o) throw new Error('OS 0656 não encontrada');
  const g = arr(data.grades).find(x => x.id === o.gradeId);
  if (!g) throw new Error('grade da 0656 não encontrada no cadastro');
  const gravada = K.map(k => o.grade[k] || 0).join(',');
  if (gravada !== '0,2,2,0,2,2,2,0,0,0,0,0,0,0,0') throw new Error('grade gravada da 0656 mudou: ' + gravada);
  if (K.map(k => g.tamanhos[k] || 0).join(',') !== '0,1,1,0,1,1,1,0,0,0,0,0,0,0,0') throw new Error('grade do cadastro mudou');
  const cam = parseInt(o.enfesto.camadas, 10);
  if (cam !== 47 || o.enfesto.totalPecas !== 470) throw new Error('enfesto da 0656 mudou');

  K.forEach(k => { o.grade[k] = parseInt(g.tamanhos[k], 10) || 0; });
  o.grade.total = K.reduce((s, k) => s + o.grade[k], 0);
  const porTam = {};
  K.forEach(k => { porTam[k] = o.grade[k] * cam * MULT; });
  const refazer = lista => (lista || []).forEach(c => {
    const porPeca = Number(c.qtdPorPeca) || 0;
    const novo = {};
    let total = 0;
    K.forEach(k => { const v = porTam[k] * porPeca; novo[k] = v; total += v; });
    console.log(`  ${c.materialNome}: ${c.qtdTotal} -> ${total}`);
    c.qtdPorTamanho = Object.assign({}, c.qtdPorTamanho, novo);
    c.qtdTotal = total;
  });
  console.log('grade: ' + gravada + ' -> ' + K.map(k => o.grade[k]).join(',') + ' (total ' + o.grade.total + ')');
  console.log('peças: ' + o.grade.total + ' × ' + cam + ' × ' + MULT + ' = ' + o.grade.total * cam * MULT);
  console.log('componentes:'); refazer(o.componentes);
  console.log('aviamentos:'); refazer(o.aviamentos);
  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }

  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-grade-0656-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const novo = Object.assign({}, data, {
    ordens: typeof data.ordens === 'string' ? JSON.stringify(ordens) : ordens,
    _device: 'grade-0656-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const o2 = arr(JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data.ordens).find(x => x.os === '0656');
  console.log('gravado: grade ' + K.map(k => o2.grade[k]).join(',') + ' · total ' + o2.grade.total);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
