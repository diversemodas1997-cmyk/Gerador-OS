/*
 * CONJUGADA DA GRADE P-G-GG-G1-G3 | CM.REC E AS OS 0669/0670.
 *
 * Junior, 20/10/2026: "Analise por que as OS670 e 669 não emitiram OS conjugada
 * cm.lisa branca" e, depois, "A CM.LISA conjugada da P-G-GG-G1-G3 é M-G Branca".
 *
 * A grade "P-G-GG-G1-G3 | CM.REC | 117cm" nunca teve OS conjugada no cadastro,
 * e não havia CM.LISA conjugada dela. Este script:
 *   1. cria a grade "M-G | CM.LISA  | 117cm (CONJUGADO P-G-GG-G1-G3 | CM.REC |
 *      117cm)" no molde das irmãs (M=1, G=1; Corpo, Gola e Viés), com o Corpo
 *      medido pela fase Corpo 2 da CM.REC (é o mesmo enfesto, e o risco é o
 *      dela);
 *   2. liga a CM.REC a ela (conjugadaGradeId);
 *   3. gera a conjugada das OS 0669 e 0670 com a função gerarConjugada DO
 *      PRÓPRIO app.js (recortada daqui), sobre os dados do servidor.
 *
 *   node servidor/conjugada-p-g-gg-g1-g3-2010.js            (só mostra)
 *   node servidor/conjugada-p-g-gg-g1-g3-2010.js --gravar   (grava)
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const GRAVAR = process.argv.includes('--gravar');
const BASE = 'http://localhost:8000';
const KEY = fs.readFileSync('C:/supabase/docker/.env', 'utf8')
  .match(/^(?:SUPABASE_)?SERVICE_ROLE_KEY=(.+)$/m)[1].trim();
const GRADE_REC = 'id_1786565692377_995';      // P-G-GG-G1-G3 | CM.REC | 117cm
const MOLDE = 'id_1788540541340_929';          // M-G | CM.LISA (CONJUGADO P ao G3)
const NOME = 'M-G | CM.LISA  | 117cm (CONJUGADO P-G-GG-G1-G3 | CM.REC | 117cm)';
// As OS pelo DESENHO na grade (os números andaram: a 0670 virou 0671 em 20/10).
const DESENHOS_OS = ['0011', '0013'];   // Recortada Roxo/Branco e Verde/Branco
const DESENHO_CONJ = 'id_1776882334182_443';   // Camiseta Básica | Branco, o mesmo da P ao G3

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

// Recorta funções do app.js para rodar a regra de verdade, e não uma cópia dela.
const src = fs.readFileSync(path.join(RAIZ, 'app.js'), 'utf8');
function corta(nome) {
  let i = src.indexOf('\nfunction ' + nome + '(');
  if (i < 0) i = src.indexOf('\nasync function ' + nome + '(');
  if (i < 0) throw new Error('não achei ' + nome + ' no app.js');
  return src.slice(i, src.indexOf('\n}', i) + 2);
}

(async () => {
  const lido = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data,updated_at'))[0];
  const data = lido.data;
  const STATE = {};
  for (const k of ['ordens', 'grades', 'desenhos', 'cores', 'tecidos', 'componentes', 'materiais', 'modelos']) STATE[k] = arr(data[k]);
  STATE.osCounter = data.osCounter;

  const rec = STATE.grades.find(g => g.id === GRADE_REC);
  const molde = STATE.grades.find(g => g.id === MOLDE);
  if (!rec || !molde) throw new Error('grade CM.REC ou molde não encontrados');
  if (rec.conjugadaGradeId) throw new Error('a CM.REC já conjuga ' + rec.conjugadaGradeId + ' — nada feito');

  // 1. A grade CM.LISA nova.
  let lisa = STATE.grades.find(g => g.nome === NOME);
  if (!lisa) {
    const corpo2 = (rec.fases || []).find(f => /corpo 2/i.test(f.nome));
    lisa = JSON.parse(JSON.stringify(molde));
    lisa.id = 'id_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    lisa.nome = NOME;
    const c = lisa.fases.find(f => f.nome === 'Corpo');
    Object.assign(c, { comp: corpo2.comp, larg: corpo2.larg, bobinas: corpo2.bobinas, risco: corpo2.risco });
    lisa.enfestoComprimento = corpo2.comp;
    lisa.enfestoLargura = corpo2.larg;
    STATE.grades.push(lisa);
    console.log('grade nova: ' + lisa.nome + '  (Corpo ' + c.comp + ' × ' + c.larg + ' m, risco ' + c.risco + ')');
  }
  // 2. A CM.REC conjuga a nova.
  rec.conjugadaGradeId = lisa.id;
  rec.conjugadaDesenhoId = DESENHO_CONJ;
  console.log('ligada: ' + rec.nome + ' -> ' + lisa.nome);

  // 3. A conjugada de cada OS, pela função do app.
  const avisos = [];
  const fn = new Function('STATE', 'avisos', `
    const toast = (m, t) => avisos.push((t || 'info') + ': ' + m);
    const saveState = async () => {};
    const atualizarCounterOS = async () => {};
    const NUMERO_OS_MAX = 9999, OS_SALTO_ESTRANHO = 100;
    ${['uid', '_sufixoTecidoNorm', '_normNome', 'corBaseNome', 'ordemCoresPorDesc', '_fotoDoDesenho', '_numeroOSCanonico', '_numerosOSExistentes', '_topoDaFilaOS', 'formatarNumeroOS', 'proximoNumeroOS', 'numeroOSLivre', 'numeroParaConjugada',
       '_congelarDesenhoNaOS', 'ordenarCoresIdsPorDesc', 'gradeConjugadaDaGrade', 'deveGerarConjugada',
       'gerarConjugada'].map(corta).join('\n')}
    return { deveGerarConjugada, gerarConjugada };
  `);
  const api = fn(STATE, avisos);
  const alvos = DESENHOS_OS.map(c => STATE.ordens.find(x => x.gradeId === GRADE_REC && x.codigo === c));
  if (alvos.some(o => !o)) throw new Error('não achei as duas OS na grade');
  const OSS = alvos.map(o => o.os);
  for (const o of alvos) {
    const n = o.os;
    if (!api.deveGerarConjugada(o)) { console.log('OS ' + n + ': não gera (já tem conjugada?)'); continue; }
    const nova = await api.gerarConjugada(o);
    console.log(`OS ${n} -> conjugada OS ${nova.os}: ${nova.codigo} · ${nova.grade.descricao} · ${nova.grade.total} pç/camada × ${nova.enfesto.camadas} camadas = ${nova.enfesto.totalPecas} pç · cor ${nova.variantes && nova.variantes[0].cor1Nome}`);
  }
  avisos.forEach(a => console.log('  aviso ' + a));

  if (!GRAVAR) { console.log('\nNada gravado (rode com --gravar).'); return; }
  const agora = new Date().toISOString();
  const arq = path.join(RAIZ, 'backups', 'shared_data-antes-conjugada-p-g-gg-g1-g3-' + agora.replace(/[:.]/g, '-') + '.json');
  fs.mkdirSync(path.dirname(arq), { recursive: true });
  fs.writeFileSync(arq, JSON.stringify(lido), 'utf8');
  console.log('\nbackup: ' + path.relative(RAIZ, arq));
  const maior = Math.max(parseInt(data.osCounter) || 0, ...STATE.ordens.map(o => parseInt(o.os) || 0));
  const enc = (k, v) => (typeof data[k] === 'string' ? JSON.stringify(v) : v);
  const novo = Object.assign({}, data, {
    ordens: enc('ordens', STATE.ordens),
    grades: enc('grades', STATE.grades),
    osCounter: typeof data.osCounter === 'string' ? String(maior) : maior,
    _device: 'conjugada-p-g-gg-g1-g3-' + Date.now()
  });
  const volta = JSON.parse(await req('PATCH',
    '/rest/v1/shared_data?id=eq.main&updated_at=eq.' + encodeURIComponent(lido.updated_at),
    { data: novo, updated_at: new Date().toISOString() }));
  if (!volta.length) throw new Error('ALGUEM GRAVOU NO SERVIDOR ENTRE A LEITURA E A ESCRITA -- nada foi alterado. Rode de novo.');
  const d2 = JSON.parse(await req('GET', '/rest/v1/shared_data?id=eq.main&select=data'))[0].data;
  const o2 = arr(d2.ordens);
  OSS.forEach(n => {
    const a = o2.find(o => o.os === n);
    const c = a && o2.find(o => o.id === a.conjugadaId);
    console.log(`gravado OS ${n} ⇄ ${c ? c.os + ' (' + c.grade.descricao + ')' : 'SEM CONJUGADA'}`);
  });
  console.log('grade ligada: ' + (arr(d2.grades).find(g => g.id === GRADE_REC) || {}).conjugadaGradeId);
  console.log('\nAgora: F5 em todas as abas do programa.');
})().catch(e => { console.error('ERRO: ' + e.message); process.exit(1); });
