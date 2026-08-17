/* Rode com:  node testes/pastas-grades-semelhanca.js

   AS PASTAS DE GRADE ORGANIZADAS POR SEMELHANÇA.

   O cadastro tem mais de cem grades. Pasta (tipo de peça) e subpasta (variação)
   dão o primeiro corte, mas dentro delas o que a fábrica pede é uma FAIXA: "uma
   que vá do M ao GG". Ordenar por semelhança já punha as parecidas lado a lado —
   e ainda era uma parede de nomes, porque nada dizia onde uma faixa acaba e a
   outra começa. Agora a faixa é o terceiro nível de pasta, automático: sai dos
   tamanhos da própria grade, então grade nova cai no grupo certo ao ser salva e
   não há pasta para manter.

   O que este teste guarda: o nome do grupo (é ele que a pessoa procura), e o
   agrupamento — mesma faixa junta, quantidade não separa, ordem preservada. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

function recorte(de, ate, oQue) {
  const i = src.indexOf(de);
  const j = src.indexOf(ate, i);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';
const cortaLinha = (nome) => recorte(nome, '\n', nome);

const api = new Function(`
  ${cortaLinha('const _ORDEM_TAM')}
  ${corta('function _gradeChaveSemelhanca')}
  ${corta('function compararGradesPorSemelhanca')}
  ${corta('function _gradeFaixaLabel')}
  ${corta('function _agruparGradesPorFaixa')}
  return { _gradeFaixaLabel, _agruparGradesPorFaixa, compararGradesPorSemelhanca, _gradeChaveSemelhanca };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome + ' → ' + JSON.stringify(esperado), got === esperado, got);

const T = (o) => Object.assign({ p: 0, m: 0, g: 0, gg: 0, g1: 0, g2: 0, g3: 0 }, o);
const grade = (nome, tam) => ({ nome, tamanhos: T(tam) });

/* ---------- 1. o nome do grupo é como a fábrica pede a grade ---------- */

eq('faixa corrida do começo ao fim',
  api._gradeFaixaLabel(T({ p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 })), 'P ao G3');
eq('faixa corrida no meio',
  api._gradeFaixaLabel(T({ m: 1, g: 1, gg: 1 })), 'M ao GG');
eq('quantidade não muda o nome da faixa',
  api._gradeFaixaLabel(T({ p: 2, m: 2, g: 2, gg: 2, g1: 2, g2: 2, g3: 2 })), 'P ao G3');
eq('faixa que pula tamanho é listada',
  api._gradeFaixaLabel(T({ m: 1, g: 1, g1: 1, g3: 1 })), 'M · G · G1 · G3');
eq('duas vizinhas ficam listadas (P ao M diria menos)',
  api._gradeFaixaLabel(T({ p: 1, m: 1 })), 'P · M');
eq('duas não vizinhas', api._gradeFaixaLabel(T({ p: 2, gg: 2 })), 'P · GG');
eq('um tamanho só', api._gradeFaixaLabel(T({ g: 1 })), 'G');
eq('grade sem tamanho nenhum não fica sem nome',
  api._gradeFaixaLabel(T({})), 'Sem tamanhos');
eq('tamanhos como texto (vem assim de import antigo)',
  api._gradeFaixaLabel({ p: '1', m: '1', g: '1' }), 'P ao G');

/* ---------- 2. o agrupamento ---------- */

// Entra a lista JÁ ordenada por semelhança, como a tela faz.
const ordenar = (gs) => gs.slice().sort(api.compararGradesPorSemelhanca);

let gs = ordenar([
  grade('2M-2G | CM.LISA | 117cm', { m: 2, g: 2 }),
  grade('P ao G3 | CM.LISA | 117cm', { p: 1, m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }),
  grade('M-G | CM.REC | 117cm', { m: 1, g: 1 }),
  grade('2X P ao G3 | CM.LISA | 117cm', { p: 2, m: 2, g: 2, gg: 2, g1: 2, g2: 2, g3: 2 })
]);
let grupos = api._agruparGradesPorFaixa(gs);
eq('duas faixas distintas viram dois grupos', grupos.length, 2);
eq('o grupo mais largo vem primeiro', grupos[0].label, 'P ao G3');
eq('e leva as duas grades daquela faixa (1x e 2x juntas)', grupos[0].itens.length, 2);
eq('a faixa curta vem depois', grupos[1].label, 'M · G');
eq('com as duas grades dela', grupos[1].itens.length, 2);
// Dentro do grupo, a ordem da semelhança é preservada: 1x antes de 2x.
eq('dentro do grupo, a menor quantidade primeiro', grupos[0].itens[0].nome, 'P ao G3 | CM.LISA | 117cm');

// Nenhuma grade se perde no caminho — é o que garante que a tela não esconde
// cadastro atrás de um agrupamento errado.
const total = grupos.reduce((s, f) => s + f.itens.length, 0);
eq('nenhuma grade fica fora de grupo', total, gs.length);

// Faixas iguais separadas na lista de entrada não podem virar dois grupos com o
// mesmo nome — é por isso que a lista entra ordenada.
gs = ordenar([
  grade('a', { m: 1, g: 1 }),
  grade('b', { p: 1, m: 1, g: 1 }),
  grade('c', { m: 2, g: 2 })
]);
grupos = api._agruparGradesPorFaixa(gs);
eq('faixas iguais não se repetem como grupos', grupos.length, 2);
ok('nenhum nome de grupo aparece duas vezes',
  new Set(grupos.map(f => f.label)).size === grupos.length, grupos.map(f => f.label));

// Uma faixa só (o caso em que a tela NÃO abre o nível): tem de ser reconhecível.
grupos = api._agruparGradesPorFaixa(ordenar([grade('x', { g: 1 }), grade('y', { g: 2 })]));
eq('subpasta de faixa única dá um grupo só', grupos.length, 1);

// Lista vazia não explode.
eq('subpasta vazia não gera grupo', api._agruparGradesPorFaixa([]).length, 0);

/* ---------- 3. a máscara é a identidade da faixa ---------- */

const masc = (t) => api._gradeChaveSemelhanca(T(t)).mascara;
ok('a mesma faixa em quantidades diferentes tem a mesma máscara',
  masc({ m: 1, g: 1 }) === masc({ m: 9, g: 3 }), [masc({ m: 1, g: 1 }), masc({ m: 9, g: 3 })]);
ok('faixas diferentes têm máscaras diferentes',
  masc({ m: 1, g: 1 }) !== masc({ m: 1, gg: 1 }), masc({ m: 1, g: 1 }));
ok('quem começa no P vem antes de quem começa no M',
  masc({ p: 1 }) > masc({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }),
  [masc({ p: 1 }), masc({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 })]);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
