/* Rode com:  node testes/importar-risco-grade.js

   Importar risco (PDF do encaixe): o seletor de GRADE.

   O programa descobre a grade pela tabela de tamanhos do relatório do CAD. Isso
   acerta na maioria das vezes e ERRA quando o encaixe não é uma grade inteira:

     CM.REC - CORPO 2 - M-G-GG-G1-G2-G3
       Encaixados: 14/30 · Modelos completos: 2 · Modelos pedidos: 6
       M  completos 1  moldes 5      GG  completos 0  moldes 1
       G  completos 1  moldes 5      G1..G3 idem

   O encaixe leva os 5 moldes de M e de G e só 1 molde de cada tamanho maior. O
   CAD conta 2 modelos completos e a tabela sai "M 1 · G 1" — correto do ponto de
   vista dele, e não é a grade: a grade é M ao G3. Filtrando o seletor pelos
   tamanhos lidos, a única candidata era "M-G (CONJUGADO) | CM.LISA", que nada
   tem a ver, e a certa não aparecia em lugar nenhum.

   Regra que este teste protege: adivinhar a grade é papel do programa; IMPEDIR a
   correção não é. O seletor traz sempre o cadastro inteiro.

   O teste recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + oQue); process.exit(1); }
  return src.slice(i, j + 2);
}

const api = new Function('STATE', `
  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  ${recorte('function _riscoGradesQueCasam', 'a busca por tamanhos')}
  ${recorte('function _riscoTamanhosTexto', 'o texto dos tamanhos')}
  ${recorte('function _riscoCelulaGrade', 'a celula da grade')}
  return { _riscoGradesQueCasam, _riscoTamanhosTexto, _riscoCelulaGrade };
`);

// O cadastro, reduzido ao que importa: a grade CERTA e a que casa por engano.
const T = (o) => Object.assign({ p: 0, m: 0, g: 0, gg: 0, g1: 0, g2: 0, g3: 0 }, o);
const STATE = { grades: [
  { id: 'g_conj', nome: 'M-G (CONJUGADO) | CM.LISA | 117cm', tamanhos: T({ m: 1, g: 1 }) },
  { id: 'g_certa', nome: 'M ao G3 | CM.REC | 117cm', tamanhos: T({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }) },
  { id: 'g_outra', nome: '2P-2GG | CM.REC | 117cm', tamanhos: T({ p: 2, gg: 2 }) }
] };
const { _riscoGradesQueCasam, _riscoTamanhosTexto, _riscoCelulaGrade } = api(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

console.log('-- os tamanhos que o PDF declarou --');
ok('1. lista so os que tem quantidade', _riscoTamanhosTexto({ m: 1, g: 1 }) === 'M 1 · G 1',
   _riscoTamanhosTexto({ m: 1, g: 1 }));
ok('2. mantem a ordem P..G3',
   _riscoTamanhosTexto({ g3: 2, p: 1, gg: 4 }) === 'P 1 · GG 4 · G3 2',
   _riscoTamanhosTexto({ g3: 2, p: 1, gg: 4 }));
ok('3. tabela vazia diz "nenhum"', _riscoTamanhosTexto({}) === 'nenhum' && _riscoTamanhosTexto(null) === 'nenhum');

console.log('');
console.log('-- o caso real do CORPO 2: a tabela nao descreve a grade --');
const lidos = { m: 1, g: 1 };                       // o que o PDF do CORPO 2 declara
const cands = _riscoGradesQueCasam(lidos);
ok('4. pelos tamanhos, a unica candidata e a CONJUGADO (a errada)',
   cands.length === 1 && cands[0].id === 'g_conj', cands.map(g => g.nome).join(', '));

const L = { tamanhos: lidos, grades: cands, grade: null };
const html = _riscoCelulaGrade(L, 0);
ok('5. a candidata aparece no grupo dos tamanhos do PDF',
   html.includes('com os tamanhos do PDF') && html.includes('corrigir: M-G (CONJUGADO)'), html);
ok('6. A GRADE CERTA tambem e oferecida, no grupo das outras',
   html.includes('todas as outras grades') && html.includes('value="g_certa"'), html);
ok('7. nao repete a candidata no grupo das outras',
   (html.match(/value="g_conj"/g) || []).length === 1, html);
ok('8. criar grade nova continua no fim', html.includes('__nova__'), html);
ok('9. havendo candidata, nao acusa "nenhuma grade"',
   !html.includes('nenhuma grade com estes tamanhos'), html);

console.log('');
console.log('-- escolhida a mao, fora das candidatas --');
const L2 = { tamanhos: lidos, grades: cands, grade: STATE.grades[1] };   // M ao G3
const html2 = _riscoCelulaGrade(L2, 0);
ok('10. a escolhida fica marcada no seletor',
   /value="g_certa" selected/.test(html2), html2);
ok('11. e a linha avisa que foi a mao, dizendo o que o PDF trazia',
   html2.includes('escolhida à mão') && html2.includes('M 1 · G 1'), html2);
const L3 = { tamanhos: lidos, grades: cands, grade: STATE.grades[0] };   // a candidata
ok('12. escolher a propria candidata NAO gera o aviso',
   !_riscoCelulaGrade(L3, 0).includes('escolhida à mão'), _riscoCelulaGrade(L3, 0));

console.log('');
console.log('-- nenhuma candidata (o CORPO 2 anterior, sem modelo completo) --');
const L4 = { tamanhos: {}, grades: [], grade: null };
const html4 = _riscoCelulaGrade(L4, 0);
ok('13. acusa que nada casou', html4.includes('nenhuma grade com estes tamanhos'), html4);
ok('14. e mesmo assim oferece o cadastro inteiro',
   html4.includes('value="g_certa"') && html4.includes('value="g_conj"')
   && html4.includes('value="g_outra"'), html4);
ok('15. sem candidatas nao ha o grupo do PDF',
   !html4.includes('com os tamanhos do PDF'), html4);

console.log('');
console.log('-- a busca por tamanhos continua exata --');
ok('16. casamento e por distribuicao inteira, nao por subconjunto',
   _riscoGradesQueCasam({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 })
     .map(g => g.id).join() === 'g_certa',
   _riscoGradesQueCasam({ m: 1, g: 1, gg: 1, g1: 1, g2: 1, g3: 1 }).map(g => g.nome).join());
ok('17. quantidade diferente nao casa', _riscoGradesQueCasam({ m: 2, g: 2 }).length === 0);
ok('18. tabela vazia nao casa com nada', _riscoGradesQueCasam({}).length === 0);

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
