/* Rode com:  node testes/encaixe-parcial-pasta.js

   ENCAIXE PARCIAL: o risco seguinte acha a grade que o anterior criou.

   O caso real de 12/08/2026, confirmado contra o banco da fábrica:

     CM.REC - CORPO 1 - G.pdf   1,07 m   o CAD conta G=1  -> criou a grade
                                                              "G | CM.REC | 117cm"
                                                              (tamanhos 0/0/1/0/0/0/0)
     CM.REC - CORPO 2 - G.pdf   0,36 m   o CAD conta G=2  -> procura 0/0/2/0/0/0/0

   O Corpo 2 é peça pequena e cabem duas no mesmo pano; o Corpo 1 cabe uma. O CAD
   conta MODELOS COMPLETOS POR ENCAIXE — está certo do ponto de vista dele, e não
   é a grade. Como o destino era escolhido só pelos tamanhos, o Corpo 2 caía em
   "criar uma grade nova" e morria em "já existe uma grade chamada...", com a
   grade certa a uma linha de distância na mesma tela.

   A pista já existia e não estava sendo usada fora das fases agregadoras: o
   assistente guarda em `gradePorPasta` em que grade cada pasta lançou.

   O que este teste protege:

   1. Achar a grade irmã pela pasta EXATA.
   2. NÃO atravessar larguras. "…/174 cm" e "…/182 cm" são duas grades, e um
      Corpo 2 mora sempre na pasta do Corpo 1 dele. Aceitar a família (um nível
      acima, que a regra da ribana aceita) mandaria o corpo de 174 para a grade
      de 182 — trocar a medida de uma grade boa pela de outro pano.
   3. Só valer quando o tamanho não achou nada. Casando por tamanho, manda o
      caminho provado.

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

function api(STATE, _pastaWiz) {
  return new Function('STATE', '_pastaWiz', `
    ${recorte('function _pastaPastaDoArquivo', 'a pasta exata do arquivo')}
    ${recorte('function _pastaPastaDaGrade', 'a pasta da familia')}
    ${recorte('function _pastaGradeIrmaDaPasta', 'a grade irma da pasta')}
    ${recorte('function _riscoGradesQueCasam', 'a busca por tamanhos')}
    return { _pastaGradeIrmaDaPasta, _riscoGradesQueCasam, _pastaPastaDoArquivo };
  `)(STATE, _pastaWiz);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const T = o => Object.assign({ p: 0, m: 0, g: 0, gg: 0, g1: 0, g2: 0, g3: 0 }, o);
// O cadastro como estava no servidor no momento do print.
const STATE = { grades: [
  { id: 'g_corpo1', nome: 'G | CM.REC | 117cm', tamanhos: T({ g: 1 }),
    fases: [{ ordem: 1, nome: 'Corpo 1', comp: '1.07', larg: '1.170' }] },
  { id: 'g_2g2gg', nome: '2G-2GG | CM.REC | 117cm', tamanhos: T({ g: 2, gg: 2 }) }
] };
const grupo = caminho => ({ itens: [{ caminho }] });

console.log('-- o tamanho sozinho nao acha (era a causa) --');
{
  const { _riscoGradesQueCasam } = api(STATE, {});
  ok('1. o Corpo 2 procura G=2 e nao ha grade nenhuma assim',
     _riscoGradesQueCasam(T({ g: 2 })).length === 0, _riscoGradesQueCasam(T({ g: 2 })).map(g => g.nome));
  ok('1b. e a grade certa existe, so que com G=1',
     _riscoGradesQueCasam(T({ g: 1 })).map(g => g.nome).join() === 'G | CM.REC | 117cm');
}

console.log('-- a pasta acha --');
{
  const wiz = { gradePorPasta: { 'CM.REC/G/117 cm': 'g_corpo1' } };
  const { _pastaGradeIrmaDaPasta } = api(STATE, wiz);
  const r = _pastaGradeIrmaDaPasta(grupo('CM.REC/G/117 cm/CM.REC - CORPO 2 - G.pdf'));
  ok('2. o Corpo 2 cai na grade que o Corpo 1 criou', !!r && r.grade.id === 'g_corpo1', r);
  ok('2b. e diz por que', !!r && /mesma pasta/.test(r.porque), r && r.porque);
}

console.log('-- sem memoria da pasta, nao inventa --');
{
  const { _pastaGradeIrmaDaPasta } = api(STATE, { gradePorPasta: {} });
  ok('3. pasta sem nada lancado devolve null',
     _pastaGradeIrmaDaPasta(grupo('CM.REC/G/117 cm/CORPO 2.pdf')) === null);
}
{
  // A memoria aponta para uma grade que nao existe mais (apagada no meio da
  // importacao): nao pode devolver objeto quebrado.
  const { _pastaGradeIrmaDaPasta } = api(STATE, { gradePorPasta: { 'X/Y': 'g_sumiu' } });
  ok('3b. id orfao devolve null, nao objeto sem grade',
     _pastaGradeIrmaDaPasta(grupo('X/Y/CORPO 2.pdf')) === null);
}
{
  const { _pastaGradeIrmaDaPasta } = api(STATE, {});
  ok('3c. wizard sem gradePorPasta nao quebra',
     _pastaGradeIrmaDaPasta(grupo('CM.REC/G/117 cm/CORPO 2.pdf')) === null);
}

console.log('-- NAO atravessa largura: e a pasta EXATA --');
{
  const S = { grades: [
    { id: 'g174', nome: 'M-G-GG | BM.LISA | 174cm', tamanhos: T({ m: 1, g: 1, gg: 1 }) },
    { id: 'g182', nome: 'M-G-GG | BM.LISA | 182cm', tamanhos: T({ m: 1, g: 1, gg: 1 }) }
  ] };
  const wiz = { gradePorPasta: {
    'BM.LISA/M-G-GG/174 cm': 'g174',
    'BM.LISA/M-G-GG/182 cm': 'g182',
    'BM.LISA/M-G-GG': 'g182'          // a familia, que a regra da ribana aceita
  } };
  const { _pastaGradeIrmaDaPasta } = api(S, wiz);
  const r174 = _pastaGradeIrmaDaPasta(grupo('BM.LISA/M-G-GG/174 cm/CORPO 2.pdf'));
  ok('4. o corpo de 174 cai na grade de 174', !!r174 && r174.grade.id === 'g174', r174 && r174.grade.nome);
  const r182 = _pastaGradeIrmaDaPasta(grupo('BM.LISA/M-G-GG/182 cm/CORPO 2.pdf'));
  ok('4b. e o de 182 na de 182', !!r182 && r182.grade.id === 'g182', r182 && r182.grade.nome);
  // A pasta de largura que ainda nao lancou nada NAO herda a familia.
  const wiz2 = { gradePorPasta: { 'BM.LISA/M-G-GG': 'g182' } };
  const { _pastaGradeIrmaDaPasta: f2 } = api(S, wiz2);
  ok('4c. pasta de 177 sem lancamento proprio NAO pega a familia (seria outro pano)',
     f2(grupo('BM.LISA/M-G-GG/177 cm/CORPO 2.pdf')) === null,
     f2(grupo('BM.LISA/M-G-GG/177 cm/CORPO 2.pdf')));
}

console.log('-- a condicao de uso, como o assistente a escreve --');
{
  // A linha do renderPastaWiz: so quando nao e agregadora, nao ha certeza, o
  // tamanho nao achou NADA e o usuario nao escolheu a mao.
  const linha = /const daPasta = \(!ehAgregadora && !jaExiste && !G\.candidatas\.length && !G\.draft\.destinoManual\)/;
  ok('5. a sugestao so entra quando o tamanho nao achou nada', linha.test(src),
     (/const daPasta = .*/.exec(src) || [''])[0]);
  ok('5b. e e sugestao: nao entra no jaExiste, entao "criar nova" continua na lista',
     !/jaExiste = .*daPasta/.test(src));
  ok('5c. o destino so e pre-escolhido se ainda nao havia um',
     /if \(daPasta && !G\.gradeId\) G\.gradeId = daPasta\.grade\.id;/.test(src));
}

console.log('-- a funcao nao olha a familia, so a pasta exata --');
{
  const corpo = recorte('function _pastaGradeIrmaDaPasta', 'a grade irma da pasta');
  ok('6. nao chama _pastaPastaDaGrade', !/_pastaPastaDaGrade/.test(corpo), corpo);
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
