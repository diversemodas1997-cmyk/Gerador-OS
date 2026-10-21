/* Rode com:  node testes/grade-convertida-os.js

   GRADE GRAVADA NA OS EM UNIDADES ANTIGAS (21/10/2026).
   A OS 0656 saía com 940 peças na OE, sendo 470: a grade gravada era a antiga,
   2M-2G-2G1-2G2-2G3 (pano dobrado), e o cadastro passou a 1 por tamanho com a
   camada tubular rendendo 2. alinharGradeConvertidaOS traz a grade da OS para o
   cadastro SÓ quando a diferença é essa conversão (gravada = cadastro × rendimento),
   e nunca mexe em OS finalizada/cancelada nem em grade mudada de verdade. */
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

const motor = [
  corta('function recomputarQuantidadesOS'),
  corta('function alinharGradeConvertidaOS')
].join('\n');

function api(grades, mult) {
  const fn = new Function('GRADES', 'MULT', `
    const STATE = { grades: GRADES };
    const multiplicadorPecaOS = () => MULT;
    ${motor}
    return { alinharGradeConvertidaOS };
  `);
  return fn(grades, mult);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const CAD = { id: 'g1', tamanhos: { p: 0, m: 1, g: 1, gg: 0, g1: 1, g2: 1, g3: 1 } };
const os656 = (extra) => Object.assign({
  os: '0656', gradeId: 'g1', statusOS: 'ensacado',
  grade: { m: 2, g: 2, g1: 2, g2: 2, g3: 2, total: 10 },
  enfesto: { camadas: 47, totalPecas: 470 },
  componentes: [{ qtdPorPeca: 1, qtdTotal: 940 }, { qtdPorPeca: 2, qtdTotal: 1880 }],
  aviamentos: [{ qtdPorPeca: 1, qtdTotal: 940 }]
}, extra || {});

{
  const o = os656();
  const r = api([CAD], 2).alinharGradeConvertidaOS(o);
  ok('0656: grade em dobro é alinhada', r === true, r);
  ok('0656: grade vira 1 por tamanho, total 5', o.grade.m === 1 && o.grade.g3 === 1 && o.grade.total === 5, o.grade);
  ok('0656: componente 940 -> 470', o.componentes[0].qtdTotal === 470, o.componentes[0].qtdTotal);
  ok('0656: componente de 2 por peça 1880 -> 940', o.componentes[1].qtdTotal === 940, o.componentes[1].qtdTotal);
  ok('0656: etiqueta 940 -> 470', o.aviamentos[0].qtdTotal === 470, o.aviamentos[0].qtdTotal);
  ok('0656: rodar de novo não muda nada', api([CAD], 2).alinharGradeConvertidaOS(o) === false);
}
{
  const g = { id: 'g2', tamanhos: { m: 1, g: 2, gg: 1 } };
  const o = os656({ os: '0659', gradeId: 'g2', grade: { m: 2, g: 4, gg: 2, total: 8 }, enfesto: { camadas: 57 } });
  ok('0659 (grade desigual 2-4-2 -> 1-2-1) é alinhada', api([g], 2).alinharGradeConvertidaOS(o) && o.grade.total === 4, o.grade);
  ok('0659: componente sai 456', o.componentes[0].qtdTotal === 456, o.componentes[0].qtdTotal);
}
ok('finalizada fica como está', api([CAD], 2).alinharGradeConvertidaOS(os656({ statusOS: 'finalizado' })) === false);
ok('cancelada fica como está', api([CAD], 2).alinharGradeConvertidaOS(os656({ statusOS: 'cancelado' })) === false);
ok('moletom (rendimento 1) não mexe', api([CAD], 1).alinharGradeConvertidaOS(os656()) === false);
{
  const o = os656({ grade: { m: 2, g: 2, g1: 2, g2: 2, g3: 4, total: 12 } });
  ok('grade mudada de verdade (não é o dobro) não é adivinhada', api([CAD], 2).alinharGradeConvertidaOS(o) === false && o.grade.g3 === 4);
}
{
  const o = os656({ grade: { m: 1, g: 1, g1: 1, g2: 1, g3: 1, total: 5 } });
  ok('grade igual ao cadastro não mexe', api([CAD], 2).alinharGradeConvertidaOS(o) === false && o.componentes[0].qtdTotal === 940);
}
ok('OS sem grade no cadastro não mexe', api([], 2).alinharGradeConvertidaOS(os656()) === false);

console.log(falhas ? `\n${falhas} FALHA(S)` : '\nTudo certo.');
process.exit(falhas ? 1 : 0);
