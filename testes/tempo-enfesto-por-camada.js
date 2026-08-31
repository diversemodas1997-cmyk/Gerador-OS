/* Rode com:  node testes/tempo-enfesto-por-camada.js

   O TEMPO PREVISTO DE UMA FASE DE ENFESTO ESCALA COM AS CAMADAS.

   Junior, 31/08/2026: "o programa deve analisar a media de tempo necessario nas
   fases de enfesto, alem de como ja e feito, tambem a media de tempo POR CAMADA
   em cada fase de enfesto, pois nem todos os enfestos produzem o limite maximo
   de camadas, mas sim, apenas algumas camadas. Logo, o tempo estimado deve
   considerar tambem fases que nao sao determinadas por preencher o maximo de
   camadas."

   O PROBLEMA. A mesma fase enfestada com 78 camadas e com 12 nao leva o mesmo
   tempo. A media chapada das duas nao descreve nenhuma das duas: aplicada ao
   enfesto curto, ela acusa "muito abaixo da media" um trabalho perfeitamente
   normal; aplicada ao cheio, promete um tempo que nao vai dar. Era o que a
   folha de OS e a tabela da grade faziam.

   A taxa por camada ja existia no programa -- `_opTempoEnfestoPrevisto`, na
   cascata de operacoes, converte cada medicao em minutos por camada antes de
   aplicar. O que mudou e que ela deixou de viver so la.

   O que este teste protege:

     · a previsao ESCALA: o dobro de camadas pede aproximadamente o dobro;
     · ela sai da TAXA, nao da media chapada -- num historico de 78/36/12
       camadas, prever 12 tem de dar perto do que 12 camadas levaram, e nao a
       media dos tres;
     · medicao sem camadas conhecidas entra na media chapada mas NAO na taxa
       (dividir por zero seria pior do que nao ter a taxa);
     · sem taxa apurada, ou sem camadas para aplicar, vale a media chapada -- o
       comportamento antigo continua sendo o piso, nunca um erro;
     · `_camadasDaFase` le o bloco da fase e so cai nas camadas da OS quando a
       fase nao tem bloco proprio.

   O teste recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function corta(nome) {
  const i = src.indexOf(nome);
  if (i < 0) { console.error('nao achei ' + nome + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  if (j < 0) { console.error('nao achei o fim de ' + nome); process.exit(1); }
  return src.slice(i, j + 2);
}

const api = (STATE) => new Function('STATE', `
  ${corta('function _normNome')}
  ${corta('function _normFaseNome')}
  ${corta('function _osGradeKey')}
  ${corta('function _opMin(')}
  // Sem pausas: o desconto de cafe/almoco tem o teste dele, e aqui ele so
  // embaralharia os numeros do que esta sendo medido.
  const _opMinutosTrabalhados = (ini, fim) => Math.max(0, fim - ini);
  ${corta('function _camadasDaFase')}
  ${corta('function _tempoEsperadoFase')}
  ${corta('function _mediaTempoFasesSimilares')}
  return { _camadasDaFase, _tempoEsperadoFase, _mediaTempoFasesSimilares };
`)(STATE);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

// Uma OS medida: uma fase "Corpo", tantas camadas, do inicio ao fim.
const osMedida = (id, num, camadas, ini, fim) => ({
  id, os: num, gradeId: 'gr1',
  grade: { descricao: 'M-G | CM.LISA', m: 1, g: 1 },
  fases: [{ ordem: 1, nome: 'Corpo' }],
  enfesto: { camadas, blocos: [{ ordem: 1, camadas }] },
  progresso: { enfestosTempos: { 1: { enfIni: ini, enfFim: fim } } }
});

// Historico coerente: 2 min por camada, em enfestos de tamanhos bem diferentes.
//   78 camadas -> 156 min      36 -> 72 min      12 -> 24 min
const historico = [
  osMedida('h1', '0401', 78, '07:00', '09:36'),
  osMedida('h2', '0402', 36, '07:00', '08:12'),
  osMedida('h3', '0403', 12, '07:00', '07:24')
];
const mediaChapada = Math.round((156 + 72 + 24) / 3);   // 84 min

console.log('-- a previsao escala com as camadas --');
{
  const alvo = osMedida('alvo', '0500', 12, '', '');
  const STATE = { ordens: historico.concat([alvo]), grades: [{ id: 'gr1', nome: 'M-G | CM.LISA' }] };
  const a = api(STATE);
  const med = a._mediaTempoFasesSimilares(alvo).corpo;
  ok('1. a media chapada continua sendo apurada (era o que existia)',
     med && med.mediaMin === mediaChapada, med && med.mediaMin);
  ok('2. e agora existe tambem a taxa por camada -- 2 min, das 3 medicoes',
     med && Math.abs(med.minPorCamada - 2) < 1e-9 && med.nCam === 3,
     med && (med.minPorCamada + ' min/cam em ' + med.nCam));
  ok('3. prever 12 camadas da 24 min, nao os 84 da media chapada',
     a._tempoEsperadoFase(med, 12) === 24, a._tempoEsperadoFase(med, 12));
  ok('4. prever 78 camadas da 156 min -- o dobro de camadas, o dobro de tempo',
     a._tempoEsperadoFase(med, 78) === 156 && a._tempoEsperadoFase(med, 39) === 78,
     a._tempoEsperadoFase(med, 78) + ' / ' + a._tempoEsperadoFase(med, 39));
}

console.log('-- o comportamento antigo continua sendo o piso --');
{
  // Medicoes SEM camadas: a taxa nao pode nascer de divisao por zero.
  const semCam = historico.map((o, i) => ({
    ...o, enfesto: { camadas: 0, blocos: [] }
  }));
  const alvo = osMedida('alvo', '0500', 12, '', '');
  const a = api({ ordens: semCam.concat([alvo]), grades: [{ id: 'gr1', nome: 'M-G | CM.LISA' }] });
  const med = a._mediaTempoFasesSimilares(alvo).corpo;
  ok('5. sem camadas conhecidas nao se apura taxa nenhuma',
     med && med.nCam === 0 && med.minPorCamada === 0, med && JSON.stringify(med));
  ok('6. e a previsao cai na media chapada, como era antes',
     a._tempoEsperadoFase(med, 12) === mediaChapada, a._tempoEsperadoFase(med, 12));
}
{
  const alvo = osMedida('alvo', '0500', 12, '', '');
  const a = api({ ordens: historico.concat([alvo]), grades: [{ id: 'gr1', nome: 'M-G | CM.LISA' }] });
  const med = a._mediaTempoFasesSimilares(alvo).corpo;
  ok('7. com taxa, mas sem camadas para aplicar, tambem vale a media chapada',
     a._tempoEsperadoFase(med, 0) === mediaChapada, a._tempoEsperadoFase(med, 0));
  ok('8. e sem linha de historico nenhuma nao inventa numero',
     a._tempoEsperadoFase(null, 12) === null, a._tempoEsperadoFase(null, 12));
}

console.log('-- de onde saem as camadas de uma fase --');
{
  const a = api({ ordens: [], grades: [] });
  const o = { enfesto: { camadas: 40, blocos: [{ ordem: 1, camadas: 78 }, { ordem: 2, camadas: 15 }] } };
  ok('9. o bloco da fase manda', a._camadasDaFase(o, 1) === 78 && a._camadasDaFase(o, 2) === 15,
     a._camadasDaFase(o, 1) + ' / ' + a._camadasDaFase(o, 2));
  ok('10. fase sem bloco proprio cai nas camadas da OS', a._camadasDaFase(o, 3) === 40,
     a._camadasDaFase(o, 3));
  ok('11. a ordem pode vir como texto (e o que o dataset do DOM entrega)',
     a._camadasDaFase(o, '1') === 78, a._camadasDaFase(o, '1'));
  ok('12. OS sem enfesto nao quebra', a._camadasDaFase({}, 1) === 0 && a._camadasDaFase(null, 1) === 0);
}

console.log('-- a taxa e a media DAS TAXAS, nao a razao dos totais --');
{
  // Um enfesto gigante e lento (78 cam, 3 min/cam) e um curto e rapido (6 cam,
  // 1 min/cam). Razao dos totais: 240/84 = 2,86 -- o gigante domina. Media das
  // taxas: 2,0 -- cada medicao pesa igual, que e como a cascata sempre fez.
  const l = [osMedida('h1', '0401', 78, '07:00', '10:54'), osMedida('h2', '0402', 6, '07:00', '07:06')];
  const alvo = osMedida('alvo', '0500', 12, '', '');
  const a = api({ ordens: l.concat([alvo]), grades: [{ id: 'gr1', nome: 'M-G | CM.LISA' }] });
  const med = a._mediaTempoFasesSimilares(alvo).corpo;
  ok('13. duas medicoes de 3 e 1 min por camada dao taxa 2, e nao 2,86',
     Math.abs(med.minPorCamada - 2) < 1e-9, med.minPorCamada);
}

console.log(falhas === 0 ? '\nTODOS OS TESTES PASSARAM' : `\n${falhas} FALHA(S)`);
process.exit(falhas === 0 ? 0 : 1);
