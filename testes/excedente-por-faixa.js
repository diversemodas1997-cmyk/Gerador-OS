/* Rode com:  node testes/excedente-por-faixa.js

   O EXCEDENTE POR FAIXA DE COMPRIMENTO.

   A sobra de enfesto acompanha o comprimento que a fase estende — um viés de
   1 m não precisa da mesma ponta que um corpo de 9 m:

       até 1,50 m ...... 10 cm
       1,50 m a 9 m .... 15 cm
       9 m a 12 m ...... 20 cm

   A faixa dos 15 cm ia até 8 m e foi esticada para 9 m em 12/08/2026: o corpo
   de 8,20 m das grades "P ao G3" é enfesto do mesmo feitio dos de 7 m, e os
   20 cm ali eram sobra a mais em pano caro. Mexeu em 20 fases do cadastro.

   O que este teste existe para proteger são as BORDAS e o SILÊNCIO:

   - As bordas ficam na faixa DE BAIXO. 1,50 m exato leva 10 cm, não 15; 8 m
     exatos levam 15 cm, não 20. O enunciado ("até 1,50", "de 1,50 até 8")
     repete os números nas duas pontas, e é aí que uma reescrita em massa de
     124 grades erra em silêncio.
   - Fase sem comprimento e fase acima de 12 m devolvem null, e null quer dizer
     NÃO MEXER. Se um dia isso virar um número, a alteração em massa passa a
     escrever excedente em cadastro que ninguém mediu — que é justamente o
     estrago que a função foi escrita para não fazer.

   O teste recorta a função do app.js de verdade. */
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

// A tabela de faixas vem do app.js tambem: se alguem mudar os limites la, este
// teste tem que falar sobre os limites NOVOS, nao sobre uma copia velha daqui.
const tabela = /const EXCEDENTE_FAIXAS = \[[\s\S]*?\];/.exec(src);
if (!tabela) { console.error('nao achei EXCEDENTE_FAIXAS no app.js'); process.exit(1); }

const api = new Function(`
  ${tabela[0]}
  ${recorte('function excedentePorComprimento', 'a regra das faixas')}
  return { EXCEDENTE_FAIXAS, excedentePorComprimento };
`)();
const { excedentePorComprimento: exc, EXCEDENTE_FAIXAS } = api;

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

console.log('-- a tabela e a que o pedido descreve --');
ok('0. tres faixas: 1,50 / 9 / 12', EXCEDENTE_FAIXAS.length === 3
   && EXCEDENTE_FAIXAS[0].ate === 1.5 && EXCEDENTE_FAIXAS[1].ate === 9 && EXCEDENTE_FAIXAS[2].ate === 12,
   EXCEDENTE_FAIXAS);
ok('0b. 10 / 15 / 20 cm', EXCEDENTE_FAIXAS.map(f => f.cm).join(',') === '10,15,20',
   EXCEDENTE_FAIXAS.map(f => f.cm));

console.log('\n-- faixa 1: ate 1,50 m -> 10 cm --');
[['0.30', 10], ['1', 10], ['1.17', 10], ['1.49', 10]].forEach(([e, s]) =>
  ok(`1. ${e} m -> ${s} cm`, exc(e) === s, exc(e)));
ok('1b. o vies padrao (1,17) cai aqui', exc('1.17') === 10, exc('1.17'));

console.log('\n-- faixa 2: de 1,50 a 9 m -> 15 cm --');
[['1.51', 15], ['2.94', 15], ['4.5493', 15], ['6.50', 15], ['7.99', 15], ['8.99', 15]].forEach(([e, s]) =>
  ok(`2. ${e} m -> ${s} cm`, exc(e) === s, exc(e)));
// O trecho que MUDOU DE FAIXA em 12/08: 20 fases do cadastro moram aqui, entre
// elas o corpo de 8,20 m das grades "P ao G3", que era o caso do pedido.
[['8.01', 15], ['8.20', 15], ['8.44', 15], ['8.50', 15]].forEach(([e, s]) =>
  ok(`2b. ${e} m -> ${s} cm (antes eram 20)`, exc(e) === s, exc(e)));

console.log('\n-- faixa 3: de 9 a 12 m -> 20 cm --');
[['9.01', 20], ['10', 20], ['11.99', 20], ['12', 20]].forEach(([e, s]) =>
  ok(`3. ${e} m -> ${s} cm`, exc(e) === s, exc(e)));

console.log('\n-- AS BORDAS: o limite fica na faixa DE BAIXO --');
ok('4. 1,50 exato -> 10 cm (nao 15)', exc('1.50') === 10, exc('1.50'));
ok('4b. 1,5 escrito curto -> 10 cm', exc('1.5') === 10, exc('1.5'));
ok('4c. 9 exato -> 15 cm (nao 20)', exc('9') === 15, exc('9'));
ok('4d. 9.00 -> 15 cm', exc('9.00') === 15, exc('9.00'));
ok('4e. 12 exato -> 20 cm (ultima faixa inclui o teto)', exc('12') === 20, exc('12'));
ok('4f. 8 exato agora e 15, e nao mais a borda', exc('8') === 15, exc('8'));

console.log('\n-- null quer dizer NAO MEXER --');
ok('5. acima de 12 m nao tem regra', exc('12.01') === null, exc('12.01'));
ok('5b. bem acima tambem', exc('30') === null, exc('30'));
ok('5c. sem comprimento (vazio)', exc('') === null, exc(''));
ok('5d. null', exc(null) === null, exc(null));
ok('5e. undefined', exc(undefined) === null, exc(undefined));
ok('5f. zero nao e comprimento', exc('0') === null, exc('0'));
ok('5g. negativo nao e comprimento', exc('-3') === null, exc('-3'));
ok('5h. texto que nao e numero', exc('abc') === null, exc('abc'));

console.log('\n-- como o comprimento chega do cadastro --');
// O campo e `type="number"`, entao o valor salvo vem com PONTO; mas o dado
// antigo (e o digitado a mao no import de JSON) aparece com virgula.
ok('6. numero de verdade, nao string', exc(6.5) === 15, exc(6.5));
ok('6b. virgula decimal (dado antigo)', exc('6,50') === 15, exc('6,50'));
ok('6c. virgula na borda de baixo', exc('1,50') === 10, exc('1,50'));
ok('6d. virgula na borda de cima', exc('8,00') === 15, exc('8,00'));
ok('6e. com espaco em volta', exc(' 10 ') === 20, exc(' 10 '));

console.log('\n-- nenhuma faixa devolve o padrao da casa por acidente --');
// 15 e o EXCEDENTE_ENFESTO_PADRAO_CM. Se a funcao devolvesse 15 para "sem
// base", a alteracao em massa carimbaria 15 cm em fase sem medida e ninguem
// veria a diferenca entre "a regra decidiu" e "a regra nao sabia".
ok('7. sem base NAO devolve 15', exc('') !== 15 && exc(null) !== 15 && exc('40') !== 15);

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
