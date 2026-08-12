/* Rode com:  node testes/fase-vies.js

   A FASE VIÉS ENTRA SOZINHA no cadastro de grade.

   Toda peça da casa leva viés, e ele quase nunca vinha no cadastro: quem
   preenchia a grade lançava o corpo, a gola e ia embora. A OS saía sem o pano
   do viés e alguém tinha que lembrar de completar depois, abrindo outra grade
   só para achar a medida — que é sempre a mesma coisa:

     largura     = 1,17 m (o rolo que a casa compra)
     comprimento = o da primeira fase, arredondado para o metro de CIMA

   As duas regras que este teste protege, e que são o coração da coisa:

   1. O automático NUNCA pisa no que foi digitado nem no que já estava salvo.
      Ele escreve no campo vazio e no que ele mesmo escreveu antes (data-sug).
      Sem isso, corrigir uma grade antiga apagaria a medida boa que estava lá.

   2. Uma grade que já tem viés não ganha um segundo — inclusive quando ele se
      chama "Gola e Viés" ou "Gola/Viés", que é como metade do cadastro escreve.

   O teste recorta as funções do app.js de verdade e roda contra um DOM de
   mentira. O que ele NÃO cobre é o desenho da linha na tela (addFaseGradeRow,
   que é HTML puro) — aqui está dublado de propósito, para o teste falar só
   sobre a regra. */
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

// ---------------------------------------------------------------- DOM falso
// O bastante para as funções da regra: um container de blocos, cada bloco com
// os três campos que interessam. `dataset` é objeto de verdade, porque é nele
// que mora a marca `sug` que separa "escrevi eu" de "digitou o usuário".
function novoInput(classe, valor) {
  return { classe, value: valor == null ? '' : String(valor), dataset: {} };
}
function novoBloco(nome, comp, larg) {
  const campos = [novoInput('fase-nome', nome), novoInput('fase-comp', comp), novoInput('fase-larg', larg)];
  return {
    campos,
    querySelector: sel => campos.find(c => c.classe === sel.replace('.', '')) || null
  };
}
function novoContainer(blocos) {
  return {
    blocos,
    querySelectorAll: () => blocos,
    querySelector: sel => (sel.indexOf('last-child') >= 0 ? blocos[blocos.length - 1] || null : null)
  };
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const campo = (bloco, cls) => bloco.querySelector('.' + cls);

// Um cadastro de grade aberto na tela, reduzido ao que a regra enxerga.
// `addFaseGradeRow` entra dublado: a linha de verdade é HTML puro, e o que este
// teste tem a dizer é sobre a REGRA, não sobre o desenho dela.
function ambiente(blocos) {
  const cont = novoContainer(blocos.slice());
  const doc = { getElementById: id => (id === 'm-fases-container' ? cont : null) };
  const fn = new Function('document', 'cont', 'novoBloco', `
    ${recorte('function _normNome', 'a normalizacao de nome')}
    ${recorte('function _normFaseNome', 'a normalizacao de nome de fase')}
    ${recorte('function _ehFaseVies', 'o reconhecedor da fase vies')}
    ${recorte('function _compViesSugerido', 'o comprimento sugerido do vies')}
    ${recorte('function atualizarFaseVies', 'o acerto da fase vies')}
    ${recorte('function garantirFaseVies', 'a garantia da fase vies')}
    const VIES_LARGURA_PADRAO_M = ${/const VIES_LARGURA_PADRAO_M = ([\d.]+)/.exec(src)[1]};
    function addFaseGradeRow(fase) { cont.blocos.push(novoBloco((fase && fase.nome) || '', '', '')); }
    return { _ehFaseVies, _compViesSugerido, atualizarFaseVies, garantirFaseVies };
  `);
  return { cont, api: fn(doc, cont, novoBloco) };
}

console.log('-- reconhecer a fase do vies pelo nome --');
{
  const { api } = ambiente([]);
  ok('1. "Viés" e vies', api._ehFaseVies('Viés') === true);
  ok('1b. sem acento tambem', api._ehFaseVies('Vies') === true);
  ok('1c. "Gola e Viés" JA e a linha do vies', api._ehFaseVies('Gola e Viés') === true);
  ok('1d. "Gola/Viés" idem (a barra vira espaco)', api._ehFaseVies('Gola/Viés') === true);
  ok('1e. "Corpo" nao e', api._ehFaseVies('Corpo') === false);
  ok('1f. palavra INTEIRA: "Enviesado" nao e vies', api._ehFaseVies('Enviesado') === false,
     api._ehFaseVies('Enviesado'));
  ok('1g. nome vazio nao e', api._ehFaseVies('') === false && api._ehFaseVies(null) === false);
}

console.log('\n-- grade nova: a linha do vies nasce preenchida --');
{
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80')]);
  api.garantirFaseVies();
  ok('2. virou 2 fases (Corpo + Viés)', cont.blocos.length === 2, cont.blocos.length);
  const v = cont.blocos[1];
  ok('2b. a nova se chama Viés', campo(v, 'fase-nome').value === 'Viés', campo(v, 'fase-nome').value);
  ok('2c. largura 1.17', campo(v, 'fase-larg').value === '1.17', campo(v, 'fase-larg').value);
  ok('2d. comprimento 6,50 -> 7 (para CIMA)', campo(v, 'fase-comp').value === '7', campo(v, 'fase-comp').value);
  ok('2e. os dois campos ficam marcados como da regra',
     campo(v, 'fase-larg').dataset.sug === '1' && campo(v, 'fase-comp').dataset.sug === '1');
  ok('2f. a fase do corpo nao foi tocada',
     campo(cont.blocos[0], 'fase-comp').value === '6.50' && campo(cont.blocos[0], 'fase-larg').value === '1.80');
}

console.log('\n-- o arredondamento --');
{
  const casos = [['6.50', '7'], ['6.01', '7'], ['6', '6'], ['6.00', '6'], ['0.4', '1'], ['12.9', '13']];
  casos.forEach(([entra, sai]) => {
    const { cont, api } = ambiente([novoBloco('Corpo', entra, '1.80')]);
    api.garantirFaseVies();
    const got = campo(cont.blocos[1], 'fase-comp').value;
    ok(`3. ${entra} m -> ${sai} m`, got === sai, got);
  });
  // Inteiro NAO vira o proximo: arredondar para cima e teto, nao "somar um".
  const { cont, api } = ambiente([novoBloco('Corpo', '8', '1.80')]);
  api.garantirFaseVies();
  ok('3b. 8 m continua 8 m (teto, nao +1)', campo(cont.blocos[1], 'fase-comp').value === '8',
     campo(cont.blocos[1], 'fase-comp').value);
}

console.log('\n-- grade que JA tem vies: nao ganha um segundo --');
{
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80'), novoBloco('Viés', '7', '1.17')]);
  api.garantirFaseVies();
  ok('4. continua com 2 fases', cont.blocos.length === 2, cont.blocos.length);
}
{
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80'), novoBloco('Gola e Viés', '3', '1.17')]);
  api.garantirFaseVies();
  ok('4b. "Gola e Viés" conta como vies — nao entra outra linha', cont.blocos.length === 2, cont.blocos.length);
  ok('4c. e o que estava salvo nela fica intacto',
     campo(cont.blocos[1], 'fase-comp').value === '3', campo(cont.blocos[1], 'fase-comp').value);
}

console.log('\n-- o automatico nunca pisa no que ja existe --');
{
  // Grade antiga sendo CORRIGIDA: o vies ja tem medidas proprias, diferentes da
  // regra. Elas mandam — foi alguem que as pos ali.
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80'), novoBloco('Viés', '4', '0.90')]);
  api.garantirFaseVies();
  ok('5. comprimento salvo (4) sobrevive, nao vira 7',
     campo(cont.blocos[1], 'fase-comp').value === '4', campo(cont.blocos[1], 'fase-comp').value);
  ok('5b. largura salva (0.90) sobrevive, nao vira 1.17',
     campo(cont.blocos[1], 'fase-larg').value === '0.90', campo(cont.blocos[1], 'fase-larg').value);
}
{
  // Vies pela metade: so a largura estava salva. O vazio se preenche, o cheio nao.
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80'), novoBloco('Viés', '', '0.90')]);
  api.garantirFaseVies();
  ok('5c. o comprimento VAZIO se preenche', campo(cont.blocos[1], 'fase-comp').value === '7',
     campo(cont.blocos[1], 'fase-comp').value);
  ok('5d. e a largura cheia continua como estava', campo(cont.blocos[1], 'fase-larg').value === '0.90',
     campo(cont.blocos[1], 'fase-larg').value);
}

console.log('\n-- quem digita vira dono do campo --');
{
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80')]);
  api.garantirFaseVies();
  const v = cont.blocos[1];
  ok('6. de saida o vies acompanha (7)', campo(v, 'fase-comp').value === '7');

  // A primeira fase muda: enquanto ninguem digitou no vies, ele acompanha.
  campo(cont.blocos[0], 'fase-comp').value = '9.20';
  api.atualizarFaseVies();
  ok('6b. corpo 9,20 -> vies acompanha para 10', campo(v, 'fase-comp').value === '10',
     campo(v, 'fase-comp').value);

  // Agora o usuario digita no vies. O oninput do campo limpa a marca `sug`,
  // e a partir daqui o automatico nao mexe mais ali.
  campo(v, 'fase-comp').value = '5';
  campo(v, 'fase-comp').dataset.sug = '';
  campo(cont.blocos[0], 'fase-comp').value = '12.30';
  api.atualizarFaseVies();
  ok('6c. digitado a mao (5) nao e mais sobrescrito', campo(v, 'fase-comp').value === '5',
     campo(v, 'fase-comp').value);
}

console.log('\n-- a base do calculo e a primeira fase QUE NAO E VIES --');
{
  const { cont, api } = ambiente([novoBloco('Viés', '', ''), novoBloco('Corpo', '5.10', '1.80')]);
  api.garantirFaseVies();
  ok('7. vies em primeiro lugar nao calcula em cima de si mesmo',
     campo(cont.blocos[0], 'fase-comp').value === '6', campo(cont.blocos[0], 'fase-comp').value);
}

console.log('\n-- primeira fase ainda sem comprimento --');
{
  const { cont, api } = ambiente([novoBloco('Corpo', '', '1.80')]);
  api.garantirFaseVies();
  const v = cont.blocos[1];
  ok('8. o vies entra assim mesmo', cont.blocos.length === 2);
  ok('8b. com a largura padrao ja posta', campo(v, 'fase-larg').value === '1.17', campo(v, 'fase-larg').value);
  ok('8c. e o comprimento em branco, esperando', campo(v, 'fase-comp').value === '', campo(v, 'fase-comp').value);
  ok('8d. continua marcado como da regra, para acompanhar depois',
     campo(v, 'fase-comp').dataset.sug === '1');

  // Digitado o comprimento do corpo, o vies se completa sozinho.
  campo(cont.blocos[0], 'fase-comp').value = '7.10';
  api.atualizarFaseVies();
  ok('8e. digitado o corpo (7,10), o vies vira 8', campo(v, 'fase-comp').value === '8',
     campo(v, 'fase-comp').value);
}

console.log('\n-- grade sem nenhuma fase de vies depois de remover --');
{
  const { cont, api } = ambiente([novoBloco('Corpo', '6.50', '1.80')]);
  // Nenhum vies no container: atualizarFaseVies nao pode explodir nem inventar.
  api.atualizarFaseVies();
  ok('9. sem linha de vies, nao faz nada e nao quebra', cont.blocos.length === 1, cont.blocos.length);
}

console.log(falhas ? `\n>>> ${falhas} FALHA(S)` : '\n>>> todos passaram');
process.exit(falhas ? 1 : 0);
