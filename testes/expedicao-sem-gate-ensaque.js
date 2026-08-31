/* Rode com:  node testes/expedicao-sem-gate-ensaque.js

   O ENSAQUE NAO DECIDE QUEM PODE SER ALOCADO NA EXPEDICAO.

   Junior, 31/08/2026: "o preenchimento do checklist ensacamento nao pode
   determinar a capacidade do usuario de alocar uma os no planejamento de
   expedicao."

   O QUE ACONTECEU. A OS 0500 tinha 3.744 pecas, grade viva, duas tonalidades,
   nenhuma viagem e as janelas ativas -- estava alocavel por todos os criterios
   que importam. Mesmo assim sumia da tela de Expedicao, porque a lista de
   pendentes exigia a etapa ENSAQUE marcada e a folha dela so tinha "Preparo
   Materia-prima" e "Corte".

   E o efeito era o INVERSO do proposito da lista. Medido no dia, das 174 OS com
   pecas e sem carga nenhuma, ela mostrava 86 -- quase todas antigas -- e
   escondia 88, entre elas as MAIS RECENTES (0504, 0503, 0500, 0498, 0497). A OS
   que acabou de sair do corte e justamente a que ainda nao tem o Ensaque
   marcado: a lista de "nao esqueca" escondia exatamente o que ninguem podia
   esquecer.

   O que este teste protege:

     · a OS entra na lista por DUAS condicoes -- tem peca, nao esta em carga --
       e o Ensaque nao e uma delas;
     · nao ensacada e ensacada saem lado a lado, e a nao ensacada nao vem
       depois nem "marcada de outro jeito": a ordem e pelo numero da OS;
     · o `ensacada` continua sendo CALCULADO e devolvido, porque a lista mostra
       a etapa numa coluna -- tirar o gate nao pode virar tirar a informacao;
     · OS ja alocada em alguma carga nao aparece, ensacada ou nao;
     · OS sem peca nao aparece, ensacada ou nao.

   O teste recorta a funcao do app.js de verdade. */
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

// `_expPecasOS` e `osEnsacada` sao os dois insumos, e aqui eles viram stubs de
// proposito: o que esta sob teste e a REGRA DE ENTRADA, nao a contagem de pecas
// nem a leitura do checklist (essas tem os testes delas).
const api = new Function(`
  const _expPecasOS = o => Number(o.pecas) || 0;
  const osEnsacada = o => !!o.ensacada;
  ${corta('function _expOsSemCarga')}
  return { _expOsSemCarga };
`)();

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const ordens = [
  { id: 'a', os: '0500', pecas: 3744, ensacada: false },   // a do relato
  { id: 'b', os: '0499', pecas: 3888, ensacada: true },
  { id: 'c', os: '0498', pecas: 6860, ensacada: false },
  { id: 'd', os: '0490', pecas: 1000, ensacada: true },    // ja embarcada
  { id: 'e', os: '0489', pecas: 1000, ensacada: false },   // ja embarcada
  { id: 'f', os: '0488', pecas: 0,    ensacada: true },    // sem peca
  { id: 'g', os: '0487', pecas: 0,    ensacada: false }    // sem peca
];
const cargas = [{ osId: 'd' }, { osId: 'e' }];

const saida = api._expOsSemCarga(ordens, cargas);
const numeros = saida.map(x => x.o.os);

console.log('-- o Ensaque nao e criterio --');
ok('1. a OS 0500, NAO ensacada, entra na lista',
   numeros.includes('0500'), numeros.join(', '));
ok('2. entram as tres com peca e sem carga, ensacadas ou nao',
   numeros.join(',') === '0500,0499,0498', numeros.join(','));
ok('3. e a ordem e pelo numero da OS, nao pelo Ensaque',
   saida[0].o.os === '0500' && saida[1].o.os === '0499' && saida[2].o.os === '0498',
   numeros.join(','));

console.log('-- o que continua fora --');
ok('4. OS ja alocada nao aparece -- nem a ensacada nem a que nao esta',
   !numeros.includes('0490') && !numeros.includes('0489'), numeros.join(','));
ok('5. OS sem peca nao aparece -- nem a ensacada nem a que nao esta',
   !numeros.includes('0488') && !numeros.includes('0487'), numeros.join(','));

console.log('-- a informacao sobrevive ao gate --');
ok('6. cada item ainda diz se esta ensacada (a lista mostra numa coluna)',
   saida.find(x => x.o.os === '0499').ensacada === true
   && saida.find(x => x.o.os === '0500').ensacada === false,
   JSON.stringify(saida.map(x => [x.o.os, x.ensacada])));
ok('7. e ainda traz as pecas, que a lista tambem mostra',
   saida.find(x => x.o.os === '0500').pecas === 3744,
   JSON.stringify(saida.map(x => [x.o.os, x.pecas])));

console.log('-- casos de borda --');
ok('8. sem cargas nenhuma, todas as com peca entram',
   api._expOsSemCarga(ordens, []).map(x => x.o.os).join(',') === '0500,0499,0498,0490,0489',
   api._expOsSemCarga(ordens, []).map(x => x.o.os).join(','));
ok('9. lista vazia nao quebra',
   api._expOsSemCarga(null, null).length === 0);

console.log(falhas === 0 ? '\nTODOS OS TESTES PASSARAM' : `\n${falhas} FALHA(S)`);
process.exit(falhas === 0 ? 0 : 1);
