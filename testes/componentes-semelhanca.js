/* Rode com:  node testes/componentes-semelhanca.js

   A lista de componentes mostra as partes da mesma peça juntas.

   O cadastro guarda os componentes na ordem em que foram criados, e com o tempo
   isso separa o que é a MESMA coisa. Nos 17 componentes da fábrica (19/08/2026),
   "Frente" era a primeira linha e "Frente Parte 2" e "Frente Parte 3" as duas
   últimas — catorze linhas abaixo —, porque nasceram depois, com a tricolor.
   Quem marca os componentes de um desenho tricolor procurava as partes da
   frente em duas pontas da lista.

   O QUE ESTE TESTE PROTEGE:

   1. FAMÍLIA JUNTA, NA ORDEM DA PARTE. "Frente", "Frente Parte 2", "Frente
      Parte 3", nessa sequência, e a peça inteira antes das partes.

   2. A ORDEM DA CASA SOBREVIVE. As famílias não são alfabéticas nem inventadas:
      valem na ordem em que aparecem no cadastro, que é a ordem da peça sendo
      montada (Frente, Costas, Capuz, Forro do capuz, Mangas...). Ordenar por
      semelhança não pode virar "ordenar alfabeticamente", que jogaria Barra e
      Bolso canguru para o começo.

   3. COMPONENTE NOVO CONTINUA NASCENDO NO FIM, onde quem acabou de cadastrar
      espera achá-lo — a menos que seja parte de uma família que já existe.

   4. NÃO AGRUPA POR PARECENÇA DE PALAVRA. "Gola" e "Cobre gola" são duas peças
      diferentes; juntá-las trocaria uma lista fora de ordem por uma errada.

   5. NÃO PERDE NEM INVENTA LINHA, e ordenar de novo não muda mais nada.

   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(deOnde, ate, oQue) {
  const i = src.indexOf(deOnde);
  const j = i < 0 ? -1 : src.indexOf(ate, i + deOnde.length);
  if (i < 0 || j < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, j);
}
const corta = (nome) => recorte(nome, '\n}', nome) + '\n}';

const motor = [
  corta('function _normNome'),
  corta('function _componenteFamilia'),
  corta('function _componentesPorSemelhanca')
].join('\n');

const ordenar = (nomes) => new Function('NOMES', `
  ${motor}
  return _componentesPorSemelhanca(NOMES.map((n, i) => ({ id: 'c' + i, nome: n }))).map(c => c.nome);
`)(nomes);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const eq = (nome, got, esperado) => ok(nome, JSON.stringify(got) === JSON.stringify(esperado), got);

/* ---------- os 17 componentes REAIS, na ordem em que estão gravados ---------- */
const REAIS = ['Frente', 'Costas', 'Capuz', 'Forro do capuz', 'Mangas', 'Bolso canguru',
  'Punho', 'Barra', 'Gola', 'Viés', 'Recorte lateral', 'Cordão', 'Ilhós',
  'Etiqueta interna', 'Tag', 'Frente Parte 2', 'Frente Parte 3'];

const r = ordenar(REAIS);

eq('a lista da fábrica sai com a frente inteira junta', r,
  ['Frente', 'Frente Parte 2', 'Frente Parte 3', 'Costas', 'Capuz', 'Forro do capuz',
   'Mangas', 'Bolso canguru', 'Punho', 'Barra', 'Gola', 'Viés', 'Recorte lateral',
   'Cordão', 'Ilhós', 'Etiqueta interna', 'Tag']);

ok('não perdeu nem inventou linha', r.length === REAIS.length && REAIS.every(n => r.includes(n)), r.length);
eq('ordenar de novo não muda mais nada', ordenar(r), r);

/* ---------- a ordem da casa não vira ordem alfabética ---------- */
ok('Barra continua depois de Punho (alfabética a jogaria para o começo)',
  r.indexOf('Barra') > r.indexOf('Punho'), r.slice(0, 10));
ok('Costas continua antes de Capuz',
  r.indexOf('Costas') < r.indexOf('Capuz'), r.slice(0, 6));

/* ---------- componente novo nasce no fim ---------- */
eq('componente novo, de família nova, fica no fim',
  ordenar(REAIS.concat(['Bolso faca'])).slice(-1), ['Bolso faca']);
eq('mas parte de família que existe sobe para junto dela',
  ordenar(REAIS.concat(['Costas Parte 2'])).slice(0, 5),
  ['Frente', 'Frente Parte 2', 'Frente Parte 3', 'Costas', 'Costas Parte 2']);

/* ---------- não agrupa por parecença de palavra ---------- */
const comCobreGola = ordenar(['Gola', 'Frente', 'Cobre gola']);
ok('"Cobre gola" NÃO é puxada para junto de "Gola"',
  comCobreGola.indexOf('Cobre gola') === 2, comCobreGola);

/* ---------- duplicado vizinho, que é como se enxerga o duplicado ---------- */
const comDup = ordenar(['Frente', 'Costas', 'Capuz', 'frente']);
ok('nomes iguais em grafia diferente ficam em linhas vizinhas',
  Math.abs(comDup.indexOf('Frente') - comDup.indexOf('frente')) === 1, comDup);

/* ---------- a peça inteira vem antes das partes ---------- */
eq('a peça inteira vem antes das partes, mesmo cadastrada depois',
  ordenar(['Frente Parte 3', 'Frente Parte 2', 'Frente']),
  ['Frente', 'Frente Parte 2', 'Frente Parte 3']);

/* ---------- NÃO MEXE NA LISTA ORIGINAL ----------
   Se ordenasse no lugar, a ordem gravada em STATE.componentes mudaria sozinha e
   o próximo save gravaria a nova ordem no servidor — uma alteração de dados que
   ninguém pediu, disparada por abrir uma tela. */
const original = ['Frente', 'Costas', 'Frente Parte 2'];
const naMao = new Function('NOMES', `
  ${motor}
  const lista = NOMES.map((n, i) => ({ id: 'c' + i, nome: n }));
  const saida = _componentesPorSemelhanca(lista);
  return { saida: saida.map(c => c.nome), entrada: lista.map(c => c.nome) };
`)(original);
eq('a lista de entrada continua na ordem em que estava', naMao.entrada, original);
eq('  ... e a ordenada sai junta', naMao.saida, ['Frente', 'Frente Parte 2', 'Costas']);

/* ---------- lista vazia e sujeira não quebram ---------- */
eq('lista vazia', ordenar([]), []);
ok('componente sem nome não derruba a ordenação',
  new Function(`${motor}\nreturn _componentesPorSemelhanca([{id:'a'},{id:'b',nome:'Frente'}]).length;`)() === 2);

console.log(falhas ? `\n${falhas} falha(s)` : '\ntudo certo');
process.exit(falhas ? 1 : 0);
