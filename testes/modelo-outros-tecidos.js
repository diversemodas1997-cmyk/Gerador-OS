/* Rode com:  node testes/modelo-outros-tecidos.js

   O TIPO DO MODELO PARA OS OUTROS TECIDOS (28/10/2026, Jones: "analise por que
   o tipo Tactel não aparece como opção para criar um novo modelo"). O Tipo do
   modelo casa com a Categoria do tecido; o Tactel é 'outro' ("Sem limite de
   camadas"), e a lista do modelo só tinha Camiseta e Moletom. O que este teste
   protege:
     · com tecido na categoria 'outro', o modelo oferece "Outros tecidos — <nomes>";
     · sem tecido nela, a opção não aparece (não há o que pedir);
     · o modelo que já é 'outro' continua com a opção, escolhida;
     · e a OS de um modelo 'outro' mostra o Tactel na lista de tecidos.
   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  return src.slice(i, src.indexOf('\n}', i) + 2);
}

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const document = { getElementById: (id) => (id === 'f-modelo' ? { value: ctx.modeloId || '' } : null) };
  ${recorte('function _opcaoModeloOutrosTecidos', 'a opcao do modelo')}
  ${recorte('function modeloCategoriaAtual', 'a categoria do modelo na OS')}
  ${recorte('function tecOptions', 'os tecidos da OS')}
  return { _opcaoModeloOutrosTecidos, tecOptions };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const TECIDOS = [
  { id: 't1', nome: 'Malha Algodão', categoria: 'malha' },
  { id: 't2', nome: 'Tactel', categoria: 'outro' },
  { id: 't3', nome: 'Moletom', categoria: 'moletom' }
];

let a = monta({ STATE: { tecidos: TECIDOS, modelos: [] } });
let op = a._opcaoModeloOutrosTecidos('');
ok('1. com o Tactel cadastrado, o modelo oferece o tipo', /value="outro"/.test(op) && /Outros tecidos — Tactel/.test(op), op);
ok('2. e ela nao vem escolhida num modelo novo', !/selected/.test(op), op);
ok('3. no modelo que ja e "outro", vem escolhida', /selected/.test(a._opcaoModeloOutrosTecidos('outro')));

a = monta({ STATE: { tecidos: TECIDOS.filter(t => t.categoria !== 'outro'), modelos: [] } });
ok('4. sem tecido nessa categoria, a opcao nao aparece', a._opcaoModeloOutrosTecidos('') === '');
ok('5. mas o modelo que ja e "outro" nao perde o tipo', /value="outro" selected/.test(a._opcaoModeloOutrosTecidos('outro')));

a = monta({ STATE: { tecidos: TECIDOS, modelos: [{ id: 'm1', categoria: 'outro' }, { id: 'm2', categoria: 'malha' }] }, modeloId: 'm1' });
const lista = a.tecOptions('');
ok('6. na OS, o modelo "outro" mostra o Tactel', /Tactel/.test(lista) && !/Malha Algodão/.test(lista), lista);
a = monta({ STATE: { tecidos: TECIDOS, modelos: [{ id: 'm2', categoria: 'malha' }] }, modeloId: 'm2' });
ok('7. e o modelo de camiseta continua sem ele', !/Tactel/.test(a.tecOptions('')), a.tecOptions(''));

ok('8. a ficha do modelo usa a opcao', /\$\{_opcaoModeloOutrosTecidos\(item\.categoria\)\}/.test(src));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
