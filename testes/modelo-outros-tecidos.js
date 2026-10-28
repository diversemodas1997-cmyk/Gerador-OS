/* Rode com:  node testes/modelo-outros-tecidos.js

   O TIPO DO MODELO PARA OS OUTROS TECIDOS (28/10/2026, Jones: "analise por que
   o tipo Tactel não aparece como opção para criar um novo modelo"). O Tipo do
   modelo casa com a Categoria do tecido; o Tactel é 'outro' ("Sem limite de
   camadas"), e a lista do modelo só tinha Camiseta e Moletom. Depois (mesmo
   dia, Jones: "Cotton, Tactel devem ser opções separadas"), cada tecido dessa
   categoria é uma opção própria. O que este teste protege:
     · cada tecido 'outro' é uma opção ("Cotton", "Tactel"), não uma só;
     · sem tecido nela, nenhuma opção aparece (não há o que pedir);
     · o modelo do Tactel vem com o Tactel escolhido;
     · o modelo 'outro' antigo, sem tecido, fica com "Outros tecidos";
     · gravar 'outro:<id>' dá categoria 'outro' + tecidoOutroId;
     · a OS do modelo do Tactel mostra só o Tactel, sem o Cotton;
     · a lista de modelos mostra o nome do tecido.
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
  ${recorte('function _tipoModeloDoValor', 'o valor do tipo')}
  ${recorte('function _rotuloTipoModelo', 'o rotulo da lista')}
  ${recorte('function modeloCategoriaAtual', 'a categoria do modelo na OS')}
  ${recorte('function tecOptions', 'os tecidos da OS')}
  return { _opcaoModeloOutrosTecidos, _tipoModeloDoValor, _rotuloTipoModelo, tecOptions };
`)(ctx);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const TECIDOS = [
  { id: 't1', nome: 'Malha Algodão', categoria: 'malha' },
  { id: 't2', nome: 'Tactel', categoria: 'outro' },
  { id: 't3', nome: 'Moletom', categoria: 'moletom' },
  { id: 't4', nome: 'Cotton', categoria: 'outro' }
];

let a = monta({ STATE: { tecidos: TECIDOS, modelos: [] } });
let op = a._opcaoModeloOutrosTecidos({});
ok('1. Cotton e Tactel sao opcoes separadas', /<option value="outro:t4" >Cotton<\/option><option value="outro:t2" >Tactel<\/option>/.test(op), op);
ok('2. e nenhuma vem escolhida num modelo novo', !/selected/.test(op) && !/value="outro"/.test(op), op);
op = a._opcaoModeloOutrosTecidos({ categoria: 'outro', tecidoOutroId: 't2' });
ok('3. o modelo do Tactel vem com o Tactel escolhido', /value="outro:t2" selected>Tactel/.test(op) && !/value="outro:t4" selected/.test(op), op);
op = a._opcaoModeloOutrosTecidos({ categoria: 'outro' });
ok('4. o modelo "outro" antigo fica com "Outros tecidos"', /^<option value="outro" selected>Outros tecidos<\/option>/.test(op), op);

a = monta({ STATE: { tecidos: TECIDOS.filter(t => t.categoria !== 'outro'), modelos: [] } });
ok('5. sem tecido nessa categoria, a opcao nao aparece', a._opcaoModeloOutrosTecidos({}) === '');
ok('6. mas o modelo que ja e "outro" nao perde o tipo', /value="outro" selected>Outros tecidos</.test(a._opcaoModeloOutrosTecidos({ categoria: 'outro' })));

ok('7. gravar "outro:t2" da categoria outro + o tecido', JSON.stringify(a._tipoModeloDoValor('outro:t2')) === '{"categoria":"outro","tecidoOutroId":"t2"}');
ok('8. gravar "malha" nao leva tecido', JSON.stringify(a._tipoModeloDoValor('malha')) === '{"categoria":"malha","tecidoOutroId":""}');

const MODELOS = [{ id: 'm1', categoria: 'outro', tecidoOutroId: 't2' }, { id: 'm2', categoria: 'malha' }, { id: 'm3', categoria: 'outro' }];
a = monta({ STATE: { tecidos: TECIDOS, modelos: MODELOS }, modeloId: 'm1' });
let lista = a.tecOptions('');
ok('9. na OS, o modelo do Tactel mostra so o Tactel', /Tactel/.test(lista) && !/Cotton/.test(lista) && !/Malha Algodão/.test(lista), lista);
a = monta({ STATE: { tecidos: TECIDOS, modelos: MODELOS }, modeloId: 'm3' });
lista = a.tecOptions('');
ok('10. o modelo "outro" antigo mostra os dois', /Tactel/.test(lista) && /Cotton/.test(lista), lista);
a = monta({ STATE: { tecidos: TECIDOS, modelos: MODELOS }, modeloId: 'm2' });
ok('11. e o modelo de camiseta continua sem eles', !/Tactel|Cotton/.test(a.tecOptions('')), a.tecOptions(''));

const LBL = { malha: 'Camiseta', moletom: 'Moletom', outro: 'Outro' };
ok('12. a lista de modelos mostra o tecido', a._rotuloTipoModelo(MODELOS[0], LBL) === 'Tactel' && a._rotuloTipoModelo(MODELOS[2], LBL) === 'Outro');

ok('13. a ficha do modelo usa a opcao', /\$\{_opcaoModeloOutrosTecidos\(item\)\}/.test(src));
ok('14. e grava pelo _tipoModeloDoValor', /const tipoModelo = _tipoModeloDoValor\(v\('m-categoria'\)\)/.test(src));

console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
process.exit(falhas ? 1 : 0);
