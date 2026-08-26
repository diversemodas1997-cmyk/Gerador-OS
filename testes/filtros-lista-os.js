/* Rode com:  node testes/filtros-lista-os.js

   A BUSCA E OS FILTROS da lista de OS Salvas.

   A lista tinha uma busca que só olhava o número da OS. Pedido do Junior em
   26/08/2026: um campo de busca com filtro por cor, grade e SKU — que somados
   ao status já existente respondem "das paradas, quais são pretas na grade P ao
   G3?" em quatro cliques.

   O que este teste protege:

     · a busca casa TODOS os termos digitados, cada um em qualquer campo — é
       assim que se procura sem saber em qual coluna a palavra está;
     · cada seletor é montado dos PRÓPRIOS dados, com a contagem de OS por valor
       (é o número que faz alguém perceber o que existe), as mais frequentes em
       cima;
     · uma OS tricolor aparece nas três cores, mas conta uma vez em cada — senão
       a soma das opções passaria do tamanho da lista;
     · a escolha que sumiu da lista volta sozinha para "todas", em vez de deixar
       a tela vazia sem explicação.

   Recorta as funções do app.js de verdade. */
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

// Os <select> de mentira: value + innerHTML e nada mais, que e tudo o que
// _filtroListaOS toca.
const monta = (ctx) => new Function('ctx', `
  const esc = (s) => String(s == null ? '' : s).replace(/"/g, '&quot;');
  const document = { getElementById: (id) => ctx.sel[id] || null };
  // As tres leituras da OS vem do app; aqui entram como dublês simples, porque
  // o que se testa e o FILTRO, nao a resolucao de cor/grade/SKU (essas ja tem
  // teste proprio).
  const coresDaPecaOS = (o) => o.cores || [];
  const _gradeNomeDaOS = (o) => o.gradeNome || '\\u2014';
  const skusDaOS = (o) => o.skus || [];
  ${recorte('function _filtroListaOS', 'o seletor de filtro')}
  ${recorte('function _textoBuscaOS', 'o texto que a busca varre')}
  return { _filtroListaOS, _textoBuscaOS };
`)(ctx);

const sel = () => ({ value: '', innerHTML: '' });
const ctxDe = (escolhas = {}) => {
  const ctx = { sel: { 'filtro-cor-os': sel(), 'filtro-grade-os': sel(), 'filtro-sku-os': sel() } };
  Object.entries(escolhas).forEach(([k, v]) => { ctx.sel[k].value = v; });
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

const lista = [
  { os: '0501', codigo: '008', modeloNome: 'Camiseta Básica', colecaoNome: 'Verão 2026',
    cores: ['Preto'], gradeNome: 'M-2G-GG | CM.LISA | 117cm', skus: ['CM.LISA-PRE'] },
  { os: '0502', codigo: '009', modeloNome: 'Camiseta Básica', colecaoNome: 'Verão 2026',
    cores: ['Grafite'], gradeNome: 'M-2G-GG | CM.LISA | 117cm', skus: ['CM.LISA-GRA'] },
  { os: '0492', codigo: '0023', modeloNome: 'Blusa Moletom Tricolor', colecaoNome: 'Inverno 2026',
    cores: ['Preto', 'Mostarda', 'Off-White'], gradeNome: '2M-2G-2GG | BM.TRI | 177,5cm',
    skus: ['BM.TRI-PRE'] },
  { os: '0400', codigo: '010', modeloNome: 'Camiseta Polo', colecaoNome: 'Verão 2026',
    cores: [], gradeNome: '', skus: [] }
];

console.log('-- a busca por texto --');
const A = ctxDe().api;
const casa = (o, termo) => {
  const alvo = A._textoBuscaOS(o);
  return termo.trim().toLowerCase().split(/\s+/).filter(Boolean).every(t => alvo.includes(t));
};
ok('1. acha pelo numero, como sempre achou', casa(lista[0], '0501') && !casa(lista[1], '0501'));
ok('2. acha pelo codigo do desenho', casa(lista[2], '0023'));
ok('3. acha pelo modelo', casa(lista[2], 'moletom'));
ok('4. acha pela colecao', casa(lista[2], 'inverno'));
ok('5. acha pela cor', casa(lista[2], 'mostarda') && !casa(lista[0], 'mostarda'));
ok('6. acha pela grade', casa(lista[0], '117cm'));
ok('7. acha pelo SKU, que nem coluna e', casa(lista[1], 'cm.lisa-gra'));
ok('8. dois termos valem JUNTOS (cada um em qualquer campo)',
   casa(lista[2], 'preto tricolor') && !casa(lista[0], 'preto tricolor'));
ok('9. termo que nao existe em lugar nenhum nao acha nada',
   !lista.some(o => casa(o, 'veludo')));
ok('10. OS sem cor, grade nem SKU nao quebra a busca', A._textoBuscaOS(lista[3]).includes('polo'));

console.log('');
console.log('-- as opcoes de cada seletor, com a contagem --');
let t = ctxDe();
const escolhidoCor = t.api._filtroListaOS('filtro-cor-os', lista, 'Todas as cores', o => o.cores);
const htmlCor = t.ctx.sel['filtro-cor-os'].innerHTML;
ok('11. sem escolha, o filtro nao corta nada', escolhidoCor === '');
ok('12. a primeira opcao conta a lista inteira', /Todas as cores \(4\)/.test(htmlCor), htmlCor);
ok('13. a cor de duas OS vem antes das de uma (mais frequente em cima)',
   htmlCor.indexOf('Preto (2)') > 0 && htmlCor.indexOf('Preto (2)') < htmlCor.indexOf('Grafite (1)'), htmlCor);
ok('14. a tricolor entra nas TRES cores dela',
   /Mostarda \(1\)/.test(htmlCor) && /Off-White \(1\)/.test(htmlCor), htmlCor);
ok('15. cor vazia nao vira opcao', !/value="">.*\(0\)/.test(htmlCor) && !/—/.test(htmlCor), htmlCor);

t = ctxDe();
t.api._filtroListaOS('filtro-grade-os', lista, 'Todas as grades', o => [o.gradeNome]);
const htmlGrade = t.ctx.sel['filtro-grade-os'].innerHTML;
ok('16. a grade de duas OS conta duas',
   /M-2G-GG \| CM\.LISA \| 117cm \(2\)/.test(htmlGrade), htmlGrade);
ok('17. OS sem grade nao inventa uma opcao vazia',
   (htmlGrade.match(/<option/g) || []).length === 3, htmlGrade);

t = ctxDe();
t.api._filtroListaOS('filtro-sku-os', lista, 'Todos os SKUs', o => o.skus);
const htmlSku = t.ctx.sel['filtro-sku-os'].innerHTML;
ok('18. cada SKU vira uma opcao', /CM\.LISA-PRE \(1\)/.test(htmlSku) && /BM\.TRI-PRE \(1\)/.test(htmlSku), htmlSku);

console.log('');
console.log('-- a escolha feita --');
t = ctxDe({ 'filtro-cor-os': 'Preto' });
ok('19. a escolha volta como chave e continua marcada',
   t.api._filtroListaOS('filtro-cor-os', lista, 'Todas as cores', o => o.cores) === 'Preto'
   && /value="Preto" selected/.test(t.ctx.sel['filtro-cor-os'].innerHTML),
   t.ctx.sel['filtro-cor-os'].innerHTML);
t = ctxDe({ 'filtro-cor-os': 'Vermelho' });
ok('20. escolha que sumiu da lista volta sozinha para "todas"',
   t.api._filtroListaOS('filtro-cor-os', lista, 'Todas as cores', o => o.cores) === ''
   && t.ctx.sel['filtro-cor-os'].value === '', t.ctx.sel['filtro-cor-os'].value);

console.log('');
console.log('-- a conta ao lado da busca --');
const contaCtx = { el: { classList: { toggle: (c, v) => { contaCtx.classe = v; } }, innerHTML: '', title: '' } };
const conta = new Function('ctx', `
  const document = { getElementById: () => ctx.el };
  ${recorte('function _contaListaOS', 'a conta da lista')}
  return _contaListaOS;
`)(contaCtx);
conta(228, 228);
ok('25. sem filtro, a conta e o tamanho da lista',
   /<b>228<\/b> OS/.test(contaCtx.el.innerHTML) && contaCtx.classe === false, contaCtx.el.innerHTML);
conta(32, 228);
ok('26. com filtro, diz quantas de quantas',
   /<b>32<\/b> de 228 OS/.test(contaCtx.el.innerHTML) && contaCtx.classe === true, contaCtx.el.innerHTML);
ok('27. e a dica explica o recorte', /32/.test(contaCtx.el.title) && /228/.test(contaCtx.el.title), contaCtx.el.title);
conta(0, 228);
ok('28. zero tambem conta (e o "de 228" que diz que nada sumiu)',
   /<b>0<\/b> de 228 OS/.test(contaCtx.el.innerHTML), contaCtx.el.innerHTML);
conta(1200, 1200);
ok('29. milhar sai com o ponto do portugues', /1\.200/.test(contaCtx.el.innerHTML), contaCtx.el.innerHTML);

console.log('');
console.log('-- na tela --');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
ok('30. os tres seletores novos estao na barra de filtros da lista',
   /id="filtro-cor-os"/.test(html) && /id="filtro-grade-os"/.test(html) && /id="filtro-sku-os"/.test(html),
   'faltou seletor');
ok('31. e o botao que devolve a lista inteira',
   /limparFiltrosListaOS\(\)/.test(html) && /function limparFiltrosListaOS/.test(src), 'sem o Limpar');
ok('32. a busca diz o que procura', /Buscar: número, código, modelo/.test(html), 'placeholder antigo');
ok('33. os quatro filtros valem junto com a busca (um filter so)',
   /!statusEscolhido \|\| _statusOS\(o\) === statusEscolhido/.test(src)
   && /!corEscolhida \|\| coresDaPecaOS\(o\)\.includes\(corEscolhida\)/.test(src)
   && /!gradeEscolhida \|\| _gradeNomeDaOS\(o\) === gradeEscolhida/.test(src)
   && /!skuEscolhido \|\| skusDaOS\(o\)\.includes\(skuEscolhido\)/.test(src), 'filtro incompleto');
ok('34. a conta fica ao lado da busca na barra de filtros',
   /id="busca-os"[\s\S]{0,600}?id="conta-os"/.test(html), 'conta fora da barra');

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
