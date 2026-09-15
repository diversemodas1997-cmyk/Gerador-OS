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
// As listas saem do app.js tambem: o teste nao pode ter a sua propria ideia de
// quais sao os quatro estados de uma OS.
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

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
  // A LINHA de SKU e o SKU sem a cor. Duble proprio porque desde 10/09/2026 o
  // filtro le a linha, e a busca le as duas.
  const linhasSkuDaOS = (o) => o.linhasSku || [];
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
    cores: ['Preto'], gradeNome: 'M-2G-GG | CM.LISA | 117cm',
    skus: ['CM.LISA-PRE'], linhasSku: ['CM.LISA'] },
  { os: '0502', codigo: '009', modeloNome: 'Camiseta Básica', colecaoNome: 'Verão 2026',
    cores: ['Grafite'], gradeNome: 'M-2G-GG | CM.LISA | 117cm',
    skus: ['CM.LISA-GRA'], linhasSku: ['CM.LISA'] },
  { os: '0492', codigo: '0023', modeloNome: 'Blusa Moletom Tricolor', colecaoNome: 'Inverno 2026',
    cores: ['Preto', 'Mostarda', 'Off-White'], gradeNome: '2M-2G-2GG | BM.TRI | 177,5cm',
    skus: ['BM.TRI-PRE'], linhasSku: ['BM.TRI'] },
  { os: '0400', codigo: '010', modeloNome: 'Camiseta Polo', colecaoNome: 'Verão 2026',
    cores: [], gradeNome: '', skus: [], linhasSku: [] },
  /* A OS DA COR SEM SIGLA. `skusDaOS` nao consegue compor o SKU completo e
     devolve vazio — era assim que ela sumia do filtro de SKU inteiro, sem nada
     na tela dizendo por que. A LINHA existe do mesmo jeito. */
  { os: '0505', codigo: '011', modeloNome: 'Camiseta Básica', colecaoNome: 'Verão 2026',
    cores: ['Verde Musgo'], gradeNome: 'M-2G-GG | CM.LISA | 117cm',
    skus: [], linhasSku: ['CM.LISA'] }
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
ok('12. a primeira opcao conta a lista inteira', /Todas as cores \(5\)/.test(htmlCor), htmlCor);
ok('13. a cor de duas OS vem antes das de uma (mais frequente em cima)',
   htmlCor.indexOf('Preto (2)') > 0 && htmlCor.indexOf('Preto (2)') < htmlCor.indexOf('Grafite (1)'), htmlCor);
ok('14. a tricolor entra nas TRES cores dela',
   /Mostarda \(1\)/.test(htmlCor) && /Off-White \(1\)/.test(htmlCor), htmlCor);
ok('15. cor vazia nao vira opcao', !/value="">.*\(0\)/.test(htmlCor) && !/—/.test(htmlCor), htmlCor);

t = ctxDe();
t.api._filtroListaOS('filtro-grade-os', lista, 'Todas as grades', o => [o.gradeNome]);
const htmlGrade = t.ctx.sel['filtro-grade-os'].innerHTML;
ok('16. a grade repetida conta todas as OS dela',
   /M-2G-GG \| CM\.LISA \| 117cm \(3\)/.test(htmlGrade), htmlGrade);
ok('17. OS sem grade nao inventa uma opcao vazia',
   (htmlGrade.match(/<option/g) || []).length === 3, htmlGrade);

t = ctxDe();
t.api._filtroListaOS('filtro-sku-os', lista, 'Todos os SKUs', o => o.linhasSku);
const htmlSku = t.ctx.sel['filtro-sku-os'].innerHTML;
ok('18. cada LINHA de SKU vira uma opcao, e a cor nao entra nela',
   /CM\.LISA \(3\)/.test(htmlSku) && /BM\.TRI \(1\)/.test(htmlSku)
   && !/CM\.LISA-/.test(htmlSku), htmlSku);

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

/* ----------------------------------------------------------------------
   O SKU NAO EXIGE COR (10/09/2026, Junior: "o cruzamento entre sku e cor deve
   estar em filtros diferentes. O filtro de sku nao deve exigir cor, apenas
   sku").

   O filtro usava o SKU COMPLETO — linha mais a sigla da cor. Escolher
   "CM.LISA" era impossivel: so existiam CM.LISA-PRE, CM.LISA-GRA... Ver todas
   as camiseta lisa exigia passar cor por cor, e a cor ja tem filtro ao lado.
   ---------------------------------------------------------------------- */
console.log('');
console.log('-- o SKU nao exige cor --');
{
  // A regra que separa a linha da cor, recortada do app de verdade.
  const linhaDe = new Function('STATE', `
    ${recorte('function _skuBaseDaOS', 'o SKU base da OS')}
    ${recorte('function linhasSkuDaOS', 'a linha de SKU da OS')}
    return linhasSkuDaOS;
  `)({ desenhos: [], modelos: [] });

  ok('37. a linha e o SKU sem a cor',
     linhaDe({ skuOverride: 'CM.LISA-PRE' }).join() === 'CM.LISA',
     linhaDe({ skuOverride: 'CM.LISA-PRE' }));
  ok('38. e quem ja e linha continua linha',
     linhaDe({ skuOverride: 'CM.LISA' }).join() === 'CM.LISA');
  ok('39. OS sem SKU nenhum nao inventa uma linha',
     linhaDe({}).length === 0, linhaDe({}));
  ok('40. minusculo e espaco sobrando nao viram outra linha',
     linhaDe({ skuOverride: '  cm.lisa-pre ' }).join() === 'CM.LISA');

  /* A COR SEM SIGLA. `skusDaOS` nao consegue compor e devolve vazio; era assim
     que a OS sumia do filtro de SKU. A linha sai do mesmo jeito. */
  const semSigla = new Function('STATE', `
    const _normNome = (x) => String(x || '').toLowerCase().trim();
    ${recorte('function _skuBaseDaOS', 'o SKU base da OS')}
    ${recorte('function linhasSkuDaOS', 'a linha de SKU da OS')}
    ${recorte('function skusDaOS', 'o SKU completo da OS')}
    return { linhasSkuDaOS, skusDaOS };
  `)({ desenhos: [], modelos: [], cores: [{ nome: 'Verde Musgo', siglaSku: '' }] });
  const osSemSigla = { skuOverride: 'CM.LISA', variantes: [{ cor1Nome: 'Verde Musgo' }] };
  ok('41. cor sem sigla: o SKU completo nao sai...',
     semSigla.skusDaOS(osSemSigla).length === 0, semSigla.skusDaOS(osSemSigla));
  ok('42. ...mas a linha sai, e a OS continua achavel pelo SKU',
     semSigla.linhasSkuDaOS(osSemSigla).join() === 'CM.LISA');

  // Na lista: escolher a linha traz as OS de TODAS as cores dela.
  const porLinha = (escolhido) => lista.filter(o => !escolhido || (o.linhasSku || []).includes(escolhido));
  ok('43. escolher CM.LISA traz as tres, de qualquer cor',
     porLinha('CM.LISA').map(o => o.os).join() === '0501,0502,0505',
     porLinha('CM.LISA').map(o => o.os));
  ok('44. e nao traz a de outra linha', !porLinha('CM.LISA').some(o => o.os === '0492'));

  // E as duas perguntas continuam cruzando quando alguem escolhe as duas.
  const cruzado = porLinha('CM.LISA').filter(o => (o.cores || []).includes('Preto'));
  ok('45. SKU e cor juntos ainda cruzam (é o filtro ao lado que faz isso)',
     cruzado.map(o => o.os).join() === '0501', cruzado.map(o => o.os));
}

/* ----------------------------------------------------------------------
   BUSCA POR DATA DE FINALIZACAO (10/09/2026, Junior: "insira no campo de
   cadastros de os, campo de busca por data de finalizacao da OS" e, logo
   depois, "para data unica, que represente a data de finalizacao da OS").

   Uma data so. A pergunta do chao e "o que foi finalizado neste dia?".

   A ARMADILHA E O FUSO. `finalizadaEm` e gravado em UTC; uma OS terminada as
   21:40 de 09/09 na fabrica esta gravada como "2026-09-10T00:40:00Z". Fatiar
   o texto do ISO a jogaria no dia 10 e ela sumiria da busca do dia 9 — o dia
   em que ela terminou para quem estava la, e o dia que a propria tela mostra.
   ---------------------------------------------------------------------- */
console.log('');
console.log('-- a busca por data de finalizacao --');
{
  const api = new Function(`
    ${constante('STATUS_OS')}
    ${constante('STATUS_FIM')}
    // Desde 15/09/2026 o status nasce do CHECKLIST: _statusOS le a etapa
    // marcada por ultimo antes de olhar o carimbo a mao. Sem estas, ela nao roda.
    ${recorte('function osEtapaMarcada', 'a etapa marcada no checklist')}
    ${src.match(/^const ETAPA_SC_RE = .+$/m)[0]}
    ${src.match(/^const _osRecebidaSC = .+$/m)[0]}
    ${recorte('function _statusDoChecklistOS', 'o status que o checklist diz')}
    ${recorte('function _ultimaMarcacaoChecklist', 'a ultima etapa marcada')}
    ${recorte('function _statusOS', 'a leitura do status')}
    ${recorte('function _dataFinalizacaoOS', 'a data de finalizacao')}
    ${recorte('function _diaFinalizacaoOS', 'o dia da finalizacao')}
    ${recorte('function _osFinalizadaNoDia', 'a OS no dia pedido')}
    return { _diaFinalizacaoOS, _osFinalizadaNoDia };
  `)();

  // A OS terminada as 21:40 do dia 9 (00:40Z do dia 10).
  const noite = { statusOS: 'estoque', finalizadaEm: new Date(2026, 8, 9, 21, 40).toISOString() };
  const dia5 = { statusOS: 'estoque', finalizadaEm: new Date(2026, 8, 5, 10, 0).toISOString() };
  const aberta = { statusOS: 'enfestando' };
  const nunca = {};

  ok('46. a OS terminada a noite fica no dia em que ela terminou',
     api._diaFinalizacaoOS(noite) === '2026-09-09', api._diaFinalizacaoOS(noite));
  ok('47. e a busca do dia 9 a encontra',
     api._osFinalizadaNoDia(noite, '2026-09-09') === true);
  ok('48. a do dia 10 nao a encontra (ela nao terminou no dia 10)',
     api._osFinalizadaNoDia(noite, '2026-09-10') === false);
  ok('49. sem data, todas passam — inclusive as que nem terminaram',
     api._osFinalizadaNoDia(aberta, '') && api._osFinalizadaNoDia(nunca, '')
     && api._osFinalizadaNoDia(dia5, ''));
  ok('50. com data, a OS que nao terminou fica de fora',
     api._osFinalizadaNoDia(aberta, '2026-09-05') === false
     && api._osFinalizadaNoDia(nunca, '2026-09-05') === false);
  ok('51. cada dia traz so quem terminou nele',
     api._osFinalizadaNoDia(dia5, '2026-09-05') === true
     && api._osFinalizadaNoDia(noite, '2026-09-05') === false);
  /* A OS PARADA QUE JA FOI ENSACADA TEM DIA DE TERMINO (15/09/2026). O fim da
     producao passou a ser o ENSAQUE, e a data deixou de se apagar quando a OS
     anda: ela e um fato (o dia em que o saco foi fechado), nao o estado atual.
     Antes isto era false porque so "Finalizado" tinha data, e sair dele
     apagava. */
  ok('52. OS parada que ja foi ensacada entra no dia em que foi ensacada',
     api._osFinalizadaNoDia({ statusOS: 'parado', finalizadaEm: dia5.finalizadaEm },
                            '2026-09-05') === true);

  // A leitura do campo da tela.
  const leitura = (valor) => new Function('campo', `
    const document = { getElementById: (id) => (id === 'filtro-fim' ? campo : null) };
    ${recorte('function _diaFiltroFinalizacaoListaOS', 'a leitura da data da tela')}
    return _diaFiltroFinalizacaoListaOS();
  `)(valor === null ? null : { value: valor });
  ok('53. le a data do campo da tela', leitura('2026-09-09') === '2026-09-09');
  ok('54. campo vazio e campo que nem existe valem "todas"',
     leitura('') === '' && leitura(null) === '');
}

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
   && /!skuEscolhido \|\| linhasSkuDaOS\(o\)\.includes\(skuEscolhido\)/.test(src), 'filtro incompleto');
ok('34. a conta fica ao lado da busca na barra de filtros',
   /id="busca-os"[\s\S]{0,600}?id="conta-os"/.test(html), 'conta fora da barra');

/* A LISTA DO NAVEGADOR NAO PODE COBRIR A COLUNA ACOES (10/09/2026, print do
   Junior). Campo de busca sem `autocomplete="off"` faz o Chrome abrir a lista
   dele — os numeros de OS ja digitados antes — logo abaixo do campo, e ela cai
   POR CIMA da primeira coluna da tabela: status, visualizar e o "..." ficam
   atras de uma janela que o programa nao desenhou e nao consegue mover.

   O teste cobre TODOS os campos de busca da tela, e nao so o da lista de OS:
   a regra e a mesma em qualquer um deles, e o proximo campo novo nasce certo
   ou o teste reclama. */
{
  const buscas = html.match(/<input[^>]*type="search"[^>]*>/g) || [];
  const semGuarda = buscas.filter(t => !/autocomplete="off"/.test(t));
  ok('35. todo campo de busca desliga o autocompletar do navegador',
     buscas.length >= 3 && semGuarda.length === 0, semGuarda.join(' | '));
  ok('36. a busca da lista de OS e uma delas (foi ela que abriu o caso)',
     /<input[^>]*id="busca-os"[^>]*autocomplete="off"/.test(html), 'busca-os sem a guarda');
}

/* O campo de data e o que o programa faz com ele: se ele nao estiver na barra,
   ou o Limpar esquecer dele, o filtro fica ligado sem ninguem ver — e a lista
   some sem explicacao. */
{
  ok('55. o campo esta na barra de filtros da lista',
     /class="lista-os-filtros"[\s\S]{0,2600}id="filtro-fim"/.test(html), 'campo de data fora da barra');
  ok('56. e e um so — o intervalo de duas datas nao voltou',
     !/filtro-fim-de|filtro-fim-ate/.test(html) && !/filtro-fim-de|filtro-fim-ate/.test(src),
     'sobrou campo do intervalo');
  ok('57. o Limpar apaga a data junto com o resto',
     /limparFiltrosListaOS[\s\S]{0,400}'filtro-fim'/.test(src), 'Limpar esqueceu a data');
  ok('58. a data entra no filtro da lista, junto com os outros',
     /_osFinalizadaNoDia\(o, diaFim\)/.test(src), 'data fora do filter');
  ok('59. e a lista vazia diz o dia que ninguem atendeu',
     /finalizadas em <b>\$\{esc\(formatDate\(diaFim\)\)\}<\/b>/.test(src), 'aviso sem o dia');
}

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
