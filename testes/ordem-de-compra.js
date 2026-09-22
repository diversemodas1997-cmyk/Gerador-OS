/* Rode com:  node testes/ordem-de-compra.js

   A ORDEM DE COMPRA - OC (22/09/2026, Junior: "no campo compra, insira uma
   ordem de compra - OC. Essa ordem de compra deve ser preenchida com os tipos
   de tecidos alocados, atraves de um botao na janela itens de compra").

   O que este teste guarda:

     - a OC leva o "A COMPRAR", e nao o bruto: mandar o bruto ao fornecedor
       seria comprar de novo o pano que ja esta na prateleira. O bruto e o
       disponivel viajam junto, para a OC saber explicar o numero que pede;
     - tecido cuja conta ja esta coberta pelo estoque NAO entra;
     - os numeros sao COPIADOS: mexer na lista de compra depois nao muda uma OC
       ja emitida. E a diferenca entre um rascunho e um documento;
     - o NUMERO sai do maior que ja existiu, mais um - apagar a OC-0003 nao pode
       fazer a proxima nascer 0003 de novo;
     - lista vazia e "nada a comprar" nao geram OC muda: avisam;
     - sem permissao, nada nasce;
     - a quantidade e editavel (o fornecedor vende bobina inteira), e o que se
       digita ali nunca vira numero negativo.

   Recorta as funcoes do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
function pegaAsync(nome) {
  const ini = src.indexOf('\nasync function ' + nome + '(');
  if (ini < 0) { console.error('nao achei a funcao async ' + nome); process.exit(1); }
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
/* Recorta uma constante ATE O FIM DA DECLARACAO, e nao ate o primeiro ';' com
   quebra de linha colada: `const _ocAbertas = new Set();   // comentario`
   termina com espaco e comentario depois do ponto-e-virgula, e o corte ingenuo
   engolia o resto do arquivo ate o proximo ponto-e-virgula em fim de linha. */
function pegaConst(nome) {
  const i = src.search(new RegExp('^const ' + nome + ' = ', 'm'));
  if (i < 0) { console.error('nao achei a constante ' + nome); process.exit(1); }
  const linhas = src.slice(i).split('\n');
  const out = [];
  for (const l of linhas) {
    out.push(l);
    if (l.replace(/\/\/.*$/, '').trimEnd().endsWith(';')) return out.join('\n');
  }
  console.error('nao achei o fim da constante ' + nome); process.exit(1);
}

/* A necessidade bruta entra DUBLADA — ela tem teste proprio
   (compra-necessidade-bruta.js) e puxa o cadastro inteiro. O que se prova aqui
   e o que a OC faz com a resposta dela. */
function monta(M) {
  return new Function('M', `
    var STATE = M.STATE;
    var window = {};
    var uid = () => 'oc' + (M.n++);
    function compraNecessidadeBruta() { return M.bruta; }
    function exigirEdicaoCompra(acao) { M.pedidos.push(acao); return !!M.pode; }
    function exigirEdicao(acao) { M.pedidos.push(acao); return !!M.pode; }
    async function saveState(k) { M.salvou.push(k); }
    function desfazerNomearAcao() {}
    function renderCompra() {}
    function toast(msg, tipo) { M.toasts.push([tipo, msg]); }
    function _cpQuemSou() { return 'quem@fabrica'; }
    function confirm() { return M.confirma !== false; }
    ${pegaConst('OC_STATUS')}
    ${pegaConst('_ocStatusDef')}
    ${pegaFuncao('_ocLista')}
    ${pegaFuncao('_ocPorId')}
    ${pegaFuncao('_ocNumeroNovo')}
    ${pegaFuncao('_ocTotais')}
    ${pegaAsync('gerarOCdaCompra')}
    ${pegaConst('_ocAbertas')}
    ${pegaAsync('ocItemCampo')}
    ${pegaAsync('ocItemRemover')}
    ${pegaAsync('ocRemover')}
    ${pegaAsync('ocCampo')}
    return { gerarOCdaCompra, ocItemCampo, ocItemRemover, ocRemover, ocCampo,
             _ocNumeroNovo, _ocTotais, ocs: () => STATE.compraOCs };
  `)(M);
}

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

/* ---------------------- o mundo do teste ----------------------
   A necessidade bruta de uma lista com tres tecidos: dois a comprar e um que o
   estoque ja cobre. */
const BRUTA = () => ([
  { tecidoNome: 'Malha Algodão', corNome: 'Azul Malha Algodão',
    kg: 217.76, disponivel: -147.177, kgComprar: 364.937, bobinasComprar: 11, semPrevisao: [] },
  { tecidoNome: 'Ribana Malha Algodão', corNome: 'Preto Ribana Malha Algodão',
    kg: 2.83, disponivel: 0, kgComprar: 2.83, bobinasComprar: null, semPrevisao: ['Barra/Punhos'] },
  { tecidoNome: 'Moletom', corNome: 'Preto Moletom',
    kg: 76.952, disponivel: 2138.434, kgComprar: 0, bobinasComprar: 0, semPrevisao: [] }
]);
const mundo = (pode) => ({
  pode: pode !== false, n: 1, pedidos: [], salvou: [], toasts: [],
  bruta: BRUTA(),
  STATE: {
    compraPlano: [{ id: 'i1', gradeId: 'g1', camadas: 30, repeticoes: 1 }],
    compraOCs: [],
    meta: {},
    fornecedores: [{ id: 'f1', nome: 'PLUMA' }]
  }
});

(async () => {
  /* ---------- 1. o botao gera a OC com o que falta comprar ---------- */
  let M = mundo(); let api = monta(M);
  await api.gerarOCdaCompra();
  let ocs = api.ocs();
  ok('1. nasceu uma OC', ocs.length === 1, ocs.length);
  const oc = ocs[0];
  ok('2. com numero OC-0001, data de hoje e situacao aberta',
     oc.numero === 'OC-0001' && /^\d{4}-\d{2}-\d{2}$/.test(oc.data) && oc.status === 'aberta', oc);
  ok('3. leva SO os tecidos que faltam comprar (o Moletom, coberto pelo estoque, fica de fora)',
     oc.itens.length === 2 && !oc.itens.some(i => i.tecidoNome === 'Moletom'),
     oc.itens.map(i => i.tecidoNome));
  ok('4. a quantidade e a A COMPRAR, nao o bruto (364,937 kg e nao 217,760)',
     oc.itens[0].kg === 364.937 && oc.itens[0].bobinas === 11, oc.itens[0]);
  ok('5. o bruto e o disponivel viajam junto, para a OC explicar o numero',
     oc.itens[0].kgBruto === 217.76 && oc.itens[0].disponivel === -147.177, oc.itens[0]);
  ok('6. tecido sem previsao de bobina entra com o quilo e a marca do que falta',
     oc.itens[1].bobinas === 0 && oc.itens[1].kg === 2.83 && oc.itens[1].semPrevisao === 'Barra/Punhos',
     oc.itens[1]);
  ok('7. gravou a OC e o contador do numero', M.salvou.join(',') === 'compraOCs,meta', M.salvou);
  const t = api._ocTotais(oc);
  ok('8. os totais somam as linhas', t.linhas === 2 && t.bobinas === 11
     && Math.abs(t.kg - 367.767) < 0.001, t);

  /* ---------- 2. a OC e uma COPIA, nao um ponteiro ---------- */
  M.bruta = [{ tecidoNome: 'Outro', corNome: 'X', kg: 1, disponivel: 0, kgComprar: 9, bobinasComprar: 1, semPrevisao: [] }];
  M.STATE.compraPlano = [];
  ok('9. mexer na lista depois nao muda a OC ja emitida',
     api.ocs()[0].itens.length === 2 && api.ocs()[0].itens[0].kg === 364.937,
     api.ocs()[0].itens);

  /* ---------- 3. o numero nao se repete ---------- */
  M = mundo(); api = monta(M);
  await api.gerarOCdaCompra();
  await api.gerarOCdaCompra();
  ok('10. a segunda OC e a OC-0002, e a mais nova fica em cima',
     api.ocs().length === 2 && api.ocs()[0].numero === 'OC-0002' && api.ocs()[1].numero === 'OC-0001',
     api.ocs().map(o => o.numero));
  await api.ocRemover(api.ocs()[0].id);
  ok('11. apagada a OC-0002, a proxima nasce OC-0003 (numero nao se reusa)',
     api._ocNumeroNovo() === 'OC-0003', api._ocNumeroNovo());

  /* ---------- 4. o que nao gera OC ---------- */
  M = mundo(); M.STATE.compraPlano = []; api = monta(M);
  await api.gerarOCdaCompra();
  ok('12. lista vazia nao gera OC — avisa', api.ocs().length === 0
     && M.toasts.some(t => t[0] === 'err' && /vazia/.test(t[1])), M.toasts);

  M = mundo(); M.bruta = BRUTA().map(b => ({ ...b, kgComprar: 0, bobinasComprar: 0 })); api = monta(M);
  await api.gerarOCdaCompra();
  ok('13. estoque cobrindo tudo nao gera OC — avisa', api.ocs().length === 0
     && M.toasts.some(t => t[0] === 'err' && /Nada a comprar/.test(t[1])), M.toasts);

  /* ---------- 5. a permissao ---------- */
  M = mundo(false); api = monta(M);
  await api.gerarOCdaCompra();
  ok('14. sem permissao nada nasce e nada e gravado',
     api.ocs().length === 0 && M.salvou.length === 0, M.salvou);
  ok('15. e a recusa passa pela permissao da compra',
     M.pedidos.join('') === 'gerar a ordem de compra', M.pedidos);

  /* ---------- 6. a quantidade editavel ---------- */
  M = mundo(); api = monta(M);
  await api.gerarOCdaCompra();
  const id = api.ocs()[0].id;
  await api.ocItemCampo(id, 0, 'bobinas', '12');
  ok('16. da para ajustar a bobina do pedido (11 -> 12)', api.ocs()[0].itens[0].bobinas === 12,
     api.ocs()[0].itens[0]);
  await api.ocItemCampo(id, 0, 'bobinas', '12,7');
  ok('17. bobina nao sai quebrada', api.ocs()[0].itens[0].bobinas === 13, api.ocs()[0].itens[0]);
  await api.ocItemCampo(id, 0, 'kg', '-5');
  ok('18. quantidade negativa vira zero', api.ocs()[0].itens[0].kg === 0, api.ocs()[0].itens[0]);
  await api.ocItemCampo(id, 0, 'kg', '');
  ok('19. campo apagado vira zero, e nao NaN', api.ocs()[0].itens[0].kg === 0, api.ocs()[0].itens[0]);
  await api.ocCampo(id, 'fornecedorId', 'f1');
  await api.ocCampo(id, 'status', 'enviada');
  ok('20. fornecedor e situacao se gravam no cabecalho',
     api.ocs()[0].fornecedorId === 'f1' && api.ocs()[0].status === 'enviada', api.ocs()[0]);
  await api.ocCampo(id, 'itens', 'coisa nenhuma');
  ok('21. campo que nao existe no cabecalho nao entra (a OC nao vira gaveta)',
     Array.isArray(api.ocs()[0].itens), api.ocs()[0].itens);
  await api.ocItemRemover(id, 1);
  ok('22. da para tirar uma linha do pedido', api.ocs()[0].itens.length === 1, api.ocs()[0].itens);

  console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
  process.exit(falhas ? 1 : 0);
})();
