/* ESTOQUE DE PECAS E DE FERRAMENTAS (25/09/2026, Junior: "Insira na barra
   lateral na pasta estoque, item Estoque de pecas e Estoque de ferramentas.
   Esses estoques devem receber o cadastro dos tipos de pecas e ferramentas em
   uso e em estoque" e "insira historico de entradas e saidas nas pecas e
   ferramentas").

   A tela e recortada do app.js e desenhada com um cadastro de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function pegaFuncao(nome, async) {
  const ini = src.indexOf('\n' + (async ? 'async ' : '') + 'function ' + nome + '(');
  if (ini < 0) throw new Error('nao achei ' + nome);
  return src.slice(ini, src.indexOf('\n}', ini) + 2);
}
const linha = re => (src.match(re) || [''])[0];
const bloco = (ini, fim) => src.slice(src.indexOf(ini), src.indexOf(fim, src.indexOf(ini)) + fim.length);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};

const paineis = {};
const campos = {};
const documento = { getElementById: id => campos[id] || (paineis[id] = paineis[id] || { innerHTML: '' }) };
const STATE = {
  pecasCad: [
    { id: 'a', nome: 'Agulha DBx1', desc: 'nº 11', unidade: 'desc', emUso: 12, emEstoque: 200 },
    { id: 'b', nome: 'Lançadeira', unidade: 'desc', emUso: 4, emEstoque: 3 },
    { id: 'c', nome: 'Agulha DBx1', desc: 'nº 11', unidade: 'sc', emUso: 8, emEstoque: 50 }
  ],
  ferramentasCad: [],
  pecasMov: [],
  ferramentasMov: []
};
const api = new Function('document', 'STATE', `
  const esc = s => String(s == null ? '' : s);
  const formatDate = d => d;
  ${linha(/const AVIAMENTO_UNIDADES = \[[\s\S]*?\];/)}
  ${linha(/const _aviUnidadeDe = [^\n]*/)}
  ${pegaFuncao('_normNome')}
  ${bloco('const ESTOQUE_ITENS = {', '\n};')}
  ${bloco('const ESTOQUE_ITENS_MOTIVOS = {', '\n};')}
  ${linha(/const _estItensMotivo = [^\n]*/)}
  const _estItensUnidade = { pecas: 'desc', ferramentas: 'desc' };
  const _estItensBusca = { pecas: '', ferramentas: '' };
  const _estItensPeriodo = { pecas: { de: '2026-09-01', ate: '2026-09-30' }, ferramentas: { de: '', ate: '' } };
  let _estItensCtx = null;
  const _aviHoje = () => '2026-09-25';
  let _n = 0; const uid = () => 'm' + (++_n);
  const saveState = async () => {}; const closeModal = () => {};
  const toasts = []; const toast = (t, k) => toasts.push(k + ':' + t);
  const exigirEstoqueTecidos = () => true;
  const confirm = () => true;
  ${pegaFuncao('_estItensHistorico')}
  ${pegaFuncao('_estItensRegistrar')}
  ${pegaFuncao('_salvarMovEstoqueItem', true)}
  ${pegaFuncao('excluirMovEstoqueItem', true)}
  ${pegaFuncao('renderEstoqueItens')}
  return { renderEstoqueItens, _estItensUnidade, _estItensBusca, _salvarMovEstoqueItem, excluirMovEstoqueItem, ESTOQUE_ITENS, toasts };
`)(documento, STATE);

(async () => {
console.log('-- a tela --');
api.renderEstoqueItens('pecas');
const desc = paineis['pecas-painel'].innerHTML;
const linhasDe = h => (h.match(/abrirEstoqueItem\('pecas'/g) || []).length;
ok('1. Descalvado mostra os dois tipos dela, e nao o de Sao Carlos', linhasDe(desc) === 2, linhasDe(desc));
ok('2. em uso e em estoque somam no rodape: 16 em uso, 203 em estoque, 219 no total',
   /em uso <b[^>]*>16<\/b>/.test(desc) && /em estoque <b[^>]*>203<\/b>/.test(desc) && />219</.test(desc));
api._estItensUnidade.pecas = 'sc';
api.renderEstoqueItens('pecas');
ok('3. Sao Carlos mostra so o dela (8 em uso, 50 em estoque)',
   /em uso <b[^>]*>8<\/b>/.test(paineis['pecas-painel'].innerHTML) && /em estoque <b[^>]*>50<\/b>/.test(paineis['pecas-painel'].innerHTML));
api._estItensUnidade.pecas = 'desc';
api._estItensBusca.pecas = 'lancadeira';
api.renderEstoqueItens('pecas');
ok('4. a busca ignora acento', linhasDe(paineis['pecas-painel'].innerHTML) === 1);
api._estItensBusca.pecas = '';
api.renderEstoqueItens('ferramentas');
ok('5. ferramentas tem cadastro proprio (vazio aqui)', /Nenhuma ferramenta cadastrada/.test(paineis['ferramentas-painel'].innerHTML));

console.log('-- entradas e saidas --');
const cfg = api.ESTOQUE_ITENS.pecas;
const lanc = async (mov, motivo, qtd) => {
  campos['mei-motivo'] = { value: motivo }; campos['mei-qtd'] = { value: String(qtd) };
  campos['mei-data'] = { value: '2026-09-25' }; campos['mei-obs'] = { value: '' };
  await api._salvarMovEstoqueItem({ tipo: 'pecas', id: 'b', mov }, cfg);
};
const lan = () => STATE.pecasCad.find(x => x.id === 'b');
await lanc('entrada', 'compra', 10);
ok('6. entrada de compra: estoque 3 + 10 = 13, uso fica 4', lan().emEstoque === 13 && lan().emUso === 4, lan());
await lanc('saida', 'uso', 5);
ok('7. saida posta em uso: estoque 8, uso 9', lan().emEstoque === 8 && lan().emUso === 9, lan());
await lanc('saida', 'baixa-uso', 2);
ok('8. baixa do uso: uso 7, estoque 8', lan().emUso === 7 && lan().emEstoque === 8, lan());
await lanc('saida', 'baixa-estoque', 50);
ok('9. nao tira do estoque mais do que ha', lan().emEstoque === 8 && /Só há 8 em estoque/.test(api.toasts.join('|')), api.toasts);
await lanc('entrada', 'voltou', 3);
ok('10. voltou do uso: uso 4, estoque 11', lan().emUso === 4 && lan().emEstoque === 11, lan());
ok('11. o historico guarda os 4 lancamentos, com o que cada um mexeu',
   STATE.pecasMov.length === 4 && STATE.pecasMov[1].dUso === 5 && STATE.pecasMov[1].dEstoque === -5, STATE.pecasMov);
api.renderEstoqueItens('pecas');
const tela = paineis['pecas-painel'].innerHTML;
ok('12. a tela mostra o historico com entradas 13 e saidas 7',
   /Histórico de entradas e saídas/.test(tela) && /entradas <b[^>]*>13<\/b>/.test(tela) && /saídas <b[^>]*>7<\/b>/.test(tela));
await api.excluirMovEstoqueItem('pecas', STATE.pecasMov[1].id);
ok('13. apagar a saida "posta em uso" deixaria o uso em -1: recusado',
   STATE.pecasMov.length === 4 && /negativa/.test(api.toasts.join('|')), { lan: lan(), n: STATE.pecasMov.length });
await api.excluirMovEstoqueItem('pecas', STATE.pecasMov[3].id);
ok('14. apagar o "voltou do uso" desfaz: uso 7, estoque 8', lan().emUso === 7 && lan().emEstoque === 8 && STATE.pecasMov.length === 3, lan());
await api.excluirMovEstoqueItem('pecas', STATE.pecasMov[0].id);
ok('15. apagar a entrada de 10 deixaria o estoque em -2: recusado', lan().emEstoque === 8 && STATE.pecasMov.length === 3, lan());

console.log('-- a costura com o resto --');
ok('16. os dois itens ficam na pasta Estoques, logo abaixo do Estoque de aviamentos',
   /Estoque de aviamentos<\/a>\s*<a class="nav-btn" href="#estoque-pecas"[^>]*>Estoque de peças<\/a>\s*<a class="nav-btn" href="#estoque-ferramentas"[^>]*>Estoque de ferramentas<\/a>/.test(html));
ok('17. as paginas e o modal existem',
   /data-page="estoque-pecas"/.test(html) && /data-page="estoque-ferramentas"/.test(html) && /id="modal-estoque-item"/.test(html));
ok('18. as rotas desenham as telas',
   /if \(page === 'estoque-pecas'\) renderEstoqueItens\('pecas'\);/.test(src) && /if \(page === 'estoque-ferramentas'\) renderEstoqueItens\('ferramentas'\);/.test(src));
ok('19. cadastro e historico sao sincronizados e carregados (as duas listas de chaves)',
   (src.match(/'pecasCad','ferramentasCad','pecasMov','ferramentasMov'/g) || []).length === 2);
ok('20. corrigir altera a linha no lugar, o mesmo tipo na mesma unidade nao duplica, e a correcao vira ajuste',
   /if \(ctx\.id\) \{\s*x = lista\.find\(i => i\.id === ctx\.id\)/.test(src) && /já está cadastrado nesta unidade/.test(src)
   && /_estItensRegistrar\(cfg, x, \{ tipo: 'ajuste', motivo: 'correcao'/.test(src));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
})();
