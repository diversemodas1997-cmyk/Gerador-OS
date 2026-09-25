/* ESTOQUE DE PECAS E DE FERRAMENTAS (25/09/2026, Junior: "Insira na barra
   lateral na pasta estoque, item Estoque de pecas e Estoque de ferramentas.
   Esses estoques devem receber o cadastro dos tipos de pecas e ferramentas em
   uso e em estoque").

   A tela e recortada do app.js e desenhada com um cadastro de mentira. */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function pegaFuncao(nome) {
  const ini = src.indexOf('\nfunction ' + nome + '(');
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
const documento = { getElementById: id => (paineis[id] = paineis[id] || { innerHTML: '' }) };
const STATE = {
  pecasCad: [
    { id: 'a', nome: 'Agulha DBx1', desc: 'nº 11', unidade: 'desc', emUso: 12, emEstoque: 200 },
    { id: 'b', nome: 'Lançadeira', unidade: 'desc', emUso: 4, emEstoque: 3 },
    { id: 'c', nome: 'Agulha DBx1', desc: 'nº 11', unidade: 'sc', emUso: 8, emEstoque: 50 }
  ],
  ferramentasCad: []
};
const api = new Function('document', 'STATE', `
  const esc = s => String(s == null ? '' : s);
  const formatDate = d => d;
  ${linha(/const AVIAMENTO_UNIDADES = \[[\s\S]*?\];/)}
  ${linha(/const _aviUnidadeDe = [^\n]*/)}
  ${pegaFuncao('_normNome')}
  ${bloco('const ESTOQUE_ITENS = {', '\n};')}
  const _estItensUnidade = { pecas: 'desc', ferramentas: 'desc' };
  const _estItensBusca = { pecas: '', ferramentas: '' };
  ${pegaFuncao('renderEstoqueItens')}
  return { renderEstoqueItens, _estItensUnidade, _estItensBusca };
`)(documento, STATE);

console.log('-- a tela --');
api.renderEstoqueItens('pecas');
const desc = paineis['pecas-painel'].innerHTML;
ok('1. Descalvado mostra os dois tipos dela, e nao o de Sao Carlos',
   (desc.match(/abrirEstoqueItem\('pecas'/g) || []).length === 2, (desc.match(/abrirEstoqueItem\('pecas'/g) || []).length);
ok('2. em uso e em estoque somam no rodape: 16 em uso, 203 em estoque, 219 no total',
   /em uso <b[^>]*>16<\/b>/.test(desc) && /em estoque <b[^>]*>203<\/b>/.test(desc) && />219</.test(desc));
api._estItensUnidade.pecas = 'sc';
api.renderEstoqueItens('pecas');
ok('3. Sao Carlos mostra so o dela (8 em uso, 50 em estoque)',
   /em uso <b[^>]*>8<\/b>/.test(paineis['pecas-painel'].innerHTML) && /em estoque <b[^>]*>50<\/b>/.test(paineis['pecas-painel'].innerHTML));
api._estItensUnidade.pecas = 'desc';
api._estItensBusca.pecas = 'lancadeira';
api.renderEstoqueItens('pecas');
ok('4. a busca ignora acento', (paineis['pecas-painel'].innerHTML.match(/abrirEstoqueItem\('pecas'/g) || []).length === 1);
api.renderEstoqueItens('ferramentas');
ok('5. ferramentas tem cadastro proprio (vazio aqui)', /Nenhuma ferramenta cadastrada/.test(paineis['ferramentas-painel'].innerHTML));

console.log('-- a costura com o resto --');
ok('6. os dois itens ficam na pasta Estoques, logo abaixo do Estoque de aviamentos',
   /Estoque de aviamentos<\/a>\s*<a class="nav-btn" href="#estoque-pecas"[^>]*>Estoque de peças<\/a>\s*<a class="nav-btn" href="#estoque-ferramentas"[^>]*>Estoque de ferramentas<\/a>/.test(html));
ok('7. as paginas e o modal existem',
   /data-page="estoque-pecas"/.test(html) && /data-page="estoque-ferramentas"/.test(html) && /id="modal-estoque-item"/.test(html));
ok('8. as rotas desenham as telas',
   /if \(page === 'estoque-pecas'\) renderEstoqueItens\('pecas'\);/.test(src) && /if \(page === 'estoque-ferramentas'\) renderEstoqueItens\('ferramentas'\);/.test(src));
ok('9. as chaves sao sincronizadas e carregadas (as duas listas de chaves)',
   (src.match(/'pecasCad','ferramentasCad'/g) || []).length === 2);
ok('10. corrigir altera a linha no lugar, e o mesmo tipo na mesma unidade nao duplica',
   /if \(ctx\.id\) \{\s*const x = lista\.find\(i => i\.id === ctx\.id\)/.test(src) && /já está cadastrado nesta unidade/.test(src));

console.log(falhas ? '\n' + falhas + ' FALHA(S)' : '\ntudo certo');
process.exit(falhas ? 1 : 0);
