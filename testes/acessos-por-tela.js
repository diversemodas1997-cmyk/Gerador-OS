/* Rode com:  node testes/acessos-por-tela.js

   VER E EDITAR POR TELA (21/10/2026, Junior: "insira nas concessões todos os
   campos do programa para que seja possível configurar o que cada usuário veja
   e edita").

   A área é o padrão das telas dela; a tela marcada ganha da área. O que este
   teste protege:
     · sem ajuste, nada muda — a tela segue a área, como sempre seguiu;
     · o ajuste de uma tela não vaza para as outras da mesma área;
     · a trava das ações (temAcesso) olha a TELA ABERTA;
     · quem edita vê; tirar o Ver tira o Editar;
     · marcar a área de novo desfaz os ajustes;
     · o admin entra em tudo.
   Recorta as funções do app.js de verdade. */
const fs = require('fs');
const path = require('path');

const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
function recorte(de, oQue) {
  const i = src.indexOf(de);
  if (i < 0) { console.error('nao achei ' + oQue + ' no app.js'); process.exit(1); }
  const j = src.indexOf('\n}', i);
  return src.slice(i, j + 2);
}
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

const monta = (ctx) => new Function('ctx', `
  const STATE = ctx.STATE;
  let currentRole = ctx.papel || 'usuario';
  const currentUser = { email: ctx.login };
  const document = { querySelector: () => ({ dataset: { page: ctx.pagina } }) };
  const podeGravar = () => true;
  const toast = (m, t) => ctx.toasts.push((t || '') + ':' + m);
  const exigirAdmin = () => ctx.papelQuemMarca === 'admin';
  const desfazerNomearAcao = () => {};
  const saveState = async () => { ctx.gravou = (ctx.gravou || 0) + 1; };
  const aplicarPermissoesUI = () => {};
  const _redesenharAcessosConta = () => {};
  const ACESSO_PADRAO = {};
  const AREAS_ACESSO = [
    { k: 'os', rotulo: 'Ordens de Serviço' }, { k: 'cadastros', rotulo: 'Cadastros' },
    { k: 'estoque-fases', rotulo: 'Estoques das fases' }, { k: 'expedicao', rotulo: 'Expedição' },
    { k: 'contas', rotulo: 'Contas de acesso', soAdmin: true }
  ];
  const AREAS_SO_VISAO = [{ k: 'ranking', rotulo: 'Ranking' }, { k: 'colorimetro', rotulo: 'Colorímetro' }];
  ${constante('VISAO_PAGINAS')}
  ${constante('VISAO_CONCEDIDA')}
  ${constante('TELAS_DA_AREA')}
  ${constante('_TELA_DA_FOLHA')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('function _chaveLogin', 'a chave do login')}
  ${recorte('function _acessoChaveConta', 'a chave da conta')}
  ${recorte('function _acessosTabela', 'a tabela de acessos')}
  ${recorte('function _visaoTabela', 'a tabela de visao')}
  ${recorte('function _areaDaPagina', 'a area de uma tela')}
  ${recorte('function contaTemAcesso', 'o acesso de uma conta')}
  ${recorte('function contaVeArea', 'a visao de uma area')}
  ${recorte('function _telaBase', 'a tela de uma folha')}
  ${recorte('function _paginaAtual', 'a tela aberta')}
  ${recorte('function _telasTabela', 'a tabela das telas')}
  ${recorte('function _ajusteTela', 'o ajuste de uma tela')}
  ${recorte('function contaEditaTela', 'a edicao de uma tela')}
  ${recorte('function contaVeTela', 'a visao de uma tela')}
  ${recorte('async function definirTela', 'a marcacao de uma tela')}
  ${recorte('function _limparAjustesTelas', 'a limpeza dos ajustes')}
  ${recorte('function _telasAjustadas', 'a conta dos ajustes')}
  ${recorte('function temAcesso', 'o acesso de quem esta logado')}
  ${recorte('function podeVerPagina', 'a visao da tela')}
  ${recorte('async function definirAcesso', 'a concessao da area')}
  return { contaEditaTela, contaVeTela, temAcesso, podeVerPagina, definirTela, definirAcesso,
           _telasAjustadas, setPagina: (p) => { ctx.pagina = p; } };
`)(ctx);

const novo = (meta, extra) => {
  const ctx = Object.assign({ login: 'nathaly@diverse.local', pagina: 'home', toasts: [],
    papelQuemMarca: 'admin', STATE: { meta: meta || {} } }, extra || {});
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '   obtido: ' + JSON.stringify(extra)));
  if (!cond) falhas++;
};
const N = 'nathaly@diverse.local';

console.log('-- sem ajuste, a tela segue a area --');
let t = novo({ acessos: { nathaly: { cadastros: true } } });
ok('1. edita Cadastros -> edita a tela Equipe', t.api.contaEditaTela('cad-equipe', N, 'usuario') === true);
ok('2. e a tela Grades', t.api.contaEditaTela('cad-grades', N, 'usuario') === true);
ok('3. nao edita Expedicao, que nao recebeu', t.api.contaEditaTela('expedicao', N, 'usuario') === false);
ok('4. e ve tudo (o padrao e ver)', t.api.contaVeTela('corte-sc', N, 'usuario') === true);

console.log('\n-- o ajuste de uma tela --');
t = novo({ acessos: { nathaly: { cadastros: true } }, acessosTelas: { nathaly: { 'cad-equipe': { editar: false } } } });
ok('5. "edita Cadastros, menos Equipe": Equipe nao edita', t.api.contaEditaTela('cad-equipe', N, 'usuario') === false);
ok('6. e as outras telas de Cadastros continuam editando', t.api.contaEditaTela('cad-tecidos', N, 'usuario') === true);
ok('7. Equipe ainda se ve (so tirou o editar)', t.api.contaVeTela('cad-equipe', N, 'usuario') === true);
t.api.setPagina('cad-equipe');
ok('8. na tela Equipe, a trava da area recusa', t.api.temAcesso('cadastros') === false);
t.api.setPagina('cad-tecidos');
ok('9. na tela Tecidos, a mesma trava deixa', t.api.temAcesso('cadastros') === true);
t.api.setPagina('home');
ok('10. fora das telas da area, vale a area', t.api.temAcesso('cadastros') === true);
ok('11. a linha da area conta 1 tela ajustada', t.api._telasAjustadas('nathaly', 'cadastros') === 1);

t = novo({ acessosTelas: { nathaly: { 'corte-sc': { ver: false } } } });
ok('12. "ve os estoques, menos Sao Carlos": corte SC some', t.api.podeVerPagina('corte-sc') === false);
ok('13. e o corte de Descalvado continua', t.api.podeVerPagina('corte') === true);

t = novo({ acessosTelas: { nathaly: { 'cad-cores': { editar: true, ver: true } } } });
ok('14. editar UMA tela sem ter a area: edita Cores', t.api.contaEditaTela('cad-cores', N, 'usuario') === true);
ok('15. e so ela', t.api.contaEditaTela('cad-marcas', N, 'usuario') === false);
t.api.setPagina('cad-cores');
ok('16. na tela Cores, a trava deixa', t.api.temAcesso('cadastros') === true);
t.api.setPagina('cad-marcas');
ok('17. na tela Marcas, recusa', t.api.temAcesso('cadastros') === false);

t = novo({ acessosTelas: { nathaly: { 'lista-os': { editar: false } } }, acessos: { nathaly: { os: true } } });
t.api.setPagina('print');
ok('18. a folha de OS (print) segue a tela da lista de OS', t.api.temAcesso('os') === false);

t = novo({ acessosVisao: { nathaly: { colorimetro: false } } });
ok('19. Colorimetro tem o seu Ver', t.api.podeVerPagina('colorimetro') === false);
t = novo({});
ok('19a. o Colorimetro nasce fechado: sem marcacao, nao ve', t.api.podeVerPagina('colorimetro') === false);
ok('19b. e as outras areas continuam abertas por padrao', t.api.podeVerPagina('ranking') === true && t.api.podeVerPagina('cad-cores') === true);
t = novo({ acessosVisao: { nathaly: { colorimetro: true } } });
ok('19c. o admin concede, e a conta passa a ver', t.api.podeVerPagina('colorimetro') === true);
ok('19d. o admin ve sem ninguem conceder', t.api.contaVeTela('colorimetro', N, 'admin') === true);
ok('20. area so de consulta nao se edita', t.api.contaEditaTela('ranking', N, 'usuario') === false);
t = novo({ acessosVisao: { nathaly: { inicio: false } } });
ok('21a. o Inicio tem o seu Ver', t.api.podeVerPagina('home') === false && t.api.podeVerPagina('lista-os') === true);
ok('21b. e e so de consulta', t.api.contaEditaTela('home', N, 'usuario') === false);
ok('21. admin edita e ve tudo', t.api.contaEditaTela('cad-equipe', N, 'admin') && t.api.contaVeTela('corte-sc', N, 'admin'));

(async () => {
  console.log('\n-- marcar --');
  t = novo({});
  await t.api.definirTela(N, 'cad-equipe', 'editar', true);
  ok('22. dar o Editar marca o Ver junto',
     JSON.stringify(t.ctx.STATE.meta.acessosTelas.nathaly['cad-equipe']) === '{"editar":true,"ver":true}',
     t.ctx.STATE.meta.acessosTelas);
  await t.api.definirTela(N, 'cad-equipe', 'ver', false);
  ok('23. tirar o Ver tira o Editar', t.api.contaEditaTela('cad-equipe', N, 'usuario') === false
     && t.api.contaVeTela('cad-equipe', N, 'usuario') === false, t.ctx.STATE.meta.acessosTelas);
  ok('24. e grava', t.ctx.gravou === 2, t.ctx.gravou);
  await t.api.definirTela(N, 'tela-que-nao-existe', 'ver', true);
  ok('25. tela desconhecida nao grava', t.ctx.gravou === 2 && /desconhecida/.test(t.ctx.toasts.join(' ')), t.ctx.toasts);

  t = novo({ acessosTelas: { nathaly: { 'cad-equipe': { editar: false, ver: true }, 'corte': { ver: false } } } });
  await t.api.definirAcesso(N, 'cadastros', true);
  ok('26. marcar a area de novo desfaz os ajustes de editar das telas dela',
     !t.ctx.STATE.meta.acessosTelas.nathaly['cad-equipe'] && t.api.contaEditaTela('cad-equipe', N, 'usuario') === true,
     t.ctx.STATE.meta.acessosTelas);
  ok('27. e nao mexe nas telas de outra area', !!t.ctx.STATE.meta.acessosTelas.nathaly.corte);

  t = novo({}, { papelQuemMarca: 'usuario' });
  await t.api.definirTela(N, 'cad-equipe', 'editar', true);
  ok('28. so o admin marca', !t.ctx.STATE.meta.acessosTelas && !t.ctx.gravou);

  console.log(falhas ? `\n${falhas} FALHA(S)` : '\ntodos os testes passaram');
  process.exit(falhas ? 1 : 0);
})();
