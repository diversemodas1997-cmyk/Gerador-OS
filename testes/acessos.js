/* Rode com:  node testes/acessos.js

   A RELAÇÃO DE ÁREAS DO PROGRAMA e o acesso de cada conta.

   Até 26/08/2026 a permissão era um interruptor só — admin escreve, o resto lê
   — com duas exceções abertas na mão, dentro do código (o status da OS para o
   Enfesto.corte, o estoque de tecidos para a Nathaly). Cada exceção nova pedia
   uma linha nova no programa. Agora as áreas são uma lista, e o admin marca em
   Configurações quem entra em cada uma.

   O que este teste protege:

     · o admin entra em tudo, e "Contas de acesso" não se concede a ninguém —
       senão uma conta se promoveria sozinha;
     · quem já podia antes desta tela existir CONTINUA podendo sem ninguém
       marcar nada (as duas exceções viraram o padrão da área);
     · o que o admin marca vale sobre o padrão, inclusive para TIRAR;
     · cada ação do programa (os textos que exigirEdicao/exigirAdmin já
       recebiam) cai na área certa — é o que faz conceder uma área ter efeito
       de verdade nos botões;
     · modo nuvem (servidor da fábrica fora do ar) continua só leitura.

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
const constante = (nome) => {
  const m = src.match(new RegExp('^const ' + nome + ' = [^;]+;', 'm'));
  if (!m) { console.error('nao achei a constante ' + nome); process.exit(1); }
  return m[0];
};

const monta = (ctx) => new Function('ctx', `
  const podeGravar = () => ctx.servidorNoAr !== false;
  const currentUser = ctx.login ? { email: ctx.login } : null;
  let currentRole = ctx.papel;
  const STATE = ctx.STATE;
  ${constante('LOGINS_STATUS_OS')}
  ${constante('LOGINS_ESTOQUE_TECIDOS')}
  ${constante('AREAS_ACESSO')}
  ${constante('ACOES_POR_AREA')}
  ${constante('ACESSO_PADRAO')}
  ${recorte('function _obsQuemSou', 'o login de quem esta logado')}
  ${recorte('function _obsNomeLogin', 'o nome do login')}
  ${recorte('function _chaveLogin', 'a chave do login')}
  ${recorte('function _acessosTabela', 'a tabela de acessos')}
  ${recorte('function _acessoChaveConta', 'a chave da conta')}
  ${recorte('function contaTemAcesso', 'o acesso de uma conta')}
  ${recorte('function temAcesso', 'o acesso de quem esta logado')}
  ${recorte('function _areaDaAcao', 'a area de uma acao')}
  ${recorte('function _acessosQuantas', 'quantas areas a conta tem')}
  ${constante('DOMINIO_CONTA')}
  ${recorte('function emailParaNome', 'o nome a partir do endereco interno')}
  ${recorte('function _nomeDoPerfil', 'o nome que a barra lateral mostra')}
  return { contaTemAcesso, temAcesso, _areaDaAcao, _acessosQuantas, AREAS_ACESSO, ACOES_POR_AREA,
           emailParaNome, _nomeDoPerfil };
`)(ctx);

const ctxDe = (papel, login, acessos = {}, servidorNoAr = true) => {
  const ctx = { papel, login, servidorNoAr, STATE: { meta: { acessos } } };
  return { ctx, api: monta(ctx) };
};

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};

console.log('-- o admin e o que nunca se concede --');
let t = ctxDe('admin', 'admin@diverse.local');
ok('1. o admin entra em todas as areas',
   t.api.AREAS_ACESSO.every(a => t.api.temAcesso(a.k)), 'faltou area para o admin');
t = ctxDe('usuario', 'costura@diverse.local', { costura: { contas: true } });
ok('2. "Contas de acesso" nao se concede nem marcando na mao',
   t.api.temAcesso('contas') === false, 'uma conta poderia se promover');

console.log('');
console.log('-- quem ja podia continua podendo, sem ninguem marcar nada --');
t = ctxDe('usuario', 'enfesto.corte@diverse.local');
ok('3. Enfesto.corte segue com o status da OS', t.api.temAcesso('os-status') === true);
ok('4. mas nao com o resto', t.api.temAcesso('expedicao') === false && t.api.temAcesso('os') === false);
t = ctxDe('usuario', 'nathaly@diverse.local');
ok('5. Nathaly segue com o estoque de tecidos', t.api.temAcesso('estoque-tecidos') === true);
ok('6. e nao com os estoques das fases', t.api.temAcesso('estoque-fases') === false);
t = ctxDe('usuario', 'costura@diverse.local');
ok('7. quem nunca recebeu nada nao tem area nenhuma',
   t.api.AREAS_ACESSO.every(a => !t.api.temAcesso(a.k)), 'sobrou area para quem nao recebeu');

console.log('');
console.log('-- o que o admin marca vale sobre o padrao --');
t = ctxDe('usuario', 'costura@diverse.local', { costura: { expedicao: true } });
ok('8. area concedida passa a valer', t.api.temAcesso('expedicao') === true);
ok('9. e so ela', t.api.temAcesso('operacoes') === false);
t = ctxDe('usuario', 'nathaly@diverse.local', { nathaly: { 'estoque-tecidos': false } });
ok('10. e o admin tambem pode TIRAR o que era padrao',
   t.api.temAcesso('estoque-tecidos') === false, 'nao deu para tirar');
t = ctxDe('usuario', 'Nathaly@diverse.local', { nathaly: { operacoes: true } });
ok('11. a conta e a mesma com qualquer caixa ou pontuacao',
   t.api.temAcesso('operacoes') === true);

console.log('');
console.log('-- o modo nuvem --');
t = ctxDe('usuario', 'nathaly@diverse.local', {}, false);
ok('12. servidor da fabrica fora do ar: ninguem escreve', t.api.temAcesso('estoque-tecidos') === false);
t = ctxDe('admin', 'admin@diverse.local', {}, false);
ok('13. nem o admin', t.api.temAcesso('os') === false);

console.log('');
console.log('-- cada acao do programa cai na area certa --');
t = ctxDe('admin', 'admin@diverse.local');
const casos = [
  ['criar ou editar OS', 'os'],
  ['excluir OS', 'os'],
  ['excluir cadastros', 'cadastros'],
  ['apagar pastas de grade', 'cadastros'],
  ['lançar no estoque de tecidos', 'estoque-tecidos'],
  ['movimentar estoque', 'estoque-fases'],
  ['alocar OS na expedição', 'expedicao'],
  ['planejar operações', 'operacoes'],
  ['limpar a lista de compra', 'compra'],
  ['restaurar snapshots', 'dados']
];
casos.forEach(([acao, area], i) => {
  ok((14 + i) + '. "' + acao + '" e da area ' + area, t.api._areaDaAcao(acao) === area,
     t.api._areaDaAcao(acao) || '(nenhuma)');
});
ok('24. acao que nao esta na relacao segue so do admin',
   t.api._areaDaAcao('gerenciar usuários') === '', t.api._areaDaAcao('gerenciar usuários'));

console.log('');
console.log('-- toda acao mapeada aponta para uma area que existe --');
const chaves = new Set(t.api.AREAS_ACESSO.map(a => a.k));
const orfas = Object.entries(t.api.ACOES_POR_AREA).filter(([, area]) => !chaves.has(area));
ok('25. nenhuma acao aponta para area inexistente', orfas.length === 0, JSON.stringify(orfas));
// Toda acao que o programa usa em exigirEdicao/exigirAdmin ou esta na relacao,
// ou e assumidamente so do admin. Este teste LISTA as que ficaram de fora, para
// a escolha ser consciente e nao esquecimento.
const usadas = new Set([...src.matchAll(/exigir(?:Edicao|Admin)\('([^']+)'\)/g)].map(m => m[1]));
const foraDaRelacao = [...usadas].filter(a => !(a in t.api.ACOES_POR_AREA)).sort();
const SO_ADMIN_ESPERADO = [
  'baixar snapshots', 'conceder acesso', 'criar contas de acesso', 'editar contas de acesso',
  'gerenciar usuários', 'ver as contas de acesso', 'ver os usuários e os papéis'
];
ok('26. as acoes fora da relacao sao so as de conta/admin',
   foraDaRelacao.every(a => SO_ADMIN_ESPERADO.includes(a)),
   JSON.stringify(foraDaRelacao));

console.log('');
console.log('-- o botao na tela --');
ok('27. a lista de contas tem o botao de acesso e a listinha de areas',
   /alternarAcessos\('\$\{esc\(u\.id\)\}'\)/.test(src) && /_acessosMenuHtml\(u\)/.test(src),
   'sem o botao na lista de contas');
ok('28. marcar uma area grava e redesenha',
   /await saveState\('meta'\)/.test(recorte('async function definirAcesso', 'a concessao'))
   && /listarContasAcesso\(\)/.test(recorte('async function definirAcesso', 'a concessao')),
   'a concessao nao grava');
ok('29. e o contador do botao conta so o que da para conceder',
   ctxDe('admin', 'admin@diverse.local').api._acessosQuantas('enfesto.corte@diverse.local') === 1,
   String(ctxDe('admin', 'admin@diverse.local').api._acessosQuantas('enfesto.corte@diverse.local')));

/* --------------------------------------------------------------------------
   O NOME QUE A BARRA LATERAL MOSTRA (Junior, 28/08/2026)

   As contas da fabrica sao por NOME: "enfesto.corte" vira o endereco interno
   `enfesto.corte@diverse.local`, que nunca recebe mensagem. A barra mostra o
   nome, e NADA aqui passa a exigir e-mail.

   `_nomeDoPerfil` e separada de `emailParaNome` de proposito: a segunda
   preenche um CAMPO que a tela de Contas edita e salva, e encurtar ali mudaria
   o dado. Esta so pinta a barra.
   -------------------------------------------------------------------------- */
console.log('');
console.log('-- o nome do perfil na barra lateral --');
{
  const A = ctxDe('admin', 'admin@diverse.local').api;
  const perfil = (email) => A._nomeDoPerfil(email ? { email } : null);
  ok('30. conta por nome mostra o nome, sem o endereco interno',
     perfil('enfesto.corte@diverse.local') === 'enfesto.corte', perfil('enfesto.corte@diverse.local'));
  ok('31. e vale para todas as contas da fabrica',
     ['admin', 'backup', 'nathaly', 'escritorio'].every(n => perfil(n + '@diverse.local') === n));
  // A conta antiga do dono e um e-mail de verdade, aceito de proposito. Dela a
  // barra mostra so o nome, nao o endereco — mas ninguem passa a precisar de um.
  ok('32. e-mail de verdade tambem vira nome na barra',
     perfil('diversemodas1997@gmail.com') === 'diversemodas1997', perfil('diversemodas1997@gmail.com'));
  ok('33. sem conta nenhuma, uma palavra em vez do id interno',
     perfil('') === 'conta sem nome' && perfil(null) === 'conta sem nome');
  // emailParaNome NAO mudou: e ela que preenche o campo editavel de Contas.
  ok('34. emailParaNome segue devolvendo o e-mail inteiro (o campo que se edita)',
     A.emailParaNome('diversemodas1997@gmail.com') === 'diversemodas1997@gmail.com',
     A.emailParaNome('diversemodas1997@gmail.com'));
}

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
