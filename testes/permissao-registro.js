/* Rode com:  node testes/permissao-registro.js

   O QUE A PRODUÇÃO REGISTRA é a exceção ao "só o admin escreve".

   O programa inteiro é "só o admin escreve" — quem não é admin consulta e
   imprime. Duas coisas abriram exceção, porque nenhuma delas é cadastro nem
   definição de OS:

     · a LISTA DE COMPRA — rascunho de quanto tecido comprar; nada do que entra
       ali muda uma grade, uma OS ou um saldo de estoque;
     · o que a FOLHA DE OS anota — quadrinhos das etapas, horários, camadas por
       tonalidade, total por tamanho e a observação. Não é o que a OS PEDE, é o
       que a produção FEZ, e quem sabe é quem estava lá.

   O que este teste protege são as BORDAS da exceção, que é onde ela estraga:

     · quem não fez login continua fora (currentRole null) — a chave anônima
       está dentro de toda página aberta na fábrica;
     · o modo nuvem (servidor da fábrica fora do ar) continua sendo só leitura
       para todo mundo, inclusive admin — senão os dois lados divergem;
     · cada pessoa tira da lista o que ELA somou; o levantamento do outro, não;
     · item sem dono (os que já estavam lá antes desta regra) fica com o admin;
     · a folha vai INTEIRA — inclusive a observação —, mas CRIAR e EDITAR a OS
       não: registrar o que a produção fez não é redefinir o que ela deve fazer.

   O teste recorta as funções do app.js de verdade. */
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

// currentRole e podeGravar() sao do app; aqui entram como controles do teste.
const api = new Function('ctx', `
  const toast = (m, t) => ctx.toasts.push(t + ': ' + m);
  const podeGravar = () => ctx.servidorNoAr;
  ${recorte('function _recusarPorModoNuvem', 'a recusa do modo nuvem')}
  ${recorte('function _contaPodeRegistrar', 'a condicao do papel')}
  ${recorte('function exigirEdicaoCompra', 'a permissao da lista de compra')}
  ${recorte('function exigirEdicaoFolha', 'a permissao da folha de OS')}
  ${recorte('function _cpPodeRemover', 'a regra de quem tira da lista')}
  return {
    exigirEdicaoCompra: (acao) => { currentRole = ctx.papel; return exigirEdicaoCompra(acao); },
    exigirEdicaoFolha: (acao) => { currentRole = ctx.papel; return exigirEdicaoFolha(acao); },
    _cpPodeRemover
  };
  var currentRole;
`);

let falhas = 0;
const ok = (nome, cond, extra) => {
  console.log((cond ? '  ok  ' : 'FALHA ') + nome + (cond ? '' : '\n       obtido: ' + extra));
  if (!cond) falhas++;
};
const monta = (papel, servidorNoAr = true) => {
  const ctx = { papel, servidorNoAr, toasts: [] };
  return { ctx, api: api(ctx) };
};

console.log('-- quem pode somar na lista de compra --');
let t = monta('usuario');
ok('1. USUARIO comum soma (era isto que estava fechado)',
   t.api.exigirEdicaoCompra('montar a lista de compra') === true, t.ctx.toasts.join(' | '));
t = monta('admin');
ok('2. admin continua somando', t.api.exigirEdicaoCompra('montar a lista de compra') === true);
t = monta(null);
ok('3. sem login, nao', t.api.exigirEdicaoCompra('montar a lista de compra') === false
   && /Faça login/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));
t = monta('visita');
ok('4. papel desconhecido, nao', t.api.exigirEdicaoCompra('montar a lista de compra') === false);

console.log('');
console.log('-- o modo nuvem continua sendo so leitura, para todos --');
t = monta('usuario', false);
ok('5. usuario com o servidor fora do ar nao soma',
   t.api.exigirEdicaoCompra('montar a lista de compra') === false
   && /nuvem/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));
t = monta('admin', false);
ok('6. nem o admin', t.api.exigirEdicaoCompra('montar a lista de compra') === false);

console.log('');
console.log('-- quem tira da lista o que --');
const A = api({ papel: null, servidorNoAr: true, toasts: [] });
const meu = { criadoPor: 'costura@diverse.com' };
const doOutro = { criadoPor: 'corte@diverse.com' };
const semDono = { criadoEm: '2026-08-01T10:00:00.000Z' };
ok('7. o admin tira qualquer um',
   A._cpPodeRemover(meu, 'admin', 'chefe@diverse.com')
   && A._cpPodeRemover(semDono, 'admin', 'chefe@diverse.com'));
ok('8. cada um tira o que somou',
   A._cpPodeRemover(meu, 'usuario', 'costura@diverse.com') === true);
ok('9. o do outro, nao',
   A._cpPodeRemover(doOutro, 'usuario', 'costura@diverse.com') === false);
ok('10. item sem dono fica com o admin',
   A._cpPodeRemover(semDono, 'usuario', 'costura@diverse.com') === false);
ok('11. e-mail com caixa/espaco diferente continua sendo a mesma pessoa',
   A._cpPodeRemover({ criadoPor: ' Costura@Diverse.com ' }, 'usuario', 'costura@diverse.com') === true);
ok('12. sem saber quem sou, nao tiro nada',
   A._cpPodeRemover(meu, 'usuario', '') === false);
ok('13. quem nao fez login nao tira nem o proprio',
   A._cpPodeRemover(meu, null, 'costura@diverse.com') === false);

console.log('');
console.log('-- quem marca a folha de OS --');
t = monta('usuario');
ok('14. USUARIO comum marca as etapas', t.api.exigirEdicaoFolha('marcar etapas da OS') === true,
   t.ctx.toasts.join(' | '));
t = monta('usuario');
ok('15. e lanca os numeros (horario, camadas, total por tamanho)',
   t.api.exigirEdicaoFolha('lançar o horário de enfesto') === true
   && monta('usuario').api.exigirEdicaoFolha('editar o total por tamanho') === true);
t = monta('admin');
ok('16. admin continua marcando', t.api.exigirEdicaoFolha('marcar etapas da OS') === true);
t = monta(null);
ok('17. sem login, nao', t.api.exigirEdicaoFolha('marcar etapas da OS') === false);
t = monta('usuario');
ok('18. e escreve a observacao DELE na folha (aberta em 20/08/2026)',
   t.api.exigirEdicaoFolha('escrever a observação da OS') === true, t.ctx.toasts.join(' | '));
t = monta('usuario', false);
ok('19. servidor da fabrica fora do ar: ninguem marca',
   t.api.exigirEdicaoFolha('marcar etapas da OS') === false
   && /nuvem/.test(t.ctx.toasts.join(' ')), t.ctx.toasts.join(' | '));

console.log('');
console.log('-- o que NAO se abriu junto (conferido no app.js, nao em copia) --');
const app = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
// Cada linha aqui e um portao que tem que continuar sendo `exigirEdicao` (so
// admin). Se alguem trocar por exigirEdicaoFolha sem pensar, o teste cai.
// A OBSERVACAO virou UMA POR PESSOA em 20/08/2026: quem escreve e
// salvarObsNota (exigirEdicaoFolha, caso 18 acima), e ninguem mexe na do outro
// — isso quem cobra e testes/observacao-por-autor.js. O que ficou aqui e o
// campo ANTIGO, sem dono, que segue do admin.
const soAdmin = [
  ['a observacao antiga (sem dono)', "exigirEdicao('editar a observação antiga da OS')"],
  ['criar ou editar OS', "exigirEdicao('criar ou editar OS')"],
  ['editar OS', "exigirEdicao('editar OS')"],
  ['duplicar OS', "exigirEdicao('duplicar OS')"],
  ['criar ou editar cadastros', "exigirEdicao('criar ou editar cadastros')"],
  ['excluir cadastros', "exigirEdicao('excluir cadastros')"],
  ['a folha de OE', "exigirEdicao('editar a folha de OE')"],
  ['limpar a lista de compra', "exigirEdicao('limpar a lista de compra')"]
];
soAdmin.forEach(([oQue, trecho], i) => {
  ok((20 + i) + '. ' + oQue + ' continua so do admin', app.includes(trecho), trecho);
});

console.log('');
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
