/* Rode com:  node testes/compra-usuario.js

   A LISTA DE COMPRA é a única tela em que quem não é admin escreve.

   O programa inteiro é "só o admin escreve" — quem não é admin consulta e
   imprime. A lista de compra abriu exceção porque não é cadastro nem OS: é o
   rascunho de quanto tecido comprar, que se monta, se confere e se joga fora.
   Nada do que entra ali muda uma grade, uma OS ou um saldo de estoque.

   O que este teste protege são as BORDAS da exceção, que é onde ela estraga:

     · quem não fez login continua fora (currentRole null) — a chave anônima
       está dentro de toda página aberta na fábrica;
     · o modo nuvem (servidor da fábrica fora do ar) continua sendo só leitura
       para todo mundo, inclusive admin — senão os dois lados divergem;
     · cada pessoa tira da lista o que ELA somou; o levantamento do outro, não;
     · item sem dono (os que já estavam lá antes desta regra) fica com o admin.

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
  ${recorte('function exigirEdicaoCompra', 'a permissao da lista de compra')}
  ${recorte('function _cpPodeRemover', 'a regra de quem tira da lista')}
  return {
    exigirEdicaoCompra: (acao) => { currentRole = ctx.papel; return exigirEdicaoCompra(acao); },
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
if (falhas) { console.log(falhas + ' FALHA(S)'); process.exit(1); }
console.log('todos os testes passaram');
